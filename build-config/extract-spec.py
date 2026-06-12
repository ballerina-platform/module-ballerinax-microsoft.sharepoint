#!/usr/bin/env python3
"""
Extract SharePoint-scoped OpenAPI specs from the Microsoft Graph spec.

Reads `docs/spec/graphexplorer.yaml` (full Microsoft Graph v1.0) and writes:

  docs/spec/sharepoint.yaml                 — consolidated SharePoint spec
  docs/spec/submodules/sites.yaml           — site lifecycle + site-level resources
  docs/spec/submodules/lists.yaml           — lists & list items
  docs/spec/submodules/pages.yaml           — SharePoint pages
  docs/spec/submodules/termstore.yaml       — taxonomy / term store
  docs/spec/submodules/onenote.yaml         — OneNote on sites
  docs/spec/submodules/embedded.yaml        — SharePoint Embedded (file storage)
  docs/spec/submodules/admin.yaml           — tenant SharePoint admin

For each output, components (schemas, responses, parameters, examples, requestBodies,
headers, securitySchemes) are pruned to the transitive closure of $refs reachable
from the selected paths.

Run from the repository root:

    python3 build-config/extract-spec.py
"""

from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path
from typing import Callable, Iterable

try:
    import yaml
except ImportError:
    sys.stderr.write("PyYAML is required: pip install pyyaml\n")
    sys.exit(1)

try:
    from yaml import CSafeLoader as YamlLoader, CSafeDumper as YamlDumper
except ImportError:
    from yaml import SafeLoader as YamlLoader, SafeDumper as YamlDumper


REPO_ROOT = Path(__file__).resolve().parent.parent
SOURCE = REPO_ROOT / "docs" / "spec" / "graphexplorer.yaml"
OUT_DIR = REPO_ROOT / "docs" / "spec"
SUBMODULE_DIR = OUT_DIR / "submodules"

# `bal openapi` parses YAML through SnakeYAML, which caps single-document size
# at 3,145,728 code points. Any spec emitted as YAML and larger than this
# fails to generate a Ballerina client — silently producing no code, which
# looks like "auth missing / everything missing" to the user. We emit JSON
# alongside the YAML for any output that crosses this threshold, since the
# Jackson-based JSON parser has no such cap.
SNAKEYAML_CODEPOINT_LIMIT = 3_145_728

# YAML sections under `components` that may contain $refs and need pruning.
COMPONENT_SECTIONS = (
    "schemas",
    "responses",
    "parameters",
    "examples",
    "requestBodies",
    "headers",
    "securitySchemes",
    "links",
    "callbacks",
)


# Path-classification predicates.
def _under_resource(path: str, root: str, resource: str) -> bool:
    """True if `path` matches /<root>/{...}/<resource> or /<root>/{...}/<resource>/..."""
    pattern = rf"^/{re.escape(root)}/\{{[^}}]+\}}/{re.escape(resource)}(/|$)"
    return re.match(pattern, path) is not None


def in_admin(path: str) -> bool:
    return path == "/admin/sharepoint" or path.startswith("/admin/sharepoint/")


def in_embedded(path: str) -> bool:
    return path == "/storage/fileStorage" or path.startswith("/storage/fileStorage/")


def in_lists(path: str) -> bool:
    return _under_resource(path, "sites", "lists")


def in_pages(path: str) -> bool:
    return _under_resource(path, "sites", "pages")


def in_termstore(path: str) -> bool:
    return _under_resource(path, "sites", "termStore") or _under_resource(
        path, "sites", "termStores"
    )


def in_onenote(path: str) -> bool:
    return _under_resource(path, "sites", "onenote")


def in_sites_remainder(path: str) -> bool:
    """Anything under /sites that isn't lists/pages/termstore/onenote."""
    if not (path == "/sites" or path.startswith("/sites/")):
        return False
    return not (in_lists(path) or in_pages(path) or in_termstore(path) or in_onenote(path))


def in_sharepoint(path: str) -> bool:
    return (
        in_sites_remainder(path)
        or in_lists(path)
        or in_pages(path)
        or in_termstore(path)
        or in_onenote(path)
        or in_embedded(path)
        or in_admin(path)
    )


# --- Security: Azure AD OAuth2 (Microsoft Entra ID) ---------------------------
#
# Microsoft Graph SharePoint endpoints are authorized via Azure AD OAuth2.
# Each spec exposes a single `azureOAuth2` security scheme with both
# `authorizationCode` (delegated) and `clientCredentials` (application) flows.
# Per-submodule we list only the scopes relevant to that submodule's surface.

OAUTH2_AUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
OAUTH2_TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

# Scope -> human-readable description. App-permission description is the same
# text suffixed with " (application permission)" to match the outlook.mail style.
SCOPE_DESCRIPTIONS: dict[str, str] = {
    # Site / list / page / content type / column / permission scopes
    "Sites.Read.All": "Read items in all site collections",
    "Sites.ReadWrite.All": "Read and write items in all site collections",
    "Sites.Manage.All": "Create, edit, and delete items and lists in all site collections",
    "Sites.FullControl.All": "Have full control of all site collections",
    "Sites.Selected": "Access selected site collections only",
    # Term store / taxonomy
    "TermStore.Read.All": "Read managed metadata in SharePoint",
    "TermStore.ReadWrite.All": "Read and write managed metadata in SharePoint",
    # OneNote (on SharePoint sites)
    "Notes.Read": "Read user OneNote notebooks",
    "Notes.ReadWrite": "Read and write user OneNote notebooks",
    "Notes.Read.All": "Read all OneNote notebooks the signed-in user can access",
    "Notes.ReadWrite.All": "Read and write all OneNote notebooks the signed-in user can access",
    "Notes.Create": "Create OneNote notebooks",
    # SharePoint Embedded
    "FileStorageContainer.Selected": "Access selected file storage containers",
    # Tenant SharePoint admin
    "SharePointTenantSettings.Read.All": "Read SharePoint and OneDrive tenant settings",
    "SharePointTenantSettings.ReadWrite.All": "Read and change SharePoint and OneDrive tenant settings",
}

# Full scope list per submodule. The first entry is the consolidated spec.
SITE_BASE_SCOPES = [
    "Sites.Read.All",
    "Sites.ReadWrite.All",
    "Sites.Manage.All",
    "Sites.FullControl.All",
    "Sites.Selected",
]
TERMSTORE_SCOPES = ["TermStore.Read.All", "TermStore.ReadWrite.All"]
ONENOTE_SCOPES = [
    "Notes.Read",
    "Notes.ReadWrite",
    "Notes.Read.All",
    "Notes.ReadWrite.All",
    "Notes.Create",
]
EMBEDDED_SCOPES = ["FileStorageContainer.Selected"]
ADMIN_SCOPES = [
    "SharePointTenantSettings.Read.All",
    "SharePointTenantSettings.ReadWrite.All",
]

MODULE_SCOPES: dict[str, list[str]] = {
    "sharepoint": (
        SITE_BASE_SCOPES + TERMSTORE_SCOPES + ONENOTE_SCOPES + EMBEDDED_SCOPES + ADMIN_SCOPES
    ),
    "sites": SITE_BASE_SCOPES,
    "lists": SITE_BASE_SCOPES,
    "pages": SITE_BASE_SCOPES,
    "termstore": TERMSTORE_SCOPES,
    "onenote": ONENOTE_SCOPES,
    "embedded": EMBEDDED_SCOPES,
    "admin": ADMIN_SCOPES,
}

# Top-level `security` advertises a representative subset (the typical
# read+write pair for the submodule's main resource).
DEFAULT_SECURITY: dict[str, list[str]] = {
    "sharepoint": ["Sites.Read.All", "Sites.ReadWrite.All"],
    "sites": ["Sites.Read.All", "Sites.ReadWrite.All"],
    "lists": ["Sites.Read.All", "Sites.ReadWrite.All"],
    "pages": ["Sites.Read.All", "Sites.ReadWrite.All", "Sites.Manage.All"],
    "termstore": ["TermStore.Read.All", "TermStore.ReadWrite.All"],
    "onenote": ["Notes.Read.All", "Notes.ReadWrite.All"],
    "embedded": ["FileStorageContainer.Selected"],
    "admin": ["SharePointTenantSettings.Read.All", "SharePointTenantSettings.ReadWrite.All"],
}


def build_security_scheme(module: str) -> dict:
    """Build the `azureOAuth2` security scheme entry for a submodule."""
    scopes = MODULE_SCOPES[module]
    delegated_scopes = {s: SCOPE_DESCRIPTIONS[s] for s in scopes}
    app_scopes = {
        s: f"{SCOPE_DESCRIPTIONS[s]} (application permission)" for s in scopes
    }
    return {
        "type": "oauth2",
        "description": (
            "OAuth 2.0 authorization using Azure Active Directory (Microsoft Entra ID). "
            "Supports both delegated and application permissions for accessing "
            "SharePoint resources via the Microsoft Graph API."
        ),
        "flows": {
            "authorizationCode": {
                "authorizationUrl": OAUTH2_AUTH_URL,
                "tokenUrl": OAUTH2_TOKEN_URL,
                "scopes": delegated_scopes,
            },
            "clientCredentials": {
                "tokenUrl": OAUTH2_TOKEN_URL,
                "scopes": app_scopes,
            },
        },
    }


# (output filename, predicate, info title, security-module-key)
OUTPUTS: list[tuple[Path, Callable[[str], bool], str, str]] = [
    (OUT_DIR / "sharepoint.yaml", in_sharepoint, "Microsoft Graph — SharePoint API", "sharepoint"),
    (SUBMODULE_DIR / "sites.yaml", in_sites_remainder, "Microsoft Graph — SharePoint Sites", "sites"),
    (SUBMODULE_DIR / "lists.yaml", in_lists, "Microsoft Graph — SharePoint Lists", "lists"),
    (SUBMODULE_DIR / "pages.yaml", in_pages, "Microsoft Graph — SharePoint Pages", "pages"),
    (SUBMODULE_DIR / "termstore.yaml", in_termstore, "Microsoft Graph — SharePoint Term Store", "termstore"),
    (SUBMODULE_DIR / "onenote.yaml", in_onenote, "Microsoft Graph — SharePoint OneNote", "onenote"),
    (SUBMODULE_DIR / "embedded.yaml", in_embedded, "Microsoft Graph — SharePoint Embedded", "embedded"),
    (SUBMODULE_DIR / "admin.yaml", in_admin, "Microsoft Graph — SharePoint Admin", "admin"),
]


REF_PREFIX = "#/components/"


def collect_refs(node) -> Iterable[str]:
    """Yield every `$ref` string value found anywhere inside `node`."""
    if isinstance(node, dict):
        for k, v in node.items():
            if k == "$ref" and isinstance(v, str):
                yield v
            else:
                yield from collect_refs(v)
    elif isinstance(node, list):
        for item in node:
            yield from collect_refs(item)


def collect_tags(paths: dict) -> set[str]:
    """Collect operation-level tag names from selected paths."""
    used: set[str] = set()
    for path_item in paths.values():
        if not isinstance(path_item, dict):
            continue
        for key, op in path_item.items():
            if not isinstance(op, dict):
                continue
            tags = op.get("tags")
            if isinstance(tags, list):
                used.update(t for t in tags if isinstance(t, str))
    return used


def _drain_queue(queue: list[str], kept: dict[str, set[str]], components: dict) -> None:
    """Transitively resolve every ref currently in `queue` into `kept`."""
    while queue:
        ref = queue.pop()
        if not ref.startswith(REF_PREFIX):
            continue
        try:
            section, name = ref[len(REF_PREFIX):].split("/", 1)
        except ValueError:
            continue
        if section not in kept:
            continue
        if name in kept[section]:
            continue
        component = components.get(section, {}).get(name)
        if component is None:
            # Dangling upstream ref — record but don't crash.
            kept[section].add(name)
            continue
        kept[section].add(name)
        queue.extend(collect_refs(component))


def _build_derived_index(schemas: dict) -> dict[str, list[str]]:
    """Map base-schema-name -> list of schema names that derive from it via allOf $ref."""
    derived_of: dict[str, list[str]] = {}
    for name, schema in schemas.items():
        if not isinstance(schema, dict):
            continue
        for entry in schema.get("allOf") or []:
            if not isinstance(entry, dict):
                continue
            ref = entry.get("$ref", "")
            if isinstance(ref, str) and ref.startswith("#/components/schemas/"):
                base = ref[len("#/components/schemas/"):]
                derived_of.setdefault(base, []).append(name)
    return derived_of


def walk_ref_closure(seed_node, components: dict) -> dict[str, set[str]]:
    """
    Starting from `$ref` values found in `seed_node`, build the set of
    components that must travel with `seed_node`.

    Three phases, in order:

      1. **Transitive `$ref` closure.** Walk every `$ref` reachable from
         `seed_node`, follow into the referenced component, and walk that
         component's `$ref`s too — until the set stabilises.

      2. **Polymorphism completion.** Microsoft Graph models inheritance via
         `allOf: [{$ref: <base>}, …]`. When a path operation directly
         references a base schema as a request/response body, the actual
         payload at runtime may be any *derived* schema, disambiguated only
         by the `@odata.type` annotation. The transitive `$ref` walk does
         not reach those derived schemas (the `$ref` edge points base→from-derived,
         not derived→from-base), so they would be dropped without this step.
         For every base schema *directly* referenced from `seed_node`, we
         include every transitively derived schema, then re-run the `$ref`
         walk so their own dependencies get pulled in.

      3. **Example payloads.** Pull in `components/examples/<name>` for every
         kept `components/schemas/<name>` (the upstream uses identical
         names). These hold the `@odata.type` annotations and are not
         reachable via `$ref` at all.
    """
    kept: dict[str, set[str]] = {s: set() for s in COMPONENT_SECTIONS}

    # Phase 1: standard `$ref` closure from the seed (paths).
    queue: list[str] = list(collect_refs(seed_node))
    _drain_queue(queue, kept, components)

    # Phase 2: polymorphism completion. Only bases that are directly referenced
    # from `seed_node` get their derived chains pulled in — transitively-included
    # schemas (e.g. `microsoft.graph.entity`, which sits at the root of nearly
    # every Graph type) are intentionally excluded to avoid dragging in the
    # entire Graph surface.
    #
    # Also excluded: bases whose name is outside the `microsoft.graph.*`
    # namespace. These are OData *infrastructure* base classes — most notably
    # `BaseCollectionPaginationCountResponse` — whose derivatives are typed
    # pagination wrappers (one per entity type), not real polymorphic
    # `@odata.type` subtypes. Treating them as polymorphic bases would
    # incorrectly drag in every `*CollectionResponse` schema in Graph (>1,200
    # entries) and through their transitive `$ref`s, essentially the whole API.
    schemas_dict = components.get("schemas", {}) or {}
    derived_of = _build_derived_index(schemas_dict)

    def _is_polymorphic_base(name: str) -> bool:
        return name.startswith("microsoft.graph.")

    bases_from_seed: set[str] = set()
    for ref in collect_refs(seed_node):
        if ref.startswith("#/components/schemas/"):
            schema_name = ref[len("#/components/schemas/"):]
            if _is_polymorphic_base(schema_name):
                bases_from_seed.add(schema_name)

    poly_queue: list[str] = list(bases_from_seed)
    while poly_queue:
        base = poly_queue.pop()
        for derived_name in derived_of.get(base, []):
            if derived_name in kept["schemas"]:
                continue
            derived_schema = schemas_dict.get(derived_name)
            if derived_schema is None:
                continue
            kept["schemas"].add(derived_name)
            queue.extend(collect_refs(derived_schema))
            # The newly added schema can itself be a base for further derivatives.
            if _is_polymorphic_base(derived_name):
                poly_queue.append(derived_name)

    # Resolve any new refs introduced by the derived schemas.
    _drain_queue(queue, kept, components)

    # Phase 3: pull in `components/examples/<name>` for every kept schema.
    # These payloads carry the `@odata.type` annotations and are not reachable
    # via $ref anywhere in the upstream spec.
    available_examples = components.get("examples", {}) or {}
    if isinstance(available_examples, dict):
        for schema_name in list(kept["schemas"]):
            if schema_name in available_examples and schema_name not in kept["examples"]:
                kept["examples"].add(schema_name)
                queue.extend(collect_refs(available_examples[schema_name]))
        _drain_queue(queue, kept, components)

    return kept


def subset_components(components: dict, kept: dict[str, set[str]]) -> dict:
    """Return a new components dict containing only items in `kept` per section."""
    out: dict = {}
    for section, names in kept.items():
        if not names:
            continue
        src = components.get(section)
        if not isinstance(src, dict):
            continue
        # Preserve original ordering from the source dict.
        out[section] = {n: src[n] for n in src if n in names}
    return out


def assert_no_dangling(paths, components) -> list[str]:
    """Return a list of refs that don't resolve inside `components`."""
    missing: list[str] = []
    available: dict[str, set[str]] = {s: set(components.get(s, {}).keys()) for s in COMPONENT_SECTIONS}
    for ref in collect_refs(paths):
        if not ref.startswith(REF_PREFIX):
            continue
        try:
            section, name = ref[len(REF_PREFIX):].split("/", 1)
        except ValueError:
            missing.append(ref)
            continue
        if section not in available or name not in available[section]:
            missing.append(ref)
    for ref in collect_refs(components):
        if not ref.startswith(REF_PREFIX):
            continue
        try:
            section, name = ref[len(REF_PREFIX):].split("/", 1)
        except ValueError:
            missing.append(ref)
            continue
        if section not in available or name not in available[section]:
            missing.append(ref)
    return missing


def count_ops(paths: dict) -> int:
    op_keys = {"get", "post", "patch", "delete", "put", "head", "options", "trace"}
    n = 0
    for item in paths.values():
        if isinstance(item, dict):
            n += sum(1 for k in item if k in op_keys)
    return n


def write_yaml(target: Path, doc: dict) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("w", encoding="utf-8") as fh:
        yaml.dump(
            doc,
            fh,
            Dumper=YamlDumper,
            sort_keys=False,
            allow_unicode=True,
            width=4096,
            default_flow_style=False,
        )


def write_json(target: Path, doc: dict) -> None:
    """Write `doc` as compact JSON. JSON is the safe fallback when YAML output
    exceeds the SnakeYAML parser cap that `bal openapi` enforces."""
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, separators=(",", ":"))


def build_output(
    source_doc: dict,
    predicate: Callable[[str], bool],
    title: str,
    security_module: str,
) -> dict:
    selected_paths = {p: ops for p, ops in source_doc["paths"].items() if predicate(p)}

    kept = walk_ref_closure(selected_paths, source_doc.get("components", {}))
    pruned_components = subset_components(source_doc.get("components", {}), kept)

    # Inject the azureOAuth2 security scheme into components.securitySchemes.
    pruned_components.setdefault("securitySchemes", {})
    pruned_components["securitySchemes"]["azureOAuth2"] = build_security_scheme(security_module)

    used_tag_names = collect_tags(selected_paths)
    source_tags = source_doc.get("tags") or []
    pruned_tags = [t for t in source_tags if isinstance(t, dict) and t.get("name") in used_tag_names]

    info = dict(source_doc.get("info", {}))
    info["title"] = title
    if "description" in info:
        info["description"] = (
            "SharePoint subset of the Microsoft Graph v1.0 OpenAPI specification "
            "(generated from graphexplorer.yaml)."
        )

    out: dict = {
        "openapi": source_doc["openapi"],
        "info": info,
    }
    if "servers" in source_doc:
        out["servers"] = source_doc["servers"]
    out["security"] = [{"azureOAuth2": list(DEFAULT_SECURITY[security_module])}]
    if pruned_tags:
        out["tags"] = pruned_tags
    out["paths"] = selected_paths
    out["components"] = pruned_components
    return out


def main() -> int:
    if not SOURCE.exists():
        sys.stderr.write(f"Source spec not found: {SOURCE}\n")
        return 1

    t0 = time.monotonic()
    print(f"Loading {SOURCE.relative_to(REPO_ROOT)} ...", flush=True)
    with SOURCE.open("r", encoding="utf-8") as fh:
        source_doc = yaml.load(fh, Loader=YamlLoader)
    print(f"  loaded in {time.monotonic() - t0:.1f}s "
          f"({len(source_doc.get('paths', {}))} paths, "
          f"{count_ops(source_doc.get('paths', {}))} operations, "
          f"{len(source_doc.get('components', {}).get('schemas', {}))} schemas)",
          flush=True)

    total_paths_emitted = 0
    submodule_paths_emitted = 0

    for target, predicate, title, security_module in OUTPUTS:
        t1 = time.monotonic()
        out_doc = build_output(source_doc, predicate, title, security_module)
        n_paths = len(out_doc["paths"])
        n_ops = count_ops(out_doc["paths"])
        n_schemas = len(out_doc.get("components", {}).get("schemas", {}))

        dangling = assert_no_dangling(out_doc["paths"], out_doc.get("components", {}))
        if dangling:
            sys.stderr.write(
                f"WARNING: {target.name}: {len(dangling)} dangling $refs "
                f"(first 5): {dangling[:5]}\n"
            )

        write_yaml(target, out_doc)
        yaml_size_bytes = target.stat().st_size
        size_mb = yaml_size_bytes / (1024 * 1024)
        extras: list[str] = []
        # If the YAML output exceeds the SnakeYAML cap, also emit JSON so
        # `bal openapi` (and other Jackson-based parsers) can still consume it.
        if yaml_size_bytes > SNAKEYAML_CODEPOINT_LIMIT:
            json_target = target.with_suffix(".json")
            write_json(json_target, out_doc)
            json_size_mb = json_target.stat().st_size / (1024 * 1024)
            extras.append(
                f"YAML > SnakeYAML cap ({yaml_size_bytes:,} > {SNAKEYAML_CODEPOINT_LIMIT:,}); "
                f"also emitted {json_target.relative_to(REPO_ROOT)} ({json_size_mb:.1f} MB)"
            )
        print(
            f"  {target.relative_to(REPO_ROOT)}: "
            f"{n_paths} paths, {n_ops} ops, {n_schemas} schemas, "
            f"{size_mb:.1f} MB ({time.monotonic() - t1:.1f}s)",
            flush=True,
        )
        for note in extras:
            print(f"    note: {note}", flush=True)

        if target.parent == OUT_DIR:
            total_paths_emitted = n_paths
        else:
            submodule_paths_emitted += n_paths

    print(f"\nConsolidated paths: {total_paths_emitted}")
    print(f"Sum of submodule paths: {submodule_paths_emitted}")
    if total_paths_emitted != submodule_paths_emitted:
        sys.stderr.write(
            "ERROR: submodule paths do not partition the consolidated spec.\n"
        )
        return 2

    print(f"\nDone in {time.monotonic() - t0:.1f}s.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
