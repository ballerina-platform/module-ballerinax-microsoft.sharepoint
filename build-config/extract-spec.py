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

import os
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


# (output filename, predicate, info title)
OUTPUTS: list[tuple[Path, Callable[[str], bool], str]] = [
    (OUT_DIR / "sharepoint.yaml", in_sharepoint, "Microsoft Graph — SharePoint API"),
    (SUBMODULE_DIR / "sites.yaml", in_sites_remainder, "Microsoft Graph — SharePoint Sites"),
    (SUBMODULE_DIR / "lists.yaml", in_lists, "Microsoft Graph — SharePoint Lists"),
    (SUBMODULE_DIR / "pages.yaml", in_pages, "Microsoft Graph — SharePoint Pages"),
    (SUBMODULE_DIR / "termstore.yaml", in_termstore, "Microsoft Graph — SharePoint Term Store"),
    (SUBMODULE_DIR / "onenote.yaml", in_onenote, "Microsoft Graph — SharePoint OneNote"),
    (SUBMODULE_DIR / "embedded.yaml", in_embedded, "Microsoft Graph — SharePoint Embedded"),
    (SUBMODULE_DIR / "admin.yaml", in_admin, "Microsoft Graph — SharePoint Admin"),
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


def walk_ref_closure(seed_node, components: dict) -> dict[str, set[str]]:
    """
    Starting from $refs found in `seed_node`, walk the components graph
    transitively. Returns {section: {name, ...}} of components to keep.
    """
    kept: dict[str, set[str]] = {s: set() for s in COMPONENT_SECTIONS}
    queue: list[str] = list(collect_refs(seed_node))

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


def build_output(
    source_doc: dict,
    predicate: Callable[[str], bool],
    title: str,
) -> dict:
    selected_paths = {p: ops for p, ops in source_doc["paths"].items() if predicate(p)}

    kept = walk_ref_closure(selected_paths, source_doc.get("components", {}))
    pruned_components = subset_components(source_doc.get("components", {}), kept)

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
    if pruned_tags:
        out["tags"] = pruned_tags
    out["paths"] = selected_paths
    if pruned_components:
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

    for target, predicate, title in OUTPUTS:
        t1 = time.monotonic()
        out_doc = build_output(source_doc, predicate, title)
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
        size_mb = target.stat().st_size / (1024 * 1024)
        print(
            f"  {target.relative_to(REPO_ROOT)}: "
            f"{n_paths} paths, {n_ops} ops, {n_schemas} schemas, "
            f"{size_mb:.1f} MB ({time.monotonic() - t1:.1f}s)",
            flush=True,
        )

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
