_Author_: Thisaru Guruge \
_Created_: 2026-06-04 \
_Updated_: 2026-06-04 \
_Edition_: Swan Lake

# Sanitation for OpenAPI specification

This document records the sanitation performed on the official Microsoft Graph v1.0 OpenAPI specification to produce the SharePoint connector specs. The upstream spec is `docs/spec/graphexplorer.yaml`, obtained from Microsoft Graph Explorer (`https://aka.ms/graph/api/v1.0/openapi.yaml`). The transforms below are applied programmatically by `build-config/extract-spec.py` so the result is reproducible.

## 1. Scope reduction — SharePoint only

The upstream document covers the entire Microsoft Graph surface (mail, calendar, teams, identity, security, etc.). Only paths relevant to SharePoint are kept:

- `/sites/**` — sites, sub-sites, lists, items, pages, content types, columns, permissions, analytics, term store, OneNote on sites, and site-level actions/functions.
- `/storage/fileStorage/**` — SharePoint Embedded (file storage containers, container types, migration jobs, recycle bin, SharePoint groups).
- `/admin/sharepoint` — tenant-level SharePoint admin settings.

OneDrive APIs (`/drives/**`, `/shares/**`, `/me/drive`, etc.) are intentionally excluded because they are covered by the separate `ballerinax/microsoft.onedrive` connector. SharePoint document libraries remain reachable via `/sites/{id}/drive` and `/sites/{id}/drives`; consumers needing item-level drive operations should compose with the OneDrive connector.

## 2. Submodule partitioning

The consolidated SharePoint spec is also partitioned into seven mutually exclusive submodule specs to keep each Ballerina submodule focused. The partition rules below are applied top to bottom — every selected path matches exactly one rule.

| Submodule (file)                 | Rule                                                                                | Paths | Ops  |
| ---                              | ---                                                                                 | ---:  | ---: |
| `submodules/admin.yaml`          | `/admin/sharepoint` or `/admin/sharepoint/**`                                       |   2   |   6  |
| `submodules/embedded.yaml`       | `/storage/fileStorage` or `/storage/fileStorage/**`                                 | 110   | 181  |
| `submodules/lists.yaml`          | `/sites/{site-id}/lists/**`                                                         |  77   | 118  |
| `submodules/pages.yaml`          | `/sites/{site-id}/pages/**`                                                         |  42   |  68  |
| `submodules/termstore.yaml`      | `/sites/{site-id}/termStore/**` or `/sites/{site-id}/termStores/**`                 | 308   | 507  |
| `submodules/onenote.yaml`        | `/sites/{site-id}/onenote/**`                                                       | 103   | 153  |
| `submodules/sites.yaml`          | Anything else under `/sites` (root, entity, sub-sites, columns, externalColumns, contentTypes, permissions, analytics, operations, items, drive, drives, createdByUser, lastModifiedByUser, `microsoft.graph.*` actions/functions) |  99   | 142  |
| `sharepoint.yaml` (consolidated) | Union of all seven                                                                  | 741   | 1175 |

## 3. Component pruning

The upstream `components` block contains 4,532 schemas plus shared `responses`, `parameters`, `examples`, and `requestBodies` covering the whole Graph API. For each output file the extractor:

1. Walks every `$ref` reachable from the selected paths.
2. Resolves each referenced component, then walks its subtree for further `$ref`s.
3. Repeats until the closure stabilises.
4. Additionally pulls in `components/examples/<name>` for every kept `components/schemas/<name>` — see §3.1 below.
5. Emits only the items in that closure under `components.<section>`.

Every emitted spec is verified to have zero dangling `$ref`s before being written out.

### 3.1 Examples and the `@odata.type` annotation

The upstream `components.examples` block (~2,490 entries) holds one sample JSON payload per Microsoft Graph type, e.g. `components/examples/microsoft.graph.sitePage`. These payloads are the **only** place where `@odata.type` annotations appear in the spec — they show callers which concrete subtype to send/expect for polymorphic OData fields. No upstream path operation references these examples via `$ref`, so a strict transitive walk would drop them entirely (and with them, every `@odata.type` occurrence — 4,164 of them).

To preserve the `@odata.type` annotations, the extractor adds an extra step: after the `$ref` closure is built, for every kept schema name `X` that also exists in `components.examples`, the matching example entry is included in the output. This keeps the example/schema correspondence intact (one-to-one by name) without inflating the output with unrelated examples.

### 3.2 Polymorphism — derived schemas for path-referenced bases

Microsoft Graph models inheritance via `allOf: [{$ref: <base>}, ...]`. The `$ref` edge points *from derived to base*, so a transitive `$ref` walk never reaches derived types. When a path operation directly references a base schema as a request/response body (e.g. `microsoft.graph.baseSitePage`, `microsoft.graph.webPart`, `microsoft.graph.sharePointMigrationEvent`), the actual payload at runtime may be any derived schema, disambiguated only by the `@odata.type` annotation in the JSON body. Without those derived schemas, the generated Ballerina types can't represent the concrete subtype payloads (e.g. `pages.yaml` would expose `webPart` but not the concrete `standardWebPart` / `textWebPart`).

The extractor's second pass closes this gap: for every base schema *directly* referenced from a selected path, all transitively derived schemas are included. Two important guards keep this from dragging in the whole Graph surface:

- **Only `microsoft.graph.*` bases participate.** OData infrastructure types such as `BaseCollectionPaginationCountResponse` are excluded because their "derivatives" are typed pagination wrappers (one `*CollectionResponse` per entity type, >1,200 entries), not real `@odata.type` polymorphic subtypes.
- **Only path-referenced bases participate.** A schema that ends up in `components/schemas` only because it was pulled in transitively (e.g. `microsoft.graph.entity`, the root of nearly every Graph type — 689 direct descendants) does not get its derived chain expanded.

## 4. Cosmetic changes

- `info.title` is set per file (e.g. *Microsoft Graph — SharePoint Lists*).
- `info.description` is replaced with `"SharePoint subset of the Microsoft Graph v1.0 OpenAPI specification (generated from graphexplorer.yaml)."`.
- Top-level `tags` are filtered to only those names referenced by the retained operations.
- All other upstream fields (`openapi`, `info.version`, `info.x-ms-generated-by`, `servers`, path/operation/component contents) are preserved verbatim.

## 5. Security scheme injection

The upstream Microsoft Graph spec ships no `securitySchemes` (Microsoft documents Graph auth out-of-band). For each emitted file the extractor injects an `azureOAuth2` Azure AD OAuth2 scheme (mirroring the style used by `ballerinax/microsoft.outlook.mail`):

- `type: oauth2`, both `authorizationCode` (delegated) and `clientCredentials` (application) flows.
- `authorizationUrl` = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize`, `tokenUrl` = `https://login.microsoftonline.com/common/oauth2/v2.0/token`.
- Scopes are tailored per submodule; the same set is listed under both flows, with the application-flow descriptions suffixed with `(application permission)`.
- A top-level `security: [{azureOAuth2: [...]}]` advertises a representative subset (typically the read + read-write pair) as a sensible default.

Scopes per submodule:

| Submodule | Scopes |
| --- | --- |
| `sites`, `lists`, `pages` | `Sites.Read.All`, `Sites.ReadWrite.All`, `Sites.Manage.All`, `Sites.FullControl.All`, `Sites.Selected` |
| `termstore` | `TermStore.Read.All`, `TermStore.ReadWrite.All` |
| `onenote` | `Notes.Read`, `Notes.ReadWrite`, `Notes.Read.All`, `Notes.ReadWrite.All`, `Notes.Create` |
| `embedded` | `FileStorageContainer.Selected` |
| `admin` | `SharePointTenantSettings.Read.All`, `SharePointTenantSettings.ReadWrite.All` |
| `sharepoint` (consolidated) | Union of all of the above |

## 6. SnakeYAML size cap and JSON fallback

`bal openapi` parses YAML through SnakeYAML, which enforces a single-document cap of **3,145,728 code points** (~3 MB). Specs larger than this fail to parse with `The incoming YAML document exceeds the limit: 3145728 code points.` and produce **no generated code at all** — which can look like "the auth configuration is missing" because every other artifact is missing too.

All seven submodule YAMLs sit comfortably under the cap and `bal openapi` consumes them without issue. The consolidated `sharepoint.yaml` (~3.8 MB) exceeds the cap, so the extractor additionally emits **`docs/spec/sharepoint.json`** (~3.1 MB, compact JSON). Jackson's JSON parser has no comparable cap, so the JSON file generates successfully. Use the JSON variant whenever you need the consolidated surface in a tool that goes through `bal openapi` (or any other Swagger-Parser-based tooling).

## 7. Reproducibility / OpenAPI CLI commands

Regenerate every spec from the upstream by running this from the repository root:

```bash
python3 build-config/extract-spec.py
```

Requirements: Python 3.10+ and PyYAML (`pip install pyyaml`). The extractor is deterministic — re-running it on an unchanged source produces byte-identical output.

To generate the Ballerina client from a spec, run from the repository root:

```bash
# Consolidated SharePoint client — use the JSON variant (YAML exceeds the SnakeYAML cap)
bal openapi -i docs/spec/sharepoint.json --mode client

# Per-submodule clients — YAML works directly
bal openapi -i docs/spec/submodules/lists.yaml --mode client
bal openapi -i docs/spec/submodules/sites.yaml --mode client
# ... and likewise for pages, termstore, onenote, embedded, admin
```

Note: The license year in generated files is hardcoded to 2024; change if necessary.
