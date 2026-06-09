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
4. Emits only the items in that closure under `components.<section>`.

Every emitted spec is verified to have zero dangling `$ref`s before being written out.

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

## 6. Reproducibility / OpenAPI CLI commands

Regenerate every spec from the upstream by running this from the repository root:

```bash
python3 build-config/extract-spec.py
```

Requirements: Python 3.10+ and PyYAML (`pip install pyyaml`). The extractor is deterministic — re-running it on an unchanged source produces byte-identical output.

To generate the Ballerina client from a spec, run from the repository root:

```bash
# Consolidated SharePoint client
bal openapi -i docs/spec/sharepoint.yaml

# Per-submodule clients
bal openapi -i docs/spec/submodules/lists.yaml
bal openapi -i docs/spec/submodules/sites.yaml
# ... and likewise for pages, termstore, onenote, embedded, admin
```

Note: The license year in generated files is hardcoded to 2024; change if necessary.
