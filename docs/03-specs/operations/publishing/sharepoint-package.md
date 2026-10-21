# Publish a standalone SharePoint package

**Status:** Prototype. This page is the behavioral contract for the committed
standalone package operation.
**Domain:** Publish

## Contract

`TeamsAppInputs["sharepoint-package-file"]` is optional. Relative paths resolve
against the project folder. The input is a standalone SPFx `.sppkg`, not a source
configuration or an agent ZIP. Existing `spfx/deploy` YAML configuration remains
compatible; both paths share the operation-scoped SharePoint HTTP client.

Combined publication returns package path and separate SharePoint and agent
states. Failures carry the same partial result. A token-free recovery record
under the project `.local` folder is written before mutations. Recovery reads
remote state and compares downloaded package bytes before skipping an upload.
Unknown agent submission outcomes require operator reconciliation; they are
never automatically retried. The prototype supports same-origin commercial/GCC
SharePoint catalogs; cross-origin/multi-geo catalogs and unsupported publishing
clouds fail closed before mutation.

## Acceptance Criteria

| ID | Runtime | Purpose | Gate | Harness | Expected result |
|---|---|---|---|---|---|
| SPO-01 | L1 | operation-integration | required | archive/temp directory | Standalone package metadata and SHA-256 come from bounded immutable bytes; malformed, duplicate, unsafe, external-entity or non-client-side archives fail before HTTP. |
| SPO-02 | L1 | operation-integration | required | HTTP adapter | Selected Graph and SPO identities match; root discovery uses configured Graph and JSON Accept; missing grants/catalog or foreign authority fail before writes. |
| SPO-03 | L1 | operation-integration | required | HTTP adapter | Independent clients preserve their own origins; binary Add captures UniqueId separately from product ID, safely escapes filenames, and Deploy uses packaged SkipFeatureDeployment. |
| SPO-04 | L1 | operation-integration | required | HTTP adapter | Bounded polling requires expected version, client-side classification, deployment flags and no package errors. Cancellation and failures prevent continuation. GET throttling is bounded; POSTs are not blindly retried. |
| SPO-05 | L1 | operation-integration | required | HTTP adapter | Same-product update requires confirmation; downgrade and filename/identity collision are refused. Exact downloaded bytes plus verified state permit reuse without upload/deploy. |
| SPO-06 | L1 | operation-integration | required | temp directory/HTTP adapter | Journal preserves independent stage states across failure/retry; ambiguous writes are explicit, no destructive rollback, no repeated ambiguous agent submission. |
| SDK-01 | L1 | compatibility | required | public entry/temp directory | Root and documented lean entry expose the same bounded read-only package inspector and types. Missing, malformed and unsafe inputs work without toolkit resources, auth or network; a caller error factory can provide host-localized messages. |
| SDK-02 | L1 | operation-integration | required | HTTP adapter | An explicit expected tenant GUID is checked against the actual prepared Graph identity before SharePoint preflight, journal creation or writes. Omission preserves existing behavior. |
| SDK-03 | L1 | scenario | required | confirmation fake | Required update confirmations identify the filename, product/manifest ID, target version and planned operation; refusal/noninteractive behavior remains unchanged. |
| SDK-04 | L1 | compatibility | required | manager/file snapshot | Optional expectedSharePointSha256 must be 64 hex characters and match the immutable package snapshot before agent preparation, authentication, remote work or journal writes; stale selector metadata cannot authorize different package bytes. |

## Flow

```mermaid
flowchart TD
  A[Read and validate immutable sppkg] --> B[Selected-account discovery and catalog readiness]
  B --> C[Reconcile product, filename, version and remote bytes]
  C --> D{Exact bytes already deployed?}
  D -->|yes| G[Verify and return]
  D -->|no| E[Confirm update if needed; record intent; Add]
  E --> F[Bounded visibility polling; record intent; Deploy]
  F --> G
```

## Boundary

No catalog creation, Teams synchronization, site installation, API permission
approval, SPFx build, host-specific integration, or automatic rollback. Remote mutation
tests require approval of the exact artifacts and actions.

## Invariants

- Never send Graph tokens to SharePoint, follow token-bearing redirects, infer a
  tenant hostname from an email, or share mutable client origins.
- No archive extraction. Limit package bytes, entry count, expanded total and
  manifest size; reject DTD/entity declarations and ambiguous metadata.
- Product ID, catalog item ID, agent ID, versions and hashes remain distinct.
- Upload success is not deployment success. Version equality is not byte equality.
- Failed/unknown writes remain observable even after cancellation.
