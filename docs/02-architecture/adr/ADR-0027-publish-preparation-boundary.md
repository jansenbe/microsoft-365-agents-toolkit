# ADR-0027 - Publish preparation boundary

- **Status:** Accepted
- **Date:** 2026-10-23
- **Source:** User-approved minimal SDK extension to
  [SharePoint-first publication](../../03-specs/scenarios/publishing/sharepoint-first.md).

## Context

A host may need the selected tenant's verified app catalog URL to finalize its
application package. Freezing that package before discovery gives the host no
safe supported point to resolve catalog-dependent content. Moving the host's
archive extraction and transformation into the engine would expand ownership
beyond publication.

## Options considered

- A separate prepare/commit API with externally held authorization state.
- A typed, operation-local finalizer before final package validation.
- Host access to internal clients or transport interception.

## Decision

Use one optional typed callback in the public SDK publish execution options.
ATK validates the original package and obtains read-only verified account/catalog
context before invoking it. The callback exposes no token or client and returns
final artifact paths and typed lineage. ATK owns validation, confirmations,
recovery, and the single SharePoint-first write flow. Host-specific archive
extraction, transformation and composition stay in the host.

No new CLI flag, authentication provider, script runner, independent publish
loop or durable preparation session is introduced. Existing provider and
operation-scoped client contracts remain unchanged.

## Consequences

- The callback is awaited through cancellation; hosts retain staging until
  publication settles rather than detaching work or racing a timeout.
- Original/derived/final digests and selected identity are checked again before
  mutation; callbacks cannot bypass version, overwrite or recovery safeguards.
- The original project owns the sole journal. Versioned lineage is persisted
  with partial results; old unknown outcomes never become implicit permission
  to retry.
- The [operation spec](../../03-specs/operations/publishing/sharepoint-package.md)
  owns the public fields, validation limits and acceptance tests.
