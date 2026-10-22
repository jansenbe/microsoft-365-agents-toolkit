# Build and test the SharePoint-first publishing prototype

This guide covers [SCN-SPO-PUBLISH](sharepoint-first.md). All paths and identities
below are placeholders or derived from the current checkout. Local validation
does not require a tenant and does not install or replace a global `atk`.

## Build from the committed source

Use Node.js 24 and the repository's declared package-manager configuration.
Restore dependencies in a fresh checkout, then build the affected package
closure:

```powershell
$repo = (git rev-parse --show-toplevel).Trim()
Set-Location $repo
pnpm install --frozen-lockfile
pnpm --filter @microsoft/app-manifest build
pnpm --filter @microsoft/teamsfx-api build
pnpm --filter @microsoft/m365-spec-parser build
pnpm --filter templates build
pnpm --filter @microsoft/teamsfx-core exec tsc -p . --incremental
pnpm --filter @microsoft/m365agentstoolkit-cli build

$atk = Join-Path $repo 'packages\cli\atk-local.ps1'
& $atk publish --help --telemetry false
```

The shim invokes `packages\cli\cli.js`, which loads `packages\cli\lib\index.js`.
The shared SDK is `packages\fx-core\build\index.js`; public contracts are in
`packages\api\build\index.js`. If PowerShell script execution is restricted, use
the Node entry point directly:

```powershell
node (Join-Path $repo 'packages\cli\cli.js') publish --help --telemetry false
```

## Safe local checks

The following command must fail for the missing SharePoint file, before sign-in
or publication:

```powershell
& $atk publish --folder $repo `
  --package-file .local\not-selected.zip `
  --sharepoint-package-file .local\missing.sppkg `
  --interactive false --telemetry false
$LASTEXITCODE # nonzero
```

Focused regression suites use synthetic archives, identities and mocked HTTP:

```powershell
pnpm --filter @microsoft/teamsfx-core exec vitest run `
  tests\component\driver\deploy\spfx `
  tests\component\driver\teamsApp\sharePointPublish.test.ts `
  tests\component\driver\teamsApp\publishAppPackage.test.ts `
  tests\component\driver\teamsApp\teamsAppMgr.test.ts `
  tests\error\publish.test.ts tests\client\graphClient.test.ts `
  tests\core\FxCoreClient.test.ts --config vitest.config.ts

pnpm --filter @microsoft/m365agentstoolkit-cli exec vitest run `
  tests\unit\commands.tests.ts tests\unit\engine.tests.ts `
  tests\unit\helper.tests.ts --config vitest.config.ts
```

## Explicitly authorized live testing

**Publishing mutates remote catalogs.** Select the exact project, agent ZIP,
SharePoint package, account and tenant before a live run. Approval of a
SharePoint deployment does not authorize an embedded agent ZIP automatically.
A deployed package does not prove that backend/API permissions are approved.

The executable is isolated to this checkout, but normal ATK sign-in uses its
per-user account cache. The following example is a separate authentication
action; replace the tenant placeholder with the intended tenant GUID:

```powershell
$tenantId = '<selected-tenant-guid>'
& $atk auth login m365 --tenant $tenantId --interactive false --telemetry false
& $atk auth show m365 --telemetry false
```

`--interactive false` preserves the explicit tenant option through the existing
CLI argument filter; it does not suppress OAuth sign-in UI. An account name
printed by `auth show` alone is not proof of the tenant or Graph grants of a
publication token. Do not clear account caches to diagnose a network error or
grant permissions merely because a request returned HTTP 403.

### Ordinary organization submission

This is an independent agent publication, not preparation for the combined
command. Do not run it first merely to test the environment.

```powershell
$project = (Resolve-Path '<selected-project-folder>').Path
$agentZip = Join-Path $project 'appPackage\build\appPackage.dev.zip'
& $atk publish --folder $project --package-file $agentZip `
  --interactive false --telemetry false
```

The legacy no-flag interactive mode discards explicit package options. The
example uses `--interactive false` to retain the chosen ZIP; that legacy mode
may implicitly confirm an existing-app update. Run it only when that exact
submission/update is intended.

### Combined SharePoint-first publication

The `.sppkg` and agent ZIP are two distinct artifacts. Relative SharePoint paths
resolve against `--folder`; absolute paths are also accepted.

```powershell
$sppkg = Join-Path $project 'sharepoint\solution\component.sppkg'
& $atk publish --folder $project --package-file $agentZip `
  --sharepoint-package-file $sppkg --interactive true --telemetry false
```

Alternatively, let ordinary ATK packaging prepare the agent from the project
manifest while supplying an already-built SharePoint package:

```powershell
& $atk publish --folder $project --env dev `
  --sharepoint-package-file .\sharepoint\solution\component.sppkg `
  --interactive true --telemetry false
```

The flagged path preserves explicit artifacts in interactive mode so a human
can confirm specific updates. Both packages and known prerequisites are checked
before catalog writes. SharePoint is uploaded, deployed and verified before
the organization-catalog agent submission. An upload response alone is not
enough. Exact already-deployed bytes are verified and reused; changed bytes
under the same version, downgrades and conflicting identities are rejected.
An upgrade requires genuine confirmation, never an implied noninteractive yes.

Successful submission is **subject to organization review**, not proof of
approval, installation or rendering. No path in the new operation creates a
catalog, approves API scopes, installs into other sites or synchronizes to Teams.

## Partial failure and safe recovery

The token-free record is under
`<project>\.local\sharepoint-publish\<tenant-and-product-hash>.json`.
Keep this directory out of source control. It retains separate artifact
identities/hashes and both stage states; `lastSuccessfulStage` preserves a
completed SharePoint action when the next request's outcome is uncertain.
There is no cross-service transaction or automatic rollback.

SharePoint failure blocks the agent. Agent failure after SharePoint verification
does not retract the package. A definitive rejected request can be retried only
after resolving the issue and reconciling prior attempts. Combined-flow mutation
POSTs are not automatically retried.

Timeouts, lost responses and cancellation during a write can leave
`unknownOutcome`. Inspect the exact app/version and any pending submission
before another attempt. Matching version alone cannot establish identical
content. Preserve the journal; do not delete or edit it to make a command appear
successful. Ordinary no-flag attempts are not tracked by the combined journal,
so an earlier `failed` record does not establish a later ordinary attempt's
outcome.

## Diagnostic and SDK contracts

Agent failures identify Graph catalog read versus submission and retain bounded,
sanitized HTTP status, service code/message, request ID and transport code.
HTTP 403 alone does not identify a missing role or policy. A transport failure
does not establish a permission problem or identify a proxy. Use authorized
read-only reconciliation and the service evidence before changing permissions
or repeating publication.

Both paths use the same configured Graph base, provider, requested
`AppCatalog.ReadWrite.All` scope, ZIP content type and `requiresReview=true`
semantics. The opt-in path prepares immutable bytes and identity, checks strict
catalog snapshots, and uses one `submitPreparedAppForReview` POST with
cancellation, a 60-second timeout and disabled redirects. The ordinary
`publishTeamsApp`/`publishTeamsAppUpdate` path retains its existing retries and
fallbacks.

SDK callers use `FxCoreClient.publish(inputs, { signal })`.
`CombinedPublishError.publishResult` retains partial results and
`agent.diagnostic` contains the sanitized request evidence. Optional
`expectedTenantId` binds the combined operation to a tenant GUID; optional
`expectedSharePointSha256` binds the selector's digest to the immutable local
snapshot before authentication. Do not mask a settled result with a host timeout.

The root and `@microsoft/teamsfx-core/build/sharepointPackage` entry expose
`readSharePointPackage` and its types. The lean entry supports offline local
inspection without lifecycle/auth/resource access, including failures. Hosts
can supply `errorFactory` for their own localized messages and must not emit
the returned raw package bytes into status output.

## Permissions and interpretation

The [Graph publish API](https://learn.microsoft.com/en-us/graph/api/teamsapp-publish?view=graph-rest-beta)
documents delegated `AppCatalog.Submit` as least privilege and
`AppCatalog.ReadWrite.All` as a supported higher privilege. The latter requires
administrator consent; this prototype preserves the existing ATK scope and
does not request another grant automatically.

The SharePoint operation obtains a separate SharePoint-audience token. Root-site
discovery requires an effective Graph site-read grant; Graph catalog permission
alone is not sufficient. The prototype qualifies same-origin commercial/GCC
catalogs and fails closed for unsupported cross-origin/cloud targets.

Personal upload and organization submission are distinct, as described by
[Teams custom-app settings](https://learn.microsoft.com/en-us/microsoftteams/teams-custom-app-policies-and-settings).
The [Graph update API](https://learn.microsoft.com/en-us/graph/api/teamsapp-update?view=graph-rest-beta)
and [submission guide](https://learn.microsoft.com/en-us/microsoftteams/submit-approve-custom-apps)
describe different role requirements for some update cases. Neither proves the
cause of a particular rejection. This prototype does not silently switch to
personal sideloading when organization publication fails.

## Source versus historical local artifacts

This committed source includes public-safe fixtures and documentation.
It is not asserted to be byte-identical to the
historical local `-spopublish.20261020.3` artifact series. Those immutable
tarballs, original local provenance snapshots and live diagnostic evidence are
not part of this branch. Build a new release identity from this checkout;
do not regenerate or overwrite the historical series.
