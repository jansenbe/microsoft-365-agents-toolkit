# Package the committed SharePoint SDK prototype

This developer recipe creates a **new** local prerelease SDK closure from a
clean committed checkout. It does not publish a feed package, replace global
tools, authenticate to Microsoft 365, or call a tenant. Restore/build prerequisites
use the repository's normal package-manager and dependency-feed configuration;
credentials are never copied into the delivery.

## Build

Use Node.js 24 and the repository's declared package manager. In a fresh checkout:

```powershell
pnpm install --frozen-lockfile
pnpm --filter @microsoft/app-manifest build
pnpm --filter @microsoft/teamsfx-api build
pnpm --filter @microsoft/m365-spec-parser build
pnpm --filter templates build
pnpm --filter @microsoft/teamsfx-core exec tsc -p . --incremental
pnpm --filter @microsoft/teamsfx-core run typecheck:scenario-tooling
```

Do not reset or clean someone else's worktree to satisfy the clean-source gate.
Use a checkout of the desired commit. Generated build files are ignored by Git.

## Pack and qualify

Choose a fresh `YYYYMMDD.N` identity. The example below is illustrative; do not
reuse an identity that already names different package bytes.

```powershell
$release = '20261022.1'
node .github\scripts\sdk-prototype\pack-sdk.cjs $release
$output = Join-Path (git rev-parse --show-toplevel).Trim() ".local\sdk-prototype\$release"

Push-Location (Join-Path $output 'consumer')
npm install --ignore-scripts --no-audit --no-fund
Pop-Location

node .github\scripts\sdk-prototype\finalize-release.cjs $output
```

`ATK_SOURCE_ROOT` can select another clean checkout. `ATK_ARTIFACT_ROOT` can
select a new output directory. Windows Node distributions with a nonstandard
npm layout can set `NPM_CLI_JS` to their installed `npm-cli.js`; other platforms
use the `npm` executable. No shell command is assembled from package inputs.

The four standard npm tarballs contain core, API, app-manifest and spec-parser
under distinct prerelease identities. Internal dependency pins are exact; other
direct runtime versions come from the restored build. Runtime `tslib` is
declared explicitly rather than relying on incidental API hoisting. Nested
declarations, resource/template/schema roots and original MIT license texts are
included. No `workspace:*` reference remains in packed manifests.

The source record names the committed SHA/tree and tracked file hashes. Dirty
or untracked working-tree content is rejected rather than copied into provenance.
The two config `.gitignore` files excluded by npm are recorded explicitly.
The actual bundled metadata inputs, `templates/metadata/allTemplates.json` and
`defaultGeneratorTemplates.json`, are required. The remote release archive
`metadata.zip` is not a local runtime input.

Finalization moves the installed consumer, runs a real frozen offline `npm ci`,
denies runtime module resolution outside that consumer, blocks network access,
checks asset-free lean inspection on success and failure, verifies one API and
`CombinedPublishError` instance, checks all packed file hashes, and compiles
typed positive/negative contract assertions. It then repeats `npm pack` and
compares tarball bytes. `qualification.json`, real package locks, inventories and
license evidence record the checks actually executed.

All outputs stay under the ignored artifact directory. Review generated
provenance before sharing it; never publish credentials, private application
packages, tenant journals, diagnostic captures or `node_modules`.

## Historical artifact boundary

The retained `-spopublish.20261020.3` artifacts were made before this portable
source publication. This commit includes public-safe documentation and synthetic
test-fixture substitutions, so it is **not** asserted to reproduce those
historical bytes or their source-content identity. The recipe refuses that
reserved identity and never writes to the historical delivery directory.

Original local tarballs, source overlays, staging snapshots and live diagnostic
evidence are deliberately not checked in. Remote readers can inspect this
commit and build a new qualified set without another developer's checkout.
Generated template ZIP timestamps are captured build inputs: exact repeat
packing uses `provenance/staged-runtime-inputs.tgz` or the retained stage tree.
A new source build may regenerate that metadata and must use a new identity
if package bytes differ. No live tenant acceptance is implied by local checks.
