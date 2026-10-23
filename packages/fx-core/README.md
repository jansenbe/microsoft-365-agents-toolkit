# TeamsFx Core

The TeamsFx Core package implements shared capabilities for Microsoft 365 Agents Toolkit IDE Extensions and the CLI through API contracts defined in the [api](/packages/api).

## SharePoint publishing prototype SDK

`FxCoreClient.publish` accepts `TeamsAppInputs["sharepoint-package-file"]` for
SharePoint-first deployment and verification before organization submission.
Optional `expectedTenantId` constrains the actual selected tenant for this
combined path; omit the SharePoint property for unchanged ordinary publishing.
Optional `expectedSharePointSha256` binds an earlier local selector result to the
actual package snapshot and is checked before authentication or remote work.
Pass an `AbortSignal` through the existing execution options. Preserve
`CombinedPublishError.publishResult` on failure instead of masking it with a
host timeout or treating a rejected submission as success.

The supported local inspector is exported from both the core root and the lean
entry `@microsoft/teamsfx-core/build/sharepointPackage`:

```typescript
readSharePointPackage(
  file: string,
  projectPath: string,
  options?: SharePointPackageReadOptions
): Promise<Result<SharePointPackage, FxError>>
```

It reads a bounded snapshot without building, extracting, authenticating,
contacting a service, writing a journal, or loading toolkit resources. The result
contains `bytes`, `fileName`, `productId`, `version`, `skipFeatureDeployment` and
`sha256`; callers must treat the returned byte snapshot as immutable and must not
include raw bytes in human/JSON status output. Error messages default to stable
codes (`InvalidPackage`, `UnreadablePackage`, `ChangedPackage`).
`options.errorFactory?: (code: SharePointPackageReadErrorCode) => FxError` lets an
offline host provide its own localized errors without any toolkit assets.
The publication pipeline supplies ATK's localized factory itself.

Hosts must provide real human confirmation for required updates; do not reuse a
headless UI that automatically accepts modal prompts. Confirmation identifies
the package, product/manifest ID, version and intended action. Noninteractive
updates remain refused. Submission is for organization review, not personal
sideloading, approval or installation.

### Catalog-dependent local preparation

`FxCoreClient.publish(inputs, { signal, prepareSharePointPublish })` can await a
typed host finalizer after read-only account/catalog discovery and before final
application validation. The original `sharepoint-package-file`,
`expectedSharePointSha256`, tenant constraint and project remain authoritative.
The callback receives only readonly identity/catalog metadata and cancellation;
tokens and clients stay inside the operation.

The host returns final application and derived SharePoint package paths plus
`SharePointPublishLineage`. ATK validates the final snapshots, rechecks the
selected context and inputs before writes, and uses its existing confirmation,
deployment, verification and organization submission flow. It does not extract
or transform embedded packages. The callback must return an error for unresolved
or invalid host content, never a fallback artifact.

Keep staged files alive until the publish promise settles, even after abort.
Partial errors retain `publishResult.lineage` and the original `projectPath`;
the original project's existing recovery location remains the sole journal.
Prepared journals use schema 2 and old schema-1 unknown submissions still block
replay. A prepared journal cannot be downgraded by omitting the finalizer.
See the [preparation contract](../../docs/03-specs/operations/publishing/sharepoint-package.md#host-preparation)
for the exact public types, digest validation and host-attested provenance boundary.

## Data Collection.

The software may collect information about you and your use of the software and send it to Microsoft. Microsoft may use this information to provide services and improve our products and services. You may turn off the telemetry as described in the repository. There are also some features in the software that may enable you and Microsoft to collect data from users of your applications. If you use these features, you must comply with applicable law, including providing appropriate notices to users of your applications together with a copy of Microsoft's privacy statement. Our privacy statement is located at https://go.microsoft.com/fwlink/?LinkID=824704. You can learn more about data collection and use in the help documentation and our privacy statement. Your use of the software operates as your consent to these practices.

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Contributing

There are many ways in which you can participate in the project, for example:

- [Submit bugs and feature requests](https://github.com/OfficeDev/TeamsFx/issues), and help us verify as they are checked in
- Review [source code changes](https://github.com/OfficeDev/TeamsFx/pulls)

If you are interested in fixing issues and contributing directly to the code base, please see the [Contributing Guide](./CONTRIBUTING.md).

## Reporting Security Issues

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, please report them to the Microsoft Security Response Center (MSRC) at [https://msrc.microsoft.com/create-report](https://msrc.microsoft.com/create-report).

If you prefer to submit without logging in, send email to [secure@microsoft.com](mailto:secure@microsoft.com). If possible, encrypt your message with our PGP key; please download it from the the [Microsoft Security Response Center PGP Key page](https://www.microsoft.com/en-us/msrc/pgp-key-msrc).

You should receive a response within 24 hours. If for some reason you do not, please follow up via email to ensure we received your original message. Additional information can be found at [microsoft.com/msrc](https://www.microsoft.com/msrc).

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.

## License

Copyright (c) Microsoft Corporation. All rights reserved.

Licensed under the [MIT](LICENSE.txt) license.
