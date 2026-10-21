// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Platform,
  TeamsAppInputs,
  CombinedPublishResult,
  PublishRequestDiagnostic,
  UserError,
} from "@microsoft/teamsfx-api";
import {
  CombinedPublishError,
  FxCoreClient,
  FxCorePublishResult,
} from "@microsoft/teamsfx-core";
import {
  readSharePointPackage,
  SharePointPackage,
  SharePointPackageReadErrorCode,
  SharePointPackageReadOptions,
} from "@microsoft/teamsfx-core/build/sharepointPackage";

const inputs: TeamsAppInputs = {
  projectPath: ".",
  platform: Platform.CLI,
  "package-file": "agent.zip",
  "sharepoint-package-file": "component.sppkg",
  expectedTenantId: "22222222-2222-4222-8222-2222222222bb",
  expectedSharePointSha256: "0".repeat(64),
};
// @ts-expect-error the declared expected tenant is a string
inputs.expectedTenantId = 42;
// @ts-expect-error the declared selector digest is a string
inputs.expectedSharePointSha256 = 42;
const factory: SharePointPackageReadOptions = {
  errorFactory: (code: SharePointPackageReadErrorCode) =>
    new UserError("host", code, code),
};
const inspected = readSharePointPackage("component.sppkg", ".", factory);
declare const snapshot: SharePointPackage;
const digest: string = snapshot.sha256;
const bytes: Buffer = snapshot.bytes;
declare const client: FxCoreClient;
const published = client.publish(inputs, {
  signal: new AbortController().signal,
});
declare const output: FxCorePublishResult;
const partial: CombinedPublishResult | undefined =
  output.sharePoint && output.agent && output.tenantId && output.recoveryFile
    ? {
        packagePath: output.packagePath,
        sharePoint: output.sharePoint,
        agent: output.agent,
        tenantId: output.tenantId,
        recoveryFile: output.recoveryFile,
      }
    : undefined;
declare const error: CombinedPublishError;
const same: CombinedPublishResult = error.publishResult;
const diagnostic: PublishRequestDiagnostic | undefined = same.agent.diagnostic;
// @ts-expect-error organization submission does not mean installed
same.agent.stage = "installed";
void [inspected, digest, bytes, published, partial, diagnostic];
