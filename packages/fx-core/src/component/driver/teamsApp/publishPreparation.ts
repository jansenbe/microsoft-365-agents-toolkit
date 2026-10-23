// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  PreparedSharePointPublishArtifacts,
  Result,
  SharePointPublishLineage,
  SharePointPublishPreparer,
  TeamsAppInputs,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import path from "path";
import { UserCancelError } from "../../../error/common";
import {
  PreparedSharePointCatalog,
  SharePointPackageService,
} from "../deploy/spfx/sharePointPackage";
import { packageError } from "../deploy/spfx/utility/packageError";
import {
  SharePointPackage,
  guidPattern,
  readPackageSnapshot,
  readSharePointPackage,
  record,
  sha256,
} from "../deploy/spfx/utility/packageFile";
import { DriverContext } from "../interface/commonArgs";

export interface PreparedSharePointPublication {
  catalog: PreparedSharePointCatalog;
  originalPackagePath: string;
  sharePointPackagePath: string;
  appPackagePath: string;
  appPackageBytes: Buffer;
  package: SharePointPackage;
  lineage: SharePointPublishLineage;
}

function digest(value: unknown): value is string {
  return typeof value === "string" && /^[a-f0-9]{64}$/i.test(value);
}

export function publishLineage(value: unknown): SharePointPublishLineage | undefined {
  if (
    !record(value) ||
    value.schemaVersion !== 1 ||
    !digest(value.originalSharePointSha256) ||
    !digest(value.derivedSharePointSha256) ||
    !digest(value.finalAppSha256) ||
    !digest(value.extractedPluginSha256) ||
    typeof value.transformationVersion !== "string" ||
    !/^[a-z0-9][a-z0-9._-]{0,63}$/i.test(value.transformationVersion) ||
    typeof value.selectedEmbeddedPackage !== "string" ||
    value.selectedEmbeddedPackage.length > 1024 ||
    /[\\:\u0000-\u001f\u007f?#]/.test(value.selectedEmbeddedPackage) ||
    value.selectedEmbeddedPackage
      .split("/")
      .some((segment) => !segment || segment === "." || segment === "..")
  ) {
    return undefined;
  }
  return Object.freeze({
    schemaVersion: 1,
    originalSharePointSha256: value.originalSharePointSha256.toLowerCase(),
    transformationVersion: value.transformationVersion,
    selectedEmbeddedPackage: value.selectedEmbeddedPackage,
    extractedPluginSha256: value.extractedPluginSha256.toLowerCase(),
    derivedSharePointSha256: value.derivedSharePointSha256.toLowerCase(),
    finalAppSha256: value.finalAppSha256.toLowerCase(),
  });
}

export async function verifyPreparedFiles(
  prepared: PreparedSharePointPublication,
  context: DriverContext
): Promise<Result<undefined, FxError>> {
  if (context.signal?.aborted) return err(new UserCancelError("SharePointPublish"));
  for (const [file, expected] of [
    [prepared.originalPackagePath, prepared.lineage.originalSharePointSha256],
    [prepared.sharePointPackagePath, prepared.lineage.derivedSharePointSha256],
    [prepared.appPackagePath, prepared.lineage.finalAppSha256],
  ]) {
    const current = await readPackageSnapshot(file, {
      errorFactory: () => packageError("PreparedPackageChanged"),
    });
    if (current.isErr()) return err(current.error);
    if (sha256(current.value) !== expected) return err(packageError("PreparedPackageChanged"));
    if (context.signal?.aborted) return err(new UserCancelError("SharePointPublish"));
  }
  return ok(undefined);
}

export async function finalizeSharePointPublication(
  inputs: TeamsAppInputs,
  original: SharePointPackage,
  prepare: SharePointPublishPreparer,
  context: DriverContext
): Promise<Result<PreparedSharePointPublication, FxError>> {
  if (inputs.expectedTenantId !== undefined && !guidPattern.test(inputs.expectedTenantId)) {
    return err(packageError("ExpectedTenantMismatch"));
  }
  const service = new SharePointPackageService(context);
  const discovery = await service.discover(undefined, inputs.expectedTenantId);
  if (discovery.isErr()) return err(discovery.error);
  const catalog = discovery.value;
  const originalPackagePath = path.resolve(inputs.projectPath, inputs["sharepoint-package-file"]!);
  if (context.signal?.aborted) return err(new UserCancelError("SharePointPublish"));
  let output: Result<PreparedSharePointPublishArtifacts, FxError>;
  try {
    output = await prepare(
      Object.freeze({
        projectPath: context.projectPath,
        tenantId: catalog.identity.tenantId,
        catalogUrl: catalog.catalogUrl,
        originalSharePoint: Object.freeze({
          packagePath: originalPackagePath,
          productId: original.productId,
          version: original.version,
          sha256: original.sha256,
          skipFeatureDeployment: original.skipFeatureDeployment,
        }),
        signal: context.signal,
      })
    );
  } catch {
    return err(
      context.signal?.aborted
        ? new UserCancelError("SharePointPublish")
        : packageError("PreparationFailed")
    );
  }
  if (context.signal?.aborted) return err(new UserCancelError("SharePointPublish"));
  if (!output || typeof output.isErr !== "function") return err(packageError("InvalidPreparation"));
  if (output.isErr()) return err(output.error);
  const value: unknown = output.value;
  if (
    !record(value) ||
    typeof value.sharePointPackagePath !== "string" ||
    !value.sharePointPackagePath ||
    typeof value.appPackagePath !== "string" ||
    !value.appPackagePath
  )
    return err(packageError("InvalidPreparation"));
  const lineage = publishLineage(value.lineage);
  if (!lineage || lineage.originalSharePointSha256 !== original.sha256) {
    return err(packageError("InvalidPreparation"));
  }
  const sharePointPackagePath = path.resolve(context.projectPath, value.sharePointPackagePath);
  const appPackagePath = path.resolve(context.projectPath, value.appPackagePath);
  const derived = await readSharePointPackage(sharePointPackagePath, context.projectPath, {
    errorFactory: packageError,
  });
  if (derived.isErr()) return err(derived.error);
  if (
    derived.value.sha256 !== lineage.derivedSharePointSha256 ||
    derived.value.productId !== original.productId ||
    derived.value.version !== original.version ||
    derived.value.skipFeatureDeployment !== original.skipFeatureDeployment
  )
    return err(packageError("InvalidPreparation"));
  const application = await readPackageSnapshot(appPackagePath, {
    errorFactory: () => packageError("InvalidPreparation"),
  });
  if (application.isErr()) return err(application.error);
  if (sha256(application.value) !== lineage.finalAppSha256) {
    return err(packageError("InvalidPreparation"));
  }
  const prepared: PreparedSharePointPublication = {
    catalog,
    originalPackagePath,
    sharePointPackagePath,
    appPackagePath,
    appPackageBytes: application.value,
    package: derived.value,
    lineage,
  };
  const current = await service.revalidate(catalog);
  if (current.isErr()) return err(current.error);
  const unchanged = await verifyPreparedFiles(prepared, context);
  if (unchanged.isErr()) return err(unchanged.error);
  return ok(prepared);
}
