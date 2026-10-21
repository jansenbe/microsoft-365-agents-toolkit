// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CombinedPublishResult, FxError, Result, UserError, err, ok } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import lockfile from "proper-lockfile";
import { getLocalizedString } from "../../../common/localizeUtils";
import { UserCancelError } from "../../../error/common";
import { PublishRemoteError, remotePublishError } from "../../../error/publish";
import { SharePointPackageService, confirmPublish } from "../deploy/spfx/sharePointPackage";
import { SharePointPackage, guidPattern, record, sha256 } from "../deploy/spfx/utility/packageFile";
import { packageError } from "../deploy/spfx/utility/packageError";
import { DriverContext } from "../interface/commonArgs";
import { prepareAgentPublication } from "./preparedPublication";
import { PublishAppPackageDriver } from "./publishAppPackage";

export class CombinedPublishError extends UserError {
  constructor(
    public readonly publishResult: CombinedPublishResult,
    cause: FxError
  ) {
    super({
      source: "SharePointPublish",
      name: "CombinedPublishIncomplete",
      message: getLocalizedString(
        "publish.sharepoint.Partial",
        publishResult.sharePoint.stage,
        publishResult.agent.stage,
        publishResult.recoveryFile,
        cause.message
      ),
      error: cause,
    });
  }
}

/** Coordinates, but never substitutes for, the existing agent publisher. */
export async function publishWithSharePoint(
  appPackagePath: string,
  pkg: SharePointPackage,
  context: DriverContext,
  publisher: PublishAppPackageDriver,
  validatedAgentHash: string,
  expectedTenantId?: string
): Promise<Result<CombinedPublishResult, FxError>> {
  if (expectedTenantId !== undefined && !guidPattern.test(expectedTenantId)) {
    return err(packageError("ExpectedTenantMismatch"));
  }
  const prepared = await prepareAgentPublication(appPackagePath, context);
  if (prepared.isErr()) return err(prepared.error);
  const agent = prepared.value;
  if (
    expectedTenantId !== undefined &&
    expectedTenantId.toLowerCase() !== agent.identity.tenantId.toLowerCase()
  ) {
    return err(packageError("ExpectedTenantMismatch"));
  }
  if (agent.sha256 !== validatedAgentHash) return err(packageError("AgentChanged"));
  const recoveryFile = path.join(
    context.projectPath,
    ".local",
    "sharepoint-publish",
    `${sha256(Buffer.from(`${agent.identity.tenantId}:${pkg.productId}`))}.json`
  );
  const result: CombinedPublishResult = {
    packagePath: appPackagePath,
    tenantId: agent.identity.tenantId,
    recoveryFile,
    sharePoint: {
      stage: "notStarted",
      productId: pkg.productId,
      version: pkg.version,
      sha256: pkg.sha256,
    },
    agent: {
      stage: "notStarted",
      manifestId: agent.manifestId,
      version: agent.version,
      sha256: agent.sha256,
    },
  };
  const fail = (cause: FxError): Result<CombinedPublishResult, FxError> =>
    err(new CombinedPublishError(result, cause));
  let release: (() => Promise<void>) | undefined;
  const save = async (): Promise<void> => {
    const temporary = `${recoveryFile}.${process.pid}.tmp`;
    await fs.writeFile(
      temporary,
      JSON.stringify(
        {
          schemaVersion: 1,
          objectId: agent.identity.objectId,
          ...result,
        },
        undefined,
        2
      ),
      { mode: 0o600 }
    );
    await fs.rename(temporary, recoveryFile);
  };
  try {
    await fs.ensureDir(path.dirname(recoveryFile));
    release = await lockfile.lock(recoveryFile, { realpath: false, retries: 0 });
    let previous: unknown;
    try {
      previous = JSON.parse(await fs.readFile(recoveryFile, "utf8"));
    } catch (error) {
      if (!(error instanceof Error && "code" in error && error.code === "ENOENT")) {
        return fail(packageError("RecoveryInvalid"));
      }
    }
    let alreadySubmitted = false;
    let reconcileAgent = false;
    if (previous !== undefined) {
      if (
        !record(previous) ||
        previous.schemaVersion !== 1 ||
        previous.tenantId !== result.tenantId ||
        previous.objectId !== agent.identity.objectId ||
        !record(previous.sharePoint) ||
        !record(previous.agent) ||
        previous.sharePoint.productId !== pkg.productId
      ) {
        return fail(packageError("RecoveryInvalid"));
      }
      const oldAgent = previous.agent;
      if (
        !["notStarted", "submitted", "failed", "cancelled", "unknownOutcome"].some(
          (stage) => stage === oldAgent.stage
        )
      ) {
        return fail(packageError("RecoveryInvalid"));
      }
      const sameAgent =
        oldAgent.manifestId === agent.manifestId && oldAgent.sha256 === agent.sha256;
      if (oldAgent.stage === "unknownOutcome") {
        // prepareAgentPublication has already read the live organization catalog.
        // Catalog version alone cannot prove that an ambiguous POST contained our bytes.
        reconcileAgent = true;
      }
      alreadySubmitted =
        oldAgent.stage === "submitted" &&
        sameAgent &&
        agent.snapshot.catalogId === oldAgent.catalogId &&
        agent.snapshot.versions.includes(agent.version);
      if (oldAgent.stage === "submitted" && !alreadySubmitted) {
        const knownPriorSubmission =
          oldAgent.manifestId === agent.manifestId &&
          typeof oldAgent.version === "string" &&
          agent.snapshot.catalogId === oldAgent.catalogId &&
          agent.snapshot.versions.includes(oldAgent.version);
        // A deliberate new version can advance a known submission, not an uncertain one.
        reconcileAgent = !knownPriorSubmission || oldAgent.version === agent.version;
      }
    }
    const service = new SharePointPackageService(context);
    const sharePoint = await service.prepare(pkg, agent.identity);
    if (sharePoint.isErr()) return fail(sharePoint.error);
    result.sharePoint = sharePoint.value.result;
    if (reconcileAgent) {
      result.agent.stage = "unknownOutcome";
      return fail(packageError("RecoveryRequired"));
    }
    if (alreadySubmitted) {
      result.agent.stage = "submitted";
      result.agent.catalogId = agent.snapshot.catalogId;
    } else if (agent.snapshot.catalogId) {
      const confirmation = await confirmPublish(
        context,
        getLocalizedString(
          "publish.sharepoint.ConfirmAgent",
          path.basename(appPackagePath),
          agent.manifestId,
          agent.version
        )
      );
      if (confirmation.isErr()) return fail(confirmation.error);
    }
    await save();
    const deployed = await service.deploy(sharePoint.value, async (state) => {
      result.sharePoint = state;
      await save();
    });
    if (deployed.isErr()) return fail(deployed.error);
    result.sharePoint = deployed.value;
    if (alreadySubmitted) {
      result.agent.stage = "submitted";
      result.agent.catalogId = agent.snapshot.catalogId;
      await save();
      context.logProvider.info(
        getLocalizedString("publish.sharepoint.Complete", result.recoveryFile)
      );
      return ok(result);
    }
    if (context.signal?.aborted) {
      result.agent.stage = "cancelled";
      await save();
      return fail(new UserCancelError("SharePointPublish"));
    }
    const current = await agent.graph.getPublishSnapshot(
      agent.token,
      agent.manifestId,
      context.signal
    );
    if (current.fingerprint !== agent.snapshot.fingerprint)
      return fail(packageError("RecoveryRequired"));
    if (sha256(await fs.readFile(appPackagePath)) !== agent.sha256)
      return fail(packageError("AgentChanged"));
    if (context.signal?.aborted) {
      result.agent.stage = "cancelled";
      await save();
      return fail(new UserCancelError("SharePointPublish"));
    }
    result.agent.stage = "unknownOutcome";
    await save();
    if (context.signal?.aborted) {
      result.agent.stage = "cancelled";
      await save();
      return fail(new UserCancelError("SharePointPublish"));
    }
    const published = await publisher.execute(
      { appPackagePath, preparedPublication: agent },
      context
    );
    if (published.result.isErr()) {
      const error = published.result.error;
      if (error instanceof PublishRemoteError) {
        result.agent.diagnostic = error.diagnostic;
        if (error.isDefinitiveRejection) result.agent.stage = "failed";
        await save();
      }
      return fail(error);
    }
    result.agent.stage = "submitted";
    result.agent.catalogId = published.result.value.get("TEAMS_APP_PUBLISHED_APP_ID");
    await save();
    context.logProvider.info(
      getLocalizedString("publish.sharepoint.Complete", result.recoveryFile)
    );
    return ok(result);
  } catch (error) {
    if (error instanceof PublishRemoteError && error.diagnostic.target !== "sharePoint") {
      result.agent.diagnostic = error.diagnostic;
      await save();
    }
    return fail(remotePublishError(error));
  } finally {
    if (release) await release();
  }
}
