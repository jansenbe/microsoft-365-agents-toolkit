// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Result, SharePointPackageResult, err, ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import { setTimeout } from "timers/promises";
import {
  GraphTeamsAppCatalogReadWriteScopes,
  ResourceServiceType,
  SPFxScopes,
  getResourceServiceEndpoint,
} from "../../../../common/constants";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { UserCancelError } from "../../../../error/common";
import { remotePublishError } from "../../../../error/publish";
import { DriverContext } from "../../interface/commonArgs";
import {
  SharePointPackage,
  guidPattern,
  record,
  sha256,
  versionPattern,
} from "./utility/packageFile";
import { packageError } from "./utility/packageError";
import { SPOClient } from "./utility/spoClient";

export interface PublishIdentity {
  tenantId: string;
  objectId: string;
}

export interface PreparedSharePointDeployment {
  package: SharePointPackage;
  client: SPOClient;
  token: string;
  result: SharePointPackageResult;
  identical: boolean;
  overwrite: boolean;
  previousHash?: string;
  previousVersion?: string;
}

export async function confirmPublish(
  context: DriverContext,
  message: string
): Promise<Result<undefined, FxError>> {
  if (context.signal?.aborted) return err(new UserCancelError("SharePointPublish"));
  if (context.nonInteractive || !context.ui) return err(packageError("ConfirmationRequired"));
  const confirm = getLocalizedString("core.option.confirm");
  const answer = await context.ui.showMessage("warn", message, true, confirm);
  if (answer.isErr()) return err(answer.error);
  return answer.value === confirm ? ok(undefined) : err(new UserCancelError("SharePointPublish"));
}

function appId(app: Record<string, unknown>): string | undefined {
  const id = app.ID ?? app.Id;
  return typeof id === "string" && guidPattern.test(id) ? id.toLowerCase() : undefined;
}

function validApp(
  app: Record<string, unknown>,
  pkg: SharePointPackage,
  deployed: boolean
): boolean {
  return (
    app.AppCatalogVersion === pkg.version &&
    app.IsClientSideSolution === true &&
    (app.ErrorMessage === "" || app.ErrorMessage === "No errors.") &&
    (!deployed || (app.Deployed === true && app.CurrentVersionDeployed === true))
  );
}

export class SharePointPackageService {
  constructor(
    private readonly context: DriverContext,
    private readonly pause: (milliseconds: number) => Promise<void> = (milliseconds) =>
      setTimeout(milliseconds, undefined, { signal: context.signal })
  ) {}

  public async prepare(
    pkg: SharePointPackage,
    identity: PublishIdentity
  ): Promise<Result<PreparedSharePointDeployment, FxError>> {
    const { m365TokenProvider, signal } = this.context;
    if (signal?.aborted) return err(new UserCancelError("SharePointPublish"));
    try {
      const graph = getResourceServiceEndpoint(ResourceServiceType.Graph);
      if (graph !== "https://graph.microsoft.com") return err(packageError("UnsupportedCloud"));
      const token = await m365TokenProvider.getAccessToken({
        scopes: GraphTeamsAppCatalogReadWriteScopes,
      });
      if (token.isErr()) return err(token.error);
      const claims = await m365TokenProvider.getJsonObject({
        scopes: GraphTeamsAppCatalogReadWriteScopes,
      });
      if (claims.isErr()) return err(claims.error);
      if (
        !this.sameIdentity(claims.value, identity) ||
        !record(claims.value) ||
        (claims.value.aud !== "00000003-0000-0000-c000-000000000000" && claims.value.aud !== graph)
      ) {
        return err(packageError("TenantMismatch"));
      }
      // A real read establishes site-read readiness without silently requesting Sites.Read.All.
      const root = await axios.get<unknown>(`${graph}/v1.0/sites/root?$select=webUrl`, {
        headers: { Authorization: `Bearer ${token.value}`, Accept: "application/json" },
        signal,
        timeout: 30000,
        maxRedirects: 0,
      });
      if (!record(root.data) || typeof root.data.webUrl !== "string")
        return err(packageError("DiscoveryFailed"));
      const origin = new URL(root.data.webUrl);
      if (
        origin.protocol !== "https:" ||
        !/^[a-z0-9-]+\.sharepoint\.com$/i.test(origin.hostname) ||
        origin.username ||
        origin.password ||
        origin.port ||
        origin.search ||
        origin.hash
      ) {
        return err(packageError("UnsupportedCatalog"));
      }
      const spo = await m365TokenProvider.getAccessToken({ scopes: SPFxScopes(origin.origin) });
      if (spo.isErr()) return err(spo.error);
      const spoClaims = await m365TokenProvider.getJsonObject({
        scopes: SPFxScopes(origin.origin),
      });
      if (spoClaims.isErr()) return err(spoClaims.error);
      if (
        !this.sameIdentity(spoClaims.value, identity) ||
        !record(spoClaims.value) ||
        (spoClaims.value.aud !== "00000003-0000-0ff1-ce00-000000000000" &&
          spoClaims.value.aud !== origin.origin)
      ) {
        return err(packageError("TenantMismatch"));
      }
      const discovery = new SPOClient(origin.origin, signal);
      const catalog = await discovery.getAppCatalogSite(spo.value);
      if (!catalog) return err(packageError("MissingCatalog"));
      const url = new URL(catalog);
      if (url.origin !== origin.origin || url.username || url.password || url.search || url.hash) {
        return err(packageError("UnsupportedCatalog"));
      }
      const client = new SPOClient(url.href.replace(/\/$/, ""), signal);
      const existing = await this.read(() => client.getApp(spo.value, pkg.productId));
      const result: SharePointPackageResult = {
        stage: "notStarted",
        productId: pkg.productId,
        version: pkg.version,
        sha256: pkg.sha256,
        catalogUrl: url.href.replace(/\/$/, ""),
      };
      const prepared: PreparedSharePointDeployment = {
        package: pkg,
        client,
        token: spo.value,
        result,
        identical: false,
        overwrite: false,
      };
      if (existing) {
        const id = appId(existing);
        const version = existing.AppCatalogVersion;
        if (!id || typeof version !== "string" || !versionPattern.test(version)) {
          return err(packageError("CatalogIdentity"));
        }
        const previous = version.split(".").map(Number);
        const next = pkg.version.split(".").map(Number);
        const difference =
          next.map((part, index) => part - previous[index]).find((part) => part !== 0) ?? 0;
        if (difference < 0) return err(packageError("Downgrade"));
        result.catalogItemId = id;
        const name = await this.read(() => client.getPackageName(spo.value, id));
        if (name !== pkg.fileName) return err(packageError("FilenameCollision"));
        const remoteBytes = await this.read(() => client.downloadPackage(spo.value, id));
        prepared.previousHash = sha256(remoteBytes);
        prepared.previousVersion = version;
        prepared.identical = prepared.previousHash === pkg.sha256;
        if (prepared.identical) {
          if (!validApp(existing, pkg, false)) return err(packageError("InvalidRemotePackage"));
          result.stage = validApp(existing, pkg, true) ? "verified" : "uploaded";
          result.lastSuccessfulStage = result.stage;
        } else {
          // Same-version replacement is not distinguishable from a stale server version.
          if (difference === 0) return err(packageError("VersionCollision"));
          const confirmation = await confirmPublish(
            this.context,
            getLocalizedString(
              "publish.sharepoint.ConfirmOverwrite",
              pkg.fileName,
              pkg.productId,
              version,
              pkg.version
            )
          );
          if (confirmation.isErr()) return err(confirmation.error);
          prepared.overwrite = true;
        }
      }
      return ok(prepared);
    } catch (error) {
      return err(remotePublishError(error));
    }
  }

  public async deploy(
    prepared: PreparedSharePointDeployment,
    save: (state: SharePointPackageResult) => Promise<void>
  ): Promise<Result<SharePointPackageResult, FxError>> {
    const { package: pkg, client, token, result } = prepared;
    try {
      if (this.context.signal?.aborted) return err(new UserCancelError("SharePointPublish"));
      if (!prepared.identical) {
        // Recheck the identity immediately before an approved overwrite.
        if (prepared.overwrite) {
          const current = await this.read(() => client.getApp(token, pkg.productId));
          if (
            !current ||
            appId(current) !== result.catalogItemId ||
            current.AppCatalogVersion !== prepared.previousVersion ||
            (await client.getPackageName(token, result.catalogItemId!)) !== pkg.fileName ||
            sha256(await client.downloadPackage(token, result.catalogItemId!)) !==
              prepared.previousHash
          ) {
            return err(packageError("CatalogIdentity"));
          }
        }
        const previousStage = result.stage;
        result.stage = "unknownOutcome";
        await save({ ...result });
        if (this.context.signal?.aborted) {
          result.stage = previousStage;
          await save({ ...result });
          return err(new UserCancelError("SharePointPublish"));
        }
        const upload = await client.uploadAppPackage(
          token,
          pkg.fileName,
          pkg.bytes,
          prepared.overwrite
        );
        if (typeof upload.UniqueId !== "string" || !guidPattern.test(upload.UniqueId)) {
          return err(packageError("UnknownUpload"));
        }
        result.catalogItemId = upload.UniqueId.toLowerCase();
        result.stage = "uploaded";
        result.lastSuccessfulStage = "uploaded";
        await save({ ...result });
        this.context.logProvider.info(getLocalizedString("publish.sharepoint.Uploaded"));
      }
      if (!result.catalogItemId) return err(packageError("CatalogIdentity"));
      const visible = await this.poll(prepared, false);
      if (visible.isErr()) return err(visible.error);
      const remoteBytes = await this.read(() =>
        client.downloadPackage(token, result.catalogItemId!)
      );
      if (sha256(remoteBytes) !== pkg.sha256) return err(packageError("VersionCollision"));
      if (result.stage !== "verified") {
        if (this.context.signal?.aborted) return err(new UserCancelError("SharePointPublish"));
        const previousStage = result.stage;
        result.stage = "unknownOutcome";
        await save({ ...result });
        if (this.context.signal?.aborted) {
          result.stage = previousStage;
          await save({ ...result });
          return err(new UserCancelError("SharePointPublish"));
        }
        await client.deployAppPackage(token, result.catalogItemId, pkg.skipFeatureDeployment);
        result.stage = "deployed";
        result.lastSuccessfulStage = "deployed";
        await save({ ...result });
        this.context.logProvider.info(getLocalizedString("publish.sharepoint.Deployed"));
      }
      const verified = await this.poll(prepared, true);
      if (verified.isErr()) return err(verified.error);
      const product = await this.read(() => client.getApp(token, pkg.productId));
      if (!product || appId(product) !== result.catalogItemId || !validApp(product, pkg, true)) {
        return err(packageError("CatalogIdentity"));
      }
      // Content and deployed-version evidence must both match, including during resume.
      if (
        sha256(await this.read(() => client.downloadPackage(token, result.catalogItemId!))) !==
        pkg.sha256
      ) {
        return err(packageError("VersionCollision"));
      }
      result.stage = "verified";
      result.lastSuccessfulStage = "verified";
      await save({ ...result });
      this.context.logProvider.info(getLocalizedString("publish.sharepoint.Verified"));
      return ok({ ...result });
    } catch (error) {
      return err(remotePublishError(error));
    }
  }

  private sameIdentity(claims: unknown, identity: PublishIdentity): boolean {
    return record(claims) && claims.tid === identity.tenantId && claims.oid === identity.objectId;
  }

  private async read<T>(action: () => Promise<T>): Promise<T> {
    for (let attempt = 0; ; attempt++) {
      try {
        return await action();
      } catch (error) {
        if (
          !axios.isAxiosError(error) ||
          ![429, 503].includes(error.response?.status ?? 0) ||
          attempt >= 3
        )
          throw error;
        const seconds = Number(error.response?.headers["retry-after"]);
        if (!Number.isFinite(seconds) || seconds > 10 || seconds < 0) throw error;
        await this.pause(Math.max(1, seconds) * 1000);
      }
    }
  }

  private async poll(
    prepared: PreparedSharePointDeployment,
    deployed: boolean
  ): Promise<Result<undefined, FxError>> {
    const { client, token, result, package: pkg } = prepared;
    for (let attempt = 0; attempt < 6; attempt++) {
      if (this.context.signal?.aborted) return err(new UserCancelError("SharePointPublish"));
      const app = await this.read(() => client.getApp(token, result.catalogItemId!));
      if (app) {
        if (appId(app) !== result.catalogItemId) return err(packageError("CatalogIdentity"));
        if (
          typeof app.ErrorMessage === "string" &&
          app.ErrorMessage !== "" &&
          app.ErrorMessage !== "No errors."
        ) {
          return err(packageError("InvalidRemotePackage"));
        }
        if (validApp(app, pkg, deployed)) return ok(undefined);
      }
      if (attempt < 5) await this.pause(2000);
    }
    return err(packageError("VerificationTimeout"));
  }
}
