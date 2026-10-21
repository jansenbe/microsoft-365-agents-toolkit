// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Result, TeamsManifestConverter, err, ok } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import { GraphClient, PublishCatalogSnapshot } from "../../../client/graphClient";
import { GraphTeamsAppCatalogReadWriteScopes } from "../../../common/constants";
import { UserCancelError } from "../../../error/common";
import { PublishIdentity } from "../deploy/spfx/sharePointPackage";
import { remotePublishError } from "../../../error/publish";
import { guidPattern, maxPackageBytes, record, sha256 } from "../deploy/spfx/utility/packageFile";
import { packageError } from "../deploy/spfx/utility/packageError";
import { DriverContext } from "../interface/commonArgs";
import { verifyLocalMCPPluginCerts } from "./utils/McpCertVerification";

export interface PreparedAgentPublication {
  bytes: Buffer;
  manifestId: string;
  version: string;
  sha256: string;
  identity: PublishIdentity;
  graph: GraphClient;
  token: string;
  snapshot: PublishCatalogSnapshot;
}

export async function prepareAgentPublication(
  file: string,
  context: DriverContext
): Promise<Result<PreparedAgentPublication, FxError>> {
  if (context.signal?.aborted) return err(new UserCancelError("SharePointPublish"));
  try {
    const bytes = await fs.readFile(file);
    if (bytes.length > maxPackageBytes) return err(packageError("AgentPreflight"));
    const entries = new AdmZip(bytes).getEntries();
    const manifests = entries.filter((entry) => entry.entryName === "manifest.json");
    if (manifests.length !== 1) return err(packageError("AgentPreflight"));
    const manifest = TeamsManifestConverter.jsonToManifest(manifests[0].getData().toString("utf8"));
    if (!guidPattern.test(manifest.id) || !manifest.version)
      return err(packageError("AgentPreflight"));
    // Reuse the existing publisher's certificate policy before either service is mutated.
    const agents =
      ("copilotAgents" in manifest ? manifest.copilotAgents?.declarativeAgents : undefined) ??
      ("copilotExtensions" in manifest
        ? manifest.copilotExtensions?.declarativeCopilots
        : undefined);
    for (const agent of agents ?? []) {
      const entry = entries.find((item) => item.entryName === agent.file);
      if (!entry) return err(packageError("AgentPreflight"));
      const content: unknown = JSON.parse(entry.getData().toString("utf8"));
      if (record(content) && Array.isArray(content.actions)) {
        for (const action of content.actions) {
          if (!record(action) || typeof action.file !== "string")
            return err(packageError("AgentPreflight"));
          const plugin = entries.find((item) => item.entryName === action.file);
          if (!plugin || !(await verifyLocalMCPPluginCerts(plugin)))
            return err(packageError("AgentPreflight"));
        }
      }
    }
    const token = await context.m365TokenProvider.getAccessToken({
      scopes: GraphTeamsAppCatalogReadWriteScopes,
    });
    if (token.isErr()) return err(token.error);
    const claims = await context.m365TokenProvider.getJsonObject({
      scopes: GraphTeamsAppCatalogReadWriteScopes,
    });
    if (claims.isErr()) return err(claims.error);
    if (
      !record(claims.value) ||
      typeof claims.value.tid !== "string" ||
      typeof claims.value.oid !== "string" ||
      !guidPattern.test(claims.value.tid) ||
      !guidPattern.test(claims.value.oid)
    )
      return err(packageError("TenantMismatch"));
    if (
      claims.value.aud !== "00000003-0000-0000-c000-000000000000" &&
      claims.value.aud !== "https://graph.microsoft.com"
    ) {
      return err(packageError("TenantMismatch"));
    }
    const identity = { tenantId: claims.value.tid, objectId: claims.value.oid };
    const graph = new GraphClient(context.m365TokenProvider, context.logProvider);
    const snapshot = await graph.getPublishSnapshot(token.value, manifest.id, context.signal);
    return ok({
      bytes,
      manifestId: manifest.id,
      version: manifest.version,
      sha256: sha256(bytes),
      identity,
      graph,
      token: token.value,
      snapshot,
    });
  } catch (error) {
    return err(remotePublishError(error, "agentCatalogRead"));
  }
}
