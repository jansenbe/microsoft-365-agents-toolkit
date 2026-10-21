// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform, err, ok } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { setTools } from "../../../../src/common/globalVars";
import { WrappedAxiosClient } from "../../../../src/common/wrappedAxiosClient";
import { teamsappMgr } from "../../../../src/component/driver/teamsApp/teamsappMgr";
import { PublishAppPackageDriver } from "../../../../src/component/driver/teamsApp/publishAppPackage";
import { ValidateAppPackageDriver } from "../../../../src/component/driver/teamsApp/validateAppPackage";
import { SharePointPackageService } from "../../../../src/component/driver/deploy/spfx/sharePointPackage";
import { SPOClient } from "../../../../src/component/driver/deploy/spfx/utility/spoClient";
import * as packageFile from "../../../../src/component/driver/deploy/spfx/utility/packageFile";
import { packageError } from "../../../../src/component/driver/deploy/spfx/utility/packageError";
import { UserCancelError } from "../../../../src/error/common";
import { MockTools } from "../../../core/utils";
import { FxCoreClient } from "../../../../src/core/FxCoreClient";
import { CombinedPublishError } from "../../../../src/component/driver/teamsApp/sharePointPublish";

const productId = "11111111-1111-4111-8111-1111111111aa";
const tenantId = "22222222-2222-4222-8222-2222222222bb";
let folder: string;
let http: MockAdapter;
const tools = new MockTools();
const pkg = {
  bytes: Buffer.from("sppkg"),
  fileName: "example.sppkg",
  productId,
  version: "1.0.0.0",
  skipFeatureDeployment: false,
  sha256: packageFile.sha256(Buffer.from("sppkg")),
};
beforeEach(async () => {
  folder = await fs.mkdtemp(path.join(os.tmpdir(), "atk-combined-"));
  await fs.copy(path.join(__dirname, "success.zip"), path.join(folder, "agent.zip"));
  setTools(tools);
  vi.spyOn(packageFile, "readSharePointPackage").mockResolvedValue(ok(pkg));
  vi.spyOn(ValidateAppPackageDriver.prototype, "execute").mockResolvedValue({
    result: ok(new Map()),
    summaries: [],
  });
  vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken").mockResolvedValue(
    ok("graph-token")
  );
  vi.spyOn(tools.tokenProvider.m365TokenProvider, "getJsonObject").mockResolvedValue(
    ok({ tid: tenantId, oid: productId, aud: "00000003-0000-0000-c000-000000000000" })
  );
  vi.spyOn(WrappedAxiosClient, "create").mockImplementation((config) => axios.create(config));
  http = new MockAdapter(axios);
  http.onGet(/appCatalogs\/teamsApps/).reply(200, { value: [] });
});
afterEach(async () => {
  http.restore();
  vi.restoreAllMocks();
  await fs.remove(folder);
});
function inputs(signal?: AbortSignal) {
  return {
    projectPath: folder,
    platform: Platform.CLI,
    "package-file": path.join(folder, "agent.zip"),
    "sharepoint-package-file": "example.sppkg",
    abortSignal: signal,
  };
}
function sharePoint(success: boolean, cancel?: AbortController) {
  vi.spyOn(SharePointPackageService.prototype, "prepare").mockResolvedValue(
    ok({
      package: pkg,
      client: new SPOClient("https://tenant.sharepoint.com"),
      token: "",
      identical: false,
      overwrite: false,
      result: {
        ...pkg,
        stage: "notStarted",
        catalogUrl: "https://tenant.sharepoint.com/catalog",
      },
    })
  );
  return vi
    .spyOn(SharePointPackageService.prototype, "deploy")
    .mockImplementation(async (prepared, save) => {
      if (!success) return err(new UserCancelError());
      const state = {
        ...prepared.result,
        stage: "verified" as const,
        catalogItemId: productId,
      };
      await save(state);
      cancel?.abort();
      return ok(state);
    });
}

describe("combined publication workflow", () => {
  it.each(["invalid-digest", "0".repeat(64)])(
    "SDK-04: invalid or stale selected digest %s blocks all remote preparation",
    async (expectedSharePointSha256) => {
      const result = await new FxCoreClient(tools).publish({
        ...inputs(),
        expectedSharePointSha256,
      });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.name).toBe("SharePointExpectedPackageMismatch");
      expect(ValidateAppPackageDriver.prototype.execute).not.toHaveBeenCalled();
      expect(tools.tokenProvider.m365TokenProvider.getAccessToken).not.toHaveBeenCalled();
      expect(http.history.post).toHaveLength(0);
      expect(await fs.pathExists(path.join(folder, ".local"))).toBe(false);
    }
  );
  it("SDK-02: expected-tenant mismatch blocks SharePoint preflight, journal and writes", async () => {
    const deploy = sharePoint(true);
    const result = await new FxCoreClient(tools).publish({
      ...inputs(),
      expectedTenantId: "11111111-1111-4111-8111-111111111111",
    });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("SharePointExpectedTenantMismatch");
    expect(SharePointPackageService.prototype.prepare).not.toHaveBeenCalled();
    expect(deploy).not.toHaveBeenCalled();
    expect(http.history.post).toHaveLength(0);
    expect(await fs.pathExists(path.join(folder, ".local"))).toBe(false);
  });

  it("SDK-02: matching expected tenant is accepted and malformed tenant is rejected", async () => {
    sharePoint(true);
    http.onPost(/requiresReview=true/).reply(201, { id: "catalog-agent-id" });
    const invalid = await new FxCoreClient(tools).publish({
      ...inputs(),
      expectedTenantId: "not-a-tenant-guid",
    });
    expect(invalid.isErr()).toBe(true);
    expect(http.history.post).toHaveLength(0);
    const valid = await new FxCoreClient(tools).publish({
      ...inputs(),
      expectedTenantId: tenantId.toUpperCase(),
      expectedSharePointSha256: pkg.sha256.toUpperCase(),
    });
    expect(valid.isOk()).toBe(true);
    expect(http.history.post).toHaveLength(1);
  });
  it("PUB-07: post-SharePoint catalog recheck failure is not mistaken for an attempted submission", async () => {
    sharePoint(true);
    let reads = 0;
    http
      .onGet(/appCatalogs\/teamsApps/)
      .reply(() =>
        ++reads === 1
          ? [200, { value: [] }]
          : [403, { error: { code: "Forbidden", message: "Catalog read denied." } }]
      );
    const result = await teamsappMgr.publishTeamsApp(inputs());
    if (result.isOk() || !(result.error instanceof CombinedPublishError))
      return expect.fail("Expected partial error");
    expect(result.error.publishResult).toMatchObject({
      sharePoint: { stage: "verified" },
      agent: {
        stage: "notStarted",
        diagnostic: { target: "agentCatalogRead", status: 403 },
      },
    });
    const journal = await fs.readJson(result.error.publishResult.recoveryFile);
    expect(journal.agent.diagnostic.target).toBe("agentCatalogRead");
    expect(http.history.post).toHaveLength(0);
  });
  it("PUB-07: agent 403 preserves verified SharePoint and sanitized Graph diagnostics in the journal", async () => {
    sharePoint(true);
    const requestId = "44444444-4444-4444-8444-4444444444dd";
    http.onPost(/requiresReview=true/).reply(
      403,
      {
        error: { code: "Forbidden", message: "The service denied organization submission." },
      },
      { "request-id": requestId }
    );
    const result = await teamsappMgr.publishTeamsApp(inputs());
    if (result.isOk()) return expect.fail("Expected rejection");
    expect(result.error).toBeInstanceOf(CombinedPublishError);
    if (!(result.error instanceof CombinedPublishError)) return;
    expect(result.error.publishResult).toMatchObject({
      sharePoint: { stage: "verified" },
      agent: {
        stage: "failed",
        diagnostic: {
          target: "agentSubmit",
          operation: "create",
          status: 403,
          serviceCode: "Forbidden",
          requestId,
        },
      },
    });
    expect(result.error.message).toContain("organization catalog");
    expect(result.error.message).not.toContain("SharePoint consent");
    const journal = await fs.readJson(result.error.publishResult.recoveryFile);
    expect(journal.agent.diagnostic.requestId).toBe(requestId);
    expect(journal.sharePoint.stage).toBe("verified");
    expect(http.history.post).toHaveLength(1);
  });

  it("PUB-07: timed-out agent submission journals unknown outcome and cannot be blindly retried", async () => {
    sharePoint(true);
    http.onPost(/requiresReview=true/).timeout();
    const result = await teamsappMgr.publishTeamsApp(inputs());
    if (result.isOk() || !(result.error instanceof CombinedPublishError))
      return expect.fail("Expected partial error");
    expect(result.error.publishResult.agent.stage).toBe("unknownOutcome");
    expect(result.error.message).toContain("outcome is unknown");
    const journal = await fs.readJson(result.error.publishResult.recoveryFile);
    expect(journal.agent.diagnostic.transportCode).toBe("ECONNABORTED");
    expect(journal.agent.diagnostic.status).toBeUndefined();
    expect(journal.sharePoint.stage).toBe("verified");
    expect((await teamsappMgr.publishTeamsApp(inputs())).isErr()).toBe(true);
    expect(http.history.post).toHaveLength(1);
  });

  it("PUB-07: Graph preflight rejection is identified as catalog read, before any SharePoint mutation", async () => {
    const deploy = sharePoint(true);
    http
      .onGet(/appCatalogs\/teamsApps/)
      .reply(403, { error: { code: "Forbidden", message: "Catalog read denied." } });
    const result = await teamsappMgr.publishTeamsApp(inputs());
    if (result.isOk()) return expect.fail("Expected rejection");
    expect(result.error).toMatchObject({ diagnostic: { target: "agentCatalogRead", status: 403 } });
    expect(result.error.message).not.toContain("SharePoint consent");
    expect(deploy).not.toHaveBeenCalled();
    expect(http.history.post).toHaveLength(0);
  });
  it("PUB-02: SDK runs the real package/service/publisher pipeline through mocked HTTP only", async () => {
    vi.mocked(packageFile.readSharePointPackage).mockRestore();
    const archive = new AdmZip();
    archive.addFile(
      "AppManifest.xml",
      Buffer.from(
        `<App ProductID="${productId}" Version="1.0.0.0" IsClientSideSolution="true" SkipFeatureDeployment="true"/>`
      )
    );
    const sppkg = archive.toBuffer();
    await fs.writeFile(path.join(folder, "example.sppkg"), sppkg);
    vi.mocked(tools.tokenProvider.m365TokenProvider.getJsonObject).mockImplementation(
      async (request) =>
        ok({
          tid: tenantId,
          oid: productId,
          aud: request.scopes[0].includes("sharepoint")
            ? "00000003-0000-0ff1-ce00-000000000000"
            : "00000003-0000-0000-c000-000000000000",
        })
    );
    const catalogItemId = "33333333-3333-4333-8333-3333333333cc";
    const catalog = "https://tenant.sharepoint.com/sites/catalog";
    let uploaded = false;
    let deployed = false;
    let downloads = 0;
    const operations: string[] = [];
    const remoteApp = () => ({
      ID: catalogItemId,
      AppCatalogVersion: "1.0.0.0",
      IsClientSideSolution: true,
      Deployed: deployed,
      CurrentVersionDeployed: deployed,
      ErrorMessage: "No errors.",
    });
    http.onGet(/\/sites\/root/).reply(200, { webUrl: "https://tenant.sharepoint.com" });
    http.onGet(/SP_TenantSettings_Current/).reply(200, { CorporateCatalogUrl: catalog });
    http.onGet(/AvailableApps/).reply(() => (uploaded ? [200, remoteApp()] : [404]));
    http.onGet(/GetFileById.*\/\$value/).reply(() => {
      downloads++;
      return [200, sppkg];
    });
    http.onPost(/\/Add\(/).reply(() => {
      uploaded = true;
      operations.push("sharepoint-upload");
      return [200, { UniqueId: catalogItemId }];
    });
    http.onPost(/\/Deploy$/).reply(() => {
      deployed = true;
      operations.push("sharepoint-deploy");
      return [200];
    });
    http.onPost(/requiresReview=true/).reply(() => {
      expect(downloads).toBe(2);
      expect(deployed).toBe(true);
      operations.push("agent-submit");
      return [201, { id: "catalog-agent-id" }];
    });
    const result = await new FxCoreClient(tools).publish(inputs());
    expect(result.isOk(), result.isErr() ? result.error.message : "").toBe(true);
    expect(operations).toEqual(["sharepoint-upload", "sharepoint-deploy", "agent-submit"]);
    if (result.isOk()) {
      expect(result.value.sharePoint).toMatchObject({
        stage: "verified",
        productId,
        catalogItemId,
      });
      const journal = await fs.readFile(result.value.recoveryFile!, "utf8");
      expect(journal).not.toContain("graph-token");
      expect(JSON.parse(journal).agent.stage).toBe("submitted");
    }
  });

  it("PUB-04: a recorded successful submission is rechecked without resubmitting", async () => {
    sharePoint(true);
    http.onPost(/requiresReview=true/).reply(201, { id: "catalog-agent-id" });
    const first = await teamsappMgr.publishTeamsApp(inputs());
    if (first.isErr() || !first.value) return expect.fail("First publication should succeed");
    http.onGet(/appCatalogs\/teamsApps/).reply(200, {
      value: [
        {
          id: "catalog-agent-id",
          appDefinitions: [
            {
              id: "definition-id",
              version: first.value.agent.version,
              publishingState: "submitted",
            },
          ],
        },
      ],
    });
    const second = await teamsappMgr.publishTeamsApp(inputs());
    expect(second.isOk()).toBe(true);
    expect(http.history.post).toHaveLength(1);
    if (second.isOk()) expect(second.value?.agent.stage).toBe("submitted");
  });

  it("PUB-04: a confirmed new agent version can advance a known submission", async () => {
    sharePoint(true);
    http.onPost(/requiresReview=true/).reply(201, { id: "catalog-agent-id" });
    const first = await teamsappMgr.publishTeamsApp(inputs());
    if (first.isErr() || !first.value) return expect.fail("First publication should succeed");
    http.onGet(/appCatalogs\/teamsApps/).reply(200, {
      value: [
        {
          id: "catalog-agent-id",
          appDefinitions: [
            {
              id: "definition-id",
              version: first.value.agent.version,
              publishingState: "submitted",
            },
          ],
        },
      ],
    });
    const agentPath = path.join(folder, "agent.zip");
    const archive = new AdmZip(await fs.readFile(agentPath));
    const manifest = JSON.parse(archive.readAsText("manifest.json"));
    manifest.version = "9.0.0";
    archive.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
    await fs.writeFile(agentPath, archive.toBuffer());
    vi.spyOn(tools.ui, "showMessage").mockResolvedValue(ok("Confirm"));
    const second = await teamsappMgr.publishTeamsApp(inputs());
    expect(second.isOk()).toBe(true);
    expect(http.history.post).toHaveLength(2);
    expect(tools.ui.showMessage).toHaveBeenCalledWith(
      "warn",
      expect.stringContaining("agent.zip"),
      true,
      "Confirm"
    );
    expect(tools.ui.showMessage).toHaveBeenCalledWith(
      "warn",
      expect.stringContaining(manifest.id),
      true,
      "Confirm"
    );
    expect(tools.ui.showMessage).toHaveBeenCalledWith(
      "warn",
      expect.stringContaining("9.0.0"),
      true,
      "Confirm"
    );
    expect(http.history.post[1].url).toContain(
      "/catalog-agent-id/appDefinitions?requiresReview=true"
    );
  });
  it("PUB-02: invalid SharePoint input fails before agent validation or authentication", async () => {
    vi.mocked(packageFile.readSharePointPackage).mockResolvedValue(
      err(packageError("InvalidPackage"))
    );
    expect((await teamsappMgr.publishTeamsApp(inputs())).isErr()).toBe(true);
    expect(ValidateAppPackageDriver.prototype.execute).not.toHaveBeenCalled();
    expect(tools.tokenProvider.m365TokenProvider.getAccessToken).not.toHaveBeenCalled();
    expect(http.history.post).toHaveLength(0);
  });

  it("PUB-03: agent validation failure never deploys SharePoint", async () => {
    const deploy = sharePoint(true);
    vi.mocked(ValidateAppPackageDriver.prototype.execute).mockResolvedValue({
      result: err(new UserCancelError()),
      summaries: [],
    });
    expect((await teamsappMgr.publishTeamsApp(inputs())).isErr()).toBe(true);
    expect(deploy).not.toHaveBeenCalled();
    expect(http.history.post).toHaveLength(0);
  });

  it("PUB-04: a definitive agent failure can retry after reconciling SharePoint", async () => {
    const deploy = sharePoint(true);
    http.onPost(/requiresReview=true/).replyOnce(400);
    http.onPost(/requiresReview=true/).reply(201, { id: "catalog-agent-id" });
    const first = await teamsappMgr.publishTeamsApp(inputs());
    expect(first.isErr()).toBe(true);
    if (first.isErr())
      expect(first.error).toMatchObject({
        publishResult: {
          sharePoint: { stage: "verified" },
          agent: { stage: "failed" },
        },
      });
    const second = await teamsappMgr.publishTeamsApp(inputs());
    expect(second.isOk()).toBe(true);
    expect(deploy).toHaveBeenCalledTimes(2);
    expect(http.history.post).toHaveLength(2);
  });

  it("PUB-01: no option does not invoke SharePoint or change the result", async () => {
    const prepare = vi.spyOn(SharePointPackageService.prototype, "prepare");
    vi.spyOn(PublishAppPackageDriver.prototype, "execute").mockResolvedValue({
      result: ok(new Map()),
      summaries: [],
    });
    const noFlag = {
      projectPath: folder,
      platform: Platform.CLI,
      "package-file": path.join(folder, "agent.zip"),
    };
    expect(await teamsappMgr.publishTeamsApp(noFlag)).toEqual(ok(undefined));
    expect(prepare).not.toHaveBeenCalled();
    expect(packageFile.readSharePointPackage).not.toHaveBeenCalled();
  });

  it("PUB-02: verifies SharePoint before submitting for review exactly once", async () => {
    const deploy = sharePoint(true);
    http.onPost(/requiresReview=true/).reply(() => {
      expect(deploy).toHaveBeenCalledOnce();
      return [201, { id: "catalog-agent-id" }];
    });
    const result = await teamsappMgr.publishTeamsApp(inputs());
    expect(result.isOk(), result.isErr() ? result.error.message : "").toBe(true);
    if (result.isOk())
      expect(result.value).toMatchObject({
        sharePoint: { stage: "verified" },
        agent: { stage: "submitted" },
      });
    expect(http.history.post).toHaveLength(1);
  });

  it("PUB-03: SharePoint failure never submits the agent", async () => {
    sharePoint(false);
    const result = await teamsappMgr.publishTeamsApp(inputs());
    expect(result.isErr()).toBe(true);
    expect(http.history.post).toHaveLength(0);
  });

  it("PUB-03: cancellation between stages preserves verified SharePoint state", async () => {
    const controller = new AbortController();
    sharePoint(true, controller);
    const result = await teamsappMgr.publishTeamsApp(inputs(controller.signal));
    expect(result.isErr()).toBe(true);
    if (result.isErr())
      expect(result.error).toMatchObject({
        publishResult: { sharePoint: { stage: "verified" }, agent: { stage: "cancelled" } },
      });
    expect(http.history.post).toHaveLength(0);
  });

  it("PUB-04 SPO-06: unknown agent response persists partial state and blocks blind retry", async () => {
    sharePoint(true);
    http.onPost(/requiresReview=true/).networkError();
    const first = await teamsappMgr.publishTeamsApp(inputs());
    expect(first.isErr()).toBe(true);
    if (first.isErr())
      expect(first.error).toMatchObject({
        publishResult: {
          sharePoint: { stage: "verified" },
          agent: { stage: "unknownOutcome" },
        },
      });
    const second = await teamsappMgr.publishTeamsApp(inputs());
    expect(second.isErr()).toBe(true);
    expect(http.history.post).toHaveLength(1);
  });
});
