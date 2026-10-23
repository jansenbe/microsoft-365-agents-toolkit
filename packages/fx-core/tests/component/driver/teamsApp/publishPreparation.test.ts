// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Platform,
  PreparedSharePointPublishArtifacts,
  SharePointPublishPreparationContext,
  SharePointPublishPreparer,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClientProvider";
import { setTools } from "../../../../src/common/globalVars";
import { WrappedAxiosClient } from "../../../../src/common/wrappedAxiosClient";
import { sha256 } from "../../../../src/component/driver/deploy/spfx/utility/packageFile";
import { CombinedPublishError } from "../../../../src/component/driver/teamsApp/sharePointPublish";
import { ValidateAppPackageDriver } from "../../../../src/component/driver/teamsApp/validateAppPackage";
import { FxCoreClient } from "../../../../src/core/FxCoreClient";
import { MockTools } from "../../../core/utils";

const productId = "11111111-1111-4111-8111-1111111111aa";
const tenantId = "22222222-2222-4222-8222-2222222222bb";
const itemId = "33333333-3333-4333-8333-3333333333cc";
const catalogUrl = "https://tenant.sharepoint.com/sites/catalog";
let folder: string;
let stage: string;
let original: Buffer;
let derived: Buffer;
let application: Buffer;
let artifacts: PreparedSharePointPublishArtifacts;
let http: MockAdapter;
let tools: MockTools;
let selectedTenant: string;
let selectedCatalog: string;
let operations: string[];
let uploaded: boolean;
let deployed: boolean;

function packageBytes(version = "1.0.0.0", skip = true, embedded = false): Buffer {
  const archive = new AdmZip();
  archive.addFile(
    "AppManifest.xml",
    Buffer.from(
      `<App ProductID="${productId}" Version="${version}" IsClientSideSolution="true" SkipFeatureDeployment="${skip}"/>`
    )
  );
  archive.addFile("client/component.js", Buffer.from("component"));
  if (embedded) archive.addFile("client/embedded.zip", Buffer.from("host-owned embedded bytes"));
  return archive.toBuffer();
}

function inputs() {
  return {
    projectPath: folder,
    platform: Platform.CLI,
    env: "dev",
    "package-file": path.join(stage, "final.zip"),
    "sharepoint-package-file": "solution.sppkg",
    expectedSharePointSha256: sha256(original),
    expectedTenantId: tenantId,
  };
}

async function finalize(context: SharePointPublishPreparationContext) {
  operations.push("finalize");
  expect(context).toEqual({
    projectPath: folder,
    tenantId,
    catalogUrl,
    originalSharePoint: {
      packagePath: path.join(folder, "solution.sppkg"),
      productId,
      version: "1.0.0.0",
      sha256: sha256(original),
      skipFeatureDeployment: true,
    },
    signal: context.signal,
  });
  expect(await fs.pathExists(path.join(folder, ".local"))).toBe(false);
  expect(operations).not.toContain("validate-final");
  expect(http.history.post).toHaveLength(0);
  return ok(artifacts);
}

beforeEach(async () => {
  folder = await fs.mkdtemp(path.join(os.tmpdir(), "atk-preparation-"));
  stage = path.join(folder, "staged");
  await fs.ensureDir(stage);
  original = packageBytes("1.0.0.0", true, true);
  derived = packageBytes();
  const app = new AdmZip(await fs.readFile(path.join(__dirname, "success.zip")));
  app.addFile("catalog-context.txt", Buffer.from(catalogUrl));
  application = app.toBuffer();
  await fs.writeFile(path.join(folder, "solution.sppkg"), original);
  await fs.writeFile(path.join(stage, "solution.sppkg"), derived);
  await fs.writeFile(path.join(stage, "final.zip"), application);
  artifacts = {
    sharePointPackagePath: path.join(stage, "solution.sppkg"),
    appPackagePath: path.join(stage, "final.zip"),
    lineage: {
      schemaVersion: 1,
      originalSharePointSha256: sha256(original),
      transformationVersion: "test-v1",
      selectedEmbeddedPackage: "client/embedded.zip",
      extractedPluginSha256: sha256(Buffer.from("host-owned plugin")),
      derivedSharePointSha256: sha256(derived),
      finalAppSha256: sha256(application),
    },
  };
  selectedTenant = tenantId;
  selectedCatalog = catalogUrl;
  operations = [];
  uploaded = false;
  deployed = false;
  tools = new MockTools();
  setTools(tools);
  vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken").mockImplementation(
    async (request) => ok(request.scopes[0].includes("sharepoint") ? "spo-token" : "graph-token")
  );
  vi.spyOn(tools.tokenProvider.m365TokenProvider, "getJsonObject").mockImplementation(
    async (request) =>
      ok({
        tid: selectedTenant,
        oid: productId,
        aud: request.scopes[0].includes("sharepoint")
          ? "00000003-0000-0ff1-ce00-000000000000"
          : "00000003-0000-0000-c000-000000000000",
      })
  );
  vi.spyOn(WrappedAxiosClient, "create").mockImplementation((config) => axios.create(config));
  vi.spyOn(teamsDevPortalClient, "partnerCenterAppPackageValidation").mockImplementation(
    async (_token, bytes) => {
      operations.push("validate-final");
      expect(sha256(bytes)).toBe(sha256(application));
      return { errors: [], warnings: [], notes: [] };
    }
  );
  http = new MockAdapter(axios);
  http.onGet(/\/sites\/root/).reply(() => {
    operations.push("discover-root");
    return [200, { webUrl: "https://tenant.sharepoint.com" }];
  });
  http.onGet(/SP_TenantSettings_Current/).reply(() => {
    operations.push("discover-catalog");
    return [200, { CorporateCatalogUrl: selectedCatalog }];
  });
  http.onGet(/appCatalogs\/teamsApps/).reply(200, { value: [] });
  http.onGet(/AvailableApps/).reply(() =>
    uploaded
      ? [
          200,
          {
            ID: itemId,
            AppCatalogVersion: "1.0.0.0",
            IsClientSideSolution: true,
            Deployed: deployed,
            CurrentVersionDeployed: deployed,
            ErrorMessage: "",
          },
        ]
      : [404]
  );
  http.onGet(/GetFileById.*\/\$value/).reply(() => [200, derived]);
  http.onGet(/GetFileById/).reply(200, { Name: "solution.sppkg" });
  http.onPost(/\/Add\(/).reply((config) => {
    expect(sha256(config.data)).toBe(sha256(derived));
    operations.push("upload-derived");
    uploaded = true;
    return [200, { UniqueId: itemId }];
  });
  http.onPost(/\/Deploy$/).reply(() => {
    operations.push("deploy-derived");
    deployed = true;
    return [200];
  });
  http.onPost(/requiresReview=true/).reply((config) => {
    expect(deployed).toBe(true);
    expect(sha256(config.data)).toBe(sha256(application));
    operations.push("submit-final");
    return [201, { id: "agent-catalog-id" }];
  });
});

afterEach(async () => {
  http.restore();
  vi.restoreAllMocks();
  await fs.remove(folder);
});

describe("public publication preparation", () => {
  it("PREP-01: original digest mismatch prevents auth, callback and journal", async () => {
    const callback = vi.fn(finalize);
    const result = await new FxCoreClient(tools).publish(
      { ...inputs(), expectedSharePointSha256: "0".repeat(64) },
      { prepareSharePointPublish: callback }
    );
    expect(result.isErr()).toBe(true);
    expect(callback).not.toHaveBeenCalled();
    expect(tools.tokenProvider.m365TokenProvider.getAccessToken).not.toHaveBeenCalled();
    expect(http.history.get).toHaveLength(0);
    expect(await fs.pathExists(path.join(folder, ".local"))).toBe(false);
  });

  it("PREP-01: expected tenant mismatch stops before SharePoint discovery and callback", async () => {
    selectedTenant = productId;
    const callback = vi.fn(finalize);
    const result = await new FxCoreClient(tools).publish(inputs(), {
      prepareSharePointPublish: callback,
    });
    expect(result.isErr()).toBe(true);
    expect(callback).not.toHaveBeenCalled();
    expect(http.history.get).toHaveLength(0);
  });

  it("PREP-01 PUB-08: discovers before finalization and validates/submits only final bytes", async () => {
    const callback = vi.fn(finalize);
    const result = await new FxCoreClient(tools).publish(inputs(), {
      prepareSharePointPublish: callback,
    });
    expect(result.isOk(), result.isErr() ? result.error.message : "").toBe(true);
    expect(callback).toHaveBeenCalledOnce();
    expect(operations.indexOf("discover-catalog")).toBeLessThan(operations.indexOf("finalize"));
    expect(operations.indexOf("finalize")).toBeLessThan(operations.indexOf("validate-final"));
    expect(operations.indexOf("validate-final")).toBeLessThan(operations.indexOf("upload-derived"));
    expect(operations.filter((operation) => /^(upload|deploy|submit)-/.test(operation))).toEqual([
      "upload-derived",
      "deploy-derived",
      "submit-final",
    ]);
    if (result.isErr()) return;
    expect(result.value).toMatchObject({
      projectPath: folder,
      packagePath: artifacts.appPackagePath,
      lineage: artifacts.lineage,
      sharePoint: { sha256: sha256(derived), stage: "verified" },
      agent: { sha256: sha256(application), stage: "submitted" },
    });
    expect(path.dirname(result.value.recoveryFile!)).toBe(
      path.join(folder, ".local", "sharepoint-publish")
    );
    const journal = await fs.readJson(result.value.recoveryFile!);
    expect(journal).toMatchObject({
      schemaVersion: 2,
      projectPath: folder,
      lineage: artifacts.lineage,
    });
    expect(JSON.stringify(journal)).not.toMatch(/graph-token|spo-token/);
    expect(await fs.pathExists(path.join(stage, ".local"))).toBe(false);
  });

  it.each(["originalSharePointSha256", "derivedSharePointSha256", "finalAppSha256"] as const)(
    "PREP-02: rejects incorrect lineage digest %s before validation/journal/mutation",
    async (field) => {
      const result = await new FxCoreClient(tools).publish(inputs(), {
        prepareSharePointPublish: async () =>
          ok({ ...artifacts, lineage: { ...artifacts.lineage, [field]: "0".repeat(64) } }),
      });
      expect(result.isErr()).toBe(true);
      expect(operations).not.toContain("validate-final");
      expect(http.history.post).toHaveLength(0);
      expect(await fs.pathExists(path.join(folder, ".local"))).toBe(false);
    }
  );

  it.each([
    { selectedEmbeddedPackage: "../unsafe.zip" },
    { selectedEmbeddedPackage: "https://example.com/package.zip" },
    { transformationVersion: "" },
    { extractedPluginSha256: "invalid" },
  ])("PREP-02: rejects malformed host lineage %j", async (invalid) => {
    const result = await new FxCoreClient(tools).publish(inputs(), {
      prepareSharePointPublish: async () =>
        ok({ ...artifacts, lineage: { ...artifacts.lineage, ...invalid } }),
    });
    expect(result.isErr()).toBe(true);
    expect(http.history.post).toHaveLength(0);
    expect(await fs.pathExists(path.join(folder, ".local"))).toBe(false);
  });

  it.each(["original", "derived-version", "derived-policy", "final-app"])(
    "PREP-02: rejects changed %s during host finalization",
    async (change) => {
      const result = await new FxCoreClient(tools).publish(inputs(), {
        prepareSharePointPublish: async () => {
          if (change === "original") {
            await fs.appendFile(path.join(folder, "solution.sppkg"), "changed");
          } else if (change === "final-app") {
            await fs.appendFile(artifacts.appPackagePath, "changed");
          } else {
            const changed = packageBytes(
              change === "derived-version" ? "2.0.0.0" : "1.0.0.0",
              change !== "derived-policy"
            );
            await fs.writeFile(artifacts.sharePointPackagePath, changed);
            artifacts = {
              ...artifacts,
              lineage: { ...artifacts.lineage, derivedSharePointSha256: sha256(changed) },
            };
          }
          return ok(artifacts);
        },
      });
      expect(result.isErr()).toBe(true);
      expect(http.history.post).toHaveLength(0);
      expect(await fs.pathExists(path.join(folder, ".local"))).toBe(false);
    }
  );

  it.each(["identity", "catalog"])(
    "PREP-03: rejects changed %s after finalization",
    async (change) => {
      const result = await new FxCoreClient(tools).publish(inputs(), {
        prepareSharePointPublish: async () => {
          if (change === "identity") selectedTenant = productId;
          else selectedCatalog = "https://tenant.sharepoint.com/sites/other-catalog";
          return ok(artifacts);
        },
      });
      expect(result.isErr()).toBe(true);
      expect(http.history.post).toHaveLength(0);
    }
  );

  it.each(["original", "derived", "application", "identity", "catalog"])(
    "PREP-03: rejects changed %s during final update confirmation",
    async (change) => {
      http.onGet(/appCatalogs\/teamsApps/).reply(200, {
        value: [
          {
            id: "existing-agent",
            appDefinitions: [{ id: "prior", version: "0.0.1", publishingState: "submitted" }],
          },
        ],
      });
      vi.spyOn(tools.ui, "showMessage").mockImplementation(async (_level, _message, modal) => {
        if (modal) {
          if (change === "identity") selectedTenant = productId;
          else if (change === "catalog") selectedCatalog += "-changed";
          else
            await fs.appendFile(
              change === "original"
                ? path.join(folder, "solution.sppkg")
                : change === "derived"
                  ? artifacts.sharePointPackagePath
                  : artifacts.appPackagePath,
              "changed"
            );
        }
        return ok("Confirm");
      });
      const result = await new FxCoreClient(tools).publish(inputs(), {
        prepareSharePointPublish: async () => ok(artifacts),
      });
      expect(result.isErr()).toBe(true);
      expect(tools.ui.showMessage).toHaveBeenCalledWith(
        "warn",
        expect.stringContaining("final.zip"),
        true,
        "Confirm"
      );
      expect(http.history.post).toHaveLength(0);
    }
  );

  it("PREP-03: final validation consumes the same snapshot despite a path change", async () => {
    const execute = ValidateAppPackageDriver.prototype.execute;
    vi.spyOn(ValidateAppPackageDriver.prototype, "execute").mockImplementation(
      async function (args, context) {
        await fs.appendFile(artifacts.appPackagePath, "changed-before-validation-read");
        return execute.call(this, args, context);
      }
    );
    const result = await new FxCoreClient(tools).publish(inputs(), {
      prepareSharePointPublish: async () => ok(artifacts),
    });
    expect(result.isErr()).toBe(true);
    expect(teamsDevPortalClient.partnerCenterAppPackageValidation).toHaveBeenCalledOnce();
    expect(http.history.post).toHaveLength(0);
  });

  it("PREP-04: pre-cancel does not acquire tokens or invoke preparation", async () => {
    const controller = new AbortController();
    controller.abort();
    const callback = vi.fn(finalize);
    const result = await new FxCoreClient(tools).publish(inputs(), {
      signal: controller.signal,
      prepareSharePointPublish: callback,
    });
    expect(result.isErr()).toBe(true);
    expect(callback).not.toHaveBeenCalled();
    expect(tools.tokenProvider.m365TokenProvider.getAccessToken).not.toHaveBeenCalled();
  });

  it("PREP-04: awaits a cancelled finalizer and does not validate or write afterward", async () => {
    const controller = new AbortController();
    let enter = (): void => {};
    let finish = (): void => {};
    const entered = new Promise<void>((resolve) => {
      enter = resolve;
    });
    const finished = new Promise<void>((resolve) => {
      finish = resolve;
    });
    let settled = false;
    const pending = new FxCoreClient(tools)
      .publish(inputs(), {
        signal: controller.signal,
        prepareSharePointPublish: async (context) => {
          expect(context.signal).toBe(controller.signal);
          enter();
          await finished;
          return ok(artifacts);
        },
      })
      .then((result) => {
        settled = true;
        return result;
      });
    await entered;
    controller.abort();
    await Promise.resolve();
    expect(settled).toBe(false);
    finish();
    const result = await pending;
    expect(result.isErr()).toBe(true);
    expect(operations).not.toContain("validate-final");
    expect(http.history.post).toHaveLength(0);
    expect(await fs.pathExists(path.join(folder, ".local"))).toBe(false);
  });

  it("PREP-02 PREP-04: rejected discovery or callback never falls back to original publication", async () => {
    const failure = new UserError("host", "PreparationRejected", "Synthetic preparation rejection");
    const callback: SharePointPublishPreparer = vi.fn(async () => err(failure));
    const client = new FxCoreClient(tools);
    http.onGet(/SP_TenantSettings_Current/).reply(403);
    expect((await client.publish(inputs(), { prepareSharePointPublish: callback })).isErr()).toBe(
      true
    );
    expect(callback).not.toHaveBeenCalled();
    http.onGet(/SP_TenantSettings_Current/).reply(200, { CorporateCatalogUrl: catalogUrl });
    const rejected = await client.publish(inputs(), { prepareSharePointPublish: callback });
    expect(rejected).toEqual(err(failure));
    expect(http.history.post).toHaveLength(0);
    expect(await fs.pathExists(path.join(folder, ".local"))).toBe(false);
  });

  it.each(["throw", "missing-file", "empty-path"])(
    "PREP-02: callback %s has no fallback",
    async (mode) => {
      const result = await new FxCoreClient(tools).publish(inputs(), {
        prepareSharePointPublish: async () => {
          if (mode === "throw") throw new Error("synthetic host failure");
          return ok({ ...artifacts, appPackagePath: mode === "empty-path" ? "" : "missing.zip" });
        },
      });
      expect(result.isErr()).toBe(true);
      expect(http.history.post).toHaveLength(0);
      expect(await fs.pathExists(path.join(folder, ".local"))).toBe(false);
    }
  );

  it("PREP-06: preparation without a SharePoint input is rejected instead of ignored", async () => {
    const result = await new FxCoreClient(tools).publish(
      { projectPath: folder, platform: Platform.CLI, "package-file": artifacts.appPackagePath },
      { prepareSharePointPublish: finalize }
    );
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("SharePointPreparationRequiresSharePoint");
    expect(http.history.post).toHaveLength(0);
  });

  it.each(["same-version", "downgrade"])(
    "PREP-06: finalization cannot bypass %s collision policy",
    async (collision) => {
      http.onGet(/AvailableApps/).reply(200, {
        ID: itemId,
        AppCatalogVersion: collision === "downgrade" ? "2.0.0.0" : "1.0.0.0",
        IsClientSideSolution: true,
        Deployed: true,
        CurrentVersionDeployed: true,
        ErrorMessage: "",
      });
      http.onGet(/GetFileById.*\/\$value/).reply(200, original);
      const result = await new FxCoreClient(tools).publish(inputs(), {
        prepareSharePointPublish: async () => ok(artifacts),
      });
      expect(result.isErr()).toBe(true);
      if (result.isErr())
        expect(result.error.message).toMatch(/newer package version|Package bytes do not match/);
      expect(http.history.post).toHaveLength(0);
    }
  );

  it.each([403, "timeout"] as const)(
    "PREP-05: preserves final lineage and verified SharePoint on Graph %s",
    async (failure) => {
      if (failure === "timeout") http.onPost(/requiresReview=true/).timeout();
      else http.onPost(/requiresReview=true/).reply(403);
      const result = await new FxCoreClient(tools).publish(inputs(), {
        prepareSharePointPublish: async () => ok(artifacts),
      });
      if (result.isOk() || !(result.error instanceof CombinedPublishError)) {
        return expect.fail("Expected combined partial error");
      }
      const partial = result.error.publishResult;
      expect(partial).toMatchObject({
        projectPath: folder,
        lineage: artifacts.lineage,
        sharePoint: { stage: "verified", sha256: sha256(derived) },
        agent: {
          stage: failure === "timeout" ? "unknownOutcome" : "failed",
          sha256: sha256(application),
        },
      });
      expect(await fs.readJson(partial.recoveryFile)).toMatchObject({
        schemaVersion: 2,
        lineage: artifacts.lineage,
      });
    }
  );

  it.each([
    [1, false],
    [2, false],
    [2, true],
  ] as const)(
    "PREP-06: schema %s unknown submission blocks replay with changed lineage %s",
    async (schemaVersion, changedLineage) => {
      const journal = path.join(
        folder,
        ".local",
        "sharepoint-publish",
        `${sha256(Buffer.from(`${tenantId}:${productId}`))}.json`
      );
      await fs.outputJson(journal, {
        schemaVersion,
        projectPath: folder,
        lineage:
          schemaVersion === 2
            ? {
                ...artifacts.lineage,
                transformationVersion: changedLineage ? "prior-v0" : "test-v1",
              }
            : undefined,
        tenantId,
        objectId: productId,
        sharePoint: { productId, sha256: artifacts.lineage.derivedSharePointSha256 },
        agent: { stage: "unknownOutcome", sha256: artifacts.lineage.finalAppSha256 },
      });
      const previous = await fs.readFile(journal);
      const result = await new FxCoreClient(tools).publish(inputs(), {
        prepareSharePointPublish: async () => ok(artifacts),
      });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.message).toContain("previous agent submission");
      if (result.isErr() && result.error instanceof CombinedPublishError) {
        expect(result.error.publishResult.agent.stage).toBe("unknownOutcome");
      }
      expect(http.history.post).toHaveLength(0);
      expect(await fs.readFile(journal)).toEqual(previous);
    }
  );
});
