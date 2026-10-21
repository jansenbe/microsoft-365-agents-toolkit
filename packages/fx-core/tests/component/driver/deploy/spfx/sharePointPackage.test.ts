// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok, err, UserError, Platform } from "@microsoft/teamsfx-api";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { SharePointPackageService } from "../../../../../src/component/driver/deploy/spfx/sharePointPackage";
import { sha256 } from "../../../../../src/component/driver/deploy/spfx/utility/packageFile";
import { SPOClient } from "../../../../../src/component/driver/deploy/spfx/utility/spoClient";
import { createDriverContext } from "../../../../../src/component/driver/util/utils";
import { setTools } from "../../../../../src/common/globalVars";
import { MockTools } from "../../../../core/utils";

const productId = "11111111-1111-4111-8111-1111111111aa";
const itemId = "33333333-3333-4333-8333-3333333333cc";
const identity = { tenantId: "22222222-2222-4222-8222-2222222222bb", objectId: productId };
const catalog = "https://tenant.sharepoint.com/sites/catalog";
const bytes = Buffer.from("immutable fixture package bytes");
const pkg = {
  bytes,
  fileName: "test's #1.sppkg",
  productId,
  version: "1.0.0.16",
  skipFeatureDeployment: false,
  sha256: sha256(bytes),
};
const app = {
  ID: itemId,
  AppCatalogVersion: pkg.version,
  IsClientSideSolution: true,
  Deployed: true,
  CurrentVersionDeployed: true,
  ErrorMessage: "No errors.",
};
let http: MockAdapter;
const tools = new MockTools();

function context(signal?: AbortSignal) {
  setTools(tools);
  const ctx = createDriverContext({
    platform: Platform.CLI,
    projectPath: ".",
    abortSignal: signal,
  });
  vi.spyOn(ctx.m365TokenProvider, "getAccessToken").mockResolvedValue(ok("spo-token"));
  vi.spyOn(ctx.m365TokenProvider, "getJsonObject").mockImplementation(async (request) =>
    ok({
      tid: identity.tenantId,
      oid: identity.objectId,
      aud: request?.scopes[0].includes("sharepoint")
        ? "00000003-0000-0ff1-ce00-000000000000"
        : "00000003-0000-0000-c000-000000000000",
    })
  );
  return ctx;
}

beforeEach(() => {
  http = new MockAdapter(axios);
  http
    .onGet("https://graph.microsoft.com/v1.0/sites/root?$select=webUrl")
    .reply(200, { webUrl: "https://tenant.sharepoint.com" });
  http
    .onGet("https://tenant.sharepoint.com/_api/SP_TenantSettings_Current")
    .reply(200, { CorporateCatalogUrl: catalog });
});
afterEach(() => {
  http.restore();
  vi.restoreAllMocks();
});

describe("SharePoint package operation", () => {
  it("SPO-04 SPO-06: deployment failure preserves the completed upload and never verifies success", async () => {
    http.onGet(/AvailableApps/).replyOnce(404);
    http.onGet(/AvailableApps/).reply(200, app);
    http.onPost(/\/Add\(/).reply(200, { UniqueId: itemId });
    http.onGet(`${catalog}/_api/web/GetFileById('${itemId}')/$value`).reply(200, bytes);
    http.onPost(/\/Deploy$/).timeout();
    const service = new SharePointPackageService(context());
    const prepared = await service.prepare(pkg, identity);
    if (prepared.isErr()) return expect.fail(prepared.error.message);
    const result = await service.deploy(prepared.value, async () => {});
    expect(result.isErr()).toBe(true);
    expect(prepared.value.result).toMatchObject({
      stage: "unknownOutcome",
      lastSuccessfulStage: "uploaded",
      catalogItemId: itemId,
    });
    expect(http.history.post).toHaveLength(2);
  });

  it("SPO-04: cancellation after recording intent but before upload records no remote write", async () => {
    const controller = new AbortController();
    http.onGet(/AvailableApps/).reply(404);
    const service = new SharePointPackageService(context(controller.signal));
    const prepared = await service.prepare(pkg, identity);
    if (prepared.isErr()) return expect.fail(prepared.error.message);
    const result = await service.deploy(prepared.value, async () => {
      controller.abort();
    });
    expect(result.isErr()).toBe(true);
    expect(prepared.value.result.stage).toBe("notStarted");
    expect(http.history.post).toHaveLength(0);
  });
  it.each([401, 403])(
    "SPO-02: denied root discovery (%s) performs no SharePoint writes",
    async (status) => {
      http.onGet("https://graph.microsoft.com/v1.0/sites/root?$select=webUrl").reply(status);
      const prepared = await new SharePointPackageService(context()).prepare(pkg, identity);
      expect(prepared.isErr()).toBe(true);
      expect(http.history.post).toHaveLength(0);
    }
  );

  it("SPO-02: missing SharePoint consent is surfaced before catalog writes", async () => {
    const ctx = context();
    vi.mocked(ctx.m365TokenProvider.getAccessToken)
      .mockResolvedValueOnce(ok("graph-token"))
      .mockResolvedValueOnce(err(new UserError("test", "ConsentRequired", "Consent required")));
    const prepared = await new SharePointPackageService(ctx).prepare(pkg, identity);
    expect(prepared.isErr()).toBe(true);
    if (prepared.isErr()) expect(prepared.error.name).toBe("ConsentRequired");
    expect(http.history.post).toHaveLength(0);
  });

  it("SPO-02: denied catalog rights are not treated as an absent package", async () => {
    http.onGet(/AvailableApps/).reply(403);
    expect((await new SharePointPackageService(context()).prepare(pkg, identity)).isErr()).toBe(
      true
    );
    expect(http.history.post).toHaveLength(0);
  });

  it.each([
    { AppCatalogVersion: "0.0.0.1" },
    { ErrorMessage: "Package validation failed" },
    { ID: productId },
    { IsClientSideSolution: false },
  ])("SPO-04: rejects invalid/stale catalog state %j within bounded reads", async (change) => {
    const ctx = context();
    http.onGet(/AvailableApps/).replyOnce(404);
    http.onGet(/AvailableApps/).reply(200, { ...app, ...change });
    http.onPost(/\/Add\(/).reply(200, { UniqueId: itemId });
    const pause = vi.fn().mockResolvedValue(undefined);
    const service = new SharePointPackageService(ctx, pause);
    const prepared = await service.prepare(pkg, identity);
    expect(prepared.isOk()).toBe(true);
    if (prepared.isErr()) return;
    expect((await service.deploy(prepared.value, async () => {})).isErr()).toBe(true);
    expect(pause.mock.calls.length).toBeLessThanOrEqual(5);
    expect(http.history.post).toHaveLength(1);
  });

  it("SPO-04 SPO-06: uncertain upload is recorded and POST is not retried", async () => {
    http.onGet(/AvailableApps/).reply(404);
    http.onPost(/\/Add\(/).timeout();
    const service = new SharePointPackageService(context());
    const prepared = await service.prepare(pkg, identity);
    if (prepared.isErr()) return expect.fail(prepared.error.message);
    const stages: string[] = [];
    const result = await service.deploy(prepared.value, async (state) => {
      stages.push(state.stage);
    });
    expect(result.isErr()).toBe(true);
    expect(stages).toEqual(["unknownOutcome"]);
    expect(http.history.post).toHaveLength(1);
  });

  it("SPO-04: honors bounded Retry-After on read-only calls", async () => {
    const pause = vi.fn().mockResolvedValue(undefined);
    http.onGet(/AvailableApps/).reply(429, {}, { "retry-after": "2" });
    const result = await new SharePointPackageService(context(), pause).prepare(pkg, identity);
    expect(result.isErr()).toBe(true);
    expect(pause).toHaveBeenCalledTimes(3);
    expect(pause).toHaveBeenCalledWith(2000);
    expect(http.history.post).toHaveLength(0);
  });

  it("SPO-04: cancellation after upload retains uploaded state and never deploys", async () => {
    const controller = new AbortController();
    http.onGet(/AvailableApps/).reply(404);
    http.onPost(/\/Add\(/).reply(200, { UniqueId: itemId });
    const service = new SharePointPackageService(context(controller.signal));
    const prepared = await service.prepare(pkg, identity);
    if (prepared.isErr()) return expect.fail(prepared.error.message);
    const stages: string[] = [];
    const result = await service.deploy(prepared.value, async (state) => {
      stages.push(state.stage);
      if (state.stage === "uploaded") controller.abort();
    });
    expect(result.isErr()).toBe(true);
    expect(stages.at(-1)).toBe("uploaded");
    expect(http.history.post).toHaveLength(1);
  });

  it.each([
    { name: "different.sppkg", version: pkg.version, content: bytes, nonInteractive: false },
    {
      name: pkg.fileName,
      version: pkg.version,
      content: Buffer.from("different bytes"),
      nonInteractive: false,
    },
    {
      name: pkg.fileName,
      version: "1.0.0.15",
      content: Buffer.from("older"),
      nonInteractive: true,
    },
  ])("SPO-05: collision or unapproved replacement cannot overwrite %j", async (existing) => {
    const ctx = context();
    ctx.nonInteractive = existing.nonInteractive;
    http.onGet(/AvailableApps/).reply(200, { ...app, AppCatalogVersion: existing.version });
    http
      .onGet(`${catalog}/_api/web/GetFileById('${itemId}')?$select=Name`)
      .reply(200, { Name: existing.name });
    http.onGet(`${catalog}/_api/web/GetFileById('${itemId}')/$value`).reply(200, existing.content);
    expect((await new SharePointPackageService(ctx).prepare(pkg, identity)).isErr()).toBe(true);
    expect(http.history.post).toHaveLength(0);
  });

  it("SPO-05 SDK-03: an explicitly approved version upgrade identifies and overwrites only the matched product", async () => {
    const ctx = context();
    vi.spyOn(ctx.ui!, "showMessage").mockResolvedValue(ok("Confirm"));
    http.onGet(/AvailableApps/).reply(200, { ...app, AppCatalogVersion: "1.0.0.15" });
    http
      .onGet(`${catalog}/_api/web/GetFileById('${itemId}')?$select=Name`)
      .reply(200, { Name: pkg.fileName });
    http
      .onGet(`${catalog}/_api/web/GetFileById('${itemId}')/$value`)
      .reply(200, Buffer.from("old"));
    http.onPost(/\/Add\(/).reply(() => {
      http.onGet(/AvailableApps/).reply(200, app);
      http.onGet(`${catalog}/_api/web/GetFileById('${itemId}')/$value`).reply(200, bytes);
      return [200, { UniqueId: itemId }];
    });
    http.onPost(/\/Deploy$/).reply(200);
    const service = new SharePointPackageService(ctx);
    const prepared = await service.prepare(pkg, identity);
    expect(prepared.isOk()).toBe(true);
    if (prepared.isErr()) return;
    for (const value of [pkg.fileName, productId, "1.0.0.15", pkg.version]) {
      expect(ctx.ui!.showMessage).toHaveBeenCalledWith(
        "warn",
        expect.stringContaining(value),
        true,
        "Confirm"
      );
    }
    expect((await service.deploy(prepared.value, async () => {})).isOk()).toBe(true);
    expect(http.history.post[0].url).toContain("overwrite=true");
  });

  it("SPO-02 SPO-03 SPO-04: separate scopes, upload -> visibility -> deploy -> verification", async () => {
    const ctx = context();
    const order: string[] = [];
    http
      .onGet(`${catalog}/_api/web/tenantappcatalog/AvailableApps/GetById('${productId}')`)
      .replyOnce(404);
    http
      .onGet(`${catalog}/_api/web/tenantappcatalog/AvailableApps/GetById('${productId}')`)
      .reply(200, app);
    http.onPost(/\/Add\(/).reply((config) => {
      order.push("upload");
      expect(config.url).toContain("overwrite=false");
      expect(config.url).toContain("test%27%27s%20%231.sppkg");
      expect(config.data).toEqual(bytes);
      return [200, { UniqueId: itemId }];
    });
    http
      .onGet(`${catalog}/_api/web/tenantappcatalog/AvailableApps/GetById('${itemId}')`)
      .reply(() => {
        order.push("read");
        return [200, app];
      });
    http.onGet(`${catalog}/_api/web/GetFileById('${itemId}')/$value`).reply(200, bytes);
    http.onPost(/\/Deploy$/).reply((config) => {
      order.push("deploy");
      expect(JSON.parse(config.data)).toEqual({ skipFeatureDeployment: false });
      return [200];
    });
    const service = new SharePointPackageService(ctx);
    const prepared = await service.prepare(pkg, identity);
    expect(prepared.isOk()).toBe(true);
    if (prepared.isErr()) return;
    const states: string[] = [];
    const deployed = await service.deploy(prepared.value, async (state) => {
      states.push(state.stage);
    });
    expect(deployed.isOk()).toBe(true);
    expect(order).toEqual(["upload", "read", "deploy", "read"]);
    expect(states).toContain("verified");
    expect(ctx.m365TokenProvider.getAccessToken).toHaveBeenCalledWith({
      scopes: ["https://tenant.sharepoint.com/Sites.FullControl.All"],
    });
    expect(http.history.get[0].headers?.Accept).toBe("application/json");
    expect(
      http.history.post.every((r) => !/Teams|SyncSolution|EnsureTenant/.test(r.url || ""))
    ).toBe(true);
  });

  it("SPO-05: exact remote bytes and verified state skip both writes", async () => {
    const service = new SharePointPackageService(context());
    http.onGet(/AvailableApps/).reply(200, app);
    http
      .onGet(`${catalog}/_api/web/GetFileById('${itemId}')?$select=Name`)
      .reply(200, { Name: pkg.fileName });
    http.onGet(`${catalog}/_api/web/GetFileById('${itemId}')/$value`).reply(200, bytes);
    const prepared = await service.prepare(pkg, identity);
    expect(prepared.isOk()).toBe(true);
    if (prepared.isErr()) return;
    const result = await service.deploy(prepared.value, async () => {});
    expect(result.isOk()).toBe(true);
    expect(http.history.post).toHaveLength(0);
  });

  it.each([
    "https://foreign.sharepoint.com/sites/catalog",
    "http://tenant.sharepoint.com/catalog",
    "",
  ])("SPO-02: rejects unsupported catalog authority %s before mutation", async (url) => {
    http
      .onGet("https://tenant.sharepoint.com/_api/SP_TenantSettings_Current")
      .reply(200, { CorporateCatalogUrl: url });
    expect((await new SharePointPackageService(context()).prepare(pkg, identity)).isErr()).toBe(
      true
    );
    expect(http.history.post).toHaveLength(0);
  });

  it("SPO-02: rejects a mismatched selected tenant and discovery denied", async () => {
    const ctx = context();
    vi.mocked(ctx.m365TokenProvider.getJsonObject).mockResolvedValue(
      ok({ tid: "other", oid: productId })
    );
    expect((await new SharePointPackageService(ctx).prepare(pkg, identity)).isErr()).toBe(true);
    expect(http.history.post).toHaveLength(0);
  });

  it("SPO-05: refuses downgrade before upload", async () => {
    http.onGet(/AvailableApps/).reply(200, { ...app, AppCatalogVersion: "2.0.0.0" });
    expect((await new SharePointPackageService(context()).prepare(pkg, identity)).isErr()).toBe(
      true
    );
    expect(http.history.post).toHaveLength(0);
  });

  it("SPO-03: clients retain independent origins", async () => {
    http
      .onGet("https://a.sharepoint.com/_api/SP_TenantSettings_Current")
      .reply(200, { CorporateCatalogUrl: "a" });
    http
      .onGet("https://b.sharepoint.com/_api/SP_TenantSettings_Current")
      .reply(200, { CorporateCatalogUrl: "b" });
    const a = new SPOClient("https://a.sharepoint.com");
    const b = new SPOClient("https://b.sharepoint.com");
    expect(
      await Promise.all([a.getAppCatalogSite("a-token"), b.getAppCatalogSite("b-token")])
    ).toEqual(["a", "b"]);
  });
});
