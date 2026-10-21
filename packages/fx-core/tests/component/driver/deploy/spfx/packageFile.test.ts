// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { UserError } from "@microsoft/teamsfx-api";
import { afterEach, describe, expect, it, vi } from "vitest";
import { readSharePointPackage } from "../../../../../src/component/driver/deploy/spfx/utility/packageFile";
import { readSharePointPackage as publicRead } from "../../../../../src/sharepointPackage";
import * as localize from "../../../../../src/common/localizeUtils";

const productId = "11111111-1111-4111-8111-1111111111aa";
const xml = `<App xmlns="http://schemas.microsoft.com/sharepoint/2012/app/manifest"
 ProductID="{${productId}}" Version="1.0.0.16" IsClientSideSolution="true"
 SkipFeatureDeployment="false"><Properties><Title>Example</Title></Properties></App>`;
const folders: string[] = [];

async function fixture(content = xml, name = "example.sppkg"): Promise<string> {
  const folder = await fs.mkdtemp(path.join(os.tmpdir(), "atk-sppkg-"));
  folders.push(folder);
  const zip = new AdmZip();
  zip.addFile("AppManifest.xml", Buffer.from(content));
  await fs.writeFile(path.join(folder, name), zip.toBuffer());
  return folder;
}

afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(folders.splice(0).map((folder) => fs.remove(folder)));
});

describe("standalone SharePoint package", () => {
  it("SDK-01: lean inspection shares the parser and handles failures without toolkit resources", async () => {
    expect(publicRead).toBe(readSharePointPackage);
    const folder = await fixture();
    const unsafeFolder = await fixture(
      `<!DOCTYPE App [<!ENTITY x SYSTEM "file:///not-read">]>${xml}`
    );
    const invalidFolder = await fixture("<App>");
    vi.spyOn(localize, "getLocalizedString").mockImplementation(() => {
      throw new Error("Toolkit resources are unavailable in this host");
    });
    const errorFactory = vi.fn((code: string) => new UserError("host", code, `host:${code}`));
    expect((await publicRead("example.sppkg", folder, { errorFactory })).isOk()).toBe(true);
    for (const [file, root, reason] of [
      ["missing.sppkg", folder, "UnreadablePackage"],
      ["example.sppkg", unsafeFolder, "InvalidPackage"],
      ["example.sppkg", invalidFolder, "InvalidPackage"],
    ]) {
      const result = await publicRead(file, root, { errorFactory });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.message).toBe(`host:${reason}`);
    }
    const defaultError = await publicRead("missing.sppkg", folder);
    expect(defaultError.isErr()).toBe(true);
    if (defaultError.isErr()) expect(defaultError.error.message).toBe("UnreadablePackage");
    expect(localize.getLocalizedString).not.toHaveBeenCalled();
    expect(await fs.readdir(folder)).toEqual(["example.sppkg"]);
  });
  it("SPO-01: resolves against folder and reads packaged metadata without source config", async () => {
    const folder = await fixture();
    const result = await readSharePointPackage("example.sppkg", folder);
    expect(result.isOk()).toBe(true);
    if (result.isOk()) {
      expect(result.value).toMatchObject({
        productId,
        version: "1.0.0.16",
        skipFeatureDeployment: false,
        fileName: "example.sppkg",
      });
      expect(result.value.sha256).toMatch(/^[a-f0-9]{64}$/);
      expect(result.value.bytes).toEqual(await fs.readFile(path.join(folder, "example.sppkg")));
    }
  });

  it.each([
    xml.replace('IsClientSideSolution="true"', 'IsClientSideSolution="false"'),
    xml.replace("1.0.0.16", "1.0"),
    xml.replace(productId, "not-a-guid"),
    `<!DOCTYPE App [<!ENTITY x SYSTEM "file:///etc/passwd">]>${xml}`,
    "<App>",
  ])("SPO-01: rejects unsafe or invalid XML: %s", async (content) => {
    const folder = await fixture(content);
    expect((await readSharePointPackage("example.sppkg", folder)).isErr()).toBe(true);
  });

  it("SPO-01: rejects missing, non-sppkg and ambiguous manifest entries", async () => {
    const folder = await fixture();
    expect((await readSharePointPackage("missing.sppkg", folder)).isErr()).toBe(true);
    expect((await readSharePointPackage("example.zip", folder)).isErr()).toBe(true);
    const zip = new AdmZip();
    zip.addFile("AppManifest.xml", Buffer.from(xml));
    zip.addFile("appmanifest.xml", Buffer.from(xml));
    await fs.writeFile(path.join(folder, "example.sppkg"), zip.toBuffer());
    expect((await readSharePointPackage("example.sppkg", folder)).isErr()).toBe(true);
  });
});
