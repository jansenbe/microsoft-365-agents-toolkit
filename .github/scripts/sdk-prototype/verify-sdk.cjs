// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const os = require("node:os");
const crypto = require("node:crypto");
const net = require("node:net");
const Module = require("node:module");

const consumer = fs.realpathSync(__dirname);
const descriptor = JSON.parse(
  fs.readFileSync(path.join(consumer, "..", "descriptor.json"), "utf8"),
);
const inventories = JSON.parse(
  fs.readFileSync(
    path.join(consumer, "..", "provenance", "runtime-inventories.json"),
    "utf8",
  ),
);
const hash = (bytes) => crypto.createHash("sha256").update(bytes).digest("hex");
const withinConsumer = (file) => {
  const resolved = path.resolve(file);
  const root = consumer + path.sep;
  return process.platform === "win32"
    ? resolved.toLowerCase().startsWith(root.toLowerCase())
    : resolved.startsWith(root);
};
const resolve = Module._resolveFilename;
Module._resolveFilename = function (...args) {
  const resolved = resolve.apply(this, args);
  if (
    typeof resolved === "string" &&
    path.isAbsolute(resolved) &&
    !withinConsumer(resolved)
  ) {
    throw new Error(`Runtime escaped the relocated consumer: ${args[0]}`);
  }
  return resolved;
};
let networkAttempts = 0;
net.Socket.prototype.connect = function () {
  networkAttempts++;
  throw new Error("Network is forbidden during SDK verification");
};
os.homedir = () => path.join(consumer, "empty-verification-home");

async function main() {
  const rawRead = fs.readFileSync;
  let resourceAttempts = 0;
  let blockResources = true;
  fs.readFileSync = function (file, ...args) {
    if (
      blockResources &&
      typeof file === "string" &&
      /[/\\](?:resource|templates|json-schemas|\.fx)[/\\]/i.test(file)
    ) {
      resourceAttempts++;
      throw new Error(
        "Toolkit resources/auth files are forbidden for the lean inspector",
      );
    }
    return rawRead.call(this, file, ...args);
  };
  const lean = require("@microsoft/teamsfx-core/build/sharepointPackage");
  const corePackage = path.dirname(
    require.resolve("@microsoft/teamsfx-core/package.json"),
  );
  const unexpectedCoreModules = Object.keys(require.cache).filter(
    (file) =>
      file.startsWith(corePackage + path.sep) &&
      !file.endsWith(path.join("build", "sharepointPackage.js")) &&
      !file.endsWith(path.join("utility", "packageFile.js")),
  );
  assert.deepEqual(unexpectedCoreModules, []);
  const api = require("@microsoft/teamsfx-api");
  const AdmZip = require("adm-zip");
  const fixtures = path.join(consumer, "verification-fixtures");
  fs.mkdirSync(fixtures, { recursive: true });
  const productId = "11111111-1111-4111-8111-1111111111aa";
  const xml = `<App ProductID="${productId}" Version="1.0.0.0" IsClientSideSolution="true" SkipFeatureDeployment="false"/>`;
  for (const [name, content] of [
    ["valid.sppkg", xml],
    ["malformed.sppkg", "<App>"],
    ["unsafe.sppkg", "<!DOCTYPE App>" + xml],
  ]) {
    const zip = new AdmZip();
    zip.addFile("AppManifest.xml", Buffer.from(content));
    fs.writeFileSync(path.join(fixtures, name), zip.toBuffer());
  }
  const good = await lean.readSharePointPackage("valid.sppkg", fixtures);
  assert.equal(good.isOk(), true);
  assert.equal(good.value.productId, productId);
  assert.equal(good.value.sha256, hash(good.value.bytes));
  const errors = {};
  for (const name of ["missing.sppkg", "malformed.sppkg", "unsafe.sppkg"]) {
    const result = await lean.readSharePointPackage(name, fixtures, {
      errorFactory: (code) => new api.UserError("host", code, `host:${code}`),
    });
    assert.equal(result.isErr(), true);
    assert.ok(result.error instanceof api.UserError);
    errors[name] = result.error.message;
  }
  assert.deepEqual(errors, {
    "missing.sppkg": "host:UnreadablePackage",
    "malformed.sppkg": "host:InvalidPackage",
    "unsafe.sppkg": "host:InvalidPackage",
  });
  assert.equal(resourceAttempts, 0);
  blockResources = false;
  fs.readFileSync = rawRead;
  const core = require("@microsoft/teamsfx-core");
  assert.equal(core.readSharePointPackage, lean.readSharePointPackage);
  assert.equal(typeof core.FxCoreClient.prototype.publish, "function");
  const internal = require("@microsoft/teamsfx-core/build/component/driver/teamsApp/sharePointPublish");
  assert.equal(core.CombinedPublishError, internal.CombinedPublishError);
  assert.equal(
    require.resolve("@microsoft/teamsfx-api", { paths: [corePackage] }),
    require.resolve("@microsoft/teamsfx-api"),
  );
  const partial = {
    packagePath: "agent.zip",
    tenantId: "22222222-2222-4222-8222-2222222222bb",
    recoveryFile: ".local/recovery.json",
    sharePoint: {
      stage: "verified",
      productId,
      version: "1.0.0.0",
      sha256: "0".repeat(64),
    },
    agent: {
      stage: "failed",
      manifestId: "33333333-3333-4333-8333-3333333333cc",
      version: "1.0.0",
      sha256: "1".repeat(64),
      diagnostic: {
        target: "agentSubmit",
        operation: "create",
        status: 403,
        serviceCode: "Forbidden",
      },
    },
  };
  const error = new core.CombinedPublishError(
    partial,
    new api.UserError("host", "Rejected", "Synthetic rejection"),
  );
  assert.ok(error instanceof core.CombinedPublishError);
  assert.ok(error instanceof api.UserError);
  assert.equal(error.publishResult, partial);
  assert.match(error.message, /SharePoint: verified; agent: failed/);
  assert.match(
    core.getLocalizedString("publish.sharepoint.ExpectedPackageMismatch"),
    /SHA-256/,
  );
  const checked = [];
  for (const artifact of descriptor.artifacts) {
    const root = path.dirname(require.resolve(`${artifact.name}/package.json`));
    const metadata = JSON.parse(
      fs.readFileSync(path.join(root, "package.json"), "utf8"),
    );
    assert.equal(metadata.version, artifact.version);
    assert.equal(
      metadata.atkPrototype.sourceIdentity,
      descriptor.source.sourceIdentity,
    );
    assert.deepEqual(metadata.dependencies, artifact.dependencies);
    const omissions = new Set(
      artifact.standardPackOmissions.map((file) => file.path),
    );
    let count = 0;
    for (const file of inventories[artifact.name]) {
      if (omissions.has(file.path)) continue;
      assert.equal(
        hash(fs.readFileSync(path.join(root, file.path))),
        file.sha256,
        `${artifact.name}: ${file.path}`,
      );
      count++;
    }
    checked.push({
      name: artifact.name,
      version: metadata.version,
      verifiedFiles: count,
    });
  }
  assert.equal(networkAttempts, 0);
  const result = {
    relocatedConsumerOnly: true,
    leanCoreModules: 2,
    leanResourceReads: resourceAttempts,
    networkAttempts,
    classIdentity: "one CombinedPublishError and one API instance",
    inspectorOutcomes: errors,
    checked,
  };
  fs.writeFileSync(
    path.join(consumer, "runtime-verification.json"),
    JSON.stringify(result, null, 2) + "\n",
  );
  console.log(JSON.stringify(result, null, 2));
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
