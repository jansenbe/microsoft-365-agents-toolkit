// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const cp = require("node:child_process");
const { npm } = require("./npm-command.cjs");

if (!process.argv[2]) throw new Error("Supply the new artifact directory.");
const release = path.resolve(process.argv[2]);
const descriptorFile = path.join(release, "descriptor.json");
const descriptor = JSON.parse(fs.readFileSync(descriptorFile, "utf8"));
if (descriptor.release === "20261020.3")
  throw new Error("Historical .3 artifacts must not be changed.");
if (fs.existsSync(path.join(release, "qualification.json")))
  throw new Error("Release is already qualified; do not overwrite it.");
const hash = (file) =>
  crypto.createHash("sha256").update(fs.readFileSync(file)).digest("hex");
const consumer = path.join(release, "relocated-consumer");
if (fs.existsSync(consumer))
  throw new Error(
    "Relocated consumer already exists; inspect it before proceeding.",
  );
fs.renameSync(path.join(release, "consumer"), consumer);
npm(
  ["ci", "--offline", "--ignore-scripts", "--no-audit", "--no-fund"],
  consumer,
  "inherit",
);
cp.execFileSync(process.execPath, ["verify-sdk.cjs"], {
  cwd: consumer,
  stdio: "inherit",
});
cp.execFileSync(
  process.execPath,
  [
    path.join("node_modules", "typescript", "bin", "tsc"),
    "-p",
    "verify-tsconfig.json",
  ],
  { cwd: consumer, stdio: "inherit" },
);
const lockFile = path.join(consumer, "package-lock.json");
const lock = JSON.parse(fs.readFileSync(lockFile, "utf8"));
const runtime = JSON.parse(
  fs.readFileSync(path.join(consumer, "runtime-verification.json"), "utf8"),
);
if (
  !runtime.relocatedConsumerOnly ||
  runtime.networkAttempts ||
  runtime.leanResourceReads
)
  throw new Error("Consumer verification failed");
const repack = path.join(release, "repack-verification");
fs.mkdirSync(repack);
for (const artifact of descriptor.artifacts) {
  if (hash(path.join(release, artifact.file)) !== artifact.sha256)
    throw new Error("Tarball changed");
  const entry = lock.packages[`node_modules/${artifact.name}`];
  if (
    !entry ||
    entry.version !== artifact.version ||
    entry.integrity !== artifact.integrity ||
    entry.resolved !== `file:../${artifact.file}`
  )
    throw new Error("Lock does not bind exact artifact");
  const repeated = JSON.parse(
    npm(
      ["pack", "--ignore-scripts", "--json", "--pack-destination", repack],
      path.join(release, "stage", artifact.folder),
    ),
  )[0];
  if (hash(path.join(repack, repeated.filename)) !== artifact.sha256)
    throw new Error("Repack is not byte-identical");
}
const licenses = [];
for (const [name, entry] of Object.entries(lock.packages)) {
  if (!name) continue;
  if (entry.resolved?.startsWith("http")) {
    const url = new URL(entry.resolved);
    if (
      url.username ||
      url.password ||
      [...url.searchParams.keys()].some((key) =>
        /sig|token|secret|password/i.test(key),
      )
    ) {
      throw new Error(
        "Credential-bearing URL in generated lock; do not distribute it",
      );
    }
  } else if (
    entry.resolved?.startsWith("file:") &&
    !entry.resolved.startsWith("file:../tarballs/")
  ) {
    throw new Error("Unexpected local dependency in generated lock");
  }
  const root = path.resolve(consumer, name);
  if (!root.startsWith(consumer + path.sep))
    throw new Error("Lock path escaped consumer");
  let metadata;
  try {
    metadata = JSON.parse(
      fs.readFileSync(path.join(root, "package.json"), "utf8"),
    );
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
  const texts = metadata
    ? fs
        .readdirSync(root)
        .filter((file) =>
          /^(LICENSE|LICENCE|COPYING|NOTICE)(?:[._-].*)?$/i.test(file),
        )
        .filter((file) => fs.statSync(path.join(root, file)).isFile())
        .map((file) => ({ file, sha256: hash(path.join(root, file)) }))
    : [];
  licenses.push({
    path: name,
    name: metadata?.name,
    version: entry.version,
    declaredLicense:
      metadata?.license ?? metadata?.licenses ?? entry.license ?? null,
    installed: !!metadata,
    optional: !!entry.optional,
    licenseFiles: texts,
  });
}
for (const [source, target] of [
  ["package-lock.json", "consumer-package-lock.json"],
  ["package.json", "consumer-package.json"],
  ["runtime-verification.json", "consumer-runtime-verification.json"],
]) {
  fs.copyFileSync(
    path.join(consumer, source),
    path.join(release, "provenance", target),
  );
}
fs.writeFileSync(
  path.join(release, "provenance", "consumer-license-inventory.json"),
  JSON.stringify(licenses, null, 2) + "\n",
);
const qualification = {
  schemaVersion: 1,
  status: "qualified",
  descriptorSha256: hash(descriptorFile),
  sourceCommit: descriptor.source.commit,
  sourceIdentity: descriptor.source.sourceIdentity,
  consumerLock: {
    file: "provenance/consumer-package-lock.json",
    sha256: hash(lockFile),
    lockfileVersion: lock.lockfileVersion,
  },
  verified: {
    frozenOfflineNpmCiAfterRelocation: true,
    strictRuntimeBoundary: true,
    typecheck: true,
    repeatNpmPackByteIdentical: true,
    classIdentity: runtime.classIdentity,
    leanResourceReads: runtime.leanResourceReads,
    networkAttempts: runtime.networkAttempts,
  },
  artifacts: descriptor.artifacts.map(
    ({ name, version, file, sha256, integrity }) => ({
      name,
      version,
      file,
      sha256,
      integrity,
    }),
  ),
};
fs.writeFileSync(
  path.join(release, "qualification.json"),
  JSON.stringify(qualification, null, 2) + "\n",
);
console.log(
  JSON.stringify(
    {
      status: "qualified",
      qualificationSha256: hash(path.join(release, "qualification.json")),
    },
    null,
    2,
  ),
);
