// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const cp = require("node:child_process");
const { createRequire } = require("node:module");
const { npm } = require("./npm-command.cjs");

const workspace = path.resolve(
  process.env.ATK_SOURCE_ROOT || path.join(__dirname, "..", "..", ".."),
);
const release = process.argv[2];
if (!release || !/^\d{8}\.\d+$/.test(release)) {
  throw new Error("Supply a new YYYYMMDD.N release identity.");
}
if (release === "20261020.3")
  throw new Error(
    "The historical .3 identity is reserved; choose a new release.",
  );
const output = path.join(
  path.resolve(
    process.env.ATK_ARTIFACT_ROOT ||
      path.join(workspace, ".local", "sdk-prototype"),
  ),
  release,
);
const git = (...args) =>
  cp.execFileSync("git", ["--no-pager", ...args], {
    cwd: workspace,
    maxBuffer: 32 * 1024 * 1024,
  });
if (
  git("status", "--porcelain", "--untracked-files=normal").toString().trim()
) {
  throw new Error(
    "Package a clean committed checkout. Do not include untracked files or private working-tree material.",
  );
}
if (fs.existsSync(output))
  throw new Error("Output already exists; never overwrite a release identity.");
const coreRequire = createRequire(
  path.join(workspace, "packages", "fx-core", "package.json"),
);
const tar = coreRequire("tar");
const AdmZip = coreRequire("adm-zip");
const semver = coreRequire("semver");
const sha = (bytes) => crypto.createHash("sha256").update(bytes).digest("hex");
const json = (file) => JSON.parse(fs.readFileSync(file, "utf8"));
const writeJson = (file, value) =>
  fs.writeFileSync(file, JSON.stringify(value, null, 2) + "\n");
const relative = (file) =>
  path.relative(workspace, file).split(path.sep).join("/");
const packages = [
  { folder: "manifest", name: "@microsoft/app-manifest", roots: ["build"] },
  { folder: "api", name: "@microsoft/teamsfx-api", roots: ["build"] },
  {
    folder: "spec-parser",
    name: "@microsoft/m365-spec-parser",
    roots: ["dist"],
  },
  {
    folder: "fx-core",
    name: "@microsoft/teamsfx-core",
    roots: ["build", "resource", "templates"],
  },
].map((entry) => {
  const dir = path.join(workspace, "packages", entry.folder);
  const manifest = json(path.join(dir, "package.json"));
  if (manifest.name !== entry.name)
    throw new Error("Unexpected SDK package identity");
  const version = `${manifest.version}-spopublish.${release}`;
  if (!semver.valid(version)) throw new Error("Invalid prerelease identity");
  return { ...entry, dir, manifest, version };
});
const versions = Object.fromEntries(
  packages.map((entry) => [entry.name, entry.version]),
);
const knownPackOmissions = new Set([
  "templates/configs/local/typescript/.gitignore",
  "templates/configs/playground/typescript/.gitignore",
]);

function filesBelow(directory) {
  const files = [];
  for (const entry of fs
    .readdirSync(directory, { withFileTypes: true })
    .sort((a, b) => a.name.localeCompare(b.name))) {
    const file = path.join(directory, entry.name);
    if (entry.isSymbolicLink())
      throw new Error(`Runtime symlink is not permitted: ${relative(file)}`);
    if (entry.isDirectory()) files.push(...filesBelow(file));
    else if (entry.isFile()) files.push(file);
  }
  return files;
}

async function main() {
  const commit = git("rev-parse", "HEAD").toString().trim();
  const tree = git("rev-parse", "HEAD^{tree}").toString().trim();
  const tracked = git("ls-files", "-z")
    .toString("utf8")
    .split("\0")
    .filter(Boolean)
    .sort();
  const sourceFiles = tracked.map((name) => {
    const bytes = fs.readFileSync(path.join(workspace, name));
    return { path: name, bytes: bytes.length, sha256: sha(bytes) };
  });
  const sourceIdentity = sha(Buffer.from(JSON.stringify(sourceFiles)));
  fs.mkdirSync(output, { recursive: true });
  for (const directory of ["provenance", "stage", "tarballs", "consumer"]) {
    fs.mkdirSync(path.join(output, directory));
  }
  for (const file of [
    "pnpm-lock.yaml",
    "package.json",
    "pnpm-workspace.yaml",
  ]) {
    fs.copyFileSync(
      path.join(workspace, file),
      path.join(output, "provenance", file),
    );
  }
  for (const file of [
    "pack-sdk.cjs",
    "npm-command.cjs",
    "verify-sdk.cjs",
    "verify-types.ts",
    "verify-tsconfig.json",
    "finalize-release.cjs",
    "README.md",
  ]) {
    fs.copyFileSync(
      path.join(__dirname, file),
      path.join(output, "provenance", file),
    );
  }
  writeJson(path.join(output, "provenance", "source-files.json"), sourceFiles);
  const artifacts = [];
  const dependencyPins = {};
  const inventories = {};
  for (const entry of packages) {
    const stage = path.join(output, "stage", entry.folder);
    fs.mkdirSync(stage);
    for (const root of entry.roots) {
      const source = path.join(entry.dir, root);
      filesBelow(source);
      fs.cpSync(source, path.join(stage, root), { recursive: true });
    }
    const metadata = structuredClone(entry.manifest);
    metadata.version = entry.version;
    metadata.files = entry.roots;
    delete metadata.gitHead;
    metadata.atkPrototype = { commit, sourceIdentity };
    if (entry.folder === "fx-core" && !metadata.dependencies.tslib) {
      metadata.dependencies.tslib = metadata.devDependencies.tslib;
      delete metadata.devDependencies.tslib;
    }
    dependencyPins[metadata.name] = {};
    for (const [name, specifier] of Object.entries(
      metadata.dependencies || {},
    )) {
      const installed = json(
        path.join(entry.dir, "node_modules", name, "package.json"),
      ).version;
      const exact = versions[name] || installed;
      if (!semver.valid(exact))
        throw new Error(`Invalid dependency version: ${name}`);
      metadata.dependencies[name] = exact;
      dependencyPins[metadata.name][name] = {
        originalSpecifier: specifier,
        installedBuildVersion: installed,
        packedVersion: exact,
      };
    }
    for (const section of [
      "dependencies",
      "optionalDependencies",
      "peerDependencies",
      "devDependencies",
    ]) {
      for (const [name, specifier] of Object.entries(metadata[section] || {})) {
        if (String(specifier).startsWith("workspace:")) {
          if (!versions[name])
            throw new Error(`Unclosed workspace dependency: ${name}`);
          metadata[section][name] = versions[name];
        }
      }
    }
    writeJson(path.join(stage, "package.json"), metadata);
    for (const name of [
      "README.md",
      "NOTICE",
      "NOTICE.txt",
      "ThirdPartyNotices.txt",
    ]) {
      const source = path.join(entry.dir, name);
      if (fs.existsSync(source))
        fs.copyFileSync(source, path.join(stage, name));
    }
    const ownLicense = path.join(entry.dir, "LICENSE.txt");
    const license = fs.existsSync(ownLicense)
      ? ownLicense
      : path.join(workspace, "LICENSE");
    fs.copyFileSync(license, path.join(stage, "LICENSE.txt"));
    const inventory = filesBelow(stage).map((file) => {
      const bytes = fs.readFileSync(file);
      const text = bytes.toString("utf8");
      if (
        text.includes(workspace) ||
        text.includes(workspace.replace(/\\/g, "/"))
      ) {
        throw new Error(
          `Checkout path in runtime: ${path.relative(stage, file)}`,
        );
      }
      if (file.endsWith(".zip")) {
        for (const entry of new AdmZip(bytes).getEntries()) {
          if (entry.isDirectory) continue;
          if (
            entry.entryName.startsWith("/") ||
            /^[A-Za-z]:/.test(entry.entryName) ||
            entry.entryName.split("/").includes("..")
          ) {
            throw new Error("Nonportable runtime ZIP entry");
          }
          const content = entry.getData().toString("utf8");
          if (
            content.includes(workspace) ||
            content.includes(workspace.replace(/\\/g, "/"))
          ) {
            throw new Error("Checkout path in runtime ZIP contents");
          }
        }
      }
      return {
        path: path.relative(stage, file).split(path.sep).join("/"),
        bytes: bytes.length,
        sha256: sha(bytes),
      };
    });
    inventories[entry.name] = inventory;
    const packed = JSON.parse(
      npm(
        [
          "pack",
          "--ignore-scripts",
          "--json",
          "--pack-destination",
          path.join(output, "tarballs"),
        ],
        stage,
      ),
    )[0];
    const bytes = fs.readFileSync(
      path.join(output, "tarballs", packed.filename),
    );
    const integrity =
      "sha512-" + crypto.createHash("sha512").update(bytes).digest("base64");
    if (integrity !== packed.integrity)
      throw new Error("npm integrity mismatch");
    const packedPaths = new Set(packed.files.map((file) => file.path));
    const missing = inventory.filter(
      (file) =>
        entry.roots.some((root) => file.path.startsWith(root + "/")) &&
        !packedPaths.has(file.path),
    );
    if (
      missing.some(
        (file) =>
          entry.folder !== "fx-core" || !knownPackOmissions.has(file.path),
      )
    ) {
      throw new Error("Unexpected runtime asset omitted by npm pack");
    }
    artifacts.push({
      name: entry.name,
      version: entry.version,
      folder: entry.folder,
      file: `tarballs/${packed.filename}`,
      sha256: sha(bytes),
      integrity,
      shasum: packed.shasum,
      bytes: bytes.length,
      unpackedBytes: packed.unpackedSize,
      entries: packed.entryCount,
      license: metadata.license,
      licenseSource: relative(license),
      licenseSha256: sha(fs.readFileSync(license)),
      dependencies: metadata.dependencies,
      runtimeRoots: entry.roots,
      standardPackOmissions: missing.map((file) => ({
        ...file,
        reason: "Standard npm exclusion of unused config .gitignore files",
      })),
    });
  }
  for (const required of [
    "templates/metadata/allTemplates.json",
    "templates/metadata/defaultGeneratorTemplates.json",
  ]) {
    if (
      !inventories["@microsoft/teamsfx-core"].some(
        (file) => file.path === required,
      )
    ) {
      throw new Error(`Required bundled metadata is missing: ${required}`);
    }
  }
  writeJson(
    path.join(output, "provenance", "dependency-pins.json"),
    dependencyPins,
  );
  writeJson(
    path.join(output, "provenance", "runtime-inventories.json"),
    inventories,
  );
  await tar.create(
    {
      file: path.join(output, "provenance", "staged-runtime-inputs.tgz"),
      cwd: path.join(output, "stage"),
      gzip: true,
      portable: true,
      mtime: new Date(0),
    },
    packages.map((entry) => entry.folder),
  );
  const compilerVersion = json(
    path.join(
      workspace,
      "packages",
      "fx-core",
      "node_modules",
      "typescript",
      "package.json",
    ),
  ).version;
  const nodeTypesVersion = json(
    path.join(
      workspace,
      "packages",
      "fx-core",
      "node_modules",
      "@types",
      "node",
      "package.json",
    ),
  ).version;
  writeJson(path.join(output, "consumer", "package.json"), {
    name: "atk-sharepoint-sdk-verification",
    version: "0.0.0",
    private: true,
    engines: { node: ">=24" },
    dependencies: Object.fromEntries(
      artifacts.map((artifact) => [artifact.name, `file:../${artifact.file}`]),
    ),
    overrides: Object.fromEntries(
      artifacts.map((artifact) => [artifact.name, `$${artifact.name}`]),
    ),
    devDependencies: {
      typescript: compilerVersion,
      "@types/node": nodeTypesVersion,
    },
  });
  fs.writeFileSync(
    path.join(output, "consumer", ".npmrc"),
    "ignore-scripts=true\naudit=false\nfund=false\n",
  );
  for (const file of [
    "verify-sdk.cjs",
    "verify-types.ts",
    "verify-tsconfig.json",
  ]) {
    fs.copyFileSync(
      path.join(__dirname, file),
      path.join(output, "consumer", file),
    );
  }
  const descriptor = {
    schemaVersion: 1,
    state: "packed-awaiting-consumer-verification",
    release,
    source: {
      upstreamRepository:
        "https://github.com/OfficeDev/microsoft-365-agents-toolkit.git",
      commit,
      tree,
      sourceIdentity,
      sourceLockSha256: sha(
        fs.readFileSync(path.join(workspace, "pnpm-lock.yaml")),
      ),
      stagedRuntimeSha256: sha(
        fs.readFileSync(
          path.join(output, "provenance", "staged-runtime-inputs.tgz"),
        ),
      ),
    },
    toolchain: {
      node: process.version,
      npm: npm(["--version"], workspace).trim(),
      typescript: compilerVersion,
    },
    packaging: {
      method: "Standard npm pack from compiled/runtime inputs",
      transformations: [
        "Distinct prerelease versions",
        "Exact SDK closure and installed direct dependency pins",
        "Full compiled roots including nested declarations",
        "Removed stale gitHead",
        "Declared runtime tslib",
        "Preserved original MIT texts",
        "Recorded committed-source identity",
      ],
    },
    artifacts,
  };
  writeJson(path.join(output, "descriptor.json"), descriptor);
  console.log(
    JSON.stringify(
      {
        output,
        sourceIdentity,
        artifacts: artifacts.map(({ name, version, file, sha256 }) => ({
          name,
          version,
          file,
          sha256,
        })),
      },
      null,
      2,
    ),
  );
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
