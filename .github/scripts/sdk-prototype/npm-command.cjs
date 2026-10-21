// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

const fs = require("node:fs");
const path = require("node:path");
const cp = require("node:child_process");

function npm(args, cwd, stdio = ["ignore", "pipe", "pipe"]) {
  const cli =
    process.env.NPM_CLI_JS ||
    (process.platform === "win32"
      ? path.join(
          path.dirname(fs.realpathSync(process.execPath)),
          "node_modules",
          "npm",
          "bin",
          "npm-cli.js",
        )
      : undefined);
  return cp.execFileSync(
    cli ? process.execPath : "npm",
    cli ? [cli, ...args] : args,
    {
      cwd,
      encoding: "utf8",
      maxBuffer: 32 * 1024 * 1024,
      stdio,
      env: {
        ...process.env,
        npm_config_audit: "false",
        npm_config_fund: "false",
      },
    },
  );
}

module.exports = { npm };
