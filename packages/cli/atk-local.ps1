# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.

$ErrorActionPreference = "Stop"
& node (Join-Path $PSScriptRoot "cli.js") @args
exit $LASTEXITCODE
