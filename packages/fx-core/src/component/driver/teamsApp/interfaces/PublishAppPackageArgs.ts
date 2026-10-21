// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { PreparedAgentPublication } from "../preparedPublication";

export interface PublishAppPackageArgs {
  /**
   * Zipped app package path
   */
  appPackagePath: string;
  /** In-memory preflight snapshot for the opt-in coordinated publication only. */
  preparedPublication?: PreparedAgentPublication;
}
