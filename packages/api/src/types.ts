// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import {
  IBot,
  IComposeExtension,
  IConfigurableTab,
  IStaticTab,
  IWebApplicationInfo,
} from "@microsoft/app-manifest";
import { Platform } from "./constants";
import { Result } from "neverthrow";
import { FxError } from "./error";

/**
 * Definition of option item in single selection or multiple selection
 */
export interface OptionItem {
  /**
   * unique identifier of the option item in the option list
   */
  id: string;
  /**
   * display name
   */
  label: string;
  /**
   * short description
   */
  description?: string;
  /**
   * detailed description
   */
  detail?: string;
  /**
   * customized user data, which is not displayed
   */
  data?: unknown;
  /**
   * @deprecated CLI display name. CLI will use `cliName` as display name, and use `id` instead if `cliName` is undefined.
   */
  cliName?: string;
  /**
   * group name. If it's set, separator will be rendered on UI between groups.
   */
  groupName?: string;

  /**
   * Actions that can be made within the item.
   * @param An array of actions
   * @param `icon` is the icon id of the action item
   * @param `tooltip` is the hint of the action item
   * @param `command` is the command name that will be executed when current action triggered
   */
  buttons?: { iconPath: string; tooltip: string; command: string }[];
}

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
export type Void = {};
export const Void = {};
/**
 * environment meta data
 */
export interface EnvMeta {
  name: string;
  local: boolean;
  sideloading: boolean;
}
export type AuthCredentialSource = "provision" | "environment";
export interface Inputs extends Record<string, any> {
  platform: Platform;
  projectPath?: string;
  projectId?: string;
  nonInteractive?: boolean;
  correlationId?: string;
  /**
   * whether the caller is triggered by @teams or @office agent
   */
  agent?: "teams" | "office";
  /**
   * Auth info about user selected APIs.
   */
  apiAuthData?: AuthInfo[];
  /**
   * Source for credentials referenced by generated authentication actions.
   */
  authCredentialSource?: AuthCredentialSource;
}

export type InputsWithProjectPath = Inputs & { projectPath: string };

export type CreateProjectInputs = Inputs & { "app-name": string; folder: string };

// This type has not been supported by TypeScript yet.
// Check here https://github.com/microsoft/TypeScript/issues/13923.
export type DeepReadonly<T> = {
  readonly [P in keyof T]: DeepReadonly<T[P]>;
};

export type MaybePromise<T> = T | Promise<T>;

/**
 * simplified tooling settings for v3
 */
export interface Settings {
  version: string;
  trackingId: string;
}

export type ManifestCapability =
  | {
      name: "staticTab";
      snippet?: IStaticTab;
      existingApp?: boolean;
    }
  | {
      name: "configurableTab";
      snippet?: IConfigurableTab;
      existingApp?: boolean;
    }
  | {
      name: "Bot";
      snippet?: IBot;
      existingApp?: boolean;
    }
  | {
      name: "MessageExtension";
      snippet?: IComposeExtension;
      existingApp?: boolean;
    }
  | {
      name: "WebApplicationInfo";
      snippet?: IWebApplicationInfo;
      existingApp?: boolean;
    };

export interface AuthInfo {
  serverUrl: string;
  authName?: string;
  authType?: "apiKey" | "oauth2";
}

export interface ApiOperation {
  id: string;
  label: string;
  groupName: string;
  data: AuthInfo;
  detail?: string;
}

export interface Warning {
  type: string;
  content: string;
  data?: any;
}

export interface CreateProjectResult {
  projectPath: string;
  warnings?: Warning[];
  shouldInvokeTeamsAgent?: boolean;
  projectId?: string;
  lastCommand?: string;
}

export interface TeamsAppInputs extends InputsWithProjectPath {
  "manifest-file"?: string;
  "package-file"?: string;
  /** Standalone SPFx package; relative paths resolve against projectPath. */
  "sharepoint-package-file"?: string;
  /** Optional tenant GUID constraint for SharePoint-enabled publication, checked before writes. */
  expectedTenantId?: string;
  /** Optional selector digest binding, checked against the local SharePoint snapshot before auth. */
  expectedSharePointSha256?: string;
  "output-package-file"?: string;
  "output-folder"?: string;
  env?: string;
  "env-file"?: string;
}

export interface SharePointPackageResult {
  stage: "notStarted" | "uploaded" | "deployed" | "verified" | "failed" | "unknownOutcome";
  lastSuccessfulStage?: "uploaded" | "deployed" | "verified";
  productId: string;
  version: string;
  sha256: string;
  catalogUrl?: string;
  catalogItemId?: string;
}

export interface SharePointPublishPreparationContext {
  readonly projectPath: string;
  readonly tenantId: string;
  readonly catalogUrl: string;
  readonly originalSharePoint: {
    readonly packagePath: string;
    readonly productId: string;
    readonly version: string;
    readonly sha256: string;
    readonly skipFeatureDeployment: boolean;
  };
  readonly signal?: AbortSignal;
}

/** Artifact hashes are verified by ATK; embedded selection/plugin provenance is host-attested. */
export interface SharePointPublishLineage {
  readonly schemaVersion: 1;
  readonly originalSharePointSha256: string;
  readonly transformationVersion: string;
  readonly selectedEmbeddedPackage: string;
  readonly extractedPluginSha256: string;
  readonly derivedSharePointSha256: string;
  readonly finalAppSha256: string;
}

export interface PreparedSharePointPublishArtifacts {
  readonly sharePointPackagePath: string;
  readonly appPackagePath: string;
  readonly lineage: SharePointPublishLineage;
}

export type SharePointPublishPreparer = (
  context: SharePointPublishPreparationContext
) => Promise<Result<PreparedSharePointPublishArtifacts, FxError>>;

export interface CombinedPublishResult {
  packagePath: string;
  tenantId: string;
  sharePoint: SharePointPackageResult;
  agent: {
    stage: "notStarted" | "submitted" | "failed" | "cancelled" | "unknownOutcome";
    manifestId: string;
    version: string;
    sha256: string;
    catalogId?: string;
    diagnostic?: PublishRequestDiagnostic;
  };
  recoveryFile: string;
  /** Present for host-prepared publication; never replaced by the staging directory. */
  projectPath?: string;
  lineage?: SharePointPublishLineage;
}

/** Allowlisted, sanitized request evidence; never contains tokens or raw HTTP objects. */
export interface PublishRequestDiagnostic {
  target: "sharePoint" | "agentCatalogRead" | "agentSubmit";
  operation?: "create" | "update";
  status?: number;
  serviceCode?: string;
  serviceMessage?: string;
  requestId?: string;
  transportCode?: string;
}
