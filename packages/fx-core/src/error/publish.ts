// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, PublishRequestDiagnostic, SystemError, UserError } from "@microsoft/teamsfx-api";
import axios from "axios";
import { getLocalizedString } from "../common/localizeUtils";
import { maskSecret } from "../common/stringUtils";
import { UserCancelError } from "./common";

const rejectedStatuses = [400, 401, 403, 404, 413, 422];
const transportCodes = new Set([
  "ETIMEDOUT",
  "ECONNABORTED",
  "ECONNRESET",
  "ECONNREFUSED",
  "ENOTFOUND",
  "EAI_AGAIN",
  "EHOSTUNREACH",
  "ENETUNREACH",
  "ERR_NETWORK",
  "UND_ERR_CONNECT_TIMEOUT",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function object(value: unknown): Record<string, unknown> {
  return isRecord(value) ? value : {};
}

function safeMessage(value: unknown): string | undefined {
  if (typeof value !== "string" || !value) return undefined;
  return maskSecret(
    value
      .slice(0, 8192)
      .replace(/https?:\/\/[^\s<>"']+/gi, "[redacted URL]")
      .replace(/\b(?:Bearer|Basic)\s+[^\s,;]+/gi, "[redacted authorization]")
      .replace(
        /\b(?:access_token|refresh_token|id_token|client_secret|authorization_code|auth_code|code_verifier|code_challenge|password|assertion|token|code)\s*[:=]\s*(?:"[^"]*"|'[^']*'|[^\s&,;]+)/gi,
        "[redacted credential]"
      )
      .replace(/\b[A-Za-z0-9_-]{12,}\.[A-Za-z0-9_-]{12,}\.[A-Za-z0-9_-]{8,}\b/g, "[redacted token]")
      .replace(/[\u0000-\u001f\u007f-\u009f]/g, " ")
  ).slice(0, 1024);
}

function safeCode(value: unknown): string | undefined {
  if (typeof value !== "string" || !/^[A-Za-z][A-Za-z0-9_]{0,95}$/.test(value)) return undefined;
  // Protocol identifiers such as Authorization_RequestDenied are not opaque credentials.
  return value.split("_").every((part) => maskSecret(part) === part) ? value : undefined;
}

function requestId(value: unknown): string | undefined {
  return typeof value === "string" &&
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value)
    ? value
    : undefined;
}

function diagnosticMessage(diagnostic: PublishRequestDiagnostic): string {
  const unknown = getLocalizedString("publish.diagnostic.Unavailable");
  const details = getLocalizedString(
    "publish.diagnostic.Details",
    diagnostic.target === "agentSubmit"
      ? diagnostic.operation === "update"
        ? "publishTeamsAppUpdate"
        : "publishTeamsApp"
      : diagnostic.target,
    diagnostic.status ?? unknown,
    diagnostic.serviceCode ?? unknown,
    diagnostic.serviceMessage ?? unknown,
    diagnostic.requestId ?? unknown,
    diagnostic.transportCode ?? unknown
  );
  if (diagnostic.target === "sharePoint") {
    return `${getLocalizedString("publish.sharepoint.RemoteFailure", diagnostic.status ?? unknown)} ${details}`;
  }
  const key =
    diagnostic.target === "agentCatalogRead"
      ? "CatalogReadFailure"
      : diagnostic.status === 403
        ? "SubmitForbidden"
        : rejectedStatuses.includes(diagnostic.status ?? 0)
          ? "SubmitRejected"
          : "SubmitUnknown";
  return getLocalizedString(`publish.agent.${key}`, details);
}

export class PublishRemoteError extends SystemError {
  constructor(public readonly diagnostic: PublishRequestDiagnostic) {
    super({
      source: diagnostic.target === "sharePoint" ? "SharePointPublish" : "GraphPublish",
      name: "PublishRemoteRequestFailed",
      message: diagnosticMessage(diagnostic),
    });
  }

  get status(): number | undefined {
    return this.diagnostic.status;
  }

  get isDefinitiveRejection(): boolean {
    return rejectedStatuses.includes(this.status ?? 0);
  }
}

export function remotePublishError(
  error: unknown,
  target: PublishRequestDiagnostic["target"] = "sharePoint",
  operation?: PublishRequestDiagnostic["operation"]
): FxError {
  if (error instanceof UserError || error instanceof SystemError) return error;
  if (axios.isCancel(error) || (error instanceof Error && error.name === "AbortError")) {
    return new UserCancelError(target === "sharePoint" ? "SharePointPublish" : "GraphPublish");
  }
  // Read only allowlisted fields. Axios config/request can include authorization and package bytes.
  const raw = object(error);
  const response = object(raw.response);
  const data = object(response.data);
  const service = object(data.error);
  const inner = object(service.innerError ?? service.innererror);
  const headers = object(response.headers);
  const headerId = Object.entries(headers).find(([key]) => key.toLowerCase() === "request-id")?.[1];
  const details = Array.isArray(inner.details) ? inner.details.slice(0, 3) : [];
  const messages = [
    service.message,
    inner.message,
    ...details.map((detail) => object(detail).message),
    data.errorMessage,
  ]
    .map(safeMessage)
    .filter((message): message is string => !!message);
  const code = raw.code ?? object(raw.cause).code;
  const status = response.status;
  return new PublishRemoteError({
    target,
    operation,
    status:
      typeof status === "number" && Number.isInteger(status) && status >= 100 && status <= 599
        ? status
        : undefined,
    serviceCode: safeCode(service.code) ?? safeCode(inner.code),
    serviceMessage: messages.length
      ? messages.join("; ").slice(0, 1024)
      : response.data === undefined
        ? safeMessage(raw.message)
        : undefined,
    requestId: requestId(headerId) ?? requestId(inner["request-id"]),
    transportCode: typeof code === "string" && transportCodes.has(code) ? code : undefined,
  });
}
