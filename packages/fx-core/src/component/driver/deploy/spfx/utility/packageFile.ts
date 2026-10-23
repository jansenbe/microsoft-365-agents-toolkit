// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Result, UserError, err, ok } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { createHash } from "crypto";
import fs from "fs-extra";
import path from "path";
import { inflateRawSync } from "zlib";
import { parseStringPromise } from "xml2js";

export const maxPackageBytes = 100 * 1024 * 1024;
export const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export const versionPattern = /^\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5}$/;

export interface SharePointPackage {
  bytes: Buffer;
  fileName: string;
  productId: string;
  version: string;
  skipFeatureDeployment: boolean;
  sha256: string;
}

export type SharePointPackageReadErrorCode =
  "InvalidPackage" | "UnreadablePackage" | "ChangedPackage";

export interface SharePointPackageReadOptions {
  /** Supplies host-localized errors without loading toolkit resources. */
  errorFactory?: (code: SharePointPackageReadErrorCode) => FxError;
}

export function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function sha256(bytes: Buffer): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function inspectionError(reason: SharePointPackageReadErrorCode): UserError {
  return new UserError({
    source: "SharePointPublish",
    name: `SharePoint${reason}`,
    message: reason,
  });
}

export async function readPackageSnapshot(
  file: string,
  options: SharePointPackageReadOptions = {}
): Promise<Result<Buffer, FxError>> {
  const packageError = options.errorFactory ?? inspectionError;
  let bytes: Buffer;
  try {
    const handle = await fs.open(file, "r");
    try {
      const stat = await fs.fstat(handle);
      if (!stat.isFile() || stat.size === 0 || stat.size > maxPackageBytes) {
        return err(packageError("InvalidPackage"));
      }
      // Bound allocation even if the file grows after fstat; keep these exact bytes for upload.
      bytes = Buffer.alloc(stat.size);
      let offset = 0;
      while (offset < bytes.length) {
        const read = await fs.read(handle, bytes, offset, bytes.length - offset, offset);
        if (read.bytesRead === 0) return err(packageError("InvalidPackage"));
        offset += read.bytesRead;
      }
      const after = await fs.fstat(handle);
      if (after.size !== stat.size || after.mtimeMs !== stat.mtimeMs) {
        return err(packageError("ChangedPackage"));
      }
    } finally {
      await fs.close(handle);
    }
  } catch {
    return err(packageError("UnreadablePackage"));
  }
  return ok(bytes);
}

export async function readSharePointPackage(
  file: string,
  projectPath: string,
  options: SharePointPackageReadOptions = {}
): Promise<Result<SharePointPackage, FxError>> {
  const packageError = options.errorFactory ?? inspectionError;
  if (path.extname(file).toLowerCase() !== ".sppkg") return err(packageError("InvalidPackage"));
  const snapshot = await readPackageSnapshot(path.resolve(projectPath, file), options);
  if (snapshot.isErr()) return err(snapshot.error);
  const bytes = snapshot.value;
  try {
    const entries = new AdmZip(bytes).getEntries();
    const names = new Set<string>();
    let expanded = 0;
    if (entries.length > 10000) return err(packageError("InvalidPackage"));
    for (const entry of entries) {
      const name = entry.entryName.replace(/\\/g, "/").toLowerCase();
      expanded += entry.header.size;
      if (
        names.has(name) ||
        name.startsWith("/") ||
        name.includes(":") ||
        name.split("/").includes("..") ||
        expanded > 512 * 1024 * 1024
      )
        return err(packageError("InvalidPackage"));
      names.add(name);
    }
    const manifests = entries.filter(
      (entry) => entry.entryName.toLowerCase() === "appmanifest.xml"
    );
    if (manifests.length !== 1) return err(packageError("InvalidPackage"));
    const entry = manifests[0];
    const limit = 1024 * 1024;
    if (entry.header.size > limit || ![0, 8].includes(entry.header.method)) {
      return err(packageError("InvalidPackage"));
    }
    const data =
      entry.header.method === 0
        ? entry.getCompressedData()
        : inflateRawSync(entry.getCompressedData(), { maxOutputLength: limit });
    if (data.length !== entry.header.size || data.length > limit)
      return err(packageError("InvalidPackage"));
    const xml = data.toString("utf8");
    if (/<!DOCTYPE|<!ENTITY|\u0000/i.test(xml)) return err(packageError("InvalidPackage"));
    const parsed: unknown = await parseStringPromise(xml, {
      strict: true,
      explicitArray: true,
    });
    if (!record(parsed) || !record(parsed.App) || !record(parsed.App.$)) {
      return err(packageError("InvalidPackage"));
    }
    const attrs = parsed.App.$;
    const productId =
      typeof attrs.ProductID === "string"
        ? attrs.ProductID.replace(/^\{|\}$/g, "").toLowerCase()
        : "";
    const version = attrs.Version;
    if (
      !guidPattern.test(productId) ||
      typeof version !== "string" ||
      !versionPattern.test(version) ||
      version.split(".").some((part) => Number(part) > 65535) ||
      attrs.IsClientSideSolution !== "true" ||
      (attrs.SkipFeatureDeployment !== "true" &&
        attrs.SkipFeatureDeployment !== "false" &&
        attrs.SkipFeatureDeployment !== undefined)
    )
      return err(packageError("InvalidPackage"));
    return ok({
      bytes,
      fileName: path.basename(file),
      productId,
      version,
      skipFeatureDeployment: attrs.SkipFeatureDeployment === "true",
      sha256: sha256(bytes),
    });
  } catch {
    return err(packageError("InvalidPackage"));
  }
}
