// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosInstance } from "axios";
import { maxPackageBytes, record } from "./packageFile";

export function odataLiteral(value: string): string {
  return encodeURIComponent(value.replace(/'/g, "''")).replace(/'/g, "%27");
}

export function unwrap(value: unknown): Record<string, unknown> {
  if (!record(value)) return {};
  return record(value.d) ? value.d : value;
}

/** One instance per deployment; no shared tenant/catalog or token state. */
export class SPOClient {
  constructor(
    private baseUrl: string,
    private readonly signal?: AbortSignal
  ) {}

  public setBaseUrl(url: string): void {
    this.baseUrl = url;
  }

  private createRequesterWithToken(spoToken: string): AxiosInstance {
    const instance = axios.create({
      baseURL: this.baseUrl,
      signal: this.signal,
      timeout: 30000,
      maxRedirects: 0,
      maxContentLength: maxPackageBytes,
    });
    instance.defaults.headers.common["Authorization"] = `Bearer ${spoToken}`;
    instance.defaults.headers.common["Client-Source"] = "teamstoolkit";
    instance.defaults.headers.common["Accept"] = "application/json;odata=nometadata";
    instance.interceptors.request.use((config) => {
      config.params = { teamstoolkit: true, ...config.params };
      return config;
    });
    return instance;
  }

  public async getAppCatalogSite(spoToken: string): Promise<string | undefined> {
    const response = await this.createRequesterWithToken(spoToken).get<unknown>(
      "/_api/SP_TenantSettings_Current"
    );
    const url = unwrap(response.data).CorporateCatalogUrl;
    return typeof url === "string" ? url : undefined;
  }

  public async uploadAppPackage(
    spoToken: string,
    fileName: string,
    file: Buffer,
    overwrite = true
  ): Promise<Record<string, unknown>> {
    const response = await this.createRequesterWithToken(spoToken).post<unknown>(
      `/_api/web/tenantappcatalog/Add(overwrite=${overwrite},url='${odataLiteral(fileName)}')`,
      file,
      {
        maxBodyLength: Infinity,
        headers: { "Content-Type": "application/octet-stream" },
      }
    );
    return unwrap(response.data);
  }

  public async getApp(spoToken: string, id: string): Promise<Record<string, unknown> | undefined> {
    try {
      const response = await this.createRequesterWithToken(spoToken).get<unknown>(
        `/_api/web/tenantappcatalog/AvailableApps/GetById('${odataLiteral(id)}')`
      );
      return unwrap(response.data);
    } catch (error) {
      if (axios.isAxiosError(error) && error.response?.status === 404) return undefined;
      throw error;
    }
  }

  public async getPackageName(spoToken: string, id: string): Promise<unknown> {
    const response = await this.createRequesterWithToken(spoToken).get<unknown>(
      `/_api/web/GetFileById('${odataLiteral(id)}')?$select=Name`
    );
    return unwrap(response.data).Name;
  }

  public async downloadPackage(spoToken: string, id: string): Promise<Buffer> {
    const response = await this.createRequesterWithToken(spoToken).get<Buffer>(
      `/_api/web/GetFileById('${odataLiteral(id)}')/$value`,
      { responseType: "arraybuffer" }
    );
    return Buffer.from(response.data);
  }

  public async deployAppPackage(
    spoToken: string,
    appId: string,
    skipFeatureDeployment = true
  ): Promise<void> {
    await this.createRequesterWithToken(spoToken).post(
      `/_api/web/tenantappcatalog/AvailableApps/GetById('${odataLiteral(appId)}')/Deploy`,
      { skipFeatureDeployment }
    );
  }

  public async createAppCatalog(spoToken: string): Promise<void> {
    await this.createRequesterWithToken(spoToken).post(
      "/_api/web/EnsureTenantAppCatalog(callerId='teamsdev')"
    );
  }
}
