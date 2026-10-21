// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosError, CanceledError } from "axios";
import { describe, expect, it } from "vitest";
import { PublishRemoteError, remotePublishError } from "../../src/error/publish";

const requestId = "44444444-4444-4444-8444-4444444444dd";

describe("publication diagnostics", () => {
  it.each(["Forbidden", "Authorization_RequestDenied", "InvalidAuthenticationToken"])(
    "PUB-07: preserves the standard Graph service code %s",
    (serviceCode) => {
      const error = remotePublishError(
        {
          response: {
            status: 403,
            data: { error: { code: serviceCode, message: "Access denied." } },
          },
        },
        "agentSubmit",
        "create"
      );
      expect(error).toMatchObject({ diagnostic: { serviceCode } });
    }
  );
  it("PUB-07: preserves whitelisted Graph details but never raw request credentials", () => {
    const error = remotePublishError(
      {
        config: { headers: { Authorization: "Bearer header-canary" }, data: "private-package" },
        request: { authorization: "request-canary" },
        response: {
          status: 403,
          headers: { "request-id": requestId, "set-cookie": "cookie-canary" },
          data: {
            error: {
              code: "Forbidden",
              message:
                "Submission is forbidden. https://login.example/authorize?code=url-canary " +
                "Bearer bearer-canary access_token=access-canary authorization_code=auth-canary " +
                "refresh_token=refresh-canary client_secret=secret-canary",
              innerError: { details: [{ message: "The service denied this operation." }] },
            },
          },
        },
      },
      "agentSubmit",
      "create"
    );
    expect(error).toBeInstanceOf(PublishRemoteError);
    if (!(error instanceof PublishRemoteError)) return;
    expect(error.diagnostic).toMatchObject({
      target: "agentSubmit",
      operation: "create",
      status: 403,
      serviceCode: "Forbidden",
      requestId,
    });
    expect(error.message).toContain("organization catalog");
    expect(error.message).toContain("The service denied this operation.");
    expect(error.message).not.toContain("SharePoint consent");
    for (const secret of [
      "header-canary",
      "private-package",
      "request-canary",
      "cookie-canary",
      "url-canary",
      "bearer-canary",
      "access-canary",
      "auth-canary",
      "refresh-canary",
      "secret-canary",
    ]) {
      expect(JSON.stringify(error)).not.toContain(secret);
    }
    expect(error.innerError).toBeUndefined();
  });

  it("PUB-07: timeout is not converted into an HTTP rejection or SharePoint advice", () => {
    const error = remotePublishError(
      new AxiosError("connect timed out", "ETIMEDOUT"),
      "agentSubmit",
      "update"
    );
    expect(error).toBeInstanceOf(PublishRemoteError);
    if (!(error instanceof PublishRemoteError)) return;
    expect(error.status).toBeUndefined();
    expect(error.diagnostic.transportCode).toBe("ETIMEDOUT");
    expect(error.message).toContain("outcome is unknown");
    expect(error.message).toContain("Reconcile");
    expect(error.message).not.toContain("SharePoint consent");
  });

  it("PUB-07: innerError request IDs survive while HTML and invalid diagnostic fields do not", () => {
    const graph = remotePublishError(
      {
        response: {
          status: 403,
          data: {
            error: {
              code: "Forbidden",
              message: "Forbidden",
              innerError: { "request-id": requestId },
            },
          },
        },
      },
      "agentCatalogRead"
    );
    expect(graph).toMatchObject({ diagnostic: { target: "agentCatalogRead", requestId } });
    const html = remotePublishError(
      {
        response: {
          status: 502,
          data: "<html>private-service-body</html>",
          headers: { "request-id": "not-a-request-id\nBearer private-header" },
        },
      },
      "agentSubmit"
    );
    expect(html.message).not.toContain("private-service-body");
    expect(JSON.stringify(html)).not.toContain("private-header");
  });

  it("PUB-07: bounds and single-lines untrusted service messages", () => {
    const error = remotePublishError(
      {
        response: {
          status: 400,
          data: {
            error: {
              code: "InvalidRequest",
              message: "x".repeat(2000) + "\r\n\u001b[31m",
            },
          },
        },
      },
      "agentSubmit"
    );
    if (!(error instanceof PublishRemoteError)) return expect.fail("Expected structured error");
    expect(error.diagnostic.serviceMessage?.length).toBeLessThanOrEqual(1024);
    expect(error.diagnostic.serviceMessage).not.toMatch(/[\r\n\u001b]/);
  });

  it("PUB-07: cancellation remains cancellation, not a permission failure", () => {
    expect(remotePublishError(new CanceledError(), "agentSubmit").name).toBe("UserCancel");
  });
});
