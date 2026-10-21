// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../../../common/localizeUtils";

export function packageError(reason: string): UserError {
  return new UserError({
    source: "SharePointPublish",
    name: `SharePoint${reason}`,
    message: getLocalizedString(`publish.sharepoint.${reason}`),
  });
}
