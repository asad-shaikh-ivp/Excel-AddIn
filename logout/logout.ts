/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as msal from 'msal';

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {

    //@ts-ignore
    const postLogoutRedirectUri=`${ADDIN_URL}/logoutcomplete/logoutcomplete.html`;
    //@ts-ignore
    const redirectUri = `${ADDIN_URL}/logoutcomplete/logoutcomplete.html`;

    const config: msal.Configuration = {
      auth: {
        //@ts-ignore
        clientId: AD_AUTH_CLIENT_ID,
        //@ts-ignore
        redirectUri: redirectUri,
        postLogoutRedirectUri: postLogoutRedirectUri

      }
    };

    const userAgentApplication = new msal.UserAgentApplication(config);
    userAgentApplication.logout();

    localStorage.clear();
    sessionStorage.clear();

  };
})();
