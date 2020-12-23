/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as msal from 'msal';


(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {
    //@ts-ignore
    const redirectUri=`${ADDIN_URL}/login/login.html`;

    console.log("redirectUri"+redirectUri);

    const config: msal.Configuration = {

      auth: {
        // @ts-ignore
        clientId: AD_AUTH_CLIENT_ID,

        // @ts-ignore
        authority: `https://login.microsoftonline.com/${AD_AUTH_TENANT_ID}`,

        redirectUri: redirectUri

      },
      cache: {
        cacheLocation: 'localStorage', // needed to avoid "login required" error
        storeAuthStateInCookie: true   // recommended to avoid certain IE/Edge issues
      }
    };

    const userAgentApp: msal.UserAgentApplication = new msal.UserAgentApplication(config);

    const authCallback = (error: msal.AuthError, response: msal.AuthResponse) => {

      if (!error) {
        if (response.tokenType === 'id_token') {
         localStorage.setItem('loggedIn', 'yes');
         let tokenSendStatus = setTokenValueInFunctionContext();
         console.log(tokenSendStatus);
        }
        else {
          // The tokenType is access_token, so send success message and token.
          Office.context.ui.messageParent( JSON.stringify({ status: 'success', result : response.accessToken, userName: response.idTokenClaims.preferred_username}) );
        }
      }
      else {
        const errorData: string = `errorMessage: ${error.errorCode}
                                   message: ${error.errorMessage}
                                   errorCode: ${error.stack}`;
        Office.context.ui.messageParent( JSON.stringify({ status: 'failure', result: errorData }));
      }
    };

    userAgentApp.handleRedirectCallback(authCallback);

    const request: msal.AuthenticationParameters = {
       //@ts-ignore
       scopes: [`${AD_AUTH_CLIENT_ID}/.default`]
    };

    if (localStorage.getItem('loggedIn') === 'yes') {
      let tokenSendStatus = setTokenValueInFunctionContext();
      console.log(tokenSendStatus);
      userAgentApp.acquireTokenRedirect(request);
    }
    else {
        // This will login the user and then the (response.tokenType === "id_token")
        // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
        // and then the dialog is redirected back to this script, so the
        // acquireTokenRedirect above runs.
        userAgentApp.loginRedirect(request);
    }
  };
})();
// token_handling
function setTokenValueInFunctionContext() {
  // const key = 'token';
  // let token = localStorage.getItem('msal.idtoken');
  // let tokenSendStatus = '';
  // OfficeRuntime.storage.removeItem('token').then(() => {
  // // login successful
  // // @ts-ignore
  // OfficeRuntime.storage.setItem(key, token).then(() => {
  //   tokenSendStatus = 'Success: Item with key \'' + key + '\' saved to Storage.';
  //   return tokenSendStatus;
  // }, (error) => {
  //   tokenSendStatus = 'Error: Unable to save item with key \'' + key + '\' to Storage. ' + error;
  //   return tokenSendStatus;
  // });

  // });
}

