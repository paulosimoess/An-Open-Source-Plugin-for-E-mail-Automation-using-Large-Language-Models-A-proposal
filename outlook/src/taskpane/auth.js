import { PublicClientApplication } from "@azure/msal-browser";
import { MSAL_CLIENT_ID } from "./msal.local";

const graphScopes = ["User.Read", "Mail.Read"];

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: MSAL_CLIENT_ID,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://localhost:3000/auth.html",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
});

function sendMessageToParent(message) {
  Office.onReady(() => {
    Office.context.ui.messageParent(JSON.stringify(message));
  });
}

async function startAuth() {
  try {
    await msalInstance.initialize();

    const redirectResult = await msalInstance.handleRedirectPromise();

    if (redirectResult && redirectResult.accessToken) {
      sendMessageToParent({
        type: "AUTH_SUCCESS",
        accessToken: redirectResult.accessToken,
        username:
          redirectResult.account?.username ||
          redirectResult.account?.name ||
          "Conta autenticada",
      });
      return;
    }

    const accounts = msalInstance.getAllAccounts();

    if (accounts.length > 0) {
      const tokenResult = await msalInstance.acquireTokenSilent({
        scopes: graphScopes,
        account: accounts[0],
      });

      sendMessageToParent({
        type: "AUTH_SUCCESS",
        accessToken: tokenResult.accessToken,
        username:
          tokenResult.account?.username ||
          tokenResult.account?.name ||
          "Conta autenticada",
      });
      return;
    }

    await msalInstance.loginRedirect({
      scopes: graphScopes,
      prompt: "select_account",
    });
  } catch (error) {
    sendMessageToParent({
      type: "AUTH_ERROR",
      message: error.message || "Erro desconhecido na autenticação Microsoft.",
    });
  }
}

startAuth();