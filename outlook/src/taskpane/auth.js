import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

import { MSAL_CLIENT_ID } from "./msal.local";

const graphScopes = ["User.Read", "Mail.Read", "Mail.ReadWrite", "MailboxSettings.ReadWrite", ];

const urlParams = new URLSearchParams(window.location.search);
const authMode = urlParams.get("mode") || "login";

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: MSAL_CLIENT_ID,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://localhost:3000/auth.html",
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
});

function sendMessageToParent(message) {
  Office.onReady(() => {
    Office.context.ui.messageParent(JSON.stringify(message));
  });
}

function getAccountName(account) {
  return account?.username || account?.name || "Conta autenticada";
}

async function sendTokenFromAccount(account) {
  const tokenResult = await msalInstance.acquireTokenSilent({
    scopes: graphScopes,
    account,
  });

  sendMessageToParent({
    type: "AUTH_SUCCESS",
    accessToken: tokenResult.accessToken,
    username: getAccountName(tokenResult.account || account),
  });
}

async function startAuth() {
  try {
    await msalInstance.initialize();

    const redirectResult = await msalInstance.handleRedirectPromise();

    if (redirectResult?.accessToken) {
      sendMessageToParent({
        type: "AUTH_SUCCESS",
        accessToken: redirectResult.accessToken,
        username: getAccountName(redirectResult.account),
      });
      return;
    }

    const accounts = msalInstance.getAllAccounts();

    if (accounts.length > 0) {
      try {
        await sendTokenFromAccount(accounts[0]);
        return;
      } catch (error) {
        if (
          authMode === "restore" ||
          error instanceof InteractionRequiredAuthError ||
          error.name === "InteractionRequiredAuthError"
        ) {
          sendMessageToParent({
            type: "AUTH_NO_SESSION",
            message: "Sessão Microsoft expirada ou sem token silencioso disponível.",
          });
          return;
        }

        throw error;
      }
    }

    if (authMode === "restore") {
      sendMessageToParent({
        type: "AUTH_NO_SESSION",
        message: "Não existe sessão Microsoft guardada.",
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