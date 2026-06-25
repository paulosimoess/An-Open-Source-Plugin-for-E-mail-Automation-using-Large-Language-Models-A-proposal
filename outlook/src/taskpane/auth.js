import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

import { MSAL_CLIENT_ID } from "./msal.local";

const graphScopes = ["User.Read", "Mail.Read", "Mail.ReadWrite", "MailboxSettings.ReadWrite", ];

const urlParams = new URLSearchParams(window.location.search);
const authMode = urlParams.get("mode") || "login";
const loginHint = (urlParams.get("login_hint") || "").trim().toLowerCase();

function normalizeAccountEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function getPreferredAccount(accounts) {
  if (!Array.isArray(accounts) || !accounts.length) {
    return null;
  }

  if (!loginHint) {
    return accounts[0];
  }

  return (
    accounts.find(
      (account) =>
        normalizeAccountEmail(account.username) === loginHint ||
        normalizeAccountEmail(account.name) === loginHint
    ) || accounts[0]
  );
}

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: MSAL_CLIENT_ID,
    authority: "https://login.microsoftonline.com/consumers",
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
        const preferredAccount = getPreferredAccount(accounts);
        await sendTokenFromAccount(preferredAccount);
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
      loginHint: loginHint || undefined,
    });
  } catch (error) {
    sendMessageToParent({
      type: "AUTH_ERROR",
      message: error.message || "Erro desconhecido na autenticação Microsoft.",
    });
  }
}

startAuth();