import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

import { MSAL_CLIENT_ID } from "./msal.local";

const DEFAULT_NAA_GRAPH_SCOPES = [
  "User.Read",
  "Mail.Read",
  "Mail.ReadWrite",
  "MailboxSettings.ReadWrite",
];

let naaMsalInstance = null;

export function isNaaSupported() {
  try {
    return Boolean(
      Office?.context?.requirements?.isSetSupported?.("NestedAppAuth", "1.1")
    );
  } catch (error) {
    console.warn("Não foi possível verificar suporte NAA:", error);
    return false;
  }
}

function validateNaaConfig() {
  if (!MSAL_CLIENT_ID || MSAL_CLIENT_ID === "COLOCAR_CLIENT_ID_AQUI") {
    throw new Error("Falta configurar o MSAL_CLIENT_ID no ficheiro msal.local.js.");
  }

  if (typeof createNestablePublicClientApplication !== "function") {
    throw new Error(
      "A função createNestablePublicClientApplication não está disponível nesta versão do @azure/msal-browser."
    );
  }
}

async function initNaaMsal() {
  validateNaaConfig();

  if (naaMsalInstance) {
    return naaMsalInstance;
  }

  naaMsalInstance = await createNestablePublicClientApplication({
    auth: {
      clientId: MSAL_CLIENT_ID,
      authority: "https://login.microsoftonline.com/common",
    },
    cache: {
      cacheLocation: "localStorage",
    },
  });

  return naaMsalInstance;
}

function getOfficeLoginHint() {
  try {
    return (
      Office?.context?.mailbox?.userProfile?.emailAddress ||
      Office?.context?.mailbox?.userProfile?.displayName ||
      ""
    )
      .trim()
      .toLowerCase();
  } catch {
    return "";
  }
}

function isInteractionRequired(error) {
  return (
    error instanceof InteractionRequiredAuthError ||
    error?.name === "InteractionRequiredAuthError" ||
    error?.errorCode === "interaction_required" ||
    error?.errorCode === "consent_required" ||
    error?.errorCode === "login_required"
  );
}

function getAccountName(account, profile) {
  return (
    profile?.mail ||
    profile?.userPrincipalName ||
    account?.username ||
    account?.name ||
    "Conta autenticada"
  );
}

export async function acquireNaaAccessToken(
  scopes = DEFAULT_NAA_GRAPH_SCOPES,
  options = {}
) {
  const { interactive = true } = options;

  if (!isNaaSupported()) {
    throw new Error("Nested App Authentication não está suportado neste cliente Outlook.");
  }

  const msalInstance = await initNaaMsal();
  const accounts = msalInstance.getAllAccounts?.() || [];
  const loginHint = getOfficeLoginHint();

  const tokenRequest = {
    scopes,
    loginHint: loginHint || undefined,
  };

  if (accounts.length > 0) {
    tokenRequest.account =
      accounts.find((account) => {
        const username = String(account.username || "").trim().toLowerCase();
        const name = String(account.name || "").trim().toLowerCase();

        return loginHint && (username === loginHint || name === loginHint);
      }) || accounts[0];
  }

  try {
    const silentResult = await msalInstance.acquireTokenSilent(tokenRequest);
    return silentResult;
  } catch (silentError) {
    console.warn("NAA acquireTokenSilent falhou:", silentError);

    if (!isInteractionRequired(silentError)) {
      throw silentError;
    }

    if (!interactive) {
      throw new Error("NAA requer interação do utilizador.");
    }

    const popupResult = await msalInstance.acquireTokenPopup(tokenRequest);
    return popupResult;
  }
}

export async function acquireNaaGraphSession(
  scopes = DEFAULT_NAA_GRAPH_SCOPES,
  options = {}
) {
  const tokenResult = await acquireNaaAccessToken(scopes, options);
  const accessToken = tokenResult.accessToken;

  if (!accessToken) {
    throw new Error("NAA não devolveu access token.");
  }

  const response = await fetch(
    "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName",
    {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText || "Não foi possível obter o perfil Microsoft Graph.");
  }

  const profile = await response.json();

  return {
    accessToken,
    account: tokenResult.account,
    profile,
    username: getAccountName(tokenResult.account, profile),
  };
}