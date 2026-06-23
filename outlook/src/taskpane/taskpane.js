/* global document, Office */

import {
  categorizeEmail,
  requestRagResponse,
  getLatestResponse,
  validateResponse,
  getEmailState,
  getCategories,
  getCategoryKeywords,
  addCategoryKeyword,
  deleteCategoryKeyword,
  createCategory,
} from "./api";

import { MSAL_CLIENT_ID } from "./msal.local";

let graphAccessToken = null;
let graphAccountUsername = null;
let lastGraphMessages = [];
let currentEmailData = null;
let lastResponseId = null;
let categoriesCache = [];
let selectedCategory = null;

const GRAPH_STATE_CACHE_KEY = "ai4ap_graph_email_states";
const GRAPH_AUTH_CACHE_KEY = "ai4ap_graph_auth_session_v3";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    hideElement("sideload-msg");
    showElement("app-body", "block");

    const runButton = document.getElementById("run");
    if (runButton) runButton.onclick = run;

    const categorizeButton = document.getElementById("categorize-button");
    if (categorizeButton) categorizeButton.onclick = handleCategorize;

    const generateResponseButton = document.getElementById("generate-response-button");
    if (generateResponseButton) generateResponseButton.onclick = handleGenerateResponse;

    const refreshResponseButton = document.getElementById("refresh-response-button");
    if (refreshResponseButton) refreshResponseButton.onclick = handleRefreshResponse;

    const validateResponseButton = document.getElementById("validate-response-button");
    if (validateResponseButton) validateResponseButton.onclick = handleValidateResponse;

    const openReplyButton = document.getElementById("open-reply-button");
    if (openReplyButton) openReplyButton.onclick = handleOpenReplyInOutlook;

    const graphLoginButton = document.getElementById("graph-login-button");
    if (graphLoginButton) graphLoginButton.onclick = handleMicrosoftLogin;

    try {
      loadGraphAuthSession();
    } catch (error) {
      console.warn("Não foi possível recuperar automaticamente a sessão Microsoft:", error);
    }

    const graphReadInboxButton = document.getElementById("graph-read-inbox-button");
    if (graphReadInboxButton) graphReadInboxButton.onclick = handleReadInboxWithGraph;

    const graphProcessInboxButton = document.getElementById("graph-process-inbox-button");
    if (graphProcessInboxButton) graphProcessInboxButton.onclick = handleProcessInboxWithGraph;

    const graphProcessUnreadButton = document.getElementById("graph-process-unread-button");
    if (graphProcessUnreadButton) graphProcessUnreadButton.onclick = handleProcessUnreadInboxWithGraph;

    const filterCategoryButton = document.getElementById("filter-category-button");
    if (filterCategoryButton) filterCategoryButton.onclick = handleFilterEmailsByCategory;

    const loadCategoriesButton = document.getElementById("load-categories-button");
    if (loadCategoriesButton) loadCategoriesButton.onclick = handleLoadCategories;

    const addKeywordButton = document.getElementById("add-keyword-button");
    if (addKeywordButton) addKeywordButton.onclick = handleAddKeyword;

    const createCategoryButton = document.getElementById("create-category-button");
    if (createCategoryButton) createCategoryButton.onclick = handleCreateCategory;

    const keywordsList = document.getElementById("keywords-list");
    if (keywordsList) {
      keywordsList.onclick = (event) => {
        const removeButton = event.target.closest(".keyword-remove");

        if (!removeButton) return;

        const keywordId = removeButton.getAttribute("data-keyword-id");
        handleDeleteKeyword(keywordId);
      };
    }

    run();

    if (!loadGraphAuthSession()) {
      updateGraphAccountUi();
    }
  }
});

function hideElement(id) {
  const el = document.getElementById(id);
  if (el) el.style.display = "none";
}

function showElement(id, displayValue = "block") {
  const el = document.getElementById(id);
  if (el) el.style.display = displayValue;
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) {
    el.textContent = value ?? "";
  }
}

function setHtml(id, value) {
  const el = document.getElementById(id);
  if (el) {
    el.innerHTML = value ?? "";
  }
}

function escapeHtml(text) {
  if (text === null || text === undefined) return "";

  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function normalizeBodyToHtml(bodyText) {
  if (!bodyText) return "";
  return escapeHtml(bodyText).replace(/\n/g, "<br>");
}

function getBodyAsync(item) {
  return new Promise((resolve, reject) => {
    if (!item?.body) {
      reject(new Error("O item atual não suporta leitura do corpo do email."));
      return;
    }

    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        reject(
          new Error(
            result.error?.message || "Não foi possível obter o corpo do email."
          )
        );
      }
    });
  });
}

function getFromDisplay(item) {
  const from = item?.from;

  if (!from) return "N/A";

  const displayName = from.displayName || "";
  const emailAddress = from.emailAddress || "";

  if (displayName && emailAddress) {
    return `${displayName} <${emailAddress}>`;
  }

  return displayName || emailAddress || "N/A";
}

function getFromEmail(item) {
  return item?.from?.emailAddress || "";
}

function getConversationId(item) {
  return item?.conversationId || item?.itemId || "N/A";
}

function getItemId(item) {
  return item?.itemId || "N/A";
}

function setStatus(message, isError = false) {
  const statusEl = document.getElementById("status-message");
  if (!statusEl) return;

  statusEl.textContent = message;
  statusEl.style.color = isError ? "#b00020" : "#333";
}

function setGraphStatus(message, isError = false) {
  const graphStatusEl = document.getElementById("graph-status");
  if (!graphStatusEl) return;

  graphStatusEl.textContent = message;
  graphStatusEl.style.color = isError ? "#b00020" : "#333";
}

function resetCategorizationUi() {
  setText("suggested-category", "Ainda não categorizado");
  setText("used-keywords", "-");
}

function resetResponseUi() {
  lastResponseId = null;
  setText("response-job-status", "Sem pedido enviado");
  setText("response-id", "-");
  setText("response-content", "Nenhuma resposta carregada.");
}

function normalizeCacheText(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function normalizeEmailAddress(value) {
  const text = String(value || "").trim().toLowerCase();

  const match = text.match(/<([^>]+)>/);
  if (match && match[1]) {
    return match[1].trim().toLowerCase();
  }

  return text;
}

function getEmailCacheKeys(emailData) {
  const keys = [];

  if (emailData?.message_id) {
    keys.push(`message:${emailData.message_id}`);
  }

  if (emailData?.thread_id) {
    keys.push(`thread:${emailData.thread_id}`);
  }

  const subject = normalizeCacheText(emailData?.assunto);
  const sender = normalizeEmailAddress(emailData?.remetente);

  if (subject && sender) {
    keys.push(`subject_sender:${subject}|${sender}`);
  }

  return keys;
}

function getGraphStateCache() {
  try {
    const rawCache = sessionStorage.getItem(GRAPH_STATE_CACHE_KEY);
    return rawCache ? JSON.parse(rawCache) : {};
  } catch (error) {
    console.warn("Erro ao ler cache de categorizações Graph:", error);
    return {};
  }
}

function saveGraphCategorizationState(emailData, categoria, keywordsUsadas) {
  const keys = getEmailCacheKeys(emailData);

  if (!keys.length) return;

  const cache = getGraphStateCache();

  const state = {
    message_id: emailData.message_id,
    thread_id: emailData.thread_id,
    remetente: emailData.remetente,
    assunto: emailData.assunto,
    categoria,
    keywords_usadas: keywordsUsadas || [],
    processed_at: new Date().toISOString(),
  };

  keys.forEach((key) => {
    cache[key] = state;
  });

  sessionStorage.setItem(GRAPH_STATE_CACHE_KEY, JSON.stringify(cache));
}

function getCachedGraphCategorization(emailData) {
  const cache = getGraphStateCache();
  const keys = getEmailCacheKeys(emailData);

  for (const key of keys) {
    if (cache[key]) {
      return cache[key];
    }
  }

  return null;
}

function applyCategorizationToCurrentEmail(categoria, keywordsUsadas) {
  setText("suggested-category", categoria || "Sem categoria");
  setText(
    "used-keywords",
    keywordsUsadas?.length ? keywordsUsadas.join(", ") : "Nenhuma keyword usada"
  );
}

async function hydrateEmailState() {
  if (!currentEmailData) return;

  try {
    const state = await getEmailState(currentEmailData);
    console.log("Estado atual do email:", state);

    let hasBackendCategory = false;

    if (state?.categoria) {
      hasBackendCategory = true;

      applyCategorizationToCurrentEmail(
        state.categoria,
        state.keywords_usadas || []
      );
    }

    if (!hasBackendCategory) {
      const cachedCategorization = getCachedGraphCategorization(currentEmailData);

      if (cachedCategorization?.categoria) {
        applyCategorizationToCurrentEmail(
          cachedCategorization.categoria,
          cachedCategorization.keywords_usadas || []
        );

        console.log(
          "Categoria recuperada da cache Graph por conversationId:",
          cachedCategorization
        );
      }
    }

    if (state?.resposta) {
      lastResponseId = state.resposta.id_resposta || null;
      setText("response-id", state.resposta.id_resposta || "-");
      setText("response-job-status", state.resposta.status || "Sem estado");
      setText(
        "response-content",
        state.resposta.conteudo && state.resposta.conteudo.trim()
          ? state.resposta.conteudo
          : "Resposta ainda sem conteúdo disponível."
      );
    }
  } catch (error) {
    console.error("Erro ao hidratar estado do email:", error);

    const cachedCategorization = getCachedGraphCategorization(currentEmailData);

    if (cachedCategorization?.categoria) {
      applyCategorizationToCurrentEmail(
        cachedCategorization.categoria,
        cachedCategorization.keywords_usadas || []
      );
    }
  }
}

export async function run() {
  try {
    setStatus("A carregar dados do email...");

    const item = Office.context.mailbox.item;

    if (!item) {
      throw new Error("Nenhum email está aberto no Outlook.");
    }

    const subject = item.subject || "(Sem assunto)";
    const from = getFromDisplay(item);
    const fromEmail = getFromEmail(item);
    const itemId = getItemId(item);
    const conversationId = getConversationId(item);
    const body = await getBodyAsync(item);

    currentEmailData = {
      message_id: itemId,
      thread_id: conversationId,
      remetente: normalizeEmailAddress(fromEmail || from),
      assunto: subject,
      corpo: body,
    };

    setText("item-subject", subject);
    setText("item-from", from);
    setText("item-id", itemId);
    setText("item-conversation-id", conversationId);
    setHtml("item-body", normalizeBodyToHtml(body));

    resetCategorizationUi();
    resetResponseUi();
    await hydrateEmailState();

    setStatus("Dados do email carregados com sucesso.");
  } catch (error) {
    console.error("Erro ao carregar dados do email:", error);
    setStatus(`Erro: ${error.message}`, true);

    setText("item-subject", "N/A");
    setText("item-from", "N/A");
    setText("item-id", "N/A");
    setText("item-conversation-id", "N/A");
    setHtml("item-body", "Não foi possível carregar o conteúdo do email.");

    resetCategorizationUi();
    resetResponseUi();
  }
}

async function handleCategorize() {
  try {
    if (!currentEmailData) {
      throw new Error("Ainda não existem dados do email carregados.");
    }

    setStatus("A categorizar email...");

    const result = await categorizeEmail(currentEmailData);
    console.log("Resultado da categorização:", result);

    const categoria = result.categoria || "Sem categoria";
    const keywordsUsadas = result.keywords_usadas || [];

    applyCategorizationToCurrentEmail(categoria, keywordsUsadas);
    saveGraphCategorizationState(currentEmailData, categoria, keywordsUsadas);

    try {
      await applyCategoryAndMarkReadToOpenEmail(categoria);
      setStatus("Email categorizado, categoria aplicada no Outlook e marcado como lido.");
    } catch (outlookError) {
      console.warn("Categoria guardada no backend, mas não aplicada no Outlook:", outlookError);
      setStatus(
        `Email categorizado no plugin, mas não foi possível atualizar o Outlook: ${outlookError.message}`,
        true
      );
    }
  } catch (error) {
    console.error("Erro ao categorizar email:", error);
    setStatus(`Erro ao categorizar: ${error.message}`, true);
    setText("suggested-category", "Erro");
    setText("used-keywords", "-");
  }
}

async function handleGenerateResponse() {
  try {
    if (!currentEmailData) {
      throw new Error("Ainda não existem dados do email carregados.");
    }

    setStatus("A pedir geração de resposta...");
    setText("response-job-status", "PENDING");
    setText("response-content", "Pedido enviado. Aguarde e depois clique em 'Atualizar resposta'.");

    const result = await requestRagResponse(currentEmailData);
    console.log("Pedido de resposta gerada:", result);

    lastResponseId = result.job_id || null;
    setText("response-id", result.job_id || "-");
    setText("response-job-status", result.status || "queued");

    setStatus("Pedido de geração de resposta enviado com sucesso.");
  } catch (error) {
    console.error("Erro ao gerar resposta:", error);
    setStatus(`Erro ao gerar resposta: ${error.message}`, true);
    setText("response-job-status", "Erro");
  }
}

async function handleRefreshResponse() {
  try {
    if (!currentEmailData) {
      throw new Error("Ainda não existem dados do email carregados.");
    }

    setStatus("A atualizar resposta gerada...");

    const result = await getLatestResponse(currentEmailData);
    console.log("Última resposta gerada:", result);

    lastResponseId = result.id_resposta || null;
    setText("response-id", result.id_resposta || "-");
    setText("response-job-status", result.status || "Sem estado");
    setText(
      "response-content",
      result.conteudo && result.conteudo.trim()
        ? result.conteudo
        : "Resposta ainda sem conteúdo disponível."
    );

    setStatus("Resposta atualizada com sucesso.");
  } catch (error) {
    console.error("Erro ao atualizar resposta:", error);
    setStatus(`Erro ao atualizar resposta: ${error.message}`, true);
  }
}

async function handleValidateResponse() {
  try {
    if (!currentEmailData) {
      throw new Error("Ainda não existem dados do email carregados.");
    }

    if (!lastResponseId) {
      throw new Error("Ainda não existe uma resposta carregada para validar.");
    }

    setStatus("A validar resposta...");

    const result = await validateResponse(currentEmailData, lastResponseId);
    console.log("Resposta validada:", result);

    setText("response-job-status", "VALIDADA");
    setText("response-content", result.resposta || "Resposta validada.");

    setStatus("Resposta validada com sucesso.");
  } catch (error) {
    console.error("Erro ao validar resposta:", error);
    setStatus(`Erro ao validar resposta: ${error.message}`, true);
  }
}

function validateMsalConfig() {
  if (!MSAL_CLIENT_ID || MSAL_CLIENT_ID === "COLOCAR_CLIENT_ID_AQUI") {
    throw new Error("Falta configurar o MSAL_CLIENT_ID no ficheiro msal.local.js.");
  }
}

function decodeJwtPayload(token) {
  try {
    const payload = token.split(".")[1];

    if (!payload) return null;

    const base64 = payload.replace(/-/g, "+").replace(/_/g, "/");
    const paddedBase64 = base64.padEnd(
      base64.length + ((4 - (base64.length % 4)) % 4),
      "="
    );

    return JSON.parse(atob(paddedBase64));
  } catch (error) {
    console.warn("Não foi possível ler expiração do token:", error);
    return null;
  }
}

function getAccessTokenExpiry(accessToken) {
  const payload = decodeJwtPayload(accessToken);

  if (payload?.exp) {
    return payload.exp * 1000;
  }

  return Date.now() + 50 * 60 * 1000;
}

function saveGraphAuthSession(accessToken, username) {
  if (!accessToken) return;

  const session = {
    accessToken,
    username: username || "Conta autenticada",
    expiresAt: getAccessTokenExpiry(accessToken),
    savedAt: Date.now(),
  };

  localStorage.setItem(GRAPH_AUTH_CACHE_KEY, JSON.stringify(session));
}

function clearGraphAuthSession() {
  localStorage.removeItem(GRAPH_AUTH_CACHE_KEY);
}

function loadGraphAuthSession() {
  try {
    const rawSession = localStorage.getItem(GRAPH_AUTH_CACHE_KEY);

    if (!rawSession) {
      return false;
    }

    const session = JSON.parse(rawSession);

    if (!session.accessToken || !session.expiresAt) {
      clearGraphAuthSession();
      return false;
    }

    if (Number(session.expiresAt) <= Date.now() + 60 * 1000) {
      clearGraphAuthSession();
      return false;
    }

    graphAccessToken = session.accessToken;
    graphAccountUsername = session.username || "Conta autenticada";

    updateGraphAccountUi();
    setGraphStatus("Sessão Microsoft recuperada automaticamente.");

    return true;
  } catch (error) {
    console.warn("Erro ao recuperar sessão Microsoft local:", error);
    clearGraphAuthSession();
    return false;
  }
}

function updateGraphAccountUi() {
  if (!graphAccountUsername) {
    setText("graph-account", "-");
    setGraphStatus("Ainda não autenticado no Microsoft Graph.");
    return;
  }

  setText("graph-account", graphAccountUsername);
  setGraphStatus("Conta Microsoft autenticada.");
}

function openMicrosoftAuthDialog(mode = "login") {
  return new Promise((resolve, reject) => {
    const url = `https://localhost:3000/auth.html?mode=${mode}`;

    Office.context.ui.displayDialogAsync(
      url,
      {
        height: mode === "restore" ? 30 : 60,
        width: mode === "restore" ? 30 : 45,
        displayInIframe: false,
      },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(
            new Error(
              asyncResult.error?.message ||
                "Não foi possível abrir a janela de autenticação."
            )
          );
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          try {
            const message = JSON.parse(arg.message);

            if (message.type === "AUTH_SUCCESS") {
              graphAccessToken = message.accessToken;
              graphAccountUsername = message.username || "Conta autenticada";

              saveGraphAuthSession(graphAccessToken, graphAccountUsername);

              updateGraphAccountUi();
              setGraphStatus("Autenticação Microsoft concluída com sucesso.");

              dialog.close();
              resolve(true);
              return;
            }

            if (message.type === "AUTH_NO_SESSION") {
              graphAccessToken = null;
              graphAccountUsername = null;
              setText("graph-account", "-");

              if (mode === "restore") {
                setGraphStatus("Ainda não autenticado no Microsoft Graph.");
              } else {
                setGraphStatus(message.message || "Sessão Microsoft não encontrada.", true);
              }

              dialog.close();
              resolve(false);
              return;
            }

            if (message.type === "AUTH_ERROR") {
              graphAccessToken = null;
              graphAccountUsername = null;
              setText("graph-account", "-");

              const errorMessage = message.message || "Erro na autenticação Microsoft.";

              if (mode === "restore") {
                setGraphStatus("Ainda não autenticado no Microsoft Graph.");
                dialog.close();
                resolve(false);
                return;
              }

              setGraphStatus(`Erro no login Microsoft: ${errorMessage}`, true);
              dialog.close();
              reject(new Error(errorMessage));
            }
          } catch (error) {
            dialog.close();
            reject(error);
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          console.warn("DialogEventReceived:", arg);

          if (mode === "restore") {
            resolve(false);
          }
        });
      }
    );
  });
}

async function handleMicrosoftLogin() {
  try {
    validateMsalConfig();

    setGraphStatus("A abrir autenticação Microsoft...");

    await openMicrosoftAuthDialog("login");
  } catch (error) {
    console.error("Erro no login Microsoft:", error);
    setGraphStatus(`Erro no login Microsoft: ${error.message}`, true);
  }
}

async function getGraphAccessToken() {
  if (!graphAccessToken) {
    const restored = loadGraphAuthSession();

    if (!restored) {
      throw new Error("Ainda não existe sessão Microsoft. Clique primeiro em 'Iniciar sessão Microsoft'.");
    }
  }

  return graphAccessToken;
}

async function fetchInboxMessagesWithGraph() {
  const accessToken = await getGraphAccessToken();

  const endpoint =
    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages" +
    "?$top=50" +
    "&$select=id,subject,from,receivedDateTime,bodyPreview,body,conversationId,isRead,categories" +
    "&$orderby=receivedDateTime desc";

  const response = await fetch(endpoint, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
      Prefer: 'outlook.body-content-type="text"',
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText || "Erro ao ler emails no Microsoft Graph.");
  }

  const data = await response.json();
  return data.value || [];
}

async function fetchUnreadInboxMessagesWithGraph() {
  const accessToken = await getGraphAccessToken();

  const endpoint =
    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages" +
    "?$top=50" +
    "&$filter=isRead eq false" +
    "&$select=id,subject,from,receivedDateTime,bodyPreview,body,conversationId,isRead,categories";

  const response = await fetch(endpoint, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
      Prefer: 'outlook.body-content-type="text"',
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText || "Erro ao ler emails não lidos no Microsoft Graph.");
  }

  const data = await response.json();
  return data.value || [];
}

async function handleReadInboxWithGraph() {
  try {
    setGraphStatus("A ler emails da Inbox via Microsoft Graph...");
    setHtml("graph-inbox-list", "A carregar emails da Inbox...");

    const messages = await fetchInboxMessagesWithGraph();
    lastGraphMessages = messages;

    renderGraphInboxMessages(messages);
    updateGraphAccountUi();

    setGraphStatus(`Leitura concluída. Emails carregados: ${messages.length}.`);
  } catch (error) {
    console.error("Erro ao ler Inbox via Graph:", error);
    setGraphStatus(`Erro ao ler Inbox: ${error.message}`, true);
    setHtml("graph-inbox-list", "Não foi possível carregar os emails da Inbox.");
  }
}

function stripHtmlToText(value) {
  const text = String(value || "");

  if (!text.trim()) return "";

  const div = document.createElement("div");
  div.innerHTML = text;

  return (div.textContent || div.innerText || text).trim();
}

function getGraphMessageBodyText(message) {
  const fullBody = stripHtmlToText(message.body?.content || "");

  if (fullBody) {
    return fullBody;
  }

  const preview = stripHtmlToText(message.bodyPreview || "");

  if (preview) {
    return preview;
  }

  const subject = stripHtmlToText(message.subject || "");

  if (subject) {
    return `Corpo não disponível no Microsoft Graph. Assunto do email: ${subject}`;
  }

  return "Corpo não disponível no Microsoft Graph.";
}

function convertGraphMessageToBackendEmail(message) {
  const sender =
    message.from?.emailAddress?.address ||
    message.from?.emailAddress?.name ||
    "remetente-desconhecido";

  return {
    message_id: message.id,
    thread_id: message.conversationId || message.id,
    remetente: normalizeEmailAddress(sender) || sender,
    assunto: message.subject || "(Sem assunto)",
    corpo: getGraphMessageBodyText(message),
  };
}

async function getExistingCategorization(emailData) {
  try {
    const state = await getEmailState(emailData);

    if (state?.categoria) {
      return {
        categoria: state.categoria,
        keywords: state.keywords_usadas?.length
          ? state.keywords_usadas.join(", ")
          : "Sem keywords guardadas",
        keywordsArray: state.keywords_usadas || [],
      };
    }

    const cachedCategorization = getCachedGraphCategorization(emailData);

    if (cachedCategorization?.categoria) {
      return {
        categoria: cachedCategorization.categoria,
        keywords: cachedCategorization.keywords_usadas?.length
          ? cachedCategorization.keywords_usadas.join(", ")
          : "Sem keywords guardadas",
        keywordsArray: cachedCategorization.keywords_usadas || [],
      };
    }

    return null;
  } catch (error) {
    console.warn("Email ainda sem estado guardado no backend:", error);

    const cachedCategorization = getCachedGraphCategorization(emailData);

    if (cachedCategorization?.categoria) {
      return {
        categoria: cachedCategorization.categoria,
        keywords: cachedCategorization.keywords_usadas?.length
          ? cachedCategorization.keywords_usadas.join(", ")
          : "Sem keywords guardadas",
        keywordsArray: cachedCategorization.keywords_usadas || [],
      };
    }

    return null;
  }
}

async function handleProcessInboxWithGraph() {
  try {
    setGraphStatus("A processar emails da Inbox...");
    setHtml("graph-process-results", "A verificar emails já processados...");

    let messages = lastGraphMessages;

    if (!messages.length) {
      messages = await fetchInboxMessagesWithGraph();
      lastGraphMessages = messages;
      renderGraphInboxMessages(messages);
    }

    if (!messages.length) {
      setHtml("graph-process-results", "Nenhum email encontrado para processar.");
      setGraphStatus("Nenhum email encontrado na Inbox.");
      return;
    }

    const results = [];

    for (const message of messages) {
      const emailData = convertGraphMessageToBackendEmail(message);

      try {
        const existingCategorization = await getExistingCategorization(emailData);

        if (existingCategorization) {
          results.push({
            subject: emailData.assunto,
            from: emailData.remetente,
            status: "Já processado",
            categoria: existingCategorization.categoria,
            keywords: existingCategorization.keywords,
          });

          renderGraphProcessingResults(results);
          continue;
        }

        const result = await categorizeEmail(emailData);
        const categoria = result.categoria || "Sem categoria";
        const keywordsUsadas = result.keywords_usadas || [];

        saveGraphCategorizationState(emailData, categoria, keywordsUsadas);

        try {
          await applyCategoryAndMarkReadByGraphId(message.id, categoria);
        } catch (outlookError) {
          console.warn("Email categorizado, mas não atualizado no Outlook:", outlookError);
        }

        results.push({
          subject: emailData.assunto,
          from: emailData.remetente,
          status: "Categorizado agora",
          categoria,
          keywords: keywordsUsadas.length
            ? keywordsUsadas.join(", ")
            : "Nenhuma keyword usada",
        });

        renderGraphProcessingResults(results);
      } catch (error) {
        console.error("Erro ao processar email Graph:", error);

        results.push({
          subject: emailData.assunto,
          from: emailData.remetente,
          status: "Erro",
          categoria: "-",
          keywords: error.message || "Erro desconhecido",
        });

        renderGraphProcessingResults(results);
      }
    }

    setGraphStatus(`Processamento concluído. Emails tratados: ${results.length}.`);
  } catch (error) {
    console.error("Erro ao processar Inbox:", error);
    setGraphStatus(`Erro ao processar Inbox: ${error.message}`, true);
    setHtml("graph-process-results", "Não foi possível processar os emails da Inbox.");
  }
}

async function handleProcessUnreadInboxWithGraph() {
  try {
    setGraphStatus("A procurar emails novos/não lidos...");
    setHtml("graph-process-results", "A carregar emails não lidos da Inbox...");

    const messages = await fetchUnreadInboxMessagesWithGraph();
    lastGraphMessages = messages;

    renderGraphInboxMessages(messages);

    if (!messages.length) {
      setHtml("graph-process-results", "Nenhum email novo/não lido encontrado.");
      setGraphStatus("Nenhum email novo/não lido encontrado na Inbox.");
      return;
    }

    setGraphStatus(`Emails não lidos encontrados: ${messages.length}. A processar...`);

    const results = [];

    for (const message of messages) {
      const emailData = convertGraphMessageToBackendEmail(message);

      try {
        const existingCategorization = await getExistingCategorization(emailData);

        if (existingCategorization) {
          results.push({
            subject: emailData.assunto,
            from: emailData.remetente,
            status: "Já processado",
            categoria: existingCategorization.categoria,
            keywords: existingCategorization.keywords,
          });

          renderGraphProcessingResults(results);
          continue;
        }

        const result = await categorizeEmail(emailData);
        const categoria = result.categoria || "Sem categoria";
        const keywordsUsadas = result.keywords_usadas || [];

        saveGraphCategorizationState(emailData, categoria, keywordsUsadas);

        try {
          await applyCategoryAndMarkReadByGraphId(message.id, categoria);
        } catch (outlookError) {
          console.warn("Email categorizado, mas não atualizado no Outlook:", outlookError);
        }

        results.push({
          subject: emailData.assunto,
          from: emailData.remetente,
          status: "Categorizado agora",
          categoria,
          keywords: keywordsUsadas.length
            ? keywordsUsadas.join(", ")
            : "Nenhuma keyword usada",
        });

        renderGraphProcessingResults(results);
      } catch (error) {
        console.error("Erro ao processar email não lido:", error);

        results.push({
          subject: emailData.assunto,
          from: emailData.remetente,
          status: "Erro",
          categoria: "-",
          keywords: error.message || "Erro desconhecido",
        });

        renderGraphProcessingResults(results);
      }
    }

    setGraphStatus(`Processamento de emails novos concluído. Emails tratados: ${results.length}.`);
  } catch (error) {
    console.error("Erro ao processar emails não lidos:", error);
    setGraphStatus(`Erro ao processar emails novos: ${error.message}`, true);
    setHtml("graph-process-results", "Não foi possível processar os emails novos/não lidos.");
  }
}

function renderGraphProcessingResults(results) {
  if (!results.length) {
    setHtml("graph-process-results", "Nenhum email processado.");
    return;
  }

  const html = results
    .map((result, index) => {
      const subject = escapeHtml(result.subject || "(Sem assunto)");
      const from = escapeHtml(result.from || "Remetente desconhecido");
      const status = escapeHtml(result.status || "-");
      const categoria = escapeHtml(result.categoria || "-");
      const keywords = escapeHtml(result.keywords || "-");

      return `
        <div style="margin-bottom: 14px; padding-bottom: 10px; border-bottom: 1px solid #ddd;">
          <strong>${index + 1}. ${subject}</strong><br>
          <span><strong>De:</strong> ${from}</span><br>
          <span><strong>Estado:</strong> ${status}</span><br>
          <span><strong>Categoria:</strong> ${categoria}</span><br>
          <span><strong>Keywords:</strong> ${keywords}</span>
        </div>
      `;
    })
    .join("");

  setHtml("graph-process-results", html);
}

function renderGraphInboxMessages(messages) {
  if (!messages.length) {
    setHtml("graph-inbox-list", "Nenhum email encontrado na Inbox.");
    return;
  }

  const html = messages
    .map((message, index) => {
      const subject = escapeHtml(message.subject || "(Sem assunto)");
      const fromName = escapeHtml(message.from?.emailAddress?.name || "Remetente desconhecido");
      const fromAddress = escapeHtml(message.from?.emailAddress?.address || "");
      const received = escapeHtml(message.receivedDateTime || "");
      const preview = escapeHtml(message.bodyPreview || "");
      const conversationId = escapeHtml(message.conversationId || "");
      const readState = message.isRead ? "Lido" : "Não lido";

      return `
        <div style="margin-bottom: 14px; padding-bottom: 10px; border-bottom: 1px solid #ddd;">
          <strong>${index + 1}. ${subject}</strong><br>
          <span><strong>De:</strong> ${fromName} ${fromAddress ? `&lt;${fromAddress}&gt;` : ""}</span><br>
          <span><strong>Recebido:</strong> ${received}</span><br>
          <span><strong>Estado:</strong> ${readState}</span><br>
          <span><strong>Conversation ID:</strong> ${conversationId}</span><br>
          <p style="margin-top: 6px;">${preview || "Sem pré-visualização disponível."}</p>
        </div>
      `;
    })
    .join("");

  setHtml("graph-inbox-list", html);
}

function setCategoryStatus(message, isError = false) {
  const statusEl = document.getElementById("category-status");
  if (!statusEl) return;

  statusEl.textContent = message;
  statusEl.style.color = isError ? "#b00020" : "#333";
}

async function handleLoadCategories() {
  try {
    setCategoryStatus("A carregar categorias...");
    setHtml("categories-list", "A carregar categorias...");
    setText("selected-category", "Nenhuma categoria selecionada.");
    setHtml("keywords-list", "Nenhuma palavra-chave carregada.");

    const categories = await getCategories();

    categoriesCache = Array.isArray(categories) ? categories : [];
    selectedCategory = null;

    renderCategoriesList(categoriesCache);
    renderCategoryFilterOptions(categoriesCache);

    setCategoryStatus(`Categorias carregadas: ${categoriesCache.length}.`);
  } catch (error) {
    console.error("Erro ao carregar categorias:", error);
    setCategoryStatus(`Erro ao carregar categorias: ${error.message}`, true);
    setHtml("categories-list", "Não foi possível carregar as categorias.");
  }
}

function renderCategoriesList(categories) {
  if (!categories.length) {
    setHtml("categories-list", "Nenhuma categoria encontrada.");
    return;
  }

  const html = categories
    .map((category) => {
      const id = escapeHtml(category.id_categoria);
      const name = escapeHtml(category.nome || "Categoria sem nome");
      const description = escapeHtml(category.para_que_serve || category.questao || "");

      return `
        <div class="category-item" data-category-id="${id}">
          <div class="category-name">${name}</div>
          ${
            description
              ? `<div class="category-meta">${description}</div>`
              : `<div class="category-meta">ID: ${id}</div>`
          }
        </div>
      `;
    })
    .join("");

  setHtml("categories-list", html);

  const categoryItems = document.querySelectorAll("#categories-list .category-item");

  categoryItems.forEach((item) => {
    item.onclick = () => {
      const categoryId = item.getAttribute("data-category-id");
      handleSelectCategory(categoryId);
    };
  });
}

async function handleSelectCategory(categoryId) {
  try {
    selectedCategory = categoriesCache.find(
      (category) => String(category.id_categoria) === String(categoryId)
    );

    if (!selectedCategory) {
      throw new Error("Categoria selecionada não encontrada.");
    }

    document.querySelectorAll("#categories-list .category-item").forEach((item) => {
      item.classList.remove("selected");
    });

    const selectedElement = document.querySelector(
      `#categories-list .category-item[data-category-id="${selectedCategory.id_categoria}"]`
    );

    if (selectedElement) {
      selectedElement.classList.add("selected");
    }

    setText("selected-category", selectedCategory.nome || "Categoria sem nome");
    setHtml("keywords-list", "A carregar palavras-chave...");
    setCategoryStatus(`Categoria selecionada: ${selectedCategory.nome}`);

    await reloadSelectedCategoryKeywords();
  } catch (error) {
    console.error("Erro ao selecionar categoria:", error);
    setCategoryStatus(`Erro ao selecionar categoria: ${error.message}`, true);
    setHtml("keywords-list", "Não foi possível carregar as palavras-chave.");
  }
}

async function reloadSelectedCategoryKeywords() {
  if (!selectedCategory) {
    setHtml("keywords-list", "Seleciona primeiro uma categoria.");
    return;
  }

  const keywords = await getCategoryKeywords(selectedCategory.id_categoria);
  renderKeywordsList(keywords);
}

function getKeywordId(keyword) {
  return keyword?.id_keyword ?? keyword?.id ?? keyword?.keyword_id ?? "";
}

function renderKeywordsList(keywords) {
  if (!keywords.length) {
    setHtml("keywords-list", "Esta categoria ainda não tem palavras-chave.");
    return;
  }

  const html = keywords
    .map((keyword) => {
      const rawKeywordId = getKeywordId(keyword);
      const keywordId = escapeHtml(rawKeywordId);
      const keywordText = escapeHtml(keyword.keyword || "");

      return `
        <div class="keyword-item">
          <span>${keywordText}</span>
          <button type="button" class="keyword-remove" data-keyword-id="${keywordId}">
            Remover
          </button>
        </div>
      `;
    })
    .join("");

  setHtml("keywords-list", html);
}

async function handleAddKeyword() {
  try {
    if (!selectedCategory) {
      throw new Error("Seleciona primeiro uma categoria.");
    }

    const input = document.getElementById("new-keyword-input");
    const keyword = input?.value?.trim();

    if (!keyword) {
      throw new Error("Escreve uma palavra-chave antes de adicionar.");
    }

    setCategoryStatus(`A adicionar palavra-chave "${keyword}"...`);

    await addCategoryKeyword(selectedCategory.id_categoria, keyword);

    if (input) {
      input.value = "";
    }

    await reloadSelectedCategoryKeywords();

    setCategoryStatus(`Palavra-chave "${keyword}" adicionada com sucesso.`);
  } catch (error) {
    console.error("Erro ao adicionar palavra-chave:", error);
    setCategoryStatus(`Erro ao adicionar palavra-chave: ${error.message}`, true);
  }
}

async function handleDeleteKeyword(keywordId) {
  try {
    if (!selectedCategory) {
      throw new Error("Seleciona primeiro uma categoria.");
    }

    const cleanKeywordId = String(keywordId || "").trim();
    const cleanCategoryId = String(selectedCategory.id_categoria || "").trim();

    setCategoryStatus(
      `DEBUG remover → categoria=${cleanCategoryId}, keyword=${cleanKeywordId}`
    );

    console.log("DEBUG remover keyword:", {
      selectedCategory,
      cleanCategoryId,
      cleanKeywordId,
    });

    if (!cleanKeywordId || !/^\d+$/.test(cleanKeywordId)) {
      throw new Error(`ID da palavra-chave inválido: ${cleanKeywordId || "vazio"}`);
    }

    if (!cleanCategoryId || !/^\d+$/.test(cleanCategoryId)) {
      throw new Error(`ID da categoria inválido: ${cleanCategoryId || "vazio"}`);
    }

    await deleteCategoryKeyword(cleanCategoryId, cleanKeywordId);

    await reloadSelectedCategoryKeywords();

    setCategoryStatus(
      `Palavra-chave ${cleanKeywordId} removida com sucesso da categoria ${cleanCategoryId}.`
    );
  } catch (error) {
    console.error("Erro ao remover palavra-chave:", error);
    setCategoryStatus(`Erro ao remover palavra-chave: ${error.message}`, true);
  }
}

async function handleCreateCategory() {
  try {
    const nameInput = document.getElementById("new-category-name-input");
    const questionInput = document.getElementById("new-category-question-input");
    const purposeInput = document.getElementById("new-category-purpose-input");

    const nome = nameInput?.value?.trim();
    const questao = questionInput?.value?.trim();
    const paraQueServe = purposeInput?.value?.trim();

    if (!nome) {
      throw new Error("Indica o nome da nova categoria.");
    }

    setCategoryStatus(`A criar categoria "${nome}"...`);

    const result = await createCategory({
      nome,
      questao,
      paraQueServe,
    });

    console.log("Categoria criada:", result);

    if (nameInput) nameInput.value = "";
    if (questionInput) questionInput.value = "";
    if (purposeInput) purposeInput.value = "";

    await handleLoadCategories();

    setCategoryStatus(
      `Categoria "${nome}" criada com sucesso. Keywords geradas: ${
        result.keywordsGeradas?.length ? result.keywordsGeradas.join(", ") : "nenhuma"
      }.`
    );
  } catch (error) {
    console.error("Erro ao criar categoria:", error);
    setCategoryStatus(`Erro ao criar categoria: ${error.message}`, true);
  }
}

function getResponseContentText() {
  const responseEl = document.getElementById("response-content");
  return responseEl?.textContent?.trim() || "";
}

function cleanGeneratedResponseText(responseText) {
  return String(responseText || "")
    .replace(/^Saudação\s*/gim, "")
    .replace(/^Corpo\s*/gim, "")
    .replace(/^Assunto\s*/gim, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function buildReplyHtml(responseText) {
  const cleanedResponse = cleanGeneratedResponseText(responseText);
  const safeResponse = escapeHtml(cleanedResponse).replace(/\n/g, "<br>");

  return `
    <div>
      <p>${safeResponse}</p>
    </div>
  `;
}

async function handleOpenReplyInOutlook() {
  try {
    const item = Office.context.mailbox.item;

    if (!item) {
      throw new Error("Nenhum email está aberto no Outlook.");
    }

    const responseText = getResponseContentText();

    if (
      !responseText ||
      responseText === "Nenhuma resposta carregada." ||
      responseText === "Resposta ainda sem conteúdo disponível."
    ) {
      throw new Error("Ainda não existe uma resposta gerada para abrir no Outlook.");
    }

    setStatus("A abrir resposta no Outlook...");

    const replyHtml = buildReplyHtml(responseText);

    item.displayReplyForm({
      htmlBody: replyHtml,
      callback: (result) => {
        if (result?.status === Office.AsyncResultStatus.Failed) {
          console.error("Erro ao abrir resposta no Outlook:", result.error);
          setStatus(`Erro ao abrir resposta: ${result.error.message}`, true);
          return;
        }

        setStatus("Resposta aberta no Outlook. Revê antes de enviar.");
      },
    });
  } catch (error) {
    console.error("Erro ao abrir resposta no Outlook:", error);
    setStatus(`Erro ao abrir resposta: ${error.message}`, true);
  }
}

function cleanOutlookCategoryName(categoryName) {
  return String(categoryName || "")
    .replace(/^\./, "")
    .trim();
}

function normalizeOutlookCategoryName(categoryName) {
  return String(categoryName || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function isValidOutlookCategory(categoryName) {
  const value = String(categoryName || "").trim();

  return (
    value &&
    value !== "Ainda não categorizado" &&
    value !== "Sem categoria" &&
    value !== "Erro"
  );
}

function getGraphMessageIdFromOfficeItem() {
  const item = Office.context.mailbox.item;

  if (!item?.itemId) {
    throw new Error("Não foi possível obter o ID do email aberto.");
  }

  if (!Office.context.mailbox.convertToRestId) {
    throw new Error("Este Outlook não suporta conversão do ID para Microsoft Graph.");
  }

  return Office.context.mailbox.convertToRestId(
    item.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );
}

function getColorForOutlookCategory(categoryName) {
  const normalized = normalizeOutlookCategoryName(categoryName);

  if (normalized.includes("espacos verdes")) {
    return "preset4";
  }

  if (normalized.includes("publicidade")) {
    return "preset1";
  }

  if (normalized.includes("apoio institucional")) {
    return "preset5";
  }

  if (normalized.includes("outro")) {
    return "preset8";
  }

  return "preset2";
}

async function getOutlookMasterCategories(accessToken) {
  const response = await fetch(
    "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
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
    throw new Error(errorText || "Erro ao obter categorias principais do Outlook.");
  }

  const data = await response.json();
  return data.value || [];
}

async function createOutlookMasterCategory(accessToken, categoryName) {
  const color = getColorForOutlookCategory(categoryName);

  const response = await fetch(
    "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        displayName: categoryName,
        color,
      }),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();

    if (
      response.status === 409 ||
      errorText.includes("ErrorItemAlreadyExists") ||
      errorText.includes("already exists")
    ) {
      return null;
    }

    throw new Error(errorText || "Erro ao criar categoria no Outlook.");
  }

  return response.json();
}

async function updateOutlookMasterCategoryColor(accessToken, categoryId, categoryName) {
  const color = getColorForOutlookCategory(categoryName);

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/outlook/masterCategories/${encodeURIComponent(categoryId)}`,
    {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        color,
      }),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    console.warn("Não foi possível atualizar a cor da categoria:", errorText);
  }
}

async function ensureOutlookMasterCategoryExists(accessToken, categoryName) {
  const categories = await getOutlookMasterCategories(accessToken);
  const normalizedCategoryName = normalizeOutlookCategoryName(categoryName);

  const existingCategory = categories.find(
    (category) =>
      normalizeOutlookCategoryName(category.displayName) === normalizedCategoryName
  );

  if (existingCategory) {
    const currentColor = String(existingCategory.color || "").toLowerCase();

    if (!currentColor || currentColor === "none") {
      await updateOutlookMasterCategoryColor(
        accessToken,
        existingCategory.id,
        categoryName
      );
    }

    return existingCategory;
  }

  return createOutlookMasterCategory(accessToken, categoryName);
}

async function applyCategoryAndMarkReadByGraphId(graphMessageId, rawCategory) {
  const categoryName = cleanOutlookCategoryName(rawCategory);

  if (!isValidOutlookCategory(categoryName)) {
    throw new Error("Categoria inválida para aplicar no Outlook.");
  }

  const accessToken = await getGraphAccessToken();

  await ensureOutlookMasterCategoryExists(accessToken, categoryName);

  const getResponse = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(graphMessageId)}?$select=categories,isRead`,
    {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    }
  );

  if (!getResponse.ok) {
    const errorText = await getResponse.text();
    throw new Error(errorText || "Erro ao obter categorias atuais do email.");
  }

  const messageData = await getResponse.json();
  const currentCategories = Array.isArray(messageData.categories)
    ? messageData.categories
    : [];

  const mergedCategories = Array.from(
    new Set([...currentCategories, categoryName])
  );

  const patchResponse = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(graphMessageId)}`,
    {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        categories: mergedCategories,
        isRead: true,
      }),
    }
  );

  if (!patchResponse.ok) {
    const errorText = await patchResponse.text();
    throw new Error(errorText || "Erro ao aplicar categoria e marcar como lido.");
  }
}

async function applyCategoryAndMarkReadToOpenEmail(rawCategory) {
  const graphMessageId = getGraphMessageIdFromOfficeItem();
  await applyCategoryAndMarkReadByGraphId(graphMessageId, rawCategory);
}

function renderCategoryFilterOptions(categories) {
  const select = document.getElementById("category-filter-select");
  if (!select) return;

  if (!categories.length) {
    select.innerHTML = `<option value="">Nenhuma categoria disponível</option>`;
    return;
  }

  const options = categories
    .map((category) => {
      const rawName = category.nome || "";
      const cleanName = cleanOutlookCategoryName(rawName);

      return `<option value="${escapeHtml(cleanName)}">${escapeHtml(cleanName)}</option>`;
    })
    .join("");

  select.innerHTML = options;
}

function normalizeCategoryForCompare(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function messageHasCategory(message, categoryName) {
  const expected = normalizeCategoryForCompare(categoryName);
  const categories = Array.isArray(message.categories) ? message.categories : [];

  return categories.some(
    (category) => normalizeCategoryForCompare(category) === expected
  );
}

async function handleFilterEmailsByCategory() {
  try {
    const select = document.getElementById("category-filter-select");
    const selectedCategory = select?.value?.trim();

    if (!selectedCategory) {
      throw new Error("Seleciona uma categoria para filtrar.");
    }

    setGraphStatus(`A procurar emails da categoria "${selectedCategory}"...`);
    setHtml("category-filter-results", "A carregar emails...");

    let categories = categoriesCache;

    if (!categories.length) {
      categories = await getCategories();
      categoriesCache = Array.isArray(categories) ? categories : [];
      renderCategoryFilterOptions(categoriesCache);
    }

    const messages = await fetchInboxMessagesWithGraph();

    const filteredMessages = messages.filter((message) =>
      messageHasCategory(message, selectedCategory)
    );

    renderFilteredCategoryMessages(filteredMessages, selectedCategory);

    setGraphStatus(
      `Filtro concluído. Emails encontrados em "${selectedCategory}": ${filteredMessages.length}.`
    );
  } catch (error) {
    console.error("Erro ao filtrar emails por categoria:", error);
    setGraphStatus(`Erro ao filtrar por categoria: ${error.message}`, true);
    setHtml("category-filter-results", "Não foi possível filtrar os emails por categoria.");
  }
}

function renderFilteredCategoryMessages(messages, categoryName) {
  if (!messages.length) {
    setHtml(
      "category-filter-results",
      `Nenhum email encontrado com a categoria "${escapeHtml(categoryName)}".`
    );
    return;
  }

  const html = messages
    .map((message, index) => {
      const subject = escapeHtml(message.subject || "(Sem assunto)");
      const fromName = escapeHtml(message.from?.emailAddress?.name || "Remetente desconhecido");
      const fromAddress = escapeHtml(message.from?.emailAddress?.address || "");
      const received = escapeHtml(message.receivedDateTime || "");
      const readState = message.isRead ? "Lido" : "Não lido";
      const categories = Array.isArray(message.categories)
        ? message.categories.join(", ")
        : "-";

      return `
        <div style="margin-bottom: 14px; padding-bottom: 10px; border-bottom: 1px solid #ddd;">
          <strong>${index + 1}. ${subject}</strong><br>
          <span><strong>De:</strong> ${fromName} ${fromAddress ? `&lt;${fromAddress}&gt;` : ""}</span><br>
          <span><strong>Recebido:</strong> ${received}</span><br>
          <span><strong>Estado:</strong> ${readState}</span><br>
          <span><strong>Categorias:</strong> ${escapeHtml(categories)}</span>
        </div>
      `;
    })
    .join("");

  setHtml("category-filter-results", html);
}