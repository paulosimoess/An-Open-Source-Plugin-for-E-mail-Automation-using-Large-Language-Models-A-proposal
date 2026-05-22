/* global document, Office */

import {
  categorizeEmail,
  requestRagResponse,
  getLatestResponse,
  validateResponse,
  getEmailState,
} from "./api";

import { MSAL_CLIENT_ID } from "./msal.local";

let graphAccessToken = null;
let graphAccountUsername = null;

let currentEmailData = null;
let lastResponseId = null;

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

    const graphLoginButton = document.getElementById("graph-login-button");
    if (graphLoginButton) graphLoginButton.onclick = handleMicrosoftLogin;

    const graphReadInboxButton = document.getElementById("graph-read-inbox-button");
    if (graphReadInboxButton) graphReadInboxButton.onclick = handleReadInboxWithGraph;

    run();
    updateGraphAccountUi();
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
  if (!text) return "";
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

async function hydrateEmailState() {
  if (!currentEmailData) return;

  try {
    const state = await getEmailState(currentEmailData);
    console.log("Estado atual do email:", state);

    if (state?.categoria) {
      setText("suggested-category", state.categoria);
      setText(
        "used-keywords",
        state.keywords_usadas?.length
          ? state.keywords_usadas.join(", ")
          : "Nenhuma keyword usada"
      );
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
      remetente: fromEmail,
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

    setText("suggested-category", result.categoria || "Sem categoria");
    setText(
      "used-keywords",
      result.keywords_usadas?.length
        ? result.keywords_usadas.join(", ")
        : "Nenhuma keyword usada"
    );

    setStatus("Email categorizado com sucesso.");
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

function updateGraphAccountUi() {
  if (!graphAccountUsername) {
    setText("graph-account", "-");
    setGraphStatus("Ainda não autenticado no Microsoft Graph.");
    return;
  }

  setText("graph-account", graphAccountUsername);
  setGraphStatus("Conta Microsoft autenticada.");
}

async function handleMicrosoftLogin() {
  try {
    validateMsalConfig();

    setGraphStatus("A abrir autenticação Microsoft...");

    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/auth.html",
      {
        height: 60,
        width: 45,
        displayInIframe: false,
      },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          setGraphStatus(
            `Erro ao abrir janela de autenticação: ${asyncResult.error.message}`,
            true
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

              updateGraphAccountUi();
              setGraphStatus("Autenticação Microsoft concluída com sucesso.");

              dialog.close();
              return;
            }

            if (message.type === "AUTH_ERROR") {
              setGraphStatus(`Erro no login Microsoft: ${message.message}`, true);
              dialog.close();
            }
          } catch (error) {
            setGraphStatus(`Erro ao processar resposta de autenticação: ${error.message}`, true);
            dialog.close();
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          console.warn("DialogEventReceived:", arg);
        });
      }
    );
  } catch (error) {
    console.error("Erro no login Microsoft:", error);
    setGraphStatus(`Erro no login Microsoft: ${error.message}`, true);
  }
}

async function getGraphAccessToken() {
  if (!graphAccessToken) {
    throw new Error("Ainda não existe sessão Microsoft. Clique primeiro em 'Iniciar sessão Microsoft'.");
  }

  return graphAccessToken;
}

async function handleReadInboxWithGraph() {
  try {
    setGraphStatus("A obter token Microsoft Graph...");
    setHtml("graph-inbox-list", "A carregar emails da Inbox...");

    const accessToken = await getGraphAccessToken();

    setGraphStatus("A ler emails da Inbox via Microsoft Graph...");

    const endpoint =
      "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages" +
      "?$top=10" +
      "&$select=id,subject,from,receivedDateTime,bodyPreview,conversationId,isRead" +
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
    const messages = data.value || [];

    renderGraphInboxMessages(messages);
    await updateGraphAccountUi();

    setGraphStatus(`Leitura concluída. Emails carregados: ${messages.length}.`);
  } catch (error) {
    console.error("Erro ao ler Inbox via Graph:", error);
    setGraphStatus(`Erro ao ler Inbox: ${error.message}`, true);
    setHtml("graph-inbox-list", "Não foi possível carregar os emails da Inbox.");
  }
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