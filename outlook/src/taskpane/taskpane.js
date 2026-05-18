/* global document, Office */

import { categorizeEmail } from "./api";

let currentEmailData = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    hideElement("sideload-msg");
    showElement("app-body", "flex");

    const runButton = document.getElementById("run");
    if (runButton) {
      runButton.onclick = run;
    }

    const categorizeButton = document.getElementById("categorize-button");
    if (categorizeButton) {
      categorizeButton.onclick = handleCategorize;
    }

    run();
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
  return text
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

    setStatus("Dados do email carregados com sucesso.");
  } catch (error) {
    console.error("Erro ao carregar dados do email:", error);
    setStatus(`Erro: ${error.message}`, true);

    setText("item-subject", "N/A");
    setText("item-from", "N/A");
    setText("item-id", "N/A");
    setText("item-conversation-id", "N/A");
    setHtml("item-body", "Não foi possível carregar o conteúdo do email.");
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