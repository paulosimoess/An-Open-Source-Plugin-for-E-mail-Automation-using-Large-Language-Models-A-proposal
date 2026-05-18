const API_BASE_URL = "http://localhost:4000";
const API_SECRET = "teste123";
const IMPLEMENTACAO_ID = 1;

let cachedToken = null;

async function getToken() {
  if (cachedToken) return cachedToken;

  const response = await fetch(`${API_BASE_URL}/auth/token`, {
    method: "POST",
    headers: {
      "x-api-secret": API_SECRET,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({}),
  });

  if (!response.ok) {
    throw new Error("Não foi possível obter token do backend.");
  }

  const data = await response.json();
  cachedToken = data.token;
  return cachedToken;
}

async function authorizedFetch(url, options = {}) {
  const token = await getToken();

  const response = await fetch(url, {
    ...options,
    headers: {
      ...(options.headers || {}),
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText || "Erro na comunicação com o backend.");
  }

  return response.json();
}

export async function categorizeEmail(emailData) {
  const { message_id, thread_id, remetente, assunto, corpo } = emailData;

  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/thread/outlook-thread/email/outlook-message/categorizar`,
    {
      method: "POST",
      body: JSON.stringify({
        message_id,
        thread_id,
        remetente,
        assunto,
        corpo,
      }),
    }
  );
}

export async function requestRagResponse(emailData) {
  const { message_id, thread_id, remetente, assunto, corpo } = emailData;

  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/thread/outlook-thread/email/outlook-message/resposta/rag`,
    {
      method: "POST",
      body: JSON.stringify({
        remetente,
        assunto,
        corpo,
        message_id,
        thread_id,
      }),
    }
  );
}

export async function getLatestResponse(emailData) {
  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/thread/outlook-thread/email/outlook-message/resposta`,
    {
      method: "GET",
    }
  );
}

export async function validateResponse(emailData, respostaId) {
  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/thread/outlook-thread/email/outlook-message/resposta/${respostaId}/validar`,
    {
      method: "POST",
      body: JSON.stringify({}),
    }
  );
}

export async function getEmailState(emailData) {
  const { message_id, thread_id } = emailData;

  const params = new URLSearchParams({
    thread_id,
    message_id,
  });

  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/estado-email?${params.toString()}`,
    {
      method: "GET",
    }
  );
}