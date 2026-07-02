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

  const responseText = await response.text();

  let responseData = null;

  if (responseText) {
    try {
      responseData = JSON.parse(responseText);
    } catch {
      responseData = null;
    }
  }

  if (!response.ok) {
    const errorMessage =
      responseData?.error ||
      responseData?.message ||
      responseText ||
      "Erro na comunicação com o backend.";

    throw new Error(errorMessage);
  }

  return responseData;
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
        message_id,
        thread_id,
        remetente,
        assunto,
        corpo,
      }),
    }
  );
}

export async function getLatestResponse(emailData) {
  const state = await getEmailState(emailData);

  if (state?.resposta) {
    return state.resposta;
  }

  return {
    id_resposta: null,
    status: "Sem resposta",
    conteudo: "",
  };
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
  const { message_id, thread_id, assunto, remetente } = emailData;

  const params = new URLSearchParams();

  if (thread_id) params.append("thread_id", thread_id);
  if (message_id) params.append("message_id", message_id);
  if (assunto) params.append("assunto", assunto);
  if (remetente) params.append("remetente", remetente);

  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/estado-email?${params.toString()}`,
    {
      method: "GET",
    }
  );
}

export async function getCategories() {
  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/categorias`,
    {
      method: "GET",
    }
  );
}

export async function getCategoryKeywords(categoryId) {
  try {
    return await authorizedFetch(
      `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/categoria/${categoryId}/keywords`,
      {
        method: "GET",
      }
    );
  } catch (error) {
    if (String(error.message || "").includes("Nenhuma keyword encontrada")) {
      return [];
    }

    throw error;
  }
}

export async function addCategoryKeyword(categoryId, keyword) {
  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/categoria/${categoryId}/keyword`,
    {
      method: "POST",
      body: JSON.stringify({ keyword }),
    }
  );
}

export async function deleteCategoryKeyword(categoryId, keywordId) {
  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/categoria/${categoryId}/keyword/${keywordId}`,
    {
      method: "DELETE",
      body: JSON.stringify({}),
    }
  );
}

export async function createCategory({ nome, questao, paraQueServe }) {
  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/categoria`,
    {
      method: "POST",
      body: JSON.stringify({
        nome,
        questao,
        paraQueServe,
      }),
    }
  );
}

export async function updateCategory(categoryId, nome) {
  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/categoria/${categoryId}`,
    {
      method: "PUT",
      body: JSON.stringify({ nome }),
    }
  );
}

export async function deleteCategory(categoryId) {
  return authorizedFetch(
    `${API_BASE_URL}/implementacao/${IMPLEMENTACAO_ID}/categoria/${categoryId}`,
    {
      method: "DELETE",
      body: JSON.stringify({}),
    }
  );
}