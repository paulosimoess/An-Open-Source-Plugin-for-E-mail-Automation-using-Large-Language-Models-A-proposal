import axios from "axios";
import { pool } from "../db.js";

const OLLAMA_URL = "http://127.0.0.1:11434";
const LLM_MODEL = "llama3";
const EMBED_MODEL = "nomic-embed-text";

const RAG_TOP_K = 5;
const MAX_CONTEXT_CHARS = 4000;

function buildInstitutionalResponsePrompt(question, context = "") {
  const contextBlock = String(context || "").trim()
    ? `
Contexto documental de apoio:
${context}
`
    : `
Contexto documental de apoio:
Não existe contexto documental adicional disponível. Responde de forma genérica, prudente e institucional, sem inventar regras, prazos ou procedimentos específicos.
`;

  return `
És um assistente institucional automático de uma entidade pública em Portugal.

TAREFA:
Redigir uma resposta pronta a enviar ao cidadão, com base no email recebido e, quando existir, no contexto documental de apoio.

REGRAS OBRIGATÓRIAS:
- Escreve exclusivamente em português de Portugal.
- Usa um tom formal, claro, institucional e natural.
- Não uses português do Brasil.
- Não uses expressões como: "você", "pedestres", "nossa equipe", "reparo", "solicitação", "obrigatório atender".
- Usa preferencialmente expressões como: "o seu contacto", "peões", "serviços responsáveis", "reparação", "pedido", "encaminhado", "analisado", "verificado".
- Não prometas prazos concretos.
- Não garantas que a intervenção vai acontecer.
- Não inventes contactos, nomes, cargos, departamentos ou legislação.
- Não incluas títulos como "Saudação", "Corpo", "Assunto" ou "Resposta".
- Não incluas fontes consultadas no corpo da resposta.
- A resposta deve ser curta, objetiva e pronta a enviar.
- Se faltar informação, pede apenas os dados essenciais de forma breve.
- Quando aplicável, usa a formulação: "para que a situação possa ser verificada e seja avaliada a intervenção necessária."
- Inclui uma frase curta e natural de enquadramento do problema, adequada à categoria do email.
- Essa frase deve mostrar atenção ao pedido, mas sem prometer resolução, prazos ou intervenção garantida.
- Não uses sempre a mesma frase em todos os temas; adapta ao conteúdo do email.
- Não uses a expressão "reparação da avaria" em pedidos sobre espaços verdes, publicidade, autorizações ou informações gerais.
- Em pedidos de informação, não fales em "intervenção necessária"; usa "prestar a informação adequada".
- Garante sempre uma linha em branco antes de "Com os melhores cumprimentos,".
- Nunca escrevas "resumo do pedido", "resumo o seu pedido", "registro que faz" ou frases semelhantes.
- Não copies frases do cidadão como se fossem da entidade.
- Não uses "Gostaria de saber" na resposta institucional, porque essa expressão pertence normalmente ao cidadão.
- Só pede informação adicional se for realmente necessária.

ESTRUTURA DA RESPOSTA:
Boa tarde,

[Escreve uma frase natural a agradecer o contacto e a identificar brevemente o assunto, sem usar a palavra "resumo" e sem copiar frases do cidadão na primeira pessoa.]

[Inclui uma frase curta de enquadramento adequada ao tema do email, sem prometer resolução.]

[Indica o encaminhamento adequado ao tipo de pedido:
- Avaria/reparação: "O pedido será encaminhado para os serviços responsáveis, para que a situação possa ser verificada e seja avaliada a intervenção necessária."
- Pedido de autorização: "O pedido será encaminhado para os serviços responsáveis, para que possa ser analisado de acordo com os procedimentos aplicáveis."
- Pedido de informação: "O pedido será encaminhado para os serviços responsáveis, para que lhe possa ser prestada a informação adequada."
- Espaços verdes/manutenção: "O pedido será encaminhado para os serviços responsáveis, para que a situação possa ser verificada e seja avaliada a intervenção necessária de limpeza ou manutenção."]

[Se for mesmo necessário, pede apenas informação adicional adequada ao pedido. Não peças informação adicional quando o email já tiver dados suficientes.]

Com os melhores cumprimentos,
Assistente automático AI4APGovernance

Email recebido:
${question}
${contextBlock}
`.trim();
}

function cleanGeneratedResponse(responseText) {
  return String(responseText || "")
    .replace(/^Saudação\s*/gim, "")
    .replace(/^Corpo\s*/gim, "")
    .replace(/^Assunto\s*/gim, "")
    .replace(/^Resposta\s*/gim, "")
    .replace(/Fontes consultadas:[\s\S]*$/gim, "")
    .replace(/\b[Vv]ocê\b/g, "o/a munícipe")
    .replace(/\bpedestres\b/gi, "peões")
    .replace(/\breparo\b/gi, "reparação")
    .replace(/\bnossa equipe\b/gi, "os serviços responsáveis")
    .replace(/\bsolicitação\b/gi, "pedido")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/* =========================
   Conversão para pgvector
========================= */
function toPgVector(vec) {
  return `[${vec.join(",")}]`;
}

/* =========================
   Embeddings via Ollama
========================= */
async function embedText(text) {
  const res = await axios.post(`${OLLAMA_URL}/api/embeddings`, {
    model: EMBED_MODEL,
    prompt: text
  });

  if (!res.data?.embedding || !Array.isArray(res.data.embedding)) {
    throw new Error("Embedding inválido devolvido pelo Ollama");
  }

  return res.data.embedding;
}

/* =========================
   Pesquisa pgvector por categoria
========================= */
async function searchRAGContext(question, categoriaNome) {
  const embedding = await embedText(question);
  const embeddingVector = toPgVector(embedding);

  const { rows } = await pool.query(
    `
    SELECT
      c.content,
      c.pdf_id,
      c.page_num,
      d.filename,
      1 - (c.embedding <=> $1::vector) AS score
    FROM rag_chunks c
    JOIN pdf_documents d ON d.pdf_id = c.pdf_id
    WHERE d.categoria_nome = $2
    ORDER BY c.embedding <=> $1::vector
    LIMIT $3
    `,
    [embeddingVector, categoriaNome, RAG_TOP_K]
  );

  return rows; // todos os chunks encontrados
}

/* =========================
   Construção do prompt
========================= */
function buildPrompt(question, chunks) {
  let context = "";
  let total = 0;
  const sources = [];

  for (let i = 0; i < chunks.length; i++) {
    const c = chunks[i];
    const block =
      `[${i + 1}] ${c.filename}#p${c.page_num}\n` +
      `${c.content}\n\n`;

    if (total + block.length > MAX_CONTEXT_CHARS) break;

    context += block;
    total += block.length;

    sources.push({
      filename: c.filename,
      page: c.page_num,
      score: c.score
    });
  }

  const prompt = buildInstitutionalResponsePrompt(question, context);

  return { prompt, sources };
}

/* =========================
   Função principal
========================= */
async function generateResponseFromRAG(question, categoriaNome) {
  try {
    let prompt;
    let sources = [];

    //Caso sem categoria ou categoria = ".Outro"
    if (!categoriaNome || categoriaNome === ".Outro") {
        prompt = buildInstitutionalResponsePrompt(question);
    } else {
      // Obter contexto RAG baseado na categoria
      const chunks = await searchRAGContext(question, categoriaNome);

      if (!chunks.length) {
        prompt = buildInstitutionalResponsePrompt(question);
      } else {
        // Construir prompt com contexto
        const built = buildPrompt(question, chunks);
        prompt = built.prompt;
        sources = built.sources;
      }
    }

    // Chamar Ollama
    const response = await axios.post(`${OLLAMA_URL}/api/generate`, {
      model: LLM_MODEL,
      prompt,
      stream: false
    });

    const json = response.data;

    if (json.error || !json.response) {
      return "⚠️ Não foi possível obter uma resposta baseada nos PDFs.";
    }

    let respostaGerada = cleanGeneratedResponse(json.response);

    /* if (sources.length > 0) {
      const fontesTexto = sources
        .map(s => `- ${s.filename}, página ${s.page}`)
        .join("\n");

      respostaGerada += `\n\nFontes consultadas:\n${fontesTexto}`;
    } */

    return respostaGerada;

  } catch (e) {
    console.error("Erro ao contactar RAG:", e);
    return "⚠️ Erro ao contactar o sistema RAG.";
  }
}

export default generateResponseFromRAG;
