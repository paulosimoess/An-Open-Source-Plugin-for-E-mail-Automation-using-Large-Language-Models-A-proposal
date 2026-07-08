import axios from "axios";
import { pool } from "../db.js";

const OLLAMA_URL = "http://127.0.0.1:11434";
const LLM_MODEL = "llama3";
const EMBED_MODEL = "nomic-embed-text";

const RAG_TOP_K = 7;
const MAX_CONTEXT_CHARS = 6500;

function buildInstitutionalResponsePrompt(question, context = "") {
  const hasContext = Boolean(String(context || "").trim());

  const contextBlock = hasContext
    ? `
Informação documental de uso interno:
${context}
`
    : `
Informação documental de uso interno:
Não existe informação documental adicional disponível. Responde de forma prudente e institucional, sem inventar regras, prazos, documentos, contactos, valores ou procedimentos específicos.
`;

  return `
És um assistente institucional automático de uma entidade pública em Portugal.

TAREFA:
Redigir uma resposta pronta a enviar ao cidadão, com base no email recebido e na informação documental disponível.

OBJETIVO:
A resposta deve ser útil, concreta, institucional e escrita em português de Portugal. Quando existir informação documental, deves usá-la para responder diretamente ao pedido, mas sem mencionar nomes de ficheiros, páginas, identificadores técnicos ou o processo de pesquisa documental.

REGRAS OBRIGATÓRIAS:
- Escreve exclusivamente em português de Portugal.
- Usa um tom formal, claro, institucional e natural.
- Não uses português do Brasil.
- Nunca uses expressões como: "você", "contatar", "contato", "nossa equipe", "reparo", "solicitação", "atenciosamente", "estamos aqui para ajudá-lo".
- Usa expressões adequadas como: "o seu contacto", "contactar", "contacto", "serviços responsáveis", "reparação", "pedido", "encaminhado", "analisado", "verificado".
- Não prometas prazos concretos.
- Não garantas deferimento, autorização, intervenção, aprovação ou resolução.
- Não prometas resposta futura, contacto posterior, aprovação formal ou comunicação posterior, a menos que isso esteja claramente indicado na informação documental.
- Não inventes contactos, nomes, cargos, departamentos, legislação, taxas, documentos, moradas, links ou procedimentos.
- Não uses placeholders como "[documentos necessários]", "[custos]", "[endereço]" ou semelhantes.
- Não incluas títulos como "Saudação", "Corpo", "Assunto", "Resposta", "1.º parágrafo" ou "2.º parágrafo".
- Não copies frases do cidadão como se fossem da entidade.
- Não uses "Gostaria de saber" na resposta institucional.
- Só pede informação adicional se for realmente necessária.
- Nunca termines a resposta com "O seu contacto,".
- Nunca escrevas expressões como "identifico o assunto", "identifico brevemente", "o seu objetivo é", "o que resta saber é" ou frases que pareçam instruções internas.
- Não escrevas "De acordo com a informação documental disponível". Usa diretamente a informação documental na resposta.
- Não digas "sugiro que o seu pedido seja encaminhado". Escreve antes "o pedido será encaminhado" ou "o pedido deverá ser analisado pelos serviços responsáveis".
- Nunca escrevas "Obrigamo-nos por...", "obrigamo-nos por ter escolhido..." ou frases promocionais.
- A primeira frase deve começar preferencialmente por "Agradecemos o seu contacto relativamente a...".
- A despedida deve ser sempre exatamente:
Com os melhores cumprimentos,
Assistente automático AI4APGovernance

REGRAS SOBRE A INFORMAÇÃO DOCUMENTAL:
- Se existir informação documental, usa-a de forma prioritária.
- Se a informação documental indicar procedimentos, documentos, condições, canais de submissão ou informação útil, inclui essa informação na resposta.
- Usa a informação dos documentos como conhecimento dos serviços, sem dizer ao cidadão que estás a consultar documentos.
- Nunca digas ao cidadão para consultar "o contexto documental", "os documentos internos", "a informação documental" ou ficheiros PDF específicos.
- Nunca menciones nomes de ficheiros, códigos de ficheiros, "#p1", "#p2", páginas internas ou identificadores técnicos no corpo da resposta.
- As fontes consultadas serão acrescentadas automaticamente pelo sistema no final da resposta. Não as escrevas no corpo principal.
- Se a informação documental incluir vários procedimentos diferentes, usa apenas os que sejam diretamente relevantes para o pedido do cidadão.
- Ignora informação sobre renovação, transmissão de licença, caducidade ou alteração de titularidade quando o pedido for apenas sobre inscrição inicial, venda, autorização ou informação geral.
- Não incluas prazos, validade de licenças ou regras específicas se não forem diretamente pedidos pelo cidadão.
- Se a informação documental não tiver detalhe suficiente para responder, indica apenas que o pedido será analisado ou encaminhado pelos serviços responsáveis.
- Nunca inventes informação que não esteja no email recebido ou na informação documental.
- Se o pedido for sobre inscrição inicial, atribuição de lugar, venda na feira ou venda ambulante, não incluas informação sobre renovação, transmissão de licença, caducidade ou alteração de titularidade.
- Usa apenas informação diretamente relacionada com o pedido concreto do cidadão.
- Não escrevas "O seu objetivo é..."; escreve de forma natural, por exemplo: "O seu pedido refere-se a..." ou "Relativamente ao pedido apresentado...".
- Se o pedido for sobre inscrição inicial, atribuição de lugar, venda na feira, venda no mercado ou venda ambulante, não incluas informação sobre renovação, transmissão de licença, caducidade, alteração de titularidade ou prazos de 5 anos.
- Não incluas informação sobre prazos, validade de licenças ou transmissão de titularidade se o cidadão não tiver perguntado especificamente por esses temas.
- Para pedidos sobre venda em feira ou mercado, foca-te em: apresentação do pedido, inscrição, documentação aplicável, análise pelos serviços responsáveis, custos e disponibilidade de lugares quando aplicável.
- Não uses expressões como "comunicar-se com o Município". Usa antes "apresentar o pedido junto do Município" ou "contactar os serviços responsáveis".

FORMATO DA RESPOSTA:
A resposta deve ser escrita diretamente, sem títulos, sem listas de instruções e sem mencionar o processo de geração.

Começa sempre por:
Boa tarde,

Depois escreve 2 a 3 parágrafos naturais:
- Agradece o contacto e refere o tema do pedido de forma simples.
- Responde ao pedido usando a informação documental disponível, quando existir.
- Se aplicável, indica que o pedido será analisado ou encaminhado pelos serviços responsáveis, sem prometer deferimento, autorização, prazos ou resolução.

Termina sempre exatamente com:

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

    // Português do Brasil / expressões indesejadas
    .replace(/\b[Vv]ocê\s+poderá\b/g, "poderá")
    .replace(/\b[Vv]ocê\s+deverá\b/g, "deverá")
    .replace(/\b[Vv]ocê\s+precisará\b/g, "será necessário")
    .replace(/\b[Vv]ocê\b/g, "o/a munícipe")
    .replace(/\bcontatar\b/gi, "contactar")
    .replace(/\bcontate\b/gi, "contacte")
    .replace(/\bcontato\b/gi, "contacto")
    .replace(/\bsolicitação\b/gi, "pedido")
    .replace(/\bpedestres\b/gi, "peões")
    .replace(/\breparo\b/gi, "reparação")
    .replace(/\bnossa equipe\b/gi, "os serviços responsáveis")
    .replace(/\bAtenciosamente\b/gi, "Com os melhores cumprimentos")

    // Frases pouco naturais ou que parecem instruções do prompt
    .replace(/Agradeço o contacto e identifico brevemente o assunto do pedido\.?/gi, "Agradecemos o seu contacto.")
    .replace(/Agradeço o contacto e identifico o assunto do pedido\.?/gi, "Agradecemos o seu contacto.")
    .replace(/Agradeço o seu contacto e identifico o assunto do pedido[^.]*\./gi, "Agradecemos o seu contacto.")
    .replace(/identifico o assunto do pedido de forma natural\.?/gi, "")
    .replace(/identifico brevemente o assunto do pedido\.?/gi, "")
    .replace(/O seu objetivo é obter informação sobre/gi, "O seu pedido refere-se a")
    .replace(/O que resta saber é se/gi, "Relativamente a")
    .replace(/Obrigamos-se pelo contacto/gi, "Agradecemos o seu contacto")
    .replace(/Obrigamos pelo contacto/gi, "Agradecemos o seu contacto")
    .replace(/Obrigada por contactar a nossa entidade/gi, "Agradecemos o seu contacto")
    .replace(/Obrigado por contactar a nossa entidade/gi, "Agradecemos o seu contacto")
    .replace(/Obrigamo-nos por ter escolhido o nosso município para exercer a sua atividade\.?/gi, "")
    .replace(/Estamos aqui para ajudá-lo e fornecer-lhe os dados necessários\./gi, "")
    .replace(/O seu contacto,\s*/gi, "Com os melhores cumprimentos,\n")
    .replace(/Com os melhores cumprimentos,\s*Com os melhores cumprimentos,/gi, "Com os melhores cumprimentos,")

    // Evitar linguagem interna do RAG
    .replace(/De acordo com a informação documental disponível,\s*/gi, "")
    .replace(/informação documental disponível/gi, "informação disponível")
    .replace(/consulte os documentos de contexto documental/gi, "consulte a informação disponível")
    .replace(/documentos de contexto documental/gi, "documentos disponíveis")
    .replace(/contexto documental/gi, "informação disponível")
    .replace(/documentos internos/gi, "informação disponível")
    .replace(/IPVC-ChatBot-[^\s,.;]+#p\d+/gi, "")

    // Evitar promessas demasiado fortes
    .replace(/Em seguida, receberá uma resposta formal sobre a aprovação ou não da sua pedido\./gi, "")
    .replace(/Em seguida, receberá uma resposta formal sobre a aprovação ou não do seu pedido\./gi, "")
    .replace(/Será contactado se for necessário solicitar mais informações ou realizar uma visita à zona afectada\./gi, "")
    .replace(/e, em seguida, receberá comunicação sobre as etapas seguintes/gi, "")
    .replace(/receberá comunicação sobre as etapas seguintes\.?/gi, "")
    .replace(/sugiro que o seu pedido seja encaminhado/gi, "o pedido será encaminhado")

    // Remover frases demasiado específicas que têm aparecido em casos de mercados/feiras
    .replace(/Além disso, é recomendável consultar a informação disponível sobre as condições necessárias para vender produtos na feira municipal\.?/gi, "")
    .replace(/Lembre-se de que o prazo de 5 anos após a emissão da licença diária é necessário para manter a atividade no Mercado Municipal\.?/gi, "")

    .replace(/O seu objetivo é saber quais são/gi, "Pretende obter informação sobre")
    .replace(/O seu objetivo é obter informação sobre/gi, "O seu pedido refere-se a")
    .replace(/apresentação de uma pedido prévia/gi, "apresentação de um pedido prévio")
    .replace(/Além disso, é importante mencionar que a transmissão da licença[\s\S]*?representantes\./gi, "")
    .replace(/transmissão da licença do titular do registo[\s\S]*?representantes\./gi, "")
    .replace(/renovação da licença diária[\s\S]*?\./gi, "")
    .replace(/caducidade[\s\S]*?\./gi, "")

    // Remover placeholders caso apareçam
    .replace(/\s*\[[^\]]+\]/g, "")

    .replace(/comunicar-se com o Município/gi, "apresentar o pedido junto do Município")
    .replace(/comunicar-se com o municipio/gi, "apresentar o pedido junto do Município")
    .replace(/pode comunicar-se com o Município/gi, "poderá apresentar o pedido junto do Município")
    .replace(/pode comunicar-se com o municipio/gi, "poderá apresentar o pedido junto do Município")

    .replace(/com 90 dias de antecedência em relação à data de\.?/gi, "")
    .replace(/com 90 dias de antecedência em relação à data do pedido\.?/gi, "")
    .replace(/com 90 dias de antecedência em relação à data pretendida\.?/gi, "")
    .replace(/, com 90 dias de antecedência em relação à data de\.?/gi, "")

    .replace(/Além disso, é importante mencionar que a transmissão da licença[\s\S]*?representantes\./gi, "")
    .replace(/A transmissão da licença[\s\S]*?representantes\./gi, "")
    .replace(/transmissão da licença do titular do registo[\s\S]*?representantes\./gi, "")

    .replace(/Pretende obter informação sobre as condições para venda de produtos no mercado, quais documentos são necessários para inscrição e se existem lugares disponíveis para novos vendedores\./gi, "")
    .replace(/O seu pedido refere-se a as condições/gi, "O seu pedido refere-se às condições")

    // Limpeza geral
    .replace(/[ \t]+\n/g, "\n")
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

function normalizeCategoryForRag(categoryName) {
  return String(categoryName || "")
    .trim()
    .replace(/^\./, "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[.,;:()/_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

/* =========================
   Pesquisa pgvector por categoria
========================= */
async function searchRAGContext(question, categoriaNome) {
  const embedding = await embedText(question);
  const embeddingVector = toPgVector(embedding);

  const categoriaNormalizada = normalizeCategoryForRag(categoriaNome);

  console.log("RAG categoria normalizada para pesquisa:", categoriaNormalizada);

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
    WHERE
      upper(
        trim(
          regexp_replace(
            translate(
              regexp_replace(d.categoria_nome, '^\\.', ''),
              'ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇáàâãäéèêëíìîïóòôõöúùûüç',
              'AAAAAEEEEIIIIOOOOOUUUUCaaaaaeeeeiiiiooooouuuuc'
            ),
            '[[:space:][:punct:]]+',
            ' ',
            'g'
          )
        )
      ) = $2
    ORDER BY c.embedding <=> $1::vector
    LIMIT $3
    `,
    [embeddingVector, categoriaNormalizada, RAG_TOP_K]
  );

  return rows;
}

function normalizeTextForSearch(text) {
  return String(text || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function hasAnyTerm(text, terms) {
  const normalized = normalizeTextForSearch(text);
  return terms.some((term) => normalized.includes(normalizeTextForSearch(term)));
}

function filterChunksForQuestion(question, categoriaNome, chunks) {
  const normalizedCategory = normalizeCategoryForRag(categoriaNome);
  const normalizedQuestion = normalizeTextForSearch(question);

  if (!Array.isArray(chunks) || chunks.length === 0) {
    return chunks;
  }

  let filteredChunks = [...chunks];

  const isMercadosFeiras =
    normalizedCategory.includes("MERCADOS") ||
    normalizedCategory.includes("FEIRAS") ||
    normalizedCategory.includes("VENDA AMBULANTE");

  if (isMercadosFeiras) {
    const questionAsksRenewalOrTransfer = hasAnyTerm(normalizedQuestion, [
      "renovação",
      "renovar",
      "transmissão",
      "transmitir",
      "titularidade",
      "caducidade",
      "prazo de 5 anos",
      "5 anos",
    ]);

    if (!questionAsksRenewalOrTransfer) {
      const withoutRenewalOrTransfer = filteredChunks.filter((chunk) => {
        const chunkText = normalizeTextForSearch(
          `${chunk.filename || ""} ${chunk.content || ""}`
        );

        return !hasAnyTerm(chunkText, [
          "renlic",
          "renovlic",
          "renovação",
          "renovar",
          "trasmlic",
          "transmissao",
          "transmissão",
          "titularidade",
          "caducidade",
          "prazo de 5 anos",
          "5 anos",
        ]);
      });

      if (withoutRenewalOrTransfer.length > 0) {
        filteredChunks = withoutRenewalOrTransfer;
      }
    }

    filteredChunks.sort((a, b) => {
      const aText = normalizeTextForSearch(`${a.filename || ""} ${a.content || ""}`);
      const bText = normalizeTextForSearch(`${b.filename || ""} ${b.content || ""}`);

      const score = (text) => {
        let value = 0;

        if (text.includes("pedlic")) value += 4;
        if (text.includes("pedido")) value += 3;
        if (text.includes("licenca")) value += 2;
        if (text.includes("vendamb")) value += 2;
        if (text.includes("feira")) value += 2;
        if (text.includes("mercado")) value += 2;
        if (text.includes("inscricao")) value += 2;
        if (text.includes("vendedor")) value += 2;

        if (text.includes("renlic")) value -= 5;
        if (text.includes("renov")) value -= 5;
        if (text.includes("trasmlic")) value -= 5;
        if (text.includes("transmiss")) value -= 5;

        return value;
      };

      return score(bText) - score(aText);
    });
  }

  return filteredChunks;
}

/* =========================
   Construção do prompt com fontes
========================= */
function buildPrompt(question, chunks) {
  let context = "";
  let total = 0;
  const sources = [];
  const seenSources = new Set();

  for (let i = 0; i < chunks.length; i += 1) {
    const c = chunks[i];

    const block =
      `[${i + 1}] ${c.filename}#p${c.page_num}\n` +
      `${c.content}\n\n`;

    if (total + block.length > MAX_CONTEXT_CHARS) break;

    context += block;
    total += block.length;

    const sourceKey = `${c.filename}#p${c.page_num}`;

    if (!seenSources.has(sourceKey)) {
      seenSources.add(sourceKey);

      sources.push({
        filename: c.filename,
        page: c.page_num,
        score: c.score
      });
    }
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

    const categoriaLimpa = String(categoriaNome || "").trim();
    const categoriaNormalizada = normalizeCategoryForRag(categoriaLimpa);

    console.log("RAG categoria recebida:", categoriaNome);
    console.log("RAG categoria normalizada:", categoriaNormalizada);

    if (!categoriaLimpa || categoriaNormalizada === "OUTRO") {
      prompt = buildInstitutionalResponsePrompt(question);
      console.log("RAG sem categoria útil. A gerar resposta institucional sem contexto documental.");
    } else {
      const chunks = await searchRAGContext(question, categoriaLimpa);

      console.log("RAG chunks encontrados:", chunks.length);

      if (!chunks.length) {
        prompt = buildInstitutionalResponsePrompt(question);
        console.log("RAG sem chunks para esta categoria. A gerar resposta institucional sem contexto documental.");
      } else {
        const filteredChunks = filterChunksForQuestion(question, categoriaLimpa, chunks);

        console.log(
          "RAG chunks após filtro de relevância:",
          filteredChunks.length
        );

        const built = buildPrompt(question, filteredChunks);
        prompt = built.prompt;
        sources = built.sources;

        console.log(
          "RAG fontes usadas:",
          sources.map((s) => `${s.filename}#p${s.page}`)
        );
      }
    }

    const response = await axios.post(`${OLLAMA_URL}/api/generate`, {
      model: LLM_MODEL,
      prompt,
      stream: false,
      options: {
        temperature: 0.2,
        top_p: 0.8
      }
    });

    const json = response.data;

    if (json.error || !json.response) {
      return "⚠️ Não foi possível obter uma resposta baseada nos PDFs.";
    }

    let respostaGerada = cleanGeneratedResponse(json.response);

    if (sources.length > 0) {
      const fontesTexto = sources
        .map((s) => `- ${s.filename}, página ${s.page}`)
        .join("\n");

      respostaGerada += `\n\nFontes consultadas:\n${fontesTexto}`;
    }

    return respostaGerada;
  } catch (e) {
    console.error("Erro ao contactar RAG:", e);
    return "⚠️ Erro ao contactar o sistema RAG.";
  }
}

export default generateResponseFromRAG;