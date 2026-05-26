import { pool } from "../db.js";
import { addJob } from "../queue/queue.js";
import { PROCESS_RAG } from "../queue/jobTypes.js";

async function ensureThread(client, thread_id, id_implementacao, assunto) {
    const threadResult = await client.query(
        `SELECT thread_id
         FROM thread
         WHERE thread_id = $1
           AND id_implementacao = $2`,
        [thread_id, id_implementacao]
    );

    if (threadResult.rowCount === 0) {
        await client.query(
            `INSERT INTO thread (thread_id, id_implementacao, assunto)
             VALUES ($1, $2, $3)`,
            [thread_id, id_implementacao, assunto]
        );
    }
}

async function upsertEmail(client, {
    message_id,
    id_implementacao,
    thread_id,
    remetente,
    corpo,
    resposta = null,
    categorizado = null
}) {
    const emailExists = await client.query(
        `SELECT message_id
         FROM email
         WHERE message_id = $1`,
        [message_id]
    );

    if (emailExists.rowCount > 0) {
        if (categorizado === null) {
            await client.query(
                `UPDATE email
                 SET remetente = $1,
                     corpo = $2,
                     resposta = $3
                 WHERE message_id = $4`,
                [remetente, corpo, resposta, message_id]
            );
        } else {
            await client.query(
                `UPDATE email
                 SET remetente = $1,
                     corpo = $2,
                     resposta = $3,
                     categorizado = $4
                 WHERE message_id = $5`,
                [remetente, corpo, resposta, categorizado, message_id]
            );
        }
    } else {
        await client.query(
            `INSERT INTO email
             (message_id, id_implementacao, thread_id, remetente, corpo, resposta, categorizado)
             VALUES ($1, $2, $3, $4, $5, $6, $7)`,
            [message_id, id_implementacao, thread_id, remetente, corpo, resposta, categorizado ?? false]
        );
    }
}

async function upsertThreadCategorization(client, {
    thread_id,
    id_categoria,
    id_tipo_categorizacao,
    keywords
}) {
    await client.query(
        `UPDATE thread_categorizacao
         SET ativa = FALSE
         WHERE thread_id = $1
           AND id_categoria <> $2`,
        [thread_id, id_categoria]
    );

    const existingCat = await client.query(
        `SELECT id_categoria, thread_id
         FROM thread_categorizacao
         WHERE thread_id = $1
           AND id_categoria = $2`,
        [thread_id, id_categoria]
    );

    if (existingCat.rowCount > 0) {
        await client.query(
            `UPDATE thread_categorizacao
             SET id_tipo_categorizacao = $3,
                 keywords = $4,
                 data = NOW(),
                 ativa = TRUE
             WHERE thread_id = $1
               AND id_categoria = $2`,
            [thread_id, id_categoria, id_tipo_categorizacao, keywords]
        );
    } else {
        await client.query(
            `INSERT INTO thread_categorizacao
             (id_categoria, thread_id, id_tipo_categorizacao, keywords, data, ativa)
             VALUES ($1, $2, $3, $4, NOW(), TRUE)`,
            [id_categoria, thread_id, id_tipo_categorizacao, keywords]
        );
    }
}

export default async function emailRoutes(fastify, opts) {
    fastify.post("/implementacao/:id_implementacao/thread/:thread_id/email/:message_id/categorizar", async (request, reply) => {
        const { id_implementacao } = request.params;
        const { message_id, thread_id, remetente, assunto, corpo, resposta } = request.body;

        if (!message_id || !thread_id || !remetente || !assunto || !corpo) {
            return reply.code(400).send({
                error: "Campos 'message_id', 'thread_id', 'remetente', 'assunto' e 'corpo' são obrigatórios"
            });
        }

        const client = await pool.connect();

        try {
            await client.query("BEGIN");

            await ensureThread(client, thread_id, id_implementacao, assunto);

            await upsertEmail(client, {
                message_id,
                id_implementacao,
                thread_id,
                remetente,
                corpo,
                resposta: resposta || null
            });

            const categoriasDB = await client.query(
                `SELECT c.id_categoria, c.nome AS categoria_nome, k.keyword
                 FROM categoria c
                 LEFT JOIN keyword k ON k.id_categoria = c.id_categoria
                 WHERE c.id_implementacao = $1`,
                [id_implementacao]
            );

            const categorias = {};
            categoriasDB.rows.forEach(row => {
                if (!categorias[row.id_categoria]) {
                    categorias[row.id_categoria] = {
                        id_categoria: row.id_categoria,
                        nome: row.categoria_nome,
                        keywords: []
                    };
                }
                if (row.keyword) categorias[row.id_categoria].keywords.push(row.keyword);
            });

            const categoriasArray = Object.values(categorias);

            const textoParaCategorizar = `${assunto || ""} ${corpo || ""}`;

            const normalized = textoParaCategorizar.toLowerCase()
                .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
                .replace(/[.,!?;:]/g, "")
                .replace(/\s+/g, " ");

            let bestCategory = null;
            let bestKeywords = [];
            let highestScore = 0;

            for (const cat of categoriasArray) {
                let score = 0;
                let matches = [];

                for (const kw of cat.keywords) {
                    const normalizedKeyword = kw.toLowerCase()
                        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
                        .replace(/[.,!?;:]/g, "")
                        .replace(/\s+/g, " ");

                    const regex = new RegExp(`\\b${normalizedKeyword}\\b`, "gi");
                    const found = normalized.match(regex);
                    if (found) {
                        score += found.length;
                        matches.push(kw);
                    }
                }

                if (score > highestScore) {
                    highestScore = score;
                    bestCategory = cat;
                    bestKeywords = matches;
                }
            }

            if (!bestCategory) {
                bestCategory = categoriasArray.find(c =>
                    c.nome.toLowerCase() === ".outro"
                );
                bestKeywords = [];
            }

            await upsertThreadCategorization(client, {
                thread_id,
                id_categoria: bestCategory.id_categoria,
                id_tipo_categorizacao: 1,
                keywords: bestKeywords.join(",")
            });

            await client.query(
                `UPDATE email
                 SET categorizado = TRUE
                 WHERE message_id = $1`,
                [message_id]
            );

            await client.query("COMMIT");

            return reply.send({
                message: "Email categorizado com sucesso",
                categoria: bestCategory.nome,
                thread_id,
                message_id,
                keywords_usadas: bestKeywords
            });

        } catch (err) {
            await client.query("ROLLBACK");
            request.log.error(err);
            return reply.code(500).send({ error: "Erro ao categorizar email" });
        } finally {
            client.release();
        }
    });

    fastify.post("/implementacao/:id_implementacao/thread/:thread_id/categorizar/manual", async (request, reply) => {
        const { id_implementacao, thread_id } = request.params;
        const { id_categoria } = request.body;

        if (!id_categoria) {
            return reply.code(400).send({
                error: "O campo 'id_categoria' é obrigatório"
            });
        }

        const client = await pool.connect();

        try {
            await client.query("BEGIN");

            const threadCheck = await client.query(
                `SELECT thread_id
                 FROM thread
                 WHERE thread_id = $1
                   AND id_implementacao = $2`,
                [thread_id, id_implementacao]
            );

            if (threadCheck.rowCount === 0) {
                await client.query("ROLLBACK");
                return reply.code(404).send({ error: "Thread não encontrada para esta implementação" });
            }

            const categoriaCheck = await client.query(
                `SELECT id_categoria
                 FROM categoria
                 WHERE id_categoria = $1
                   AND id_implementacao = $2`,
                [id_categoria, id_implementacao]
            );

            if (categoriaCheck.rowCount === 0) {
                await client.query("ROLLBACK");
                return reply.code(404).send({ error: "Categoria não encontrada nesta implementação" });
            }

            await upsertThreadCategorization(client, {
                thread_id,
                id_categoria,
                id_tipo_categorizacao: 2,
                keywords: null
            });

            await client.query("COMMIT");

            return reply.send({
                message: "Categorização manual aplicada com sucesso",
                categorizacao: {
                    id_categoria,
                    thread_id,
                    id_tipo_categorizacao: 2,
                    ativa: true
                }
            });

        } catch (err) {
            await client.query("ROLLBACK");
            request.log.error(err);
            return reply.code(500).send({ error: "Erro ao aplicar categorização manual" });
        } finally {
            client.release();
        }
    });

    fastify.post("/implementacao/:id_implementacao/thread/:thread_id/email/:message_id/resposta/rag", async (request, reply) => {
        const { id_implementacao } = request.params;
        const {
            remetente,
            assunto,
            corpo,
            message_id: body_message_id,
            thread_id: body_thread_id
        } = request.body;

        const message_id = body_message_id || request.params.message_id;
        const thread_id = body_thread_id || request.params.thread_id;

        if (!remetente || !assunto || !corpo || !message_id || !thread_id) {
            return reply.code(400).send({
                error: "Campos 'message_id', 'thread_id', 'remetente', 'assunto' e 'corpo' são obrigatórios"
            });
        }

        const client = await pool.connect();

        try {
            await client.query("BEGIN");

            await ensureThread(client, thread_id, id_implementacao, assunto);

            await upsertEmail(client, {
                message_id,
                id_implementacao,
                thread_id,
                remetente,
                corpo,
                resposta: null,
                categorizado: true
            });

            const tipoResposta = 1;
            const resposta = await client.query(
                `INSERT INTO resposta_gerada
                 (message_id, id_tipo_resposta, conteudo, status, data_criacao)
                 VALUES ($1, $2, '', 'PENDING', NOW())
                 RETURNING id_resposta`,
                [message_id, tipoResposta]
            );

            const jobId = resposta.rows[0].id_resposta;

            await client.query("COMMIT");

            await addJob({
                type: PROCESS_RAG,
                job_id: jobId,
                message_id,
                remetente,
                assunto,
                corpo,
                thread_id,
                id_implementacao
            });

            return reply.code(202).send({
                status: "queued",
                job_id: jobId,
                message_id
            });

        } catch (err) {
            await client.query("ROLLBACK");
            request.log.error(err);
            return reply.code(500).send({ error: "Erro ao processar request" });
        } finally {
            client.release();
        }
    });

    fastify.get("/implementacao/:id_implementacao/thread/:thread_id/email/:message_id/resposta", async (request, reply) => {
        const message_id = request.query.message_id || request.params.message_id;

        try {
            const result = await pool.query(
                `SELECT *
                 FROM resposta_gerada
                 WHERE message_id = $1
                 ORDER BY data_criacao DESC
                 LIMIT 1`,
                [message_id]
            );

            if (result.rows.length === 0) {
                return reply.code(404).send({ error: "Nenhuma resposta gerada encontrada" });
            }

            return reply.send(result.rows[0]);

        } catch (err) {
            request.log.error(err);
            return reply.code(500).send({ error: "Erro ao buscar resposta gerada" });
        }
    });

    fastify.post("/implementacao/:id_implementacao/thread/:thread_id/email/:message_id/resposta/:resposta_id/validar",
        async (request, reply) => {
            const { id_implementacao, resposta_id } = request.params;

            const message_id = request.body?.message_id || request.params.message_id;
            const thread_id = request.body?.thread_id || request.params.thread_id;

            try {
                const result = await pool.query(
                    `SELECT conteudo
                     FROM resposta_gerada
                     WHERE id_resposta = $1`,
                    [resposta_id]
                );

                if (result.rows.length === 0) {
                    return reply.code(404).send({ error: "Resposta não encontrada." });
                }

                const respostaConteudo = result.rows[0].conteudo;

                await pool.query(
                    `UPDATE email
                     SET resposta = $1,
                         resposta_validada = TRUE,
                         respondido = TRUE
                     WHERE message_id = $2
                       AND thread_id = $3
                       AND id_implementacao = $4`,
                    [respostaConteudo, message_id, thread_id, id_implementacao]
                );

                await pool.query(
                    `UPDATE resposta_gerada
                     SET data_validacao = NOW()
                     WHERE id_resposta = $1`,
                    [resposta_id]
                );

                return reply.code(200).send({
                    ok: true,
                    resposta: respostaConteudo,
                    message: "Resposta validada e marcada com data de validação."
                });

            } catch (err) {
                console.error("⚠️ Erro ao validar resposta:", err);
                return reply.code(500).send({ error: "Erro interno ao validar resposta." });
            }
        }
    );

    fastify.get("/implementacao/:id_implementacao/estado-email", async (request, reply) => {
        const { id_implementacao } = request.params;
        const { thread_id, message_id, assunto, remetente } = request.query;

        if (!thread_id && !message_id && (!assunto || !remetente)) {
            return reply.code(400).send({
                error: "É necessário enviar pelo menos 'message_id', 'thread_id' ou 'assunto' + 'remetente'"
            });
        }

        try {
            let resolvedEmail = null;

            if (message_id) {
                const emailByMessage = await pool.query(
                    `SELECT e.message_id, e.thread_id, e.remetente, t.assunto
                    FROM email e
                    JOIN thread t ON t.thread_id = e.thread_id
                    WHERE e.message_id = $1
                    AND e.id_implementacao = $2
                    LIMIT 1`,
                    [message_id, id_implementacao]
                );

                if (emailByMessage.rowCount > 0) {
                    resolvedEmail = emailByMessage.rows[0];
                }
            }

            if (!resolvedEmail && thread_id) {
                const emailByThread = await pool.query(
                    `SELECT e.message_id, e.thread_id, e.remetente, t.assunto
                    FROM email e
                    JOIN thread t ON t.thread_id = e.thread_id
                    WHERE e.thread_id = $1
                    AND e.id_implementacao = $2
                    ORDER BY e.message_id DESC
                    LIMIT 1`,
                    [thread_id, id_implementacao]
                );

                if (emailByThread.rowCount > 0) {
                    resolvedEmail = emailByThread.rows[0];
                }
            }

            if (!resolvedEmail && assunto && remetente) {
                const emailBySubjectSender = await pool.query(
                    `SELECT e.message_id, e.thread_id, e.remetente, t.assunto
                    FROM email e
                    JOIN thread t ON t.thread_id = e.thread_id
                    WHERE e.id_implementacao = $1
                    AND LOWER(TRIM(t.assunto)) = LOWER(TRIM($2))
                    AND LOWER(TRIM(e.remetente)) = LOWER(TRIM($3))
                    ORDER BY e.message_id DESC
                    LIMIT 1`,
                    [id_implementacao, assunto, remetente]
                );

                if (emailBySubjectSender.rowCount > 0) {
                    resolvedEmail = emailBySubjectSender.rows[0];
                }
            }

            const resolvedThreadId = resolvedEmail?.thread_id || thread_id || null;
            const resolvedMessageId = resolvedEmail?.message_id || message_id || null;

            let categoria = null;
            let resposta = null;

            if (resolvedThreadId) {
                const categoriaResult = await pool.query(
                    `SELECT c.id_categoria, c.nome, tc.keywords
                    FROM thread_categorizacao tc
                    JOIN categoria c ON c.id_categoria = tc.id_categoria
                    JOIN thread t ON t.thread_id = tc.thread_id
                    WHERE tc.thread_id = $1
                    AND t.id_implementacao = $2
                    AND tc.ativa = TRUE
                    ORDER BY tc.data DESC NULLS LAST
                    LIMIT 1`,
                    [resolvedThreadId, id_implementacao]
                );

                categoria = categoriaResult.rows[0] || null;
            }

            if (resolvedMessageId) {
                const respostaResult = await pool.query(
                    `SELECT id_resposta, conteudo, status, data_criacao, data_validacao
                    FROM resposta_gerada
                    WHERE message_id = $1
                    ORDER BY data_criacao DESC
                    LIMIT 1`,
                    [resolvedMessageId]
                );

                resposta = respostaResult.rows[0] || null;
            }

            return reply.send({
                categoria: categoria ? categoria.nome : null,
                keywords_usadas: categoria?.keywords
                    ? categoria.keywords.split(",").map(k => k.trim()).filter(Boolean)
                    : [],
                resposta: resposta || null,
                email_resolvido: resolvedEmail
                    ? {
                        message_id: resolvedEmail.message_id,
                        thread_id: resolvedEmail.thread_id,
                        remetente: resolvedEmail.remetente,
                        assunto: resolvedEmail.assunto
                    }
                    : null
            });

        } catch (err) {
            request.log.error(err);
            return reply.code(500).send({ error: "Erro ao obter estado do email" });
        }
    });
}