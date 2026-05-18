import { pool } from "../db.js";
import { addJob } from "../queue/queue.js";
import { PROCESS_RAG } from "../queue/jobTypes.js";

export default async function emailRoutes(fastify, opts) {

    // POST /implementacao/:id_implementacao/thread/:thread_id/email/:message_id/categorizar
    fastify.post("/implementacao/:id_implementacao/thread/:thread_id/email/:message_id/categorizar", async (request, reply) => {
        const { id_implementacao } = request.params;
        const { message_id, thread_id, remetente, assunto, corpo, resposta } = request.body;

        if (!message_id || !thread_id || !remetente || !assunto || !corpo) {
            return reply.code(400).send({
                error: "Campos 'message_id', 'thread_id', 'remetente', 'assunto' e 'corpo' são obrigatórios"
            });
        }

        try {

            const threadResult = await pool.query(
                `SELECT thread_id FROM thread WHERE thread_id = $1 AND id_implementacao = $2`,
                [thread_id, id_implementacao]
            );

            if (threadResult.rowCount === 0) {
                await pool.query(
                    `INSERT INTO thread (thread_id, id_implementacao, assunto)
                    VALUES ($1, $2, $3)`,
                    [thread_id, id_implementacao, assunto]
                );
            }

            const emailExists = await pool.query(
                `SELECT message_id FROM email WHERE message_id = $1`,
                [message_id]
            );

            if (emailExists.rowCount > 0) {
                await pool.query(
                    `UPDATE email
                    SET remetente=$1, corpo=$2, resposta=$3
                    WHERE message_id=$4`,
                    [remetente, corpo, resposta || null, message_id]
                );
            } else {
                await pool.query(
                    `INSERT INTO email
                    (message_id, id_implementacao, thread_id, remetente, corpo, resposta)
                    VALUES ($1,$2,$3,$4,$5,$6)`,
                    [message_id, id_implementacao, thread_id, remetente, corpo, resposta || null]
                );
            }

            const categoriasDB = await pool.query(
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

            const normalized = corpo.toLowerCase()
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


            await pool.query(
                `UPDATE thread_categorizacao
                SET ativa = FALSE
                WHERE thread_id = $1
                    AND id_categoria <> $2`,
                [thread_id, bestCategory.id_categoria]
            );

            const existingCat = await pool.query(
                `SELECT id_categoria, thread_id
                FROM thread_categorizacao
                WHERE thread_id = $1
                    AND id_categoria = $2`,
                [thread_id, bestCategory.id_categoria]
            );

            if (existingCat.rowCount > 0) {
                await pool.query(
                    `UPDATE thread_categorizacao
                    SET id_tipo_categorizacao = 1,
                        keywords = $3,
                        data = NOW(),
                        ativa = TRUE
                    WHERE thread_id = $1
                        AND id_categoria = $2`,
                    [thread_id, bestCategory.id_categoria, bestKeywords.join(",")]
                );
            } else {
                await pool.query(
                    `INSERT INTO thread_categorizacao
                    (id_categoria, thread_id, id_tipo_categorizacao, keywords, data, ativa)
                    VALUES ($1, $2, 1, $3, NOW(), TRUE)`,
                    [bestCategory.id_categoria, thread_id, bestKeywords.join(",")]
                );
            }

            await pool.query(
                `UPDATE email
                SET categorizado = TRUE
                WHERE message_id = $1`,
                [message_id]
            );

            return reply.send({
                message: "Email categorizado com sucesso",
                categoria: bestCategory.nome,
                thread_id,
                message_id,
                keywords_usadas: bestKeywords
            });

        } catch (err) {
            request.log.error(err);
            return reply.code(500).send({ error: "Erro ao categorizar email" });
        }
    });



    // POST /implementacao/:id_implementacao/thread/:thread_id/categorizar/manual
    fastify.post("/implementacao/:id_implementacao/thread/:thread_id/categorizar/manual", async (request, reply) => {
        const { id_implementacao, thread_id } = request.params;
        const { id_categoria } = request.body;

        if (!id_categoria) {
            return reply.code(400).send({
                error: "O campo 'id_categoria' é obrigatório"
            });
        }

        try {
            const threadCheck = await pool.query(
                `SELECT thread_id
                FROM thread
                WHERE thread_id = $1 AND id_implementacao = $2`,
                [thread_id, id_implementacao]
            );

            if (threadCheck.rowCount === 0) {
                return reply.code(404).send({ error: "Thread não encontrada para esta implementação" });
            }

            const categoriaCheck = await pool.query(
                `SELECT id_categoria
                FROM categoria
                WHERE id_categoria = $1 AND id_implementacao = $2`,
                [id_categoria, id_implementacao]
            );

            if (categoriaCheck.rowCount === 0) {
                return reply.code(404).send({ error: "Categoria não encontrada nesta implementação" });
            }

            await pool.query(
                `UPDATE thread_categorizacao
                SET ativa = FALSE
                WHERE thread_id = $1`,
                [thread_id]
            );

            const novaCategorizacao = await pool.query(
                `INSERT INTO thread_categorizacao (
                    id_categoria,
                    thread_id,
                    id_tipo_categorizacao,
                    keywords,
                    data,
                    ativa
                )
                VALUES ($1, $2, 2, NULL, NOW(), TRUE)
                RETURNING *`,
                [id_categoria, thread_id]
            );

            return reply.send({
                message: "Categorização manual aplicada com sucesso",
                categorizacao: novaCategorizacao.rows[0]
            });

        } catch (err) {
            request.log.error(err);
            return reply.code(500).send({ error: "Erro ao aplicar categorização manual" });
        }
    });


    fastify.post("/implementacao/:id_implementacao/thread/:thread_id/email/:message_id/resposta/rag", async (request, reply) => {
        const { id_implementacao, thread_id, message_id } = request.params;
        const { remetente, assunto, corpo } = request.body;

        if (!remetente || !assunto || !corpo) {
        return reply.code(400).send({
            error: "Campos 'remetente', 'assunto' e 'corpo' são obrigatórios"
        });
        }

        try {
        let threadResult = await pool.query(
            `SELECT thread_id FROM thread WHERE thread_id = $1 AND id_implementacao = $2`,
            [thread_id, id_implementacao]
        );

        if (threadResult.rowCount === 0) {
            await pool.query(
            `INSERT INTO thread (thread_id, id_implementacao, assunto)
            VALUES ($1, $2, $3)`,
            [thread_id, id_implementacao, assunto]
            );
        }

        let emailResult = await pool.query(
            `SELECT * FROM email WHERE message_id = $1 AND id_implementacao = $2`,
            [message_id, id_implementacao]
        );

        if (emailResult.rowCount === 0) {
            await pool.query(
            `INSERT INTO email 
            (message_id, id_implementacao, thread_id, remetente, corpo, categorizado)
            VALUES ($1, $2, $3, $4, $5, TRUE)`,
            [message_id, id_implementacao, thread_id, remetente, corpo]
            );
        } else {
            await pool.query(
            `UPDATE email
            SET remetente=$1, corpo=$2
            WHERE message_id=$3`,
            [remetente, corpo, message_id]
            );
        }

        const tipoResposta = 1;
        const resposta = await pool.query(
            `INSERT INTO resposta_gerada 
            (message_id, id_tipo_resposta, conteudo, status, data_criacao)
            VALUES ($1, $2, '', 'PENDING', NOW())
            RETURNING id_resposta`,
            [message_id, tipoResposta]
        );

        const jobId = resposta.rows[0].id_resposta;

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
        request.log.error(err);
        return reply.code(500).send({ error: "Erro ao processar request" });
        }
    });

    // GET /implementacao/:id_implementacao/thread/:thread_id/email/:message_id/resposta
    fastify.get("/implementacao/:id_implementacao/thread/:thread_id/email/:message_id/resposta", async (request, reply) => {
        const { message_id } = request.params;

        try {

            const query = `
                SELECT *
                FROM resposta_gerada
                WHERE message_id = $1
                ORDER BY data_criacao DESC
                LIMIT 1;
            `;

            const result = await pool.query(query, [message_id]);

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
        const { id_implementacao, thread_id, message_id, resposta_id } = request.params;

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
    });

    fastify.get("/implementacao/:id_implementacao/estado-email", async (request, reply) => {
        const { id_implementacao } = request.params;
        const { thread_id, message_id } = request.query;

        if (!thread_id || !message_id) {
            return reply.code(400).send({
                error: "Os parâmetros 'thread_id' e 'message_id' são obrigatórios"
            });
        }

        try {
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
                [thread_id, id_implementacao]
            );

            const respostaResult = await pool.query(
                `SELECT id_resposta, conteudo, status, data_criacao, data_validacao
                 FROM resposta_gerada
                 WHERE message_id = $1
                 ORDER BY data_criacao DESC
                 LIMIT 1`,
                [message_id]
            );

            const categoria = categoriaResult.rows[0] || null;
            const resposta = respostaResult.rows[0] || null;

            return reply.send({
                categoria: categoria ? categoria.nome : null,
                keywords_usadas: categoria?.keywords
                    ? categoria.keywords.split(",").map(k => k.trim()).filter(Boolean)
                    : [],
                resposta: resposta || null
            });

        } catch (err) {
            request.log.error(err);
            return reply.code(500).send({ error: "Erro ao obter estado do email" });
        }
    });

}