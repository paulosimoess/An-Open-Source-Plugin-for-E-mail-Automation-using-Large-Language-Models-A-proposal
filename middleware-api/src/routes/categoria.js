import { pool } from "../db.js";
import { generateKeywords } from "../services/generateKeywords.js";

function normalizeText(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

export default async function categoriaRoutes(fastify, opts) {
  // GET /implementacao/:id_implementacao/categorias
  fastify.get("/implementacao/:id_implementacao/categorias", async (request, reply) => {
    const { id_implementacao } = request.params;

    try {
      const { rows } = await pool.query(
        "SELECT * FROM categoria WHERE id_implementacao = $1",
        [id_implementacao]
      );

      if (rows.length === 0) {
        return reply.code(404).send({ error: "Nenhuma categoria encontrada" });
      }

      reply.send(rows);
    } catch (err) {
      request.log.error(err);
      reply.code(500).send({ error: "Erro ao listar categorias" });
    }
  });

  // GET /implementacao/:id_implementacao/categorias/keywords
  fastify.get("/implementacao/:id_implementacao/categorias/keywords", async (request, reply) => {
    const { id_implementacao } = request.params;

    try {
      //query
      const categoriasResult = await pool.query(
        `SELECT c.id_categoria, c.nome AS label, k.keyword
        FROM categoria c
        LEFT JOIN keyword k ON c.id_categoria = k.id_categoria
        WHERE c.id_implementacao = $1`,
        [id_implementacao]
      );

      // Agrupando as categorias e as respetivas keywords
      const categoriasJson = categoriasResult.rows.reduce((acc, row) => {
        // Encontra a categoria existente ou cria uma nova
        let categoria = acc.find(c => c.label === row.label);
        
        if (!categoria) {
          categoria = { label: row.label, keywords: [] };
          acc.push(categoria);
        }

        if (row.keyword) {
          categoria.keywords.push(row.keyword);
        }

        return acc;
      }, []);

      // Retorna o JSON no formato desejado
      reply.send(categoriasJson);
    } catch (err) {
      request.log.error(err);
      reply.code(500).send({ error: "Erro ao listar categorias e keywords" });
    }
  });

  // POST /implementacao/:id_implementacao/categoria
  fastify.post("/implementacao/:id_implementacao/categoria", async (request, reply) => {
    const client = await pool.connect();

    try {
      const { id_implementacao } = request.params;
      let { nome, questao, paraQueServe } = request.body;

      if (!nome) {
        return reply.code(400).send({ error: "Campo 'nome' é obrigatório" });
      }

      nome = nome.trim();

      if (!nome.startsWith(".")) {
        nome = "." + nome;
      }

      const categoriasExistentes = await client.query(
        `SELECT id_categoria, nome
        FROM categoria
        WHERE id_implementacao = $1`,
        [id_implementacao]
      );

      const normalizedNewCategory = normalizeText(nome);

      const duplicateCategory = categoriasExistentes.rows.find(
        row => normalizeText(row.nome) === normalizedNewCategory
      );

      if (duplicateCategory) {
        return reply.code(409).send({
          error: "Esta categoria já existe nesta implementação",
          categoria: duplicateCategory
        });
      }

      const textoCompleto = [nome, questao, paraQueServe].filter(Boolean).join(" ");

      const keywords = await generateKeywords(textoCompleto, 12);

      await client.query("BEGIN");

      const categoriaResult = await client.query(
        `INSERT INTO categoria (id_implementacao, nome)
        VALUES ($1, $2)
        RETURNING id_categoria`,
        [id_implementacao, nome]
      );

      const categoriaId = categoriaResult.rows[0]?.id_categoria;

      if (!categoriaId) {
        throw new Error("ID da categoria não foi gerado");
      }

      if (keywords.length > 0) {
        await client.query(
          `INSERT INTO keyword (id_categoria, keyword)
          SELECT $1, unnest($2::text[])`,
          [categoriaId, keywords]
        );
      }

      await client.query("COMMIT");

      reply.code(201).send({
        message: "Categoria criada com sucesso",
        id_categoria: categoriaId,
        nome,
        keywordsGeradas: keywords
      });

    } catch (err) {
      await client.query("ROLLBACK");
      request.log.error(err);
      reply.code(500).send({ error: "Erro ao criar categoria" });
    } finally {
      client.release();
    }
  });



  // PUT /implementacao/:id_implementacao/categoria/:id_categoria
  fastify.put("/implementacao/:id_implementacao/categoria/:id_categoria", async (request, reply) => {
    const { id_implementacao, id_categoria } = request.params;
    let { nome } = request.body;

    if (!nome) {
      return reply.code(400).send({ error: "Campo 'nome' é obrigatório" });
    }

    nome = "." + nome;
    
    try {
      const result = await pool.query(
        `UPDATE categoria
         SET nome = $1
         WHERE id_categoria = $2 AND id_implementacao = $3
         RETURNING *`,
        [nome, id_categoria, id_implementacao]
      );

      if (result.rows.length === 0) {
        return reply.code(404).send({ error: "Categoria não encontrada" });
      }

      reply.send({
        message: "Categoria atualizada com sucesso",
        categoria: result.rows[0],
      });
    } catch (err) {
      request.log.error(err);
      reply.code(500).send({ error: "Erro ao atualizar categoria" });
    }
  });
  // DELETE /implementacao/:id_implementacao/categoria/:id_categoria
fastify.delete(
  "/implementacao/:id_implementacao/categoria/:id_categoria",
  async (request, reply) => {
    const { id_implementacao, id_categoria } = request.params;
    const client = await pool.connect();

    console.log(
      `Recebido para excluir categoria: Implementação ${id_implementacao}, Categoria ${id_categoria}`
    );

    try {
      await client.query("BEGIN");

      // 1. Verificar se a categoria existe
      const categoriaCheck = await client.query(
        `SELECT id_categoria, nome
         FROM categoria
         WHERE id_categoria = $1
           AND id_implementacao = $2
         FOR UPDATE`,
        [id_categoria, id_implementacao]
      );

      if (categoriaCheck.rowCount === 0) {
        await client.query("ROLLBACK");

        return reply.code(404).send({
          error: "Categoria não encontrada"
        });
      }

      const categoriaAEliminar = categoriaCheck.rows[0];

      // 2. Impedir a eliminação da categoria de fallback ".Outro"
      if (normalizeText(categoriaAEliminar.nome).replace(/^\./, "") === "outro") {
        await client.query("ROLLBACK");

        return reply.code(400).send({
          error: "A categoria '.Outro' não pode ser eliminada"
        });
      }

      // 3. Obter a categoria ".Outro" da mesma implementação
      const outroResult = await client.query(
        `SELECT id_categoria
         FROM categoria
         WHERE id_implementacao = $1
           AND LOWER(TRIM(nome)) IN ('.outro', 'outro')
         LIMIT 1
         FOR UPDATE`,
        [id_implementacao]
      );

      if (outroResult.rowCount === 0) {
        await client.query("ROLLBACK");

        return reply.code(400).send({
          error: "Categoria '.Outro' não encontrada nesta implementação"
        });
      }

      const idCategoriaOutro = outroResult.rows[0].id_categoria;

      // 4. Obter as threads atualmente classificadas com a categoria
      const threadsParaReclassificar = await client.query(
        `SELECT DISTINCT thread_id
         FROM thread_categorizacao
         WHERE id_categoria = $1
           AND ativa = true`,
        [id_categoria]
      );

      // 5. Reclassificar cada thread para ".Outro"
      for (const row of threadsParaReclassificar.rows) {
        const threadId = row.thread_id;

        // 5.1 Desativar todas as categorizações ativas da thread
        await client.query(
          `UPDATE thread_categorizacao
           SET ativa = false
           WHERE thread_id = $1
             AND ativa = true`,
          [threadId]
        );

        // 5.2 Criar ou reativar a categorização ".Outro"
        await client.query(
          `INSERT INTO thread_categorizacao
             (
               thread_id,
               id_categoria,
               id_tipo_categorizacao,
               ativa,
               data
             )
           VALUES ($1, $2, 3, true, NOW())
           ON CONFLICT (id_categoria, thread_id)
           DO UPDATE SET
             ativa = true,
             id_tipo_categorizacao = 3,
             data = NOW()`,
          [threadId, idCategoriaOutro]
        );
      }

      // 6. Apagar todas as relações antigas com a categoria eliminada
      // Inclui relações ativas e inativas.
      await client.query(
        `DELETE FROM thread_categorizacao
         WHERE id_categoria = $1`,
        [id_categoria]
      );

      // 7. Apagar as keywords associadas
      await client.query(
        `DELETE FROM keyword
         WHERE id_categoria = $1`,
        [id_categoria]
      );

      // 8. Apagar a categoria
      const result = await client.query(
        `DELETE FROM categoria
         WHERE id_categoria = $1
           AND id_implementacao = $2
         RETURNING *`,
        [id_categoria, id_implementacao]
      );

      if (result.rowCount === 0) {
        throw new Error("A categoria não foi eliminada");
      }

      await client.query("COMMIT");

      return reply.send({
        message:
          "Categoria eliminada com sucesso e threads reclassificadas para '.Outro'",
        categoria: result.rows[0],
        threads_reclassificadas: threadsParaReclassificar.rowCount
      });
    } catch (err) {
      try {
        await client.query("ROLLBACK");
      } catch (rollbackError) {
        request.log.error(rollbackError);
      }

      request.log.error(err);

      return reply.code(500).send({
        error: "Erro ao eliminar categoria"
      });
    } finally {
      client.release();
    }
  }
);
}