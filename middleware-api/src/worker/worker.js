import amqp from "amqplib";
import { pool } from "../db.js";
import generateResponseFromRAG from "../services/generateResponseFromRAG.js";
import { PROCESS_RAG } from "../queue/jobTypes.js";

const RABBITMQ_URL = process.env.RABBITMQ_URL || "amqp://localhost";
const QUEUE_NAME = "emailTasks";

function isCategoriaUtil(categoria) {
  const normalizada = String(categoria || "")
    .trim()
    .replace(/^\./, "")
    .toLowerCase();

  return Boolean(normalizada && normalizada !== "outro");
}

async function startWorker() {
  try {
    const connection = await amqp.connect(RABBITMQ_URL);
    const channel = await connection.createChannel();

    await channel.assertQueue(QUEUE_NAME, { durable: true });

    console.log("Worker initialized and ready to process jobs");

    channel.consume(
      QUEUE_NAME,
      async (msg) => {
        if (msg === null) return;

        const jobData = JSON.parse(msg.content.toString());
        console.log("Worker started processing job with data:", jobData);

        if (jobData.type === PROCESS_RAG) {
          const {
            job_id,
            message_id,
            remetente,
            assunto,
            corpo,
            thread_id,
            categoria,
            categoria_nome
          } = jobData;

          try {
            // Atualiza status do job para PROCESSING
            await pool.query(
              `UPDATE resposta_gerada SET status = 'PROCESSING' WHERE id_resposta = $1`,
              [job_id]
            );

            console.log(`Job ${job_id} status updated to 'PROCESSING'`);

            const emailContent = `Remetente: ${remetente}\nAssunto: ${assunto}\n${corpo}`;

            const categoriaRes = await pool.query(
              `
              SELECT c.nome 
              FROM thread_categorizacao tc
              JOIN categoria c ON tc.id_categoria = c.id_categoria
              WHERE tc.thread_id = $1 AND tc.ativa = TRUE
              ORDER BY
                CASE
                  WHEN lower(trim(replace(c.nome, '.', ''))) = 'outro' THEN 1
                  ELSE 0
                END
              LIMIT 1
              `,
              [thread_id]
            );

            const categoriaDoFrontend = categoria_nome || categoria || null;
            const categoriaDaBd = categoriaRes.rows[0]?.nome || ".Outro";

            let categoriaNome;

            if (isCategoriaUtil(categoriaDoFrontend)) {
              categoriaNome = categoriaDoFrontend;
            } else if (isCategoriaUtil(categoriaDaBd)) {
              categoriaNome = categoriaDaBd;
            } else {
              categoriaNome = categoriaDoFrontend || categoriaDaBd || ".Outro";
            }

            console.log("Categoria recebida do frontend:", categoriaDoFrontend || "(não enviada)");
            console.log(`Categoria identificada para thread ${thread_id}: ${categoriaDaBd}`);
            console.log("Categoria usada para RAG:", categoriaNome);

            const resposta = await generateResponseFromRAG(emailContent, categoriaNome);

            await pool.query(
              `UPDATE resposta_gerada 
               SET status = 'DONE', conteudo = $2
               WHERE id_resposta = $1`,
              [job_id, resposta]
            );

            await pool.query(
              `UPDATE email
               SET resposta_gerada = TRUE, resposta = $2
               WHERE message_id = $1`,
              [message_id, resposta]
            );

            console.log(`Job ${job_id} completed successfully`);
            channel.ack(msg);
          } catch (err) {
            console.error(`Error processing job ${job_id}:`, err.message);

            // Atualiza status do job para ERROR
            await pool.query(
              `UPDATE resposta_gerada
               SET status = 'ERROR', conteudo = $2
               WHERE id_resposta = $1`,
              [job_id, err.message]
            );

            channel.nack(msg, false, false);
            console.log(`Job ${job_id} marked as ERROR`);
          }
        } else {
          console.log("Job type not recognized:", jobData.type);
          channel.ack(msg);
        }
      },
      { noAck: false }
    );

    pool
      .connect()
      .then((client) => {
        console.log("Database connection established successfully");
        client.release();
      })
      .catch((err) => {
        console.error("Error connecting to the database:", err.message);
      });
  } catch (err) {
    console.error("Error initializing RabbitMQ worker:", err.message);
  }
}

startWorker();