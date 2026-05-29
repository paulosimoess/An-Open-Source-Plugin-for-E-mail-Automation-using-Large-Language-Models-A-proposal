# Configuração das tabelas RAG

Para a geração de respostas com RAG funcionar, é necessário criar as tabelas `pdf_documents` e `rag_chunks` na base de dados PostgreSQL.

Estas tabelas são usadas para guardar os PDFs processados e os respetivos chunks com embeddings.

## 1. Criar tabelas

Executar o ficheiro:

```sql
sql/create_rag_tables.sql
```

na base de dados PostgreSQL.

## 2. Inserir PDFs e embeddings

Depois de criar as tabelas, correr na pasta `middleware-api`:

```bash
node insertPdfsWithEmbeddings.js
```

## 3. Confirmar dados inseridos

Executar na base de dados:

```sql
SELECT 'pdf_documents' AS tabela, COUNT(*) AS total
FROM pdf_documents
UNION ALL
SELECT 'rag_chunks' AS tabela, COUNT(*) AS total
FROM rag_chunks;
```

Ambas as tabelas devem apresentar valores superiores a zero.

## Nota

A coluna `embedding` usa `vector(768)`, porque o modelo de embeddings usado no projeto gera vetores com 768 dimensões.
