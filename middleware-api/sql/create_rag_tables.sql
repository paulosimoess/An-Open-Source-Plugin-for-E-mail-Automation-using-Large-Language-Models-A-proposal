CREATE EXTENSION IF NOT EXISTS vector;

CREATE TABLE IF NOT EXISTS pdf_documents (
    pdf_id SERIAL PRIMARY KEY,
    filename TEXT NOT NULL,
    categoria_nome TEXT NOT NULL,
    created_at TIMESTAMP DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS rag_chunks (
    chunk_id SERIAL PRIMARY KEY,
    pdf_id INTEGER NOT NULL REFERENCES pdf_documents(pdf_id) ON DELETE CASCADE,
    content TEXT NOT NULL,
    page_num INTEGER,
    embedding vector(768),
    created_at TIMESTAMP DEFAULT NOW()
);