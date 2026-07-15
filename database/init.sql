-- Complete database initialization script.
-- Generated from schema.sql and followed by the required seed data.

--
-- PostgreSQL database dump
--


-- Dumped from database version 17.9
-- Dumped by pg_dump version 17.9

SET statement_timeout = 0;
SET lock_timeout = 0;
SET idle_in_transaction_session_timeout = 0;
SET transaction_timeout = 0;
SET client_encoding = 'UTF8';
SET standard_conforming_strings = on;
SELECT pg_catalog.set_config('search_path', '', false);
SET check_function_bodies = false;
SET xmloption = content;
SET client_min_messages = warning;
SET row_security = off;

--
-- Name: vector; Type: EXTENSION; Schema: -; Owner: -
--

CREATE EXTENSION IF NOT EXISTS vector WITH SCHEMA public;


--
-- Name: EXTENSION vector; Type: COMMENT; Schema: -; Owner: -
--

COMMENT ON EXTENSION vector IS 'vector data type and ivfflat and hnsw access methods';


--
-- Name: inserir_categoria_outro(); Type: FUNCTION; Schema: public; Owner: -
--

CREATE FUNCTION public.inserir_categoria_outro() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
BEGIN
    INSERT INTO categoria (id_implementacao, nome)
    VALUES (NEW.id_implementacao, '.Outro');
    RETURN NEW;
END;
$$;


SET default_tablespace = '';

SET default_table_access_method = heap;

--
-- Name: categoria; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.categoria (
    id_categoria integer NOT NULL,
    id_implementacao integer NOT NULL,
    nome character varying(255)
);


--
-- Name: categoria_id_categoria_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.categoria_id_categoria_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: categoria_id_categoria_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.categoria_id_categoria_seq OWNED BY public.categoria.id_categoria;


--
-- Name: email; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.email (
    message_id text NOT NULL,
    id_implementacao integer NOT NULL,
    thread_id text NOT NULL,
    remetente character varying(255),
    corpo text,
    resposta text,
    categorizado boolean DEFAULT false,
    resposta_gerada boolean DEFAULT false,
    resposta_validada boolean DEFAULT false,
    respondido boolean DEFAULT false
);


--
-- Name: implementacao; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.implementacao (
    id_implementacao integer NOT NULL,
    email character varying(255),
    plataforma character varying(100)
);


--
-- Name: implementacao_id_implementacao_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.implementacao_id_implementacao_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: implementacao_id_implementacao_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.implementacao_id_implementacao_seq OWNED BY public.implementacao.id_implementacao;


--
-- Name: keyword; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.keyword (
    id_keyword integer NOT NULL,
    id_categoria integer NOT NULL,
    keyword character varying(255)
);


--
-- Name: keyword_id_keyword_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.keyword_id_keyword_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: keyword_id_keyword_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.keyword_id_keyword_seq OWNED BY public.keyword.id_keyword;


--
-- Name: pdf_documents; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.pdf_documents (
    pdf_id integer NOT NULL,
    filename text NOT NULL,
    categoria_nome text NOT NULL,
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: pdf_documents_pdf_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.pdf_documents_pdf_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: pdf_documents_pdf_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.pdf_documents_pdf_id_seq OWNED BY public.pdf_documents.pdf_id;


--
-- Name: rag_chunks; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.rag_chunks (
    chunk_id integer NOT NULL,
    pdf_id integer NOT NULL,
    content text NOT NULL,
    page_num integer,
    embedding public.vector(768),
    created_at timestamp without time zone DEFAULT now()
);


--
-- Name: rag_chunks_chunk_id_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.rag_chunks_chunk_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: rag_chunks_chunk_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.rag_chunks_chunk_id_seq OWNED BY public.rag_chunks.chunk_id;


--
-- Name: resposta_gerada; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.resposta_gerada (
    id_resposta integer NOT NULL,
    message_id text NOT NULL,
    id_tipo_resposta integer NOT NULL,
    conteudo text,
    status character varying(100),
    data_criacao timestamp without time zone,
    data_validacao timestamp without time zone
);


--
-- Name: resposta_gerada_id_resposta_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.resposta_gerada_id_resposta_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: resposta_gerada_id_resposta_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.resposta_gerada_id_resposta_seq OWNED BY public.resposta_gerada.id_resposta;


--
-- Name: thread; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.thread (
    thread_id text NOT NULL,
    id_implementacao integer NOT NULL,
    assunto character varying(500)
);


--
-- Name: thread_categorizacao; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.thread_categorizacao (
    id_categoria integer NOT NULL,
    thread_id text NOT NULL,
    id_tipo_categorizacao integer,
    keywords text,
    data timestamp without time zone,
    ativa boolean DEFAULT true
);


--
-- Name: tipo_categorizacao; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tipo_categorizacao (
    id_tipo_categorizacao integer NOT NULL,
    descricao character varying(255)
);


--
-- Name: tipo_categorizacao_id_tipo_categorizacao_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tipo_categorizacao_id_tipo_categorizacao_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tipo_categorizacao_id_tipo_categorizacao_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tipo_categorizacao_id_tipo_categorizacao_seq OWNED BY public.tipo_categorizacao.id_tipo_categorizacao;


--
-- Name: tipo_resposta; Type: TABLE; Schema: public; Owner: -
--

CREATE TABLE public.tipo_resposta (
    id_tipo_resposta integer NOT NULL,
    descricao character varying(255)
);


--
-- Name: tipo_resposta_id_tipo_resposta_seq; Type: SEQUENCE; Schema: public; Owner: -
--

CREATE SEQUENCE public.tipo_resposta_id_tipo_resposta_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


--
-- Name: tipo_resposta_id_tipo_resposta_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: -
--

ALTER SEQUENCE public.tipo_resposta_id_tipo_resposta_seq OWNED BY public.tipo_resposta.id_tipo_resposta;


--
-- Name: categoria id_categoria; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.categoria ALTER COLUMN id_categoria SET DEFAULT nextval('public.categoria_id_categoria_seq'::regclass);


--
-- Name: implementacao id_implementacao; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.implementacao ALTER COLUMN id_implementacao SET DEFAULT nextval('public.implementacao_id_implementacao_seq'::regclass);


--
-- Name: keyword id_keyword; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.keyword ALTER COLUMN id_keyword SET DEFAULT nextval('public.keyword_id_keyword_seq'::regclass);


--
-- Name: pdf_documents pdf_id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.pdf_documents ALTER COLUMN pdf_id SET DEFAULT nextval('public.pdf_documents_pdf_id_seq'::regclass);


--
-- Name: rag_chunks chunk_id; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.rag_chunks ALTER COLUMN chunk_id SET DEFAULT nextval('public.rag_chunks_chunk_id_seq'::regclass);


--
-- Name: resposta_gerada id_resposta; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.resposta_gerada ALTER COLUMN id_resposta SET DEFAULT nextval('public.resposta_gerada_id_resposta_seq'::regclass);


--
-- Name: tipo_categorizacao id_tipo_categorizacao; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tipo_categorizacao ALTER COLUMN id_tipo_categorizacao SET DEFAULT nextval('public.tipo_categorizacao_id_tipo_categorizacao_seq'::regclass);


--
-- Name: tipo_resposta id_tipo_resposta; Type: DEFAULT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tipo_resposta ALTER COLUMN id_tipo_resposta SET DEFAULT nextval('public.tipo_resposta_id_tipo_resposta_seq'::regclass);


--
-- Name: categoria categoria_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.categoria
    ADD CONSTRAINT categoria_pkey PRIMARY KEY (id_categoria);


--
-- Name: email email_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.email
    ADD CONSTRAINT email_pkey PRIMARY KEY (message_id);


--
-- Name: implementacao implementacao_email_key; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.implementacao
    ADD CONSTRAINT implementacao_email_key UNIQUE (email);


--
-- Name: implementacao implementacao_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.implementacao
    ADD CONSTRAINT implementacao_pkey PRIMARY KEY (id_implementacao);


--
-- Name: keyword keyword_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.keyword
    ADD CONSTRAINT keyword_pkey PRIMARY KEY (id_keyword);


--
-- Name: pdf_documents pdf_documents_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.pdf_documents
    ADD CONSTRAINT pdf_documents_pkey PRIMARY KEY (pdf_id);


--
-- Name: rag_chunks rag_chunks_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.rag_chunks
    ADD CONSTRAINT rag_chunks_pkey PRIMARY KEY (chunk_id);


--
-- Name: resposta_gerada resposta_gerada_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.resposta_gerada
    ADD CONSTRAINT resposta_gerada_pkey PRIMARY KEY (id_resposta);


--
-- Name: thread_categorizacao thread_categorizacao_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.thread_categorizacao
    ADD CONSTRAINT thread_categorizacao_pkey PRIMARY KEY (id_categoria, thread_id);


--
-- Name: thread thread_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.thread
    ADD CONSTRAINT thread_pkey PRIMARY KEY (thread_id);


--
-- Name: tipo_categorizacao tipo_categorizacao_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tipo_categorizacao
    ADD CONSTRAINT tipo_categorizacao_pkey PRIMARY KEY (id_tipo_categorizacao);


--
-- Name: tipo_resposta tipo_resposta_pkey; Type: CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.tipo_resposta
    ADD CONSTRAINT tipo_resposta_pkey PRIMARY KEY (id_tipo_resposta);


--
-- Name: implementacao trigger_inserir_categoria_outro; Type: TRIGGER; Schema: public; Owner: -
--

CREATE TRIGGER trigger_inserir_categoria_outro AFTER INSERT ON public.implementacao FOR EACH ROW EXECUTE FUNCTION public.inserir_categoria_outro();


--
-- Name: categoria fk_categoria_implementacao; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.categoria
    ADD CONSTRAINT fk_categoria_implementacao FOREIGN KEY (id_implementacao) REFERENCES public.implementacao(id_implementacao);


--
-- Name: email fk_email_implementacao; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.email
    ADD CONSTRAINT fk_email_implementacao FOREIGN KEY (id_implementacao) REFERENCES public.implementacao(id_implementacao);


--
-- Name: email fk_email_thread; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.email
    ADD CONSTRAINT fk_email_thread FOREIGN KEY (thread_id) REFERENCES public.thread(thread_id);


--
-- Name: keyword fk_keyword_categoria; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.keyword
    ADD CONSTRAINT fk_keyword_categoria FOREIGN KEY (id_categoria) REFERENCES public.categoria(id_categoria);


--
-- Name: resposta_gerada fk_resposta_email; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.resposta_gerada
    ADD CONSTRAINT fk_resposta_email FOREIGN KEY (message_id) REFERENCES public.email(message_id);


--
-- Name: resposta_gerada fk_resposta_tipo; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.resposta_gerada
    ADD CONSTRAINT fk_resposta_tipo FOREIGN KEY (id_tipo_resposta) REFERENCES public.tipo_resposta(id_tipo_resposta);


--
-- Name: thread fk_thread_implementacao; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.thread
    ADD CONSTRAINT fk_thread_implementacao FOREIGN KEY (id_implementacao) REFERENCES public.implementacao(id_implementacao);


--
-- Name: thread_categorizacao fk_threadcat_categoria; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.thread_categorizacao
    ADD CONSTRAINT fk_threadcat_categoria FOREIGN KEY (id_categoria) REFERENCES public.categoria(id_categoria);


--
-- Name: thread_categorizacao fk_threadcat_thread; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.thread_categorizacao
    ADD CONSTRAINT fk_threadcat_thread FOREIGN KEY (thread_id) REFERENCES public.thread(thread_id);


--
-- Name: thread_categorizacao fk_threadcat_tipo; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.thread_categorizacao
    ADD CONSTRAINT fk_threadcat_tipo FOREIGN KEY (id_tipo_categorizacao) REFERENCES public.tipo_categorizacao(id_tipo_categorizacao);


--
-- Name: rag_chunks rag_chunks_pdf_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: -
--

ALTER TABLE ONLY public.rag_chunks
    ADD CONSTRAINT rag_chunks_pdf_id_fkey FOREIGN KEY (pdf_id) REFERENCES public.pdf_documents(pdf_id) ON DELETE CASCADE;


--
-- PostgreSQL database dump complete
--


INSERT INTO public.tipo_categorizacao (
    id_tipo_categorizacao,
    descricao
)
VALUES
    (1, 'AutomÃ¡tica'),
    (2, 'Manual')
ON CONFLICT (id_tipo_categorizacao) DO NOTHING;


INSERT INTO public.tipo_resposta (
    id_tipo_resposta,
    descricao
)
VALUES
    (1, 'RAG')
ON CONFLICT (id_tipo_resposta) DO NOTHING;
