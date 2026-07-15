-- Required initial records used by the application.

INSERT INTO public.tipo_categorizacao (
    id_tipo_categorizacao,
    descricao
)
VALUES
    (1, 'Automática'),
    (2, 'Manual')
ON CONFLICT (id_tipo_categorizacao) DO NOTHING;


INSERT INTO public.tipo_resposta (
    id_tipo_resposta,
    descricao
)
VALUES
    (1, 'RAG')
ON CONFLICT (id_tipo_resposta) DO NOTHING;