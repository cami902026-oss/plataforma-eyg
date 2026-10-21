-- ═══════════════════════════════════════════════════════════════════════════
--  FERREINDUSTRIAL — alistamiento de pedidos (21-oct-2026)
--  Estado nuevo «alistando» (quién y cuándo tomó el pedido) y FALTANTES: lo que
--  se pidió y no salió en la remisión. Se corre UNA vez en Supabase → SQL Editor.
--  Es idempotente y no toca datos existentes.
-- ═══════════════════════════════════════════════════════════════════════════
alter table fe_pedidos
  add column if not exists alistado_por text,
  add column if not exists alistado_at  timestamptz,
  add column if not exists faltantes    jsonb;   -- [{codigo, descripcion, pedida, enviada}]

-- Comprobación
select column_name from information_schema.columns
where table_name = 'fe_pedidos' and column_name in ('alistado_por', 'alistado_at', 'faltantes');
