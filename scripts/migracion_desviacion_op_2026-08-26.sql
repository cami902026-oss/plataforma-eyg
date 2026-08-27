-- ============================================================
-- E&G ENERGY GROUP — OP: columna «desviacion_tecnica» en op_items
-- (la desviación técnica de la cotización vuelve a llegar a la O.C., v223)
--
-- POR QUÉ: desde que el Plan de Compras se genera DESDE LA OP y ya no desde la
-- cotización, la desviación técnica se perdía dos veces:
--   1. `op_items` no tenía dónde guardarla y al crear la OP no se copiaba.
--   2. `opGenerarPlan` escribía `nota: null` en las líneas de compra.
-- Resultado medido el 26-ago-2026: 14 de las 23 líneas de los 6 planes nacidos
-- de una OP salieron SIN la desviación (61%). Entre ellas "SE OFERTA TUBO X 6
-- MTRS" (×3, LM1645), "LAM HR 6.0MM 1200 x 2400MM" (LM1831) y "SS-304"
-- (LM2015). Esa nota es la que imprime la O.C. en la columna "Referencia,
-- Especificaciones o datos de ingeniería" del E6-FC-01: sin ella el proveedor
-- despacha lo que le parece y no lo que se le ofreció al cliente.
--
-- Es la misma corrección que ya se había hecho en v182 para el camino viejo
-- (`_pcomLineasDeCotiz`); el camino nuevo por OP no la había heredado.
--
-- MIENTRAS NO SE CORRA: la plataforma NO se rompe, pero la desviación sigue
-- perdiéndose igual que hoy. El INSERT de op_items manda el campo; sin la
-- columna, Supabase responde 400 y la OP no se crea desde la cotización.
--   >>> Correr ANTES de que salga la v223. <<<
--
-- CÓMO CORRERLO (una sola vez):
--   1. Entra a https://supabase.com  ->  proyecto juprjevxkcitqpsnemto
--   2. Menú izquierdo:  SQL Editor  ->  New query
--   3. Pega esto  ->  botón RUN
--   4. Debe decir "Success. No rows returned"
--
-- Es ADITIVO e IDEMPOTENTE: se puede correr dos veces sin daño. No toca datos
-- existentes; las OP ya creadas quedan con la desviación vacía y se rellenan
-- aparte con scripts/backfill_desviacion_op.py
-- ============================================================

alter table public.op_items
  add column if not exists desviacion_tecnica text;

comment on column public.op_items.desviacion_tecnica is
  'Desviación técnica que viene del ítem de la cotización (cotizacion_items.desviacion_tecnica). Baja al Plan de Compras como parte de la nota y se imprime en la O.C. del proveedor (E6-FC-01, columna "Referencia, Especificaciones o datos de ingeniería").';

-- Verificación: debe devolver una fila con desviacion_tecnica / text
select column_name, data_type
  from information_schema.columns
 where table_schema = 'public'
   and table_name   = 'op_items'
   and column_name  = 'desviacion_tecnica';
