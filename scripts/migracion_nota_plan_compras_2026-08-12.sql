-- ============================================================
-- E&G ENERGY GROUP — Plan de Compras: columna «nota»
-- (desviación técnica que baja de la cotización, v174)
--
-- POR QUÉ: al traer una cotización al Plan de Compras, la DESVIACIÓN TÉCNICA
-- de cada ítem ahora baja con la línea y se imprime en la OC del proveedor,
-- en la columna "Referencia, Especificaciones o datos de ingeniería" del
-- formato E6-FC-01 (que hasta hoy salía vacía). Así el proveedor despacha lo
-- que se le ofreció al cliente y no un equivalente por su cuenta.
--
-- MIENTRAS NO SE CORRA: la plataforma NO se rompe. Al guardar reintenta sin la
-- columna y avisa que la desviación no se conservó. El texto igual se ve al
-- armar el plan y sale impreso en la OC (oc_compras.items es jsonb y no
-- necesita migración) — lo único que falta es que sobreviva al reabrir el plan.
--
-- CÓMO CORRERLO (una sola vez):
--   1. Entra a https://supabase.com  ->  proyecto juprjevxkcitqpsnemto
--   2. Menú izquierdo:  SQL Editor  ->  New query
--   3. Pega esto  ->  botón RUN
--   4. Debe decir "Success. No rows returned"
--
-- Es ADITIVO: no toca datos existentes. Los planes ya guardados quedan con
-- nota vacía, que es justo lo que corresponde.
-- ============================================================

alter table plan_compras
  add column if not exists nota text;

comment on column plan_compras.nota is
  'Desviación técnica / nota del ítem. Baja de la cotización (cotizacion_items.desviacion_tecnica) al crear el plan y se imprime en la OC del proveedor (E6-FC-01, columna de especificaciones).';
