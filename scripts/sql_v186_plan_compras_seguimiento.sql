-- ════════════════════════════════════════════════════════════════
-- v186 · Plan de Compras — seguimiento de compra por línea
-- Ejecutar UNA sola vez en Supabase → SQL Editor → RUN
-- Es 100% aditivo: no toca ni borra nada de lo que ya existe.
-- ════════════════════════════════════════════════════════════════

-- ✔ Compró: la línea ya se le pidió al proveedor (OC puesta)
alter table plan_compras add column if not exists comprado boolean not null default false;

-- Recibida: cuánto de esa línea ya llegó a bodega (parciales).
-- Cuando iguala la cantidad, la línea queda COMPLETA.
alter table plan_compras add column if not exists recibida numeric;

-- (la columna `cotizacion` ya existía: es la que permite que un mismo plan
--  junte ítems de varias cotizaciones, cada línea recordando de cuál vino)

-- Comprobación rápida
select column_name, data_type
from information_schema.columns
where table_name = 'plan_compras'
  and column_name in ('comprado','recibida','cotizacion')
order by column_name;
