-- ============================================================================
-- E&G ENERGY GROUP — Cierre de la OP con su factura
-- Fecha: 2026-08-24
--
-- POR QUÉ
-- El cierre definitivo del pedido es cuando ya se facturó: ahí terminó el
-- negocio de verdad. La OP ya exige despacho y certificados; falta guardar
-- CUÁL factura la cerró, para que el soporte quede pegado a la OP aunque
-- después alguien toque la cotización.
--
-- El número NO se teclea en la OP: lo trae la cotización cuando el equipo la
-- marca Facturada. Esta columna solo lo copia en el momento del cierre.
--
-- CÓMO USAR: Supabase → SQL Editor → New query → pegar → RUN.
-- Es aditivo y se puede correr varias veces.
-- ============================================================================

alter table ops add column if not exists factura text;

-- Verificación
select 'ops.factura' as chequeo,
       coalesce((select 'creada' from information_schema.columns
                  where table_schema='public' and table_name='ops'
                    and column_name='factura'), 'FALTA') as detalle;
