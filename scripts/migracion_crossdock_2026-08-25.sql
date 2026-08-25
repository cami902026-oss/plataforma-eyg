-- ============================================================================
-- E&G ENERGY GROUP — Recibir y despachar en un solo acto (cross-docking)
-- Fecha: 2026-08-25
--
-- POR QUÉ
-- Buena parte de lo que se compra en Bogotá no se queda en Chía: llega y sale
-- derecho para el cliente, a veces sin pasar por la bodega. Hoy eso obliga a
-- hacer dos actos separados —recibir y después despachar— o a no registrar
-- nada, y entonces la compra no aparece en ningún análisis.
--
-- Ahora la pantalla de "Recibir en bodega" tiene una columna más: SALE DE UNA.
-- Entra todo lo que llegó, sale lo que se despacha, y el saldo queda en
-- inventario. Se registran DOS movimientos de verdad, no uno "neto":
--   · la ENTRADA dice a quién se le compró, a qué precio y con qué factura
--   · la SALIDA dice a quién se entregó y con qué remisión
-- El stock neto no se mueve, pero la huella queda completa.
--
-- QUÉ FALTA EN LA BASE
-- Un ítem puede salir por partes: pidieron 10, llegaron 6 y salieron 6, y
-- después llegan y salen los otros 4. Sin llevar la cuenta de cuánto ha salido,
-- la línea nunca queda bien cerrada. Eso es esta columna.
--
-- CÓMO USAR: Supabase → SQL Editor → New query → pegar → RUN.
-- Es aditivo y se puede correr varias veces.
-- ============================================================================

alter table op_items add column if not exists despachada numeric;

-- Lo que ya salió antes de hoy: si la línea quedó despachada, salió completa.
-- No se inventa nada — solo se pone al día lo que el propio estado ya decía.
update op_items
   set despachada = cantidad
 where despachada is null and estado = 'despachado';


-- ─── VERIFICACIÓN ───────────────────────────────────────────────────────────
select 'op_items.despachada' as chequeo,
       coalesce((select 'creada' from information_schema.columns
                  where table_schema='public' and table_name='op_items'
                    and column_name='despachada'), 'FALTA') as detalle
union all
select 'líneas ya despachadas con su cantidad',
       count(*)::text from op_items where despachada is not null;
