-- ============================================================
-- E&G ENERGY GROUP — Remisiones: columna «item_cot»
-- (N° del ítem en la cotización adjudicada, v183)
--
-- POR QUÉ: en cotizaciones largas, el cliente y bodega hablan por el NÚMERO DE
-- ÍTEM de la cotización ("me falta el 47"). Al traer los ítems a una remisión se
-- renumeraban desde 1, y más aún en adjudicaciones PARCIALES, donde el ítem 47
-- podía quedar como el 3 de la remisión. Ahora cada línea conserva el número tal
-- como salió en el PDF de la cotización (incluidas las alternativas 3.1 / 3.2) y
-- se ve en el editor, en el PDF de la remisión y en el Excel.
--
-- MIENTRAS NO SE CORRA: la plataforma NO se rompe. Al guardar, la plataforma
-- pregunta una vez si la columna existe; si no está, guarda exactamente igual que
-- antes (sin ese campo). El número se ve en pantalla y sale en el PDF/Excel de esa
-- sesión — lo único que falta es que sobreviva al reabrir la remisión.
--
-- CÓMO CORRERLO (una sola vez):
--   1. Entra a https://supabase.com  ->  proyecto juprjevxkcitqpsnemto
--   2. Menú izquierdo:  SQL Editor  ->  New query
--   3. Pega esto  ->  botón RUN
--   4. Debe decir "Success. No rows returned"
--
-- Es ADITIVO: no toca datos existentes. Las remisiones ya guardadas quedan con
-- item_cot vacío, que es justo lo que corresponde.
-- ============================================================

alter table remisiones
  add column if not exists item_cot text;

comment on column remisiones.item_cot is
  'N° del ítem en la COTIZACIÓN adjudicada (texto: admite 3.1 / 3.2 de las alternativas). Se llena solo al usar "Traer de cotización" y es editable a mano. No confundir con "item", que es el consecutivo de la línea dentro de la remisión.';
