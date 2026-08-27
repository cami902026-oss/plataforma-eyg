-- ============================================================
-- E&G ENERGY GROUP — OP: el ENVÍO como paso propio (v225)
--
-- POR QUÉ: hasta hoy `despachada` significaba dos cosas a la vez:
--   (a) la remisión está hecha y el inventario descontado, y
--   (b) el material salió físicamente para el cliente.
-- Son momentos distintos. En la OP-2026-0020 (ANDES OPERATING) pasó (a) el
-- 27-ago —remisión 26280, los 2 ítems despachados y descontados— pero el
-- material seguía alistado en bodega esperando quién lo llevara. La plataforma
-- no tenía cómo decir eso, así que la OP se veía en un limbo.
--
-- Y con 92 destinos a nivel nacional, POR DÓNDE se mandó no es un detalle: es
-- lo único que permite rastrear un pedido o reclamar cuando se pierde.
--
-- QUÉ AGREGA (4 columnas, ninguna obligatoria):
--   enviada_at     cuándo salió de verdad
--   enviada_por    quién lo confirmó
--   envio_medio    CAMION_EG | BUS | TRANSPORTADORA | CLIENTE_RECOGE | OTRO
--   envio_detalle  empresa, guía, placa, conductor — texto libre
--
-- Una OP con remisión + ítems despachados pero con `enviada_at` en null es una
-- «pendiente de envío»: alistada y sin salir. Ese es el hueco que se destapa.
--
-- NO cambia ningún estado existente. `despachada` sigue significando lo mismo y
-- la máquina de estados queda igual: esto se calcula de lo que ya hay.
--
-- MIENTRAS NO SE CORRA: la plataforma NO se rompe, pero el botón «Confirmar
-- envío» responde 400 al guardar y el chip de pendiente de envío nunca aparece.
--
-- CÓMO CORRERLO (una sola vez):
--   1. https://supabase.com  ->  proyecto juprjevxkcitqpsnemto
--   2. SQL Editor  ->  New query
--   3. Pega esto  ->  RUN
--   4. Debe decir "Success. No rows returned"
--
-- Es ADITIVO e IDEMPOTENTE: se puede correr dos veces sin daño y no toca ni un
-- dato existente. Las OP ya enviadas quedan con el envío vacío; se llena cuando
-- alguien confirme, o se deja así — no estorba.
-- ============================================================

alter table public.ops
  add column if not exists enviada_at    timestamptz,
  add column if not exists enviada_por   text,
  add column if not exists envio_medio   text,
  add column if not exists envio_detalle text;

comment on column public.ops.enviada_at is
  'Cuándo salió FÍSICAMENTE el material para el cliente. Distinto de despachada: esa es la remisión + el descuento de inventario. Null con los ítems ya despachados = pendiente de envío.';
comment on column public.ops.envio_medio is
  'Por dónde se mandó: CAMION_EG | BUS | TRANSPORTADORA | CLIENTE_RECOGE | OTRO.';
comment on column public.ops.envio_detalle is
  'Empresa, número de guía, placa o conductor. Texto libre: es lo que se usa para rastrear o reclamar.';

-- Verificación: deben salir las 4 filas
select column_name, data_type
  from information_schema.columns
 where table_schema = 'public'
   and table_name   = 'ops'
   and column_name in ('enviada_at','enviada_por','envio_medio','envio_detalle')
 order by column_name;
