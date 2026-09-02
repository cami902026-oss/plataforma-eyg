-- ============================================================
-- E&G ENERGY GROUP — REMISIONES: ENTREGAS PARCIALES Y DEFINITIVA (v240)
--
-- POR QUÉ: los envíos parciales YA se hacen, a mano y de dos formas distintas.
-- Medido el 1-sep sobre las 308 remisiones:
--   · 26 órdenes tienen más de una remisión.
--   · 21 remisiones usan sufijo (26266 → 26266-1); otras veces la segunda
--     entrega toma un consecutivo nuevo (26005 → 26006). No hay convención.
--   · 32 remisiones tienen líneas en CANTIDAD 0: así se marca hoy "esto no va
--     en este envío". El cliente firma un papel con renglones en cero.
--   · Que una entrega sea la última vive en texto libre ("SE ENTREGA ORDEN A
--     SATISFACCION"), y solo 9 de 308 lo dicen. En el resto no queda rastro.
--
-- QUÉ HACE: cuatro columnas para que la entrega parcial sea un dato y no una
-- frase. Nada más. Ninguna remisión existente cambia.
--
-- 🔑 POR QUÉ NO BASTA EL NÚMERO: el sufijo «-1» hoy significa DOS cosas —
-- segunda entrega, y también renombre de una remisión (de los 21 sufijos, 19
-- tienen número base y 2 no). El número no puede distinguirlas; el dato sí:
-- una entrega parcial llena `remision_base` y `entrega_n`, un renombre no.
--
-- ADITIVO E IDEMPOTENTE. Si no se corre, la plataforma NO se rompe: el flujo
-- de entregas parciales avisa que falta la migración y el despacho sigue
-- funcionando como hasta hoy.
-- ============================================================

alter table public.remisiones
  add column if not exists remision_base   text,
  add column if not exists entrega_n       integer,
  add column if not exists es_definitiva   boolean not null default false,
  add column if not exists motivo_cierre   text,
  add column if not exists cantidad_pedida numeric;

comment on column public.remisiones.remision_base is
  'Número de la PRIMERA entrega de la orden. Amarra la familia: 26290, 26290-1, '
  '26290-2 comparten remision_base = 26290. Vacío = remisión de una sola entrega.';
comment on column public.remisiones.entrega_n is
  '1 para la primera entrega, 2 para la segunda… Se llena solo desde el flujo de '
  'entregas parciales, nunca al renombrar: es lo que distingue una cosa de la otra.';
comment on column public.remisiones.es_definitiva is
  'true = con esta entrega se cierra la orden. El papel sale marcado DEFINITIVA '
  'con el cuadro de lo entregado en todas las entregas y el saldo.';
comment on column public.remisiones.motivo_cierre is
  'Por qué se cerró quedando saldo (el cliente desistió, se canceló…). Obligatorio '
  'para marcar definitiva con pendientes. Sale impreso en el papel.';
comment on column public.remisiones.cantidad_pedida is
  'Cuánto pidió el cliente EN TOTAL de esa línea. Solo se usa en remisiones que NO '
  'nacen de una OP: ahí no hay de dónde sacar el pedido y se escribe una vez, en la '
  'primera entrega. Las que vienen de una OP sacan el pedido de op_items.';

-- Recorrer una familia de entregas es la consulta nueva más frecuente.
create index if not exists remisiones_base_idx
  on public.remisiones (remision_base)
  where remision_base is not null;

-- Comprobación
select column_name, data_type, column_default
  from information_schema.columns
 where table_schema = 'public' and table_name = 'remisiones'
   and column_name in ('remision_base', 'entrega_n', 'es_definitiva', 'motivo_cierre')
 order by column_name;
