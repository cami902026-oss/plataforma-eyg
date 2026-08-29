-- ============================================================
-- E&G ENERGY GROUP — Tabla `avisos_enviados`: memoria de qué ya se avisó
--
-- POR QUÉ: el aviso de "cotización en revisión de gerencia" guardaba su
-- anti-duplicado en data/aviso_gerencia_estado.json, un archivo del repo que se
-- escribe con `git push`. Ese push viene fallando desde el 25-ago-2026 porque
-- el equipo empuja commits todo el día y los 3 reintentos no alcanzan. Con el
-- archivo congelado, el workflow (que corre cada 15 min, ~64 veces al día)
-- nunca supo que ya había avisado: LM2055, LM2070 y LM2049-1 salieron por
-- correo una y otra vez, todo el día.
--
-- El aviso de OP nunca tuvo ese problema porque marca `aviso_enviado` en
-- op_eventos, o sea en la base. Esto le da a las cotizaciones lo mismo.
--
-- POR QUÉ GENÉRICA: `tipo` + `clave` en vez de una tabla por aviso. El próximo
-- aviso automático que haya (remisiones sin facturar, OC vencidas, lo que sea)
-- reusa esta tabla sin migración nueva.
--
-- LA CLAVE ESTÁ EN EL UNIQUE: (tipo, clave) único hace el anti-duplicado a
-- prueba de carreras. Si dos corridas coinciden, la segunda choca contra el
-- índice y no manda el correo — que es justo lo que se quiere. Un archivo en
-- disco nunca pudo dar esa garantía.
--
-- CÓMO CORRERLO (una sola vez):
--   1. https://supabase.com  ->  proyecto juprjevxkcitqpsnemto
--   2. SQL Editor  ->  New query
--   3. Pega esto  ->  RUN
--   4. Debe decir "Success. No rows returned"
--
-- Es ADITIVO e IDEMPOTENTE: se puede correr dos veces sin daño y no toca nada
-- existente. Al final siembra las 13 cotizaciones que hoy están en revisión,
-- para que al prenderlo NO le entre un correo por cada una a gerencia.
-- ============================================================

create table if not exists public.avisos_enviados (
  id         bigserial primary key,
  tipo       text        not null,
  clave      text        not null,
  enviado_at timestamptz not null default now(),
  detalle    text
);

comment on table public.avisos_enviados is
  'Memoria de avisos automáticos ya enviados, para no repetirlos. tipo = qué aviso (ej. cotiz_gerencia); clave = a qué se refiere (ej. LM2055).';

-- El corazón del anti-duplicado. Sin esto la tabla no sirve de nada.
create unique index if not exists avisos_enviados_tipo_clave_idx
  on public.avisos_enviados (tipo, clave);

-- Siembra: las que hoy están en revisión ya fueron avisadas (con creces).
insert into public.avisos_enviados (tipo, clave, detalle)
select 'cotiz_gerencia', c, 'sembrada en la migración del 29-ago-2026'
  from unnest(array[
    'LM1707','LM1711-1','LM1731','LM1921','LM1956-1','LM1996','LM1999',
    'LM2045','LM2048','LM2049','LM2049-1','LM2055','LM2070'
  ]) as c
on conflict (tipo, clave) do nothing;

-- Verificación: debe devolver 13
select count(*) as sembradas
  from public.avisos_enviados
 where tipo = 'cotiz_gerencia';
