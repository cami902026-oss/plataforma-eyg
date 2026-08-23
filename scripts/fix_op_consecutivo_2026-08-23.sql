-- ============================================================================
-- OP — blindaje del consecutivo   ·   23-ago-2026
--
-- QUÉ PASÓ: durante las pruebas se reinició `op_consecutivos` mientras el equipo
-- ya estaba creando OP reales. El contador quedó DETRÁS del número más alto que
-- existía, así que la función entregaba un número ya usado y la base lo rechazaba
-- con 409 (numero es único). El síntoma que vio el usuario: "error 409 al borrar".
--
-- QUÉ HACE ESTO:
--   1. Pone el contador por encima del número más alto que exista (idempotente).
--   2. Cambia `op_nuevo_numero()` para que SALTE los números ocupados en vez de
--      devolverlos. Así, aunque el contador se desincronice otra vez —por una
--      restauración, una carga manual o una prueba— nunca entrega un duplicado.
--
-- Un número saltado NO se reutiliza: en una serie de documentos un hueco vale
-- más que un consecutivo repetido.
--
-- CÓMO USAR: Supabase → SQL Editor → New query → pegar → RUN.
-- Es seguro correrlo varias veces.
-- ============================================================================

-- 1) Alinear el contador con la realidad
insert into op_consecutivos(anio, ultimo)
select  extract(year from (now() at time zone 'America/Bogota'))::int, 0
on conflict (anio) do nothing;

update op_consecutivos c
   set ultimo = greatest(
         c.ultimo,
         coalesce((select max((regexp_match(o.numero, '^OP-' || c.anio || '-(\d+)$'))[1]::int)
                     from ops o
                    where o.numero ~ ('^OP-' || c.anio || '-\d+$')), 0))
 where c.anio = extract(year from (now() at time zone 'America/Bogota'))::int;

-- 2) La función salta lo ocupado
create or replace function op_nuevo_numero()
returns text
language plpgsql
as $$
declare
  a     int := extract(year from (now() at time zone 'America/Bogota'))::int;
  n     int;
  cand  text;
  vuelt int := 0;
begin
  loop
    insert into op_consecutivos(anio, ultimo) values (a, 1)
      on conflict (anio) do update set ultimo = op_consecutivos.ultimo + 1
      returning ultimo into n;

    cand := 'OP-' || a || '-' || lpad(n::text, 4, '0');
    exit when not exists (select 1 from ops where numero = cand);

    -- El número estaba ocupado: se descarta y se sigue. Freno por si acaso,
    -- para que un dato raro no deje esto girando para siempre.
    vuelt := vuelt + 1;
    if vuelt > 500 then
      raise exception 'op_nuevo_numero: 500 números seguidos ocupados a partir de %. Revisa op_consecutivos.', cand;
    end if;
  end loop;

  return cand;
end $$;

-- 3) Comprobación (no gasta número)
select 'CONTADOR'      as chequeo,
       (select ultimo::text from op_consecutivos
         where anio = extract(year from (now() at time zone 'America/Bogota'))::int) as valor
union all
select 'OP MÁS ALTA',   coalesce((select max(numero) from ops), '(ninguna)')
union all
select 'PRÓXIMA SERÍA', 'OP-' || extract(year from (now() at time zone 'America/Bogota'))::int || '-'
       || lpad(((select ultimo from op_consecutivos
                  where anio = extract(year from (now() at time zone 'America/Bogota'))::int) + 1)::text, 4, '0');
