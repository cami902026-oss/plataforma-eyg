-- ============================================================================
-- E&G ENERGY GROUP — SQL PENDIENTE  ·  24-ago-2026
--
-- Todo lo que falta correr, en un solo bloque. Es ADITIVO y se puede correr
-- varias veces sin dañar nada: cada paso comprueba antes de actuar.
--
-- CÓMO USAR:  Supabase → SQL Editor → New query → pegar TODO → RUN.
--             Al final sale una tabla de verificación; los tres renglones
--             deben decir "creada" / "OK".
-- ============================================================================


-- ─── 1) La factura que cierra la OP ─────────────────────────────────────────
-- El cierre definitivo del pedido es cuando ya se facturó: ahí terminó el
-- negocio de verdad. La OP ya exige despacho y certificados; falta guardar
-- CUÁL factura la cerró, para que el soporte quede pegado a la OP aunque
-- después alguien toque la cotización.
--
-- El número NO se teclea en la OP: lo trae la cotización cuando el equipo la
-- marca Facturada. Esta columna solo lo copia en el momento del cierre.
alter table ops add column if not exists factura text;


-- ─── 2) Blindaje del consecutivo de OP ──────────────────────────────────────
-- QUÉ PASÓ: durante las pruebas se reinició `op_consecutivos` mientras el
-- equipo ya estaba creando OP reales. El contador quedó DETRÁS del número más
-- alto que existía, la función entregaba un número ya usado y la base lo
-- rechazaba con 409 (numero es único).
--
-- Esto (a) alinea el contador con la realidad y (b) hace que la función SALTE
-- los números ocupados en vez de devolverlos. Un número saltado NO se
-- reutiliza: en una serie de documentos un hueco vale más que un repetido.

insert into op_consecutivos(anio, ultimo)
select extract(year from (now() at time zone 'America/Bogota'))::int, 0
on conflict (anio) do nothing;

update op_consecutivos c
   set ultimo = greatest(
         c.ultimo,
         coalesce((select max((regexp_match(o.numero, '^OP-' || c.anio || '-(\d+)$'))[1]::int)
                     from ops o
                    where o.numero ~ ('^OP-' || c.anio || '-\d+$')), 0))
 where c.anio = extract(year from (now() at time zone 'America/Bogota'))::int;

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

    -- Ese número ya está ocupado: se salta y se pide el siguiente.
    vuelt := vuelt + 1;
    if vuelt > 500 then
      raise exception 'op_nuevo_numero: no se encontró un número libre después de 500 intentos';
    end if;
  end loop;

  return cand;
end $$;


-- ─── 3) Ligar remisiones hechas por fuera de la OP ──────────────────────────
-- El caso MONTITEC: la remisión 26269 se hizo por el módulo de Remisiones y la
-- OP-2026-0004 quedó al lado sin enterarse. La columna ya existe; esto es solo
-- por si la base viniera de una versión anterior. No borra ni cambia datos.
alter table remisiones add column if not exists op_numero text;
create index if not exists idx_remisiones_op on remisiones(op_numero);


-- ─── VERIFICACIÓN ───────────────────────────────────────────────────────────
select 'ops.factura' as chequeo,
       coalesce((select 'creada' from information_schema.columns
                  where table_schema='public' and table_name='ops'
                    and column_name='factura'), 'FALTA') as detalle
union all
select 'op_nuevo_numero() salta ocupados',
       case when exists (select 1 from pg_proc p
                          where p.proname = 'op_nuevo_numero'
                            and pg_get_functiondef(p.oid) like '%no se encontró un número libre%')
            then 'OK' else 'FALTA' end
union all
select 'contador vs OP más alta',
       (select ultimo::text from op_consecutivos
         where anio = extract(year from (now() at time zone 'America/Bogota'))::int)
       || ' / ' ||
       coalesce((select max((regexp_match(numero, '^OP-\d{4}-(\d+)$'))[1]::int)::text
                   from ops where numero ~ '^OP-\d{4}-\d+$'), '0')
union all
select 'remisiones.op_numero',
       coalesce((select 'creada' from information_schema.columns
                  where table_schema='public' and table_name='remisiones'
                    and column_name='op_numero'), 'FALTA');
