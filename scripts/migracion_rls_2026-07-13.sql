-- ============================================================
-- E&G ENERGY GROUP — Migración cotizaciones + cierre de RLS
-- Fecha: 2026-07-13
--
-- CÓMO USAR (una sola vez):
--   1. https://supabase.com → proyecto juprjevxkcitqpsnemto
--   2. SQL Editor → New query → pega TODO este archivo → RUN
--   3. Debe terminar con la tabla de verificación de políticas.
--
-- ⚠️ IMPORTANTE: correr esto SOLO cuando el proxy ya tenga la
--    propiedad SUPABASE_SECRET configurada y el Index.html nuevo
--    esté publicado. Desde ese momento la key publishable del
--    HTML queda SOLO-LECTURA; toda escritura entra por el proxy.
-- ============================================================

-- ─── 1) COLUMNAS NUEVAS — cabecera `cotizaciones` ───────────────
-- Campos del registro web que la tabla no conocía. `extra` guarda el
-- registro web COMPLETO (jsonb) para no perder jamás un campo.
alter table cotizaciones add column if not exists motivo_rechazo text;
alter table cotizaciones add column if not exists fecha_envio    timestamptz;
alter table cotizaciones add column if not exists facturada_at   date;
alter table cotizaciones add column if not exists adjudicada_at  date;
alter table cotizaciones add column if not exists factura        text;
alter table cotizaciones add column if not exists solicitud_id   text;
alter table cotizaciones add column if not exists deleted        boolean default false;
alter table cotizaciones add column if not exists created_by     text;
alter table cotizaciones add column if not exists updated_by     text;
alter table cotizaciones add column if not exists extra          jsonb;

-- ─── 2) COLUMNAS NUEVAS — `cotizacion_items` (colaboración por ítem) ───
-- uid = identidad estable de cada ítem (permite que 2 personas guarden
-- ítems distintos de la MISMA cotización sin chocar).
alter table cotizacion_items add column if not exists uid         text;
alter table cotizacion_items add column if not exists foto_grande boolean;
alter table cotizacion_items add column if not exists updated_at  timestamptz default now();
alter table cotizacion_items add column if not exists updated_by  text;
alter table cotizacion_items add column if not exists extra       jsonb;
-- Índice único COMPLETO (sin WHERE): lo exige el upsert on_conflict de PostgREST.
-- Los ítems históricos con uid NULL no chocan entre sí (NULL ≠ NULL en unique).
create unique index if not exists idx_cotitems_cotid_uid
  on cotizacion_items(cotizacion_id, uid);

-- ─── 3) TIMESTAMPS DEL SERVIDOR ─────────────────────────────────
-- updated_at lo pone SIEMPRE la base de datos (elimina la clase de
-- errores de reloj/fecha envenenada de los equipos).
create or replace function set_updated_at() returns trigger
language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end $$;

drop trigger if exists trg_cot_updated on cotizaciones;
create trigger trg_cot_updated before update on cotizaciones
  for each row execute function set_updated_at();

drop trigger if exists trg_cotitems_updated on cotizacion_items;
create trigger trg_cotitems_updated before update on cotizacion_items
  for each row execute function set_updated_at();

-- ─── 4) CERRAR RLS — las 11 tablas de la plataforma ─────────────
-- Borra TODAS las políticas actuales (las "using(true)" que permitían
-- hasta DELETE con la key pública) y deja SOLO lectura (select).
-- Las escrituras quedan reservadas a la key secreta del proxy
-- (service role: se salta la RLS).
do $$
declare
  t text;
  p record;
begin
  foreach t in array array[
    'productos','kardex','familias','conteos','conteo_items',
    'remisiones','cotizaciones','cotizacion_items',
    'plan_compras','oc_compras','proveedores'
  ] loop
    execute format('alter table public.%I enable row level security', t);
    for p in
      select policyname from pg_policies
      where schemaname = 'public' and tablename = t
    loop
      execute format('drop policy %I on public.%I', p.policyname, t);
    end loop;
    execute format(
      'create policy %I on public.%I for select using (true)',
      t || '_solo_lectura', t
    );
  end loop;
end $$;

-- ─── 5) REALTIME — publicar cambios de cotizaciones e ítems ─────
-- (si la tabla ya estaba en la publicación, se ignora el error)
do $$
begin
  begin
    alter publication supabase_realtime add table cotizaciones;
  exception when duplicate_object then null;
  end;
  begin
    alter publication supabase_realtime add table cotizacion_items;
  exception when duplicate_object then null;
  end;
end $$;

-- Realtime DELETE llega con el registro completo solo con replica identity full
alter table cotizaciones     replica identity full;
alter table cotizacion_items replica identity full;

-- ─── 6) VERIFICACIÓN ────────────────────────────────────────────
-- Debe salir UNA política por tabla, todas cmd = SELECT.
select tablename, policyname, cmd
from pg_policies
where schemaname = 'public'
order by tablename;
