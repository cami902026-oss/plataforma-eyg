-- ============================================================
-- E&G ENERGY GROUP — Cotizaciones en Supabase
-- Tablas: cotizaciones (cabecera) + cotizacion_items (lineas)
--
-- COMO USAR (una sola vez):
--   1. Entra a https://supabase.com  ->  tu proyecto (juprjevxkcitqpsnemto)
--   2. Menu izquierdo:  SQL Editor  ->  New query
--   3. Pega TODO este archivo  ->  boton  RUN
--   4. Debe decir "Success. No rows returned"
-- ============================================================

-- ---------- CABECERA ----------
create table if not exists cotizaciones (
  id            text primary key,         -- ej. 'LM1718'
  fecha         date,
  fecha_venc    date,
  cliente       text,
  contacto      text,
  correo        text,
  estado        text,
  subtotal      numeric,
  iva           numeric,
  total         numeric,
  forma_pago    text,
  sitio_entrega text,
  validez       text,
  observaciones text,
  vendedor      text,
  realizada_por text,
  aprobada_por  text,
  fuente        text,                     -- 'Plataforma' | 'LIBRO' | 'Historico'
  created_at    timestamptz default now(),
  updated_at    timestamptz default now()
);

-- ---------- LINEAS / ITEMS ----------
create table if not exists cotizacion_items (
  id                 bigint generated always as identity primary key,
  cotizacion_id      text references cotizaciones(id) on delete cascade,
  item               int,
  descripcion        text,
  udm                text,
  qty                numeric,
  v_unit             numeric,
  v_total            numeric,
  marca              text,
  proveedor          text,
  precio_proveedor   numeric,
  tiempo_entrega     text,
  desviacion_tecnica text
);

create index if not exists idx_cotitems_cotid  on cotizacion_items(cotizacion_id);
create index if not exists idx_cot_fecha        on cotizaciones(fecha);
create index if not exists idx_cot_cliente      on cotizaciones(cliente);

-- ---------- PERMISOS (mismo esquema que tus otras tablas: clave publishable) ----------
alter table cotizaciones      enable row level security;
alter table cotizacion_items  enable row level security;

drop policy if exists "cot_all"      on cotizaciones;
drop policy if exists "cotitems_all" on cotizacion_items;

create policy "cot_all"      on cotizaciones      for all using (true) with check (true);
create policy "cotitems_all" on cotizacion_items  for all using (true) with check (true);
