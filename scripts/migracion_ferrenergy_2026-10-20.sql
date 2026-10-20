-- ═══════════════════════════════════════════════════════════════════════════
--  FERRENERGY — 20-oct-2026
--  Local de ferretería + comerciales que visitan ferreterías. Es OTRO negocio
--  dentro de la misma plataforma: tablas propias (prefijo fe_), inventario
--  propio (bodega aparte, no toca productos/kardex industrial), remisiones con
--  consecutivo propio FE-0001.
--  Se corre UNA vez en Supabase → SQL Editor → RUN. Es idempotente.
-- ═══════════════════════════════════════════════════════════════════════════

-- 1) Catálogo + lista de precios (3 niveles: público, ferretería, mayorista)
create table if not exists fe_productos (
  id                bigserial primary key,
  codigo            text not null unique,          -- código corto (FE-0001 o el del proveedor)
  descripcion       text not null,
  categoria         text,                          -- Tubería, Eléctricos, Herramienta, Pintura…
  marca             text,
  unidad            text not null default 'UND',   -- UND, MT, KG, CAJA, PAR…
  costo             numeric(16,2) not null default 0,   -- último costo de compra, ANTES de IVA
  precio_publico    numeric(16,2) not null default 0,
  precio_ferreteria numeric(16,2) not null default 0,
  precio_mayorista  numeric(16,2) not null default 0,
  stock_min         numeric(14,3) not null default 0,   -- avisa cuando el stock baja de aquí
  proveedor         text,
  ubicacion         text,                          -- estante / pasillo en el local
  activo            boolean not null default true,
  notas             text,
  created_by        text,
  created_at        timestamptz not null default now(),
  updated_by        text,
  updated_at        timestamptz not null default now()
);
create index if not exists idx_fe_prod_desc on fe_productos (lower(descripcion));
create index if not exists idx_fe_prod_cat  on fe_productos (categoria);

-- 2) Kardex propio: cada movimiento con signo. El stock es la suma.
create table if not exists fe_kardex (
  id           bigserial primary key,
  producto_id  bigint not null references fe_productos(id) on delete restrict,
  fecha        date not null default current_date,
  tipo         text not null default 'entrada',    -- inicial / entrada / salida / ajuste / anulacion
  cantidad     numeric(14,3) not null,             -- CON SIGNO: entrada +, salida −
  costo_unit   numeric(16,2),                      -- costo de compra en entradas (antes de IVA)
  referencia   text,                               -- FE-0001, factura del proveedor, conteo…
  motivo       text,
  usuario      text,
  created_at   timestamptz not null default now()
);
create index if not exists idx_fe_kardex_prod on fe_kardex (producto_id);
create index if not exists idx_fe_kardex_ref  on fe_kardex (referencia);

-- 3) Clientes (ferreterías) con su zona, día de visita y comercial
create table if not exists fe_clientes (
  id           bigserial primary key,
  nombre       text not null,
  nit          text,
  contacto     text,
  telefono     text,
  email        text,
  direccion    text,
  ciudad       text,
  zona         text,                               -- ruta / barrio / municipio
  dia_visita   text,                               -- Lunes…Sábado
  comercial    text,                               -- usuario que la atiende (vivian…)
  lista_precio text not null default 'ferreteria', -- publico / ferreteria / mayorista
  cupo         numeric(16,2) not null default 0,   -- cupo de crédito (0 = de contado)
  activo       boolean not null default true,
  notas        text,
  created_by   text,
  created_at   timestamptz not null default now(),
  updated_by   text,
  updated_at   timestamptz not null default now()
);
create index if not exists idx_fe_cli_nombre on fe_clientes (lower(nombre));
create index if not exists idx_fe_cli_nit    on fe_clientes (nit);

-- 4) Pedidos de los comerciales (lo que toman en la visita) → luego se remisionan
create table if not exists fe_pedidos (
  id            bigserial primary key,
  numero        text not null unique,              -- PE-0001
  fecha         date not null default current_date,
  cliente_id    bigint references fe_clientes(id) on delete set null,
  cliente       text,                              -- nombre al momento (por si cambia)
  comercial     text,
  estado        text not null default 'pendiente', -- pendiente / remisionado / anulado
  observaciones text,
  total         numeric(16,2) not null default 0,  -- antes de IVA
  remision      text,                              -- FE-0001 cuando ya salió
  created_by    text,
  created_at    timestamptz not null default now(),
  updated_by    text,
  updated_at    timestamptz not null default now()
);
create table if not exists fe_pedido_items (
  id           bigserial primary key,
  pedido_id    bigint not null references fe_pedidos(id) on delete cascade,
  producto_id  bigint references fe_productos(id) on delete set null,
  codigo       text,
  descripcion  text not null,
  unidad       text,
  cantidad     numeric(14,3) not null default 1,
  precio       numeric(16,2) not null default 0,   -- precio unitario según la lista del cliente
  total        numeric(16,2) not null default 0
);
create index if not exists idx_fe_ped_items on fe_pedido_items (pedido_id);

-- 5) Remisiones FE-0001: lo que sale del local. Descuenta fe_kardex.
create table if not exists fe_remisiones (
  id            bigserial primary key,
  numero        text not null unique,              -- FE-0001
  fecha         date not null default current_date,
  cliente_id    bigint references fe_clientes(id) on delete set null,
  cliente       text not null,
  nit           text,
  direccion     text,
  pedido_id     bigint references fe_pedidos(id) on delete set null,
  oc            text,                              -- orden de compra del cliente, si la hay
  forma_pago    text not null default 'contado',   -- contado / credito
  observaciones text,
  estado        text not null default 'emitida',   -- emitida / anulada
  subtotal      numeric(16,2) not null default 0,
  iva           numeric(16,2) not null default 0,
  total         numeric(16,2) not null default 0,
  created_by    text,
  created_at    timestamptz not null default now(),
  updated_by    text,
  updated_at    timestamptz not null default now()
);
create table if not exists fe_remision_items (
  id           bigserial primary key,
  remision_id  bigint not null references fe_remisiones(id) on delete cascade,
  producto_id  bigint references fe_productos(id) on delete set null,
  codigo       text,
  descripcion  text not null,
  unidad       text,
  cantidad     numeric(14,3) not null default 1,
  precio       numeric(16,2) not null default 0,
  total        numeric(16,2) not null default 0
);
create index if not exists idx_fe_rem_items on fe_remision_items (remision_id);
create index if not exists idx_fe_rem_cli   on fe_remisiones (cliente_id);

-- 6) Consecutivos por serie (FE = remisiones, PE = pedidos). Misma idea que op_consecutivos.
create table if not exists fe_consecutivos (
  serie   text primary key,
  ultimo  int not null default 0
);
create or replace function fe_nuevo_numero(p_serie text)
returns text
language plpgsql
as $$
declare
  n     int;
  cand  text;
  vuelt int := 0;
begin
  loop
    insert into fe_consecutivos(serie, ultimo) values (p_serie, 1)
      on conflict (serie) do update set ultimo = fe_consecutivos.ultimo + 1
      returning ultimo into n;
    cand := p_serie || '-' || lpad(n::text, 4, '0');
    if p_serie = 'FE' then
      exit when not exists (select 1 from fe_remisiones where numero = cand);
    else
      exit when not exists (select 1 from fe_pedidos where numero = cand);
    end if;
    vuelt := vuelt + 1;
    if vuelt > 500 then
      raise exception 'fe_nuevo_numero: 500 números seguidos ocupados en la serie %', p_serie;
    end if;
  end loop;
  return cand;
end;
$$;

-- 7) Vista de stock: suma del kardex por producto + costo promedio de las entradas.
create or replace view fe_stock as
select p.id as producto_id,
       coalesce(sum(k.cantidad), 0)                                   as stock,
       coalesce(sum(case when k.cantidad > 0 and k.costo_unit is not null then k.cantidad * k.costo_unit end)
                / nullif(sum(case when k.cantidad > 0 and k.costo_unit is not null then k.cantidad end), 0), p.costo) as costo_prom,
       max(k.fecha)                                                   as ultimo_mov
from fe_productos p
left join fe_kardex k on k.producto_id = p.id
group by p.id, p.costo;

-- 8) RLS igual que el resto: lectura con la key pública, escritura SOLO por el proxy.
do $$
declare t text; p record;
begin
  foreach t in array array['fe_productos','fe_kardex','fe_clientes','fe_pedidos','fe_pedido_items',
                           'fe_remisiones','fe_remision_items','fe_consecutivos'] loop
    execute format('alter table public.%I enable row level security', t);
    for p in select policyname from pg_policies where schemaname='public' and tablename=t loop
      execute format('drop policy %I on public.%I', p.policyname, t);
    end loop;
    execute format('create policy %I on public.%I for select using (true)', t || '_solo_lectura', t);
  end loop;
end $$;

-- Comprobación
select 'fe_productos' as tabla, count(*) from fe_productos
union all select 'fe_clientes', count(*) from fe_clientes
union all select 'fe_remisiones', count(*) from fe_remisiones;
