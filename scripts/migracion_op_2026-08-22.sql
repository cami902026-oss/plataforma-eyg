-- ============================================================================
-- E&G ENERGY GROUP — OP (Orden de Pedido) · esqueleto de base de datos
-- Fecha: 2026-08-22
--
-- CÓMO USAR (una sola vez):
--   1. https://supabase.com → proyecto juprjevxkcitqpsnemto
--   2. SQL Editor → New query → pega TODO este archivo → RUN
--   3. Termina imprimiendo una tabla de verificación. Revísala.
--
-- ES ADITIVO: crea tablas NUEVAS y agrega columnas opcionales a `remisiones`
-- y `kardex`. NO modifica ni borra un solo dato existente. Se puede correr
-- varias veces sin romper nada (todo lleva IF NOT EXISTS).
--
-- ⚠️ RLS: se sigue la misma convención del 2026-07-13 — las tablas quedan
--    SOLO LECTURA para la key publishable; toda escritura entra por el proxy
--    con la secret key (service role). Si esto no se respeta, las escrituras
--    fallan EN SILENCIO desde el navegador.
-- ============================================================================


-- ─────────────────────────────────────────────────────────────────────────
-- 1) CONSECUTIVO PROPIO — OP-2026-0001
-- Se reserva en el SERVIDOR, de forma atómica. Es la lección del incidente
-- LM1790: dos personas sin conexión generaron el mismo número y el merge
-- borró una en silencio. Con esto es imposible que dos OP compartan número.
-- ─────────────────────────────────────────────────────────────────────────
create table if not exists op_consecutivos (
  anio   int primary key,
  ultimo int not null default 0
);

create or replace function op_nuevo_numero()
returns text
language plpgsql
as $$
declare
  a int := extract(year from (now() at time zone 'America/Bogota'))::int;
  n int;
begin
  insert into op_consecutivos(anio, ultimo) values (a, 1)
    on conflict (anio) do update set ultimo = op_consecutivos.ultimo + 1
    returning ultimo into n;
  return 'OP-' || a || '-' || lpad(n::text, 4, '0');
end $$;


-- ─────────────────────────────────────────────────────────────────────────
-- 2) OPS — la cabecera
-- Absorbe la O.C. del cliente: su número es una columna de aquí, y se
-- conservan las 4 etapas (Compra · Entrega · Certificado · Facturación)
-- en `stages`, con el mismo formato que hoy usa ordenes.json.
-- ─────────────────────────────────────────────────────────────────────────
create table if not exists ops (
  id                bigserial primary key,
  numero            text not null unique,           -- OP-2026-0001
  cotizacion_id     text not null,                  -- SIEMPRE hay cotización detrás
  oc_cliente        text,                           -- N° de la O.C. del cliente
  oc_archivo        text,                           -- PDF/imagen de la O.C. (Storage)
  cliente           text,
  -- borrador · pendiente_aprobacion · aprobada · devuelta · en_ejecucion
  -- · despachada · cerrada · anulada
  estado            text not null default 'borrador',
  -- Aprobación de gerencia: sin esto NO pasa a compras ni a bodega
  aprobada_por      text,
  aprobada_at       timestamptz,
  devuelta_por      text,
  devuelta_at       timestamptz,
  nota_devolucion   text,                           -- por qué la devolvió
  -- Cierre: la OP NO se puede cerrar sin certificados (ver función de abajo)
  requiere_certificados boolean not null default true,
  cerrada_at        timestamptz,
  cerrada_por       text,
  valor_venta       numeric,                        -- lo adjudicado por el cliente
  costo_compras     numeric,
  stages            jsonb,                          -- 4 etapas heredadas de la O.C.
  observaciones     text,
  deleted           boolean not null default false,
  created_by        text,
  created_at        timestamptz not null default now(),
  updated_by        text,
  updated_at        timestamptz not null default now(),
  extra             jsonb                           -- registro web completo, por si acaso
);
create index if not exists idx_ops_cotizacion on ops(cotizacion_id);
create index if not exists idx_ops_estado     on ops(estado) where deleted = false;
create index if not exists idx_ops_oc_cliente on ops(oc_cliente);


-- ─────────────────────────────────────────────────────────────────────────
-- 3) OP_ITEMS — las líneas (lo que la O.C. de hoy NO tiene)
-- `uid` da identidad estable a cada línea para que dos personas puedan
-- editar ítems distintos de la misma OP sin pisarse (igual que cotizacion_items).
-- ─────────────────────────────────────────────────────────────────────────
create table if not exists op_items (
  id                  bigserial primary key,
  op_id               bigint not null references ops(id) on delete cascade,
  uid                 text not null,
  item                int,
  descripcion         text,
  udm                 text,
  cantidad            numeric,
  v_unit              numeric,                      -- precio de venta
  v_total             numeric,
  -- BODEGA · COMPRA · PENDIENTE (todavía sin decidir)
  origen              text not null default 'PENDIENTE',
  -- Cruce con inventario: se guarda también la confianza del match automático
  -- y QUIÉN lo confirmó, para poder auditar los errores del cruce por texto.
  producto_codigo     text,
  producto_confianza  numeric,
  confirmado_por      text,
  confirmado_at       timestamptz,
  -- Compra
  proveedor           text,
  sede_id             bigint,
  costo_unit          numeric,
  -- pendiente · comprado · recibido · alistado · despachado
  estado              text not null default 'pendiente',
  colada              text,
  lote                text,
  -- Trazabilidad hacia atrás y hacia adelante
  cotizacion_item_uid text,
  remision            text,
  observaciones       text,
  created_at          timestamptz not null default now(),
  updated_at          timestamptz not null default now(),
  updated_by          text
);
create unique index if not exists idx_opitems_op_uid on op_items(op_id, uid);
create index if not exists idx_opitems_op      on op_items(op_id);
create index if not exists idx_opitems_prod    on op_items(producto_codigo);
create index if not exists idx_opitems_estado  on op_items(estado);


-- ─────────────────────────────────────────────────────────────────────────
-- 4) OP_RESERVAS — apartar el stock
-- Una reserva NO toca `productos.stock_actual`: el stock físico sigue igual
-- hasta que la pieza sale de verdad. Lo que baja es el DISPONIBLE, que se
-- calcula como stock_actual − reservas activas (ver vista más abajo).
-- Así una reserva mal hecha nunca corrompe el inventario real.
-- ─────────────────────────────────────────────────────────────────────────
create table if not exists op_reservas (
  id               bigserial primary key,
  op_id            bigint not null references ops(id) on delete cascade,
  op_item_id       bigint references op_items(id) on delete cascade,
  producto_codigo  text not null,
  cantidad         numeric not null,
  estado           text not null default 'activa',   -- activa · consumida · liberada
  creada_por       text,
  creada_at        timestamptz not null default now(),
  liberada_at      timestamptz,
  liberada_por     text,
  motivo           text                              -- por qué se liberó
);
create index if not exists idx_reservas_prod on op_reservas(producto_codigo) where estado = 'activa';
create index if not exists idx_reservas_op   on op_reservas(op_id);

-- Disponible real = stock físico − lo comprometido con otras OP.
-- Es lo que debe mirar la plataforma antes de prometer una pieza.
create or replace view v_stock_disponible as
select
  p.codigo,
  p.descripcion,
  coalesce(p.stock_actual, 0)                              as stock_fisico,
  coalesce(r.reservado, 0)                                 as reservado,
  coalesce(p.stock_actual, 0) - coalesce(r.reservado, 0)   as disponible
from productos p
left join (
  select producto_codigo, sum(cantidad) as reservado
  from op_reservas where estado = 'activa'
  group by producto_codigo
) r on r.producto_codigo = p.codigo
where p.activo is not false;


-- ─────────────────────────────────────────────────────────────────────────
-- 5) PROVEEDOR_SEDES — dónde se recoge cada compra
-- La tabla `proveedores` ya existe con direccion/ciudad/telefono, pero está
-- VACÍA y además un proveedor puede tener varias sedes (CODIFER tiene 3).
-- La zona es lo que ordena la ruta; la dirección exacta puede llegar después.
-- ─────────────────────────────────────────────────────────────────────────
create table if not exists proveedor_sedes (
  id            bigserial primary key,
  proveedor     text not null,                     -- nombre tal como se usa hoy
  nombre        text not null,                     -- 'CODIFER Zona Industrial'
  zona          text,                              -- lo que ordena la ruta
  direccion     text,
  ciudad        text,
  telefono      text,
  contacto      text,
  horario       text,
  llamar_antes  boolean default false,
  -- RECOGEMOS · DESPACHA · MENSAJERIA — los que despachan no entran a la ruta
  tipo_entrega  text not null default 'RECOGEMOS',
  notas         text,                              -- parqueo, "preguntar por…", etc.
  principal     boolean default false,
  activo        boolean not null default true,
  created_by    text,
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now()
);
create index if not exists idx_sedes_prov on proveedor_sedes(proveedor) where activo = true;

-- Orden en que se recorre la ciudad. Sin mapas: una lista ordenada a mano
-- da el 90% del beneficio y no cuesta nada.
create table if not exists zonas_ruta (
  zona   text primary key,
  orden  int not null default 100,
  activo boolean not null default true
);

-- El aprendizaje: "las bridas de CODIFER se recogen en Zona Industrial".
-- `clave` es el tipo de pieza o la familia; así no vuelve a preguntar.
create table if not exists proveedor_sede_memoria (
  id         bigserial primary key,
  proveedor  text not null,
  clave      text not null,                        -- tipo de pieza / familia
  sede_id    bigint not null references proveedor_sedes(id) on delete cascade,
  veces      int not null default 1,
  ultima_at  timestamptz not null default now()
);
create unique index if not exists idx_sedemem on proveedor_sede_memoria(proveedor, clave);


-- ─────────────────────────────────────────────────────────────────────────
-- 6) OP_CERTIFICADOS — los MTR
-- Hoy los certificados viven en Downloads\Certificados organizados por
-- CLIENTE, no por colada, y por eso buscarlos es lento. Aquí quedan colgados
-- de la OP y de la colada. La OP no se cierra sin ellos.
-- ─────────────────────────────────────────────────────────────────────────
create table if not exists op_certificados (
  id          bigserial primary key,
  op_id       bigint not null references ops(id) on delete cascade,
  op_item_id  bigint references op_items(id) on delete set null,
  colada      text,
  archivo     text not null,                       -- ruta en Supabase Storage
  nombre      text,
  tipo        text default 'MTR',                  -- MTR · ficha técnica · otro
  subido_por  text,
  subido_at   timestamptz not null default now()
);
create index if not exists idx_cert_op     on op_certificados(op_id);
create index if not exists idx_cert_colada on op_certificados(colada);


-- ─────────────────────────────────────────────────────────────────────────
-- 7) OP_EVENTOS — bitácora
-- Quién creó, quién mandó a aprobar, quién aprobó o devolvió y con qué nota.
-- Es lo que permite responder "¿por qué se compró esto?" seis meses después.
-- ─────────────────────────────────────────────────────────────────────────
create table if not exists op_eventos (
  id       bigserial primary key,
  op_id    bigint not null references ops(id) on delete cascade,
  evento   text not null,                          -- creada · enviada_aprobacion · aprobada…
  detalle  text,
  usuario  text,
  at       timestamptz not null default now()
);
create index if not exists idx_opeventos on op_eventos(op_id, at desc);


-- ─────────────────────────────────────────────────────────────────────────
-- 8) ENLACE con lo que YA existe — columnas nuevas, opcionales
-- La remisión y el kardex pasan a saber de qué OP salieron. Es lo que hoy
-- se resuelve cruzando TEXTO (kardex.remision ILIKE '%numero%'), que se
-- rompe con solo escribirlo distinto.
-- ─────────────────────────────────────────────────────────────────────────
alter table remisiones   add column if not exists op_numero text;
alter table kardex       add column if not exists op_numero text;
alter table plan_compras add column if not exists op_numero text;
create index if not exists idx_remisiones_op on remisiones(op_numero);
create index if not exists idx_kardex_op     on kardex(op_numero);
create index if not exists idx_plan_op       on plan_compras(op_numero);


-- ─────────────────────────────────────────────────────────────────────────
-- 9) REGLA DE CIERRE — no se cierra sin certificados
-- Se hace en la BASE DE DATOS, no en el navegador: así la regla se cumple
-- venga de donde venga la escritura (plataforma, proxy, script o a mano).
-- ─────────────────────────────────────────────────────────────────────────
create or replace function op_valida_cierre() returns trigger
language plpgsql as $$
declare
  n int;
begin
  if new.estado = 'cerrada' and coalesce(old.estado, '') <> 'cerrada'
     and coalesce(new.requiere_certificados, true) then
    select count(*) into n from op_certificados where op_id = new.id;
    if n = 0 then
      raise exception
        'La OP % no se puede cerrar: no tiene certificados cargados. '
        'Sube el MTR o marca requiere_certificados = false si de verdad no aplica.',
        new.numero;
    end if;
  end if;
  return new;
end $$;

drop trigger if exists trg_op_cierre on ops;
create trigger trg_op_cierre before update on ops
  for each row execute function op_valida_cierre();


-- ─────────────────────────────────────────────────────────────────────────
-- 10) TIMESTAMPS del servidor — reusa la función que ya existe (2026-07-13)
-- ─────────────────────────────────────────────────────────────────────────
create or replace function set_updated_at() returns trigger
language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end $$;

drop trigger if exists trg_ops_updated on ops;
create trigger trg_ops_updated before update on ops
  for each row execute function set_updated_at();

drop trigger if exists trg_opitems_updated on op_items;
create trigger trg_opitems_updated before update on op_items
  for each row execute function set_updated_at();

drop trigger if exists trg_sedes_updated on proveedor_sedes;
create trigger trg_sedes_updated before update on proveedor_sedes
  for each row execute function set_updated_at();


-- ─────────────────────────────────────────────────────────────────────────
-- 11) RLS — misma convención del 2026-07-13: SOLO LECTURA para la key
-- pública. Toda escritura entra por el proxy con la secret key.
-- ─────────────────────────────────────────────────────────────────────────
do $$
declare
  t text;
  p record;
begin
  foreach t in array array[
    'ops','op_items','op_reservas','op_certificados','op_eventos',
    'op_consecutivos','proveedor_sedes','proveedor_sede_memoria','zonas_ruta'
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


-- ─────────────────────────────────────────────────────────────────────────
-- 12) REALTIME — para que el equipo vea las OP al instante
-- ─────────────────────────────────────────────────────────────────────────
do $$
begin
  begin alter publication supabase_realtime add table ops;      exception when duplicate_object then null; end;
  begin alter publication supabase_realtime add table op_items; exception when duplicate_object then null; end;
end $$;
alter table ops      replica identity full;
alter table op_items replica identity full;


-- ─────────────────────────────────────────────────────────────────────────
-- 13) ZONAS — se llenan solas, no se siembran
-- Decisión del usuario (22-ago): la tabla arranca VACÍA. Las zonas son
-- barrios de Bogotá (Fontibón, Puente Aranda, Ricaurte…) y se van creando
-- a medida que alguien registra una sede, igual que las direcciones.
-- Sembrar zonas genéricas inventadas solo produciría basura que nadie usa.
--
-- El `orden` (en qué secuencia se recorren) se ajusta después, cuando ya
-- se sepa cuáles son las que de verdad aparecen.
-- ─────────────────────────────────────────────────────────────────────────
-- (intencionalmente sin INSERT)


-- ─────────────────────────────────────────────────────────────────────────
-- 14) VERIFICACIÓN — esto es lo que debe aparecer al final
-- ─────────────────────────────────────────────────────────────────────────
select 'TABLAS CREADAS' as chequeo, string_agg(table_name, ', ' order by table_name) as detalle
from information_schema.tables
where table_schema = 'public'
  and table_name in ('ops','op_items','op_reservas','op_certificados','op_eventos',
                     'op_consecutivos','proveedor_sedes','proveedor_sede_memoria','zonas_ruta')
union all
select 'COLUMNAS op_numero', string_agg(table_name || '.' || column_name, ', ' order by table_name)
from information_schema.columns
where table_schema = 'public' and column_name = 'op_numero'
union all
select 'POLITICAS RLS', string_agg(tablename || ' → ' || policyname, ', ' order by tablename)
from pg_policies
where schemaname = 'public'
  and (tablename like 'op%' or tablename like 'proveedor_sede%' or tablename = 'zonas_ruta')
union all
select 'FUNCIONES', string_agg(proname, ', ' order by proname)
from pg_proc where proname in ('op_nuevo_numero', 'op_valida_cierre')
union all
select 'PROXIMA OP SERIA', 'OP-'
  || extract(year from (now() at time zone 'America/Bogota'))::int || '-'
  || lpad((coalesce((select ultimo from op_consecutivos
       where anio = extract(year from (now() at time zone 'America/Bogota'))::int), 0) + 1)::text, 4, '0');

-- Para PROBAR el consecutivo (ojo: gasta un número; después reinicia con el update):
--   select op_nuevo_numero();
--   update op_consecutivos set ultimo = 0 where anio = 2026;
