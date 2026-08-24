-- ============================================================================
-- E&G ENERGY GROUP — Recepción de compras en bodega + ORIGEN del movimiento
-- Fecha: 2026-08-24
--
-- POR QUÉ EL ORIGEN
-- El usuario pidió que "todo entre a bodega y especifique DE DÓNDE", para poder
-- hacer un análisis correcto de importación. Hoy eso no se puede: el kardex
-- guarda fecha, producto, cantidad, costo, colada y lote — pero NO de dónde vino.
-- Lo único que marca "importado" es el prefijo IMP- en el CÓDIGO DEL PRODUCTO,
-- que es del producto y no del movimiento. Si la misma brida se compra un mes
-- importada y al otro a CODIFER, no hay forma de separarlas.
--
-- Con estas columnas el origen queda en la ENTRADA, que es donde pertenece.
--
-- CÓMO USAR: Supabase → SQL Editor → New query → pegar → RUN.
-- Es aditivo y se puede correr varias veces.
-- ============================================================================

-- ─── 1) De dónde vino cada movimiento ───────────────────────────────────────
alter table kardex add column if not exists proveedor text;
-- PLAZA (compra nacional) · IMPORTACION · DEVOLUCION · AJUSTE · TRASLADO
alter table kardex add column if not exists origen    text;
alter table kardex add column if not exists factura   text;   -- la del proveedor
create index if not exists idx_kardex_origen    on kardex(origen);
create index if not exists idx_kardex_proveedor on kardex(proveedor);

-- ─── 2) Sembrar el origen de lo que YA existe ───────────────────────────────
-- No se inventa nada: solo se marca lo que el propio dato ya decía.
--   · entradas de productos IMP-  -> IMPORTACION
--   · lotes/notas de devolución   -> DEVOLUCION
--   · correcciones                -> AJUSTE
-- El resto se deja en NULL a propósito: "no se sabe" es una respuesta honesta,
-- y marcarlo todo como PLAZA sería inventar historia.
update kardex set origen = 'IMPORTACION'
 where origen is null and upper(tipo) like 'ENTRADA%' and codigo_producto like 'IMP-%';

update kardex set origen = 'DEVOLUCION'
 where origen is null and upper(tipo) like 'ENTRADA%'
   and (upper(coalesce(lote,'')) like 'DEV-%' or upper(coalesce(notas,'')) like '%DEVOLUCI%');

update kardex set origen = 'AJUSTE'
 where origen is null and (upper(coalesce(lote,'')) like 'CORR-%'
   or upper(coalesce(notas,'')) like '%CORRECCI%' or upper(coalesce(notas,'')) like '%AJUSTE%');

-- ─── 3) Recepciones: la cabecera de cada llegada ────────────────────────────
-- Una compra puede llegar por partes. Cada llegada es una recepción con su
-- fecha, su factura y quién la recibió; el detalle vive en el kardex, que es
-- donde ya vive todo movimiento de inventario.
create table if not exists op_recepciones (
  id           bigserial primary key,
  op_numero    text,
  cc           text,                                  -- centro de costos / plan
  proveedor    text,
  factura      text,
  origen       text not null default 'PLAZA',
  fecha        date not null default current_date,
  observaciones text,
  recibido_por text,
  created_at   timestamptz not null default now()
);
create index if not exists idx_op_recep_op on op_recepciones(op_numero);
create index if not exists idx_op_recep_cc on op_recepciones(cc);

-- Enlace del movimiento con su recepción
alter table kardex add column if not exists recepcion_id bigint;

-- ─── 4) Qué se recibió de cada línea de la OP ───────────────────────────────
-- Sin esto no hay saldo: pediste 100, llegaron 40, faltan 60.
alter table op_items add column if not exists recibida numeric;

-- ─── 5) RLS: misma convención — solo lectura para la key publishable ────────
do $$
declare p record;
begin
  execute 'alter table public.op_recepciones enable row level security';
  for p in select policyname from pg_policies
            where schemaname='public' and tablename='op_recepciones' loop
    execute format('drop policy %I on public.op_recepciones', p.policyname);
  end loop;
  execute 'create policy op_recepciones_solo_lectura on public.op_recepciones for select using (true)';
end $$;

-- ─── 6) Verificación ────────────────────────────────────────────────────────
select 'COLUMNAS kardex' as chequeo,
       string_agg(column_name, ', ' order by column_name) as detalle
  from information_schema.columns
 where table_schema='public' and table_name='kardex'
   and column_name in ('proveedor','origen','factura','recepcion_id')
union all
select 'TABLA op_recepciones',
       coalesce((select 'creada' from information_schema.tables
                  where table_schema='public' and table_name='op_recepciones'), 'FALTA')
union all
select 'op_items.recibida',
       coalesce((select 'creada' from information_schema.columns
                  where table_schema='public' and table_name='op_items' and column_name='recibida'), 'FALTA')
union all
select 'ORIGEN sembrado', origen || ': ' || count(*)::text
  from kardex where origen is not null group by origen
union all
select 'SIN ORIGEN (queda en null)', count(*)::text from kardex where origen is null;
