-- ============================================================================
-- E&G ENERGY GROUP — Centro de costos del plan de compras
-- Fecha: 2026-08-24
--
-- POR QUÉ
-- El plan de compras ya lleva el material: lo que se compra a proveedores y
-- (desde v209) lo que sale de bodega a costo de inventario. Pero un pedido
-- cuesta más que su material: fletes, recogidas, empaque, permisos.
-- Hoy el transporte se escribe A MANO en una casilla del Excel de utilidad —
-- se pierde al regenerar el reporte y nadie más lo ve.
--
-- Esta tabla hace que esos gastos vivan en el plan, junto al material, para que
-- el centro de costos sea el costo REAL del pedido.
--
-- CÓMO USAR: Supabase → SQL Editor → New query → pegar → RUN.
-- Es aditivo y se puede correr varias veces.
-- ============================================================================

create table if not exists plan_gastos (
  id          bigserial primary key,
  cc          text not null,                       -- centro de costos (el plan)
  concepto    text not null,                       -- 'Flete a Yopal', 'Recogida Fontibón'…
  proveedor   text,                                -- transportadora, mensajero… opcional
  valor       numeric not null default 0,
  -- Algunos gastos vienen con IVA y otros no (un mensajero informal, por ejemplo).
  -- Se guarda la decisión en vez de asumirla.
  aplica_iva  boolean not null default false,
  nota        text,
  op_numero   text,                                -- si el plan viene de una OP
  created_by  text,
  created_at  timestamptz not null default now()
);
create index if not exists idx_plan_gastos_cc on plan_gastos(cc);
create index if not exists idx_plan_gastos_op on plan_gastos(op_numero);

-- RLS: misma convención de siempre — SOLO LECTURA para la key publishable.
-- Toda escritura entra por el proxy con la secret key.
do $$
declare p record;
begin
  execute 'alter table public.plan_gastos enable row level security';
  for p in select policyname from pg_policies
            where schemaname='public' and tablename='plan_gastos' loop
    execute format('drop policy %I on public.plan_gastos', p.policyname);
  end loop;
  execute 'create policy plan_gastos_solo_lectura on public.plan_gastos for select using (true)';
end $$;

-- Verificación
select 'TABLA' as chequeo, table_name as detalle
  from information_schema.tables
 where table_schema='public' and table_name='plan_gastos'
union all
select 'POLITICA', policyname from pg_policies
 where schemaname='public' and tablename='plan_gastos';
