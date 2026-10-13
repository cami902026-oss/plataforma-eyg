-- ═══════════════════════════════════════════════════════════════════════════
--  SALDOS A FAVOR DE PROVEEDORES — 13-oct-2026
--  Réplica del Excel "Saldo_Favor_Proveedor.xlsx" (una hoja por proveedor:
--  saldo inicial, pedidos que lo consumen, saldo restante) + ABONOS.
--  Se corre UNA vez en Supabase → SQL Editor → RUN. Es idempotente.
--  Valores ANTES de IVA (decisión del usuario, 13-oct-2026).
-- ═══════════════════════════════════════════════════════════════════════════

-- 1) Un registro por proveedor con saldo a favor
create table if not exists saldos_proveedor (
  id          bigserial primary key,
  proveedor   text not null unique,        -- nombre como está en la libreta / plan de compras (MAYÚSCULAS)
  nit         text,
  notas       text,
  created_by  text,
  created_at  timestamptz not null default now(),
  updated_by  text,
  updated_at  timestamptz not null default now()
);

-- 2) Movimientos: Saldo inicial (+), Abono (+), Consumo (−). `valor` va CON SIGNO.
create table if not exists saldos_movimientos (
  id             bigserial primary key,
  saldo_id       bigint not null references saldos_proveedor(id) on delete cascade,
  fecha          date not null default current_date,
  tipo           text not null default 'Consumo',   -- Saldo inicial / Abono / Consumo
  documento      text,                              -- No. factura del proveedor / nota crédito
  descripcion    text,                              -- descripción del pedido
  motivo         text,                              -- solo abonos: devolución, nota crédito, anticipo, pago doble…
  valor          numeric(16,2) not null default 0,  -- ANTES de IVA; consumo en NEGATIVO
  origen         text not null default 'manual',    -- manual / plan (sugerido desde Plan de Compras)
  plan_cc        text,                              -- CC del plan de compras cuando viene de ahí
  registrado_por text,
  created_at     timestamptz not null default now()
);
create index if not exists idx_spv_mov_saldo on saldos_movimientos(saldo_id);
create index if not exists idx_spv_mov_plan  on saldos_movimientos(plan_cc);

-- 3) RLS igual que el resto: lectura con la key pública, escritura SOLO por el proxy.
do $$
declare t text; p record;
begin
  foreach t in array array['saldos_proveedor','saldos_movimientos'] loop
    execute format('alter table public.%I enable row level security', t);
    for p in select policyname from pg_policies where schemaname='public' and tablename=t loop
      execute format('drop policy %I on public.%I', p.policyname, t);
    end loop;
    execute format('create policy %I on public.%I for select using (true)', t || '_solo_lectura', t);
  end loop;
end $$;

-- 4) SEMILLA: lo que traía el Excel. Solo CODIFER tenía saldo real ($994.447 al 13-10-2026);
--    BELTHEL e HIERROS Y PERFILES LA 13 quedan creados en $0 para que les carguen el suyo.
insert into saldos_proveedor (proveedor, created_by) values
  ('CODIFER', 'excel'), ('BELTHEL', 'excel'), ('HIERROS Y PERFILES LA 13', 'excel')
on conflict (proveedor) do nothing;
insert into saldos_movimientos (saldo_id, fecha, tipo, documento, descripcion, valor, origen, registrado_por)
select id, '2026-10-13', 'Saldo inicial', null, 'Saldo inicial a favor (Excel, corte 13-10-2026)', 994447, 'manual', 'excel'
from saldos_proveedor where proveedor = 'CODIFER'
  and not exists (select 1 from saldos_movimientos m where m.saldo_id = saldos_proveedor.id and m.tipo = 'Saldo inicial');

-- Comprobación
select p.proveedor, coalesce(sum(m.valor), 0) as saldo
from saldos_proveedor p left join saldos_movimientos m on m.saldo_id = p.id
group by p.proveedor order by p.proveedor;
