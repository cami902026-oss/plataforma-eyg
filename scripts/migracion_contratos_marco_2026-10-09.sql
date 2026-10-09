-- ═══════════════════════════════════════════════════════════════════════════
--  CONTRATOS MARCO — 9-oct-2026
--  Réplica del Excel "Control_Contratos_Marco_EYG": hoja CONTRATOS → tabla
--  contratos_marco · hoja CONSUMOS → tabla contratos_consumos.
--  Se corre UNA vez en Supabase → SQL Editor → RUN. Es idempotente.
--  Cupo y consumos CON IVA incluido (decisión del usuario, 9-oct-2026).
-- ═══════════════════════════════════════════════════════════════════════════

-- 1) CONTRATOS (hoja CONTRATOS del Excel)
create table if not exists contratos_marco (
  id            bigserial primary key,
  numero        text not null unique,          -- No. Contrato (= OC abierta del cliente)
  cliente       text not null,                 -- razón social como está en el contrato
  nit           text,
  cliente_clave text,                          -- palabra para cruzar con Cartera (HUPECOL / ANDES / TOC)
  objeto        text,
  cupo          numeric(16,2) not null default 0,   -- Valor Total / Cupo, CON IVA
  fecha_inicio  date,
  fecha_fin     date,
  estado        text not null default 'Activo',     -- Activo / Suspendido / Cerrado / Vencido
  notas         text,
  created_by    text,
  created_at    timestamptz not null default now(),
  updated_by    text,
  updated_at    timestamptz not null default now()
);

-- 2) CONSUMOS (hoja CONSUMOS del Excel)
create table if not exists contratos_consumos (
  id             bigserial primary key,
  contrato_id    bigint not null references contratos_marco(id) on delete cascade,
  fecha          date not null default current_date,
  tipo           text not null default 'Factura',  -- Factura / Remisión / Orden de servicio / Nota crédito
  documento      text,                             -- No. Documento
  descripcion    text,                             -- Descripción / Ítems
  valor          numeric(16,2) not null default 0, -- CON IVA; una nota crédito va en NEGATIVO
  factura_id     bigint,                           -- cartera_facturas.id cuando viene del cruce con Cartera
  origen         text not null default 'manual',   -- manual / cartera
  registrado_por text,
  created_at     timestamptz not null default now()
);
create index if not exists idx_ctr_cons_contrato on contratos_consumos(contrato_id);
create index if not exists idx_ctr_cons_factura  on contratos_consumos(factura_id);

-- 3) RLS igual que el resto de la plataforma: lectura con la key pública,
--    escritura SOLO por el proxy (key secreta, que se salta la RLS).
do $$
declare t text; p record;
begin
  foreach t in array array['contratos_marco','contratos_consumos'] loop
    execute format('alter table public.%I enable row level security', t);
    for p in select policyname from pg_policies where schemaname='public' and tablename=t loop
      execute format('drop policy %I on public.%I', p.policyname, t);
    end loop;
    execute format('create policy %I on public.%I for select using (true)', t || '_solo_lectura', t);
  end loop;
end $$;

-- 4) SEMILLA: los 3 contratos del Excel (si ya existen, no se tocan)
insert into contratos_marco (numero, cliente, nit, cliente_clave, objeto, cupo, fecha_inicio, fecha_fin, estado, created_by)
values
 ('COM-002275', 'HUPECOL OPERATING CO LLC', '900148720-6', 'HUPECOL',
  'OC Abierta bajo demanda - Suministro de materiales y consumibles para mantenimiento preventivo y correctivo de la facilidad. AFE P1000030 / CC003. Entrega: Puerto Lopez.',
  100000000, '2026-07-23', '2027-12-31', 'Activo', 'excel'),
 ('COM-000394', 'ANDES OPERATING COMPANY LLC SUCURSAL COL', '901217782-2', 'ANDES',
  'OC Abierta bajo demanda - Suministro de materiales y consumibles para mantenimiento preventivo y correctivo de la facilidad. AFE G0000030 / CC003. Entrega: Puerto Lopez.',
  100000000, '2026-07-23', '2027-12-31', 'Activo', 'excel'),
 ('COM-000012', 'TOC ENERGIA SUCURSAL COLOMBIA', '901450935-1', 'TOC',
  'OC Abierta bajo demanda - Suministro de materiales y consumibles para mantenimiento preventivo y correctivo de equipos de la facilidad. AFE T0000030 / CC003. Entrega: Paz de Ariporo.',
  100000000, '2026-07-24', '2026-12-31', 'Activo', 'excel')
on conflict (numero) do nothing;

-- Comprobación
select numero, cliente, cupo, fecha_inicio, fecha_fin, estado from contratos_marco order by numero;
