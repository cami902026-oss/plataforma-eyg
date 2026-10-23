-- ═══════════════════════════════════════════════════════════════════════════
--  PORTAL DE VINCULACIÓN (clientes y proveedores) — 22-oct-2026
--  Un cliente o proveedor recibe un correo con el link del portal y un código,
--  crea su clave, y desde entonces entra con NIT + clave a subir sus documentos
--  (RUT, cámara de comercio, cédula del representante legal, certificación
--  bancaria). Los archivos viven en Supabase Storage (bucket privado
--  `vinculacion`, carpeta por NIT); los metadatos en estas tablas.
--  Se corre UNA vez en Supabase → SQL Editor → RUN. Es idempotente.
-- ═══════════════════════════════════════════════════════════════════════════

-- 1) Una cuenta por NIT (cliente o proveedor)
create table if not exists portal_cuentas (
  id                    bigserial primary key,
  nit                   text not null unique,          -- solo dígitos, sin dígito de verificación
  razon_social          text not null,
  tipo                  text not null default 'cliente',   -- cliente / proveedor
  correo                text not null,
  contacto              text,
  telefono              text,
  estado                text not null default 'invitado',  -- invitado / activo / aprobado / bloqueado
  invitado_por          text,
  invitado_at           timestamptz not null default now(),
  invitacion_enviada_at timestamptz,
  activado_at           timestamptz,
  ultimo_ingreso        timestamptz,
  notas                 text,
  updated_at            timestamptz not null default now()
);

-- 2) Credenciales: SIN política de lectura (solo la key secreta del proxy las ve).
--    Van aparte de la cuenta a propósito: portal_cuentas se lee con la key pública
--    como el resto de la plataforma, y el hash de la clave no puede viajar ahí.
create table if not exists portal_credenciales (
  nit                text primary key references portal_cuentas(nit) on delete cascade,
  codigo_invitacion  text,               -- 6 caracteres; se borra al crear la clave
  clave_hash         text,               -- sha256(sal + ':' + clave)
  sal                text,
  intentos_fallidos  int not null default 0,
  bloqueado_hasta    timestamptz,
  updated_at         timestamptz not null default now()
);

-- 3) Documentos subidos. Nunca se borran: al subir uno nuevo del mismo tipo, el
--    anterior pasa a 'reemplazado' y queda el historial.
create table if not exists portal_documentos (
  id               bigserial primary key,
  nit              text not null references portal_cuentas(nit) on delete cascade,
  tipo_doc         text not null,          -- rut / camara / cedula / bancaria
  nombre_archivo   text not null,
  ruta             text not null,          -- ruta dentro del bucket `vinculacion`
  tamano           bigint,
  mime             text,
  fecha_expedicion date,
  vence            date,                   -- cámara de comercio: expedición + 30 días
  estado           text not null default 'pendiente',   -- pendiente / aprobado / rechazado / reemplazado
  observacion      text,                   -- motivo cuando se pide corrección
  revisado_por     text,
  revisado_at      timestamptz,
  subido_at        timestamptz not null default now()
);
create index if not exists idx_portal_doc_nit  on portal_documentos(nit);
create index if not exists idx_portal_doc_est  on portal_documentos(estado);

-- 4) RLS. Cuentas y documentos: lectura con la key pública, escritura SOLO por el
--    proxy (igual que el resto). Credenciales: NINGUNA política → solo service role.
do $$
declare t text; p record;
begin
  foreach t in array array['portal_cuentas','portal_documentos','portal_credenciales'] loop
    execute format('alter table public.%I enable row level security', t);
    for p in select policyname from pg_policies where schemaname='public' and tablename=t loop
      execute format('drop policy %I on public.%I', p.policyname, t);
    end loop;
    if t <> 'portal_credenciales' then
      execute format('create policy %I on public.%I for select using (true)', t || '_solo_lectura', t);
    end if;
  end loop;
end $$;

-- 5) Bucket PRIVADO para los archivos. Sin políticas de storage: el portal sube con
--    URL firmada (la emite el proxy con la key secreta) y la plataforma descarga
--    con URL firmada de 10 minutos. Nadie llega al archivo sin pasar por el proxy.
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('vinculacion', 'vinculacion', false, 15728640, array['application/pdf','image/jpeg','image/png'])
on conflict (id) do update set public = false, file_size_limit = 15728640,
  allowed_mime_types = array['application/pdf','image/jpeg','image/png'];

-- Comprobación
select 'portal_cuentas' as tabla, count(*) from portal_cuentas
union all select 'portal_documentos', count(*) from portal_documentos
union all select 'bucket vinculacion', count(*) from storage.buckets where id = 'vinculacion';
