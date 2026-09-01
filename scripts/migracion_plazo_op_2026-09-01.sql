-- ============================================================
-- E&G ENERGY GROUP — OP: FECHA COMPROMETIDA (v239)
--
-- POR QUÉ: la lista de OP se llenó (2 nuevas al día) y las cerradas tapaban
-- las que piden algo. Al sacarlas a 🗄️ Historial hizo falta lo contrario:
-- una forma de decir "esta lleva días quieta". Un plazo fijo no sirve.
--
-- 🔑 EL PLAZO DEPENDE DEL PRODUCTO. Hay material de plaza que llega en
-- "1-2 DIAS" y material de FABRICACIÓN cotizado a "24 SEMANAS" (medio año) y
-- "15 SEMANAS" — están en `cotizacion_items.tiempo_entrega`, que hoy trae 25
-- valores distintos y está lleno en 63 de las 72 líneas de OP. Con un plazo
-- corto, toda OP con fabricación se pintaría atrasada desde el primer día, y
-- una bandera que sale siempre deja de mirarse.
--
-- QUÉ HACE: una columna. La plataforma la calcula al crear la OP (hoy + el
-- tiempo de entrega de la línea MÁS LENTA, porque el pedido no sale completo
-- hasta que llegue la última) y se puede corregir a mano en la cabecera de la
-- OP cuando el proveedor da una fecha real.
--
-- Solo manda en las etapas que ESPERAN MATERIAL (en_compras, en_ejecucion).
-- Que algo sea de fabricación no justifica que se demore la aprobación ni la
-- factura: esas dependen de nosotros y siguen con el plazo corto.
--
-- ADITIVO E IDEMPOTENTE. Nada se borra, nada se reescribe. Si no se corre, la
-- plataforma NO se rompe: la columna llega vacía, el aviso de "no se pudo
-- guardar la fecha" salta al editar, y el atraso cae al plazo corto por etapa.
-- ============================================================

alter table public.ops
  add column if not exists fecha_comprometida date;

comment on column public.ops.fecha_comprometida is
  'Fecha en que se comprometió la entrega. La calcula la plataforma al crear la OP '
  '(hoy + el tiempo_entrega más largo de sus líneas de cotización) y se corrige a mano. '
  'Mientras no pase, la OP no se marca atrasada en en_compras/en_ejecucion.';

-- Para el filtro ⏰ Atrasadas y para los informes: se consulta por fecha sobre
-- las OP vivas, que son pocas frente al total.
create index if not exists ops_fecha_comprometida_idx
  on public.ops (fecha_comprometida)
  where deleted = false;

-- Comprobación
select column_name, data_type
  from information_schema.columns
 where table_schema = 'public' and table_name = 'ops'
   and column_name = 'fecha_comprometida';
