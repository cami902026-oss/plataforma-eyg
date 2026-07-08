# 📝 MEMORIA DEL PROYECTO ENERGY BOT — PLATAFORMA EYG

**Última actualización:** 22 de mayo de 2026
**Mantenido por:** Andrea (ADMIN)
**Propósito:** Permite continuar trabajando con contexto completo en futuras sesiones.

---

## 🌐 URLs y endpoints activos

| Servicio | URL |
|---|---|
| Asistente web (frontend) | https://cami902026-oss.github.io/plataforma-eyg/Index.html |
| Repo GitHub (datos + código) | https://github.com/cami902026-oss/plataforma-eyg |
| Apps Script — Backend Sheets (órdenes) | `AKfycbwicJkz0AAKD9an0KB5ViKvtuYoHilT0yd3CmXJ0NwGX-0HOli7w2z9Fbf0LPoCpearnQ/exec` |
| Apps Script — Email→OC (Gmail polling) | proyecto "ENERGY Email to OC" |
| Apps Script — Proxy Claude IA | `AKfycbxrq2VO3GkTapQwXOgyn8EOX6VRnQcoduZ5XRQbsrqXAkI-_OXgSTRuQ4NxNF1O0GE/exec` |
| Apps Script — Proxy GitHub (sin PAT) | mismo que arriba |
| Cartera tunnel (cambia al reiniciar PC) | `data/cartera_url.json` en el repo |

## 🔐 Credenciales

- **GitHub PAT actual:** guardado en `C:\Users\Lenovo\Desktop\config_cartera.json` (campo `gh_token`)
- **El token viejo** fue auto-revocado por GitHub (se compartió públicamente — no usarlo)
- **Sheet ID de órdenes:** `1QjeJiCQ8fND57r0MNw1vRNdcnukqpBNHDfxjbeW7Lq0` (pestaña `Pedido_Seguimiento`)
- **Claves del equipo y código de acceso:** ahora se sincronizan automáticamente vía `data/config.json` (campo `passwords` + `accessCode`), todos hasheados SHA-256

## 👥 Usuarios del sistema

| Usuario | Rol | Email | WhatsApp |
|---|---|---|---|
| Alberto | JEFE | alberto | +573113134451 |
| Andrea | ADMIN | andrea | +573107574110 |
| Sheila Baron | COLABORADOR | sheila | +573204947227 |
| Alexandra | COLABORADOR | alexandra | +573144858382 |
| Nelsy | COLABORADOR | nelsy | +573125099056 |

**Restricciones de módulos:**
- 📚 Históricos: solo JEFE/ADMIN
- 💰 Cartera: solo Andrea y Alberto (hardcoded en `userCanSeeModule`)
- 🤝 Visitas Comerciales: Alexandra, Nelsy, Lina NO ven por default

---

## 📁 Archivos y ubicaciones clave

### Proyecto principal — `C:\Users\Lenovo\OneDrive\Escritorio\Energy_bot\`

| Archivo | Función |
|---|---|
| `Index.html` | Frontend completo del asistente (también se publica a GitHub) |
| `apps-script.gs` | Backend Sheets (pegar en el proyecto Apps Script "orden de pedido") |
| `plataforma-eyg/email-to-oc.gs` | Polling de Gmail → crea OCs auto (pegar en proyecto "ENERGY Email to OC") |
| `scripts/op_report.py` | Reporte diario OP (corre vía GitHub Actions L-V 5PM) |
| `scripts/inventory_report.py` | Reporte de inventario |
| `scripts/sync_inventory.py` | Sync de inventario desde OneDrive |
| `ordenes.json` | Backup local del JSON de órdenes |
| `.github/workflows/op-report.yml` | Workflow cron del reporte de OP |
| `.github/workflows/inventory-report.yml` | Workflow cron del reporte de inventario |
| `.github/workflows/sync-inventory.yml` | Workflow cron del sync de inventario |

### Aplicativo de Cartera — `C:\Users\Lenovo\Desktop\`

| Archivo | Función |
|---|---|
| `cartera_agente.py` | Flask app de cartera (puerto 5050) |
| `cartera_config.json` | Config del aplicativo (API key Claude, datos empresa) |
| `iniciar_cartera_silencioso.vbs` | Arranque automático: cartera + cloudflared + sync URL |
| `actualizar_url_github.ps1` | Sube la URL del túnel a `data/cartera_url.json` |
| `config_cartera.json` | Token GitHub + owner + repo (lo usa el .ps1) |
| `cloudflared.exe` | Cliente del túnel Cloudflare |
| `cartera_url.log` | Output de cloudflared (se borra y regenera cada arranque) |
| `cartera_url_actual.txt` | Solo la URL actual, sin nada más |
| `ver_url_cartera.bat` | Muestra la URL en pantalla y la copia al portapapeles |

### Fuente de datos cartera
- `C:\Users\Lenovo\Downloads\Cartera\Copia de CONTROL DE CARTERA.xlsx`
- El aplicativo detecta cambios automáticamente (cache 60s + mtime)

### Extractor de facturas — `C:\Users\Lenovo\Desktop\Facturas\`
- `procesar_facturas.py` — monitorea 3 correos Outlook
- `enviar_consolidado.py` — envía consolidado diario a contabilidad
- ⚠️ El buzón de `contabilidad@eygenergygroup.com` está corrupto (StoreDriver exception). El consolidado rebota. Solución pendiente: que Microsoft repare el buzón o quitar ese destinatario temporalmente.

---

## 🚀 Cambios y mejoras hechas el 22/may/2026

### email-to-oc.gs (Apps Script)
- ✅ Búsqueda flexible: `subject:"ORDEN DE COMPRA"` (antes era "ORDEN DE COMPRA ENERGY" — no capturaba la mayoría)
- ✅ Ventana ampliada de 14d a 30d
- ✅ Claude ahora extrae también el campo `valor` de la OC

### Index.html
- ✅ **Cartera** — nuevo módulo solo para Andrea/Alberto. Auto-fetch URL desde GitHub.
- ✅ **Cotizaciones IA** — extractor ahora usa proxy seguro (antes pedía `claude_api_key`); número consecutivo editable + botón "🔄 Auto"
- ✅ **Procesos O.C.** — campo nuevo `fechaIngreso` para KPIs de compras/logística (días transcurridos visible en la card con código de colores)
- ✅ **Procesos O.C.** — etapa visual "📋 Hoja Entrada" en el timeline (solo si la OC requiere HE)
- ✅ **Procesos O.C.** — detección de duplicados al crear OC manual
- ✅ **Auto-completar OC** — ahora requiere las 4 etapas con fecha + HE done si es requerida (antes marcaba "completado" prematuramente)
- ✅ **Sync de contraseñas** — al guardar password de usuario, se sube a `data/config.json` cifrada (SHA-256). Sirve para todos los dispositivos
- ✅ **Bug fix** — `_loadCompanyConfig` ahora siempre aplica el `ghProxyUrl` del servidor (antes solo si local estaba vacío)

### apps-script.gs (orden de pedido — Sheets backend)
- ✅ 5 columnas nuevas: Valor, HE Requerida, HE Fecha, HE Estado, **Full JSON**
- ✅ Al leer, usa el Full JSON si existe (preserva TODOS los campos para siempre)
- 🆕 URL del deployment cambió a `AKfycbwicJkz0AAKD9an0KB5ViKvtuYoHilT0yd3CmXJ0NwGX-0HOli7w2z9Fbf0LPoCpearnQ`

### scripts/op_report.py (reporte diario L-V 5PM)
- ✅ Filtra entradas con `deleted:true` (antes contaban)
- ✅ Columnas nuevas en tabla principal: 📅 Ingreso, ⏱ Días, 💵 Valor, 📋 HE
- ✅ 2 stat cards nuevas: Valor Activas, Valor Completadas
- ✅ Sección "📋 Órdenes Pendientes de Hoja de Entrada"
- ✅ Sección "📊 Histórico Completo de OPs" — Activas + Completadas + Canceladas
- ✅ Incluye OCs marcadas "completado" que todavía tienen etapas pendientes

### ordenes.json (en GitHub)
- ✅ Limpiado: 50+ duplicados eliminados, quedaron 31 entradas válidas
- ✅ Guardias permanentes para LM7777, LM1551, LM1527 (timestamp 2099 — siempre ganan el merge para que el script de correo nunca las recree)
- ✅ 5 OCs marcadas "completado" prematuramente → vueltas a "activo":
  - LM1434 (NIKOIL ENERGY) — falta entrega
  - 2026005441 (NEW GRANADA) — falta facturación + HE
  - NGEC-2026005442 — falta facturación + HE
  - C_PC000009201-1 (SERTECPET) — falta cert + facturación
  - 13080 (HIDROCARBUROS DEL CASANARE) — falta cert + factura + HE
- ✅ Unificación de cliente: HIDROCASANARE → HIDROCARBUROS DEL CASANARE SAS

### Cartera — acceso desde cualquier lugar
- ✅ Cloudflare Tunnel configurado (Quick Tunnel — URL cambia al reiniciar PC)
- ✅ Arranque 100% automático con Windows (vía Startup → iniciar_cartera_silencioso.vbs)
- ✅ La URL se sube automática a `data/cartera_url.json` cada vez que cambia
- ✅ El asistente lee esa URL automáticamente cuando alguien abre el módulo Cartera

---

## ⚠️ Bugs y pendientes conocidos

1. **El VBS no llamó al PowerShell en el primer arranque** — Probablemente porque `config_cartera.json` o `actualizar_url_github.ps1` no existían aún. Verificar en próximo reinicio que sí se ejecute (revisar que `data/cartera_url.json` en GitHub se actualice automáticamente)

2. **Excel que "no llega"** — Andrea mencionó que un Excel que llegaba antes ya no llega. NO se identificó cuál Excel exactamente. Pendiente aclarar:
   - ¿El de cartera (Downloads\Cartera\)?
   - ¿Uno adjunto a algún correo automático?
   - ¿Otro Excel distinto?

3. **Buzón de `contabilidad@eygenergygroup.com` corrupto** — Microsoft 365 devuelve `StoreDriver.Submit.Exception: CorruptDataException`. El `enviar_consolidado.py` rebota a ese destinatario. Solución: pedirle a Microsoft que repare el buzón o quitar ese destinatario del script.

4. **Mejoras opcionales pendientes para cartera tunnel:**
   - No suspender PC durante el día laboral
   - Auto-reinicio si cloudflared se cae
   - URL permanente con dominio propio (Named Tunnel + DNS en `eygenergygroup.com`)

5. **Reporte de OP** — la sección "Pendientes de HE" puede crecer mucho. Considerar agregar paginación si supera 30 entradas.

---

## 📚 Cómo retomar el trabajo en una nueva sesión

1. Lee este archivo (`MEMORIA_PROYECTO.md`) para entender el contexto
2. Verifica el estado actual en GitHub: https://github.com/cami902026-oss/plataforma-eyg/commits/main
3. Para cualquier cambio que requiera subir a GitHub usa el token de `config_cartera.json`
4. Antes de modificar `Index.html`, verifica que no haya cambios pendientes del usuario
5. **Regla de oro:** "No modifiques nada de lo que ya sirve" — todo cambio debe ser aditivo o explícitamente arreglar un bug

---

## 🛠️ Comandos útiles

### Subir un archivo a GitHub vía Python
```python
import base64, json, urllib.request
token = '<leer de config_cartera.json>'
path = 'Index.html'
with open(r'C:\Users\Lenovo\OneDrive\Escritorio\Energy_bot\Index.html','rb') as f:
    b64 = base64.b64encode(f.read()).decode('ascii')
req = urllib.request.Request(f'https://api.github.com/repos/cami902026-oss/plataforma-eyg/contents/{path}',
    headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github+json'})
with urllib.request.urlopen(req) as r: sha = json.loads(r.read())['sha']
body = json.dumps({'message':'msg', 'content':b64, 'sha':sha, 'branch':'main'}).encode('utf-8')
req2 = urllib.request.Request(f'https://api.github.com/repos/cami902026-oss/plataforma-eyg/contents/{path}',
    data=body, method='PUT',
    headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github+json', 'Content-Type': 'application/json'})
urllib.request.urlopen(req2)
```

### Disparar reporte manualmente (genera correo)
```python
url = 'https://api.github.com/repos/cami902026-oss/plataforma-eyg/actions/workflows/op-report.yml/dispatches'
body = json.dumps({'ref': 'main'}).encode('utf-8')
req = urllib.request.Request(url, data=body, method='POST', headers={...})
```

⚠️ **NO disparar workflows sin avisar al usuario** — cada disparo envía correo a 5 destinatarios.

### Limpiar localStorage del navegador
```javascript
// En consola del navegador (F12)
localStorage.removeItem('procesos_oc');
localStorage.removeItem('energy_sync_queue');
localStorage.removeItem('_lastSync_ordenes.json');
location.reload();
```

---

## 📞 Contactos importantes

| Quién | Para qué |
|---|---|
| Andrea Bernal (andrea.bernal@eygenergygroup.com) | Owner del proyecto, administradora del asistente |
| Alberto (gerenciageneral@) | JEFE, usa Cartera, recibe reportes |
| Sheila (comercial1@) | Cotizaciones |
| Alexandra (asistente.administrativo@) | Apoyo administrativo |
| Nelsy (contabilidad@) | Contabilidad — su buzón está corrupto, pendiente arreglar |

---

---

## 🚀 Cambios y mejoras hechas el 23/may/2026

### Módulo Cotizaciones — campos de seguimiento

- ✅ **3 nuevos campos en el formulario:** Vendedor, Realizada por, Aprobada por
  - "Realizada por" se pre-llena automáticamente con el usuario que tiene sesión
  - "Aprobada por" se pre-llena con Alberto por defecto
  - Se guardan en `bd_cotizaciones` y se recuperan al editar
- ✅ **Filtro de vendedor en la Base de Datos** (solo visible para ADMIN/JEFE)
- ✅ **Columna VENDEDOR** en la tabla de la BD
- ✅ **Control de visibilidad por rol de equipo:**
  - Usuarios con `role:'Comercial'` en el módulo Equipo solo ven sus propias cotizaciones
  - La restricción aplica automáticamente a futuros comerciales sin tocar código
  - Alexandra, Nelsy y otros roles no comerciales ven todo
- ✅ **Email dinámico en encabezado del PDF:** se toma del campo `email` del vendedor en los datos de equipo. Fallback: `comercial1@eygenergygroup.com`
- ✅ **Fila Vendedor / Elaboró / Aprobó** al pie del PDF

---

## 🚀 Cambios y mejoras hechas el 07/jul/2026 (SW v86 → v94)

### 🗓️ Cronograma de Sábados (nuevo, SW v86-v87)
- Pestaña **"🗓️ Sábados"** en Visitas Comerciales: rotación continua **Alexandra → Lina → Sandra** (inicia 11/jul/2026), un registro por sábado en `data/sabados_plan.json` (id `sab_YYYY-MM-DD`, merge anti-pisadas).
- Estados: Programada / Cumplida / Cambiada (con reemplazo) / No laboral (corre la rotación sin consumir turno). Botón "⚙️ Generar rotación" crea solo los sábados faltantes hasta fin de año.
- Editan ADMIN/JEFE + **Mario**; el resto solo ve. Contador de equidad anual + export .ics por mes.
- **Tarjeta Dashboard** "Este sábado trabaja: X" + **workflow `sabado-reminder.yml`** (viernes 4PM: correo a la asistente de turno + dirección; alerta si el sábado está sin asignar).
- Excel de aprobación: `Escritorio\Cronograma_Sabados_2026.xlsx` (25 sábados jul-dic).

### 🛡️ Sincronización blindada (SW v88, v93, v94)
- **Fix pisadas al reconectar**: la cola offline ahora RE-MEZCLA con el remoto antes de subir (`_mergeAntesDeSubir` — cotizaciones, solicitudes, semanario, sábados). Al volver la red también baja de inmediato lo de los demás.
- **Timeouts de red** (`_fetchT`, AbortController): 10-20s por petición en todo el guardado/lectura GitHub. Fin del "Sincronizando…" eterno (caso Andrea) — ahora pasa a cola con reintento.
- **Botón 🔄 Actualizar** en barra superior: sube la cola + fuerza descarga de TODOS los datos + indicador "hace Xs" + detector de versión nueva con oferta de recarga.

### 📋 Cotizaciones (SW v89-v92, v94)
- **Extraer con IA ya no borra ítems**: si la tabla tiene ítems pregunta AGREGAR al final o reemplazar; botón "↩️ Recuperar ítems anteriores" (respaldo en `energy_cotiz_items_undo`).
- **Consecutivo vivo**: si otra persona usa el número que está en pantalla, el campo se actualiza solo al siguiente libre (sufijos incluidos: LM1751-1→-2); al renumerar por colisión el formulario y `cotizEditId` se corrigen al instante ("Genera el PDF como LMxxxx").
- **Abrir sin colgarse**: verificación pre-apertura con tope 2,5s (internet flojo → abre con copia local + aviso); mensaje claro si el número no existe.
- **Motivo de rechazo OBLIGATORIO**: desplegable al marcar Rechazada (Sin stock disponible / Stock insuficiente / Falta soporte técnico / Precio no competitivo / Tiempo de entrega no competitivo); visible en rojo en la BD; se guarda en `motivoRechazo`.
- **Ítems en $0 = "NO COTIZADO"** en PDF y Excel (los $0 son intencionales: ítems no encontrados).
- **BD de Clientes auto-alimentada**: `_cliUpsertDesdeCotiz` al guardar (solo llena campos vacíos, nunca pisa lo manual) + seed `data/clientes.json` con 59 clientes del histórico. El autocompletar del formulario ya funciona.
- **Semáforo revivido**: `fechaEnvio` se estampa al marcar Enviada (antes nunca se guardaba → columna muerta).
- **Auto-Vencidas**: Enviadas con `fechaVenc` pasada se marcan Vencida solas (`updatedBy:'auto'`); ~33 corregidas al primer ciclo.

### 💸 Informe de Pagos (Egresos) migrado a la NUBE
- Estaba en tarea local de Windows (`Informe_Pagos_EYG`, PC debía estar prendido a las 6PM) — falló los días 3-6/jul. **Ahora**: workflow `egresos-report.yml` (diario 6PM Colombia) + `scripts/egresos_report.py`: descarga el xlsm vía Graph del OneDrive de Andrea, **reporta desde el último informe enviado** (estado en `data/informe_pagos_estado.json`, nada queda sin reportar), **envía SIEMPRE** (día sin correo = alarma), remitente fijo **info@eygenergygroup.com** (pedido: que no salga del correo de Andrea). Tarea local DESHABILITADA. Probado con 3 corridas + informe real enviado (recuperó 3-7/jul).

### 👥 Clientes: Excel + anti-duplicados (SW v95-v96)
- Botón **📗 Excel** en pestaña Clientes (`cliDescargarExcel`): .xls con logo y colores E&G, respeta el filtro de búsqueda activo.
- **Emparejador anti-duplicados** (`_cliNorm`/`_cliMismo`): normaliza SAS/LTDA/SUCURSAL/paréntesis/"/ODS…"/NITs antes de comparar; prefijo mínimo 5 letras con palabra completa (METAL≈METAL INDUSTRIALES, METAL≠METALCO).
- **Limpieza única**: 59→50 fichas. 8 fusiones (Andes, CR Ingeniería, Petrolabin×3, Petroleum Blending×3, Skid Experts, TOC Energía, ARROW→ARROW EXPLORATION LLANOS SUCURSAL COLOMBIA, METAL→METAL INDUSTRIALES — las 2 últimas confirmadas por el usuario). Variantes anotadas en `notas` de cada ficha.

### 🔍 Análisis integral de TODOS los módulos (solo informativo, sin implementar)
Hecho con 3 exploradores. Hallazgos clave — **roto/falso**: "Actividad reciente" del Dashboard hardcodeada, "7 mensajes sin leer" fijo, calendario Reuniones no navega meses (prevMonth/nextMonth vacías), alerta "cotiz sin seguimiento" lee clave inexistente (nunca dispara), botón Teams de Equipo muerto, presencia online falsa, WhatsApp "todo el equipo" sin destinatario, workflows `schedule-reports.yml` (duplica reportes 5PM) y `teams-notify.yml` (roto) para apagar. **Riesgo de pérdida**: remiGuardar/pcomGuardar hacen DELETE+POST sin transacción; consecutivos de remisión/OC en memoria local (colisiones); localStorage 4,6MB sin manejo de cuota llena (falla silenciosa); Visitas/Mensajería cap 100 registros. **Seguridad**: RLS Supabase abierta + key en HTML, _proxySecret hardcodeado, hashes sin sal en config.json, defaults EYG2026. **Doble digitación**: OC cliente y Plan de Compras derivan ítems por separado; factura se digita en Procesos Y Cartera; remisión no descuenta stock. Prioridades propuestas: 1) arreglos rotos, 2) anti-pérdida, 3) Paso A Supabase, 4) seguridad, 5) cadena comercial ligada.

### 📌 Roadmap pendiente (aprobado en concepto, sin construir)
- **C** Precios sugeridos al cotizar (histórico del cliente + costo/margen). **B** Borrador de correo Outlook con PDF adjunto + marcar Enviada auto (usuario dijo SÍ). **E** Indicadores gerencia (conversión por vendedor, motivos de rechazo). **D** Alertas de vencimiento: "todavía no".
- **Paso A Supabase**: migrar cotizaciones (90 días) a Supabase con consecutivo del servidor + Realtime. Fases 1-2 sin riesgo con gente trabajando; corte final 15 min fuera de horario.
- Monitor de cotizaciones sin subir por usuario (heartbeat + panel dirección): propuesto, sin decidir.

### 🌙 Segunda parte de la sesión 07/jul (SW v97 → v103)
- **Arreglos rotos** (v97): actividad reciente del Dashboard ahora REAL, tarjeta "Alertas activas" (adiós al "7 mensajes" fijo), calendario de Reuniones navega meses, alerta de cotizaciones sin seguimiento revivida (una alerta resumen con valor en juego), botón Teams del Equipo abre chat real, WhatsApp "todo el equipo" copia al portapapeles, `schedule-reports.yml` eliminado (duplicaba reportes 5PM).
- **Chat ENERGY asesor técnico** (v98-v99): responde dudas de material mecánico/eléctrico/instrumentación; detector automático → Sonnet solo en técnicas; sugiere inventario E&G; ⚠️ de seguridad en temas críticos; reintento automático si el proxy chispea.
- **Módulo 🗓️ Programación Equipo** (v100-v101, v103): TODO el equipo lo VE, solo dirección (Alberto/Andrea/Mario) EDITA. Semana Lun-Sáb con ✅ Presencial/🏠 Remoto/🏖️ Permiso por persona, "⚡ Semana estándar", roster editable (presenciales: Andrea, Nelsy, Yesid, Alexandra, Lina, Nicolás · híbridos: Mario, Sandra), resumen mensual + Excel. La pestaña Sábados se mudó aquí desde Visitas. Datos: `data/programacion_equipo.json`. Cola de sync ahora reintenta sola cada 60s.
- **FIX layout histórico** (v102): un `</div>` sobrante tras Configuración cerraba `.content` antes de tiempo — 8 módulos (Inventario→Compras) vivían fuera del contenedor sin padding, con un hueco gigante arriba (el "espacio en blanco" del semanario). Corregido y verificado con parser HTML5.
- **Clientes** (v95-v96): botón 📗 Excel con marca, emparejador anti-duplicados normalizado, base depurada 59→50 (ARROW y METAL fusionados con confirmación).

### 📌 Pendientes al cierre (prioridad sugerida)
1) Cotizaciones: C precios sugeridos → B correo Outlook con PDF → E indicadores. 2) Paso A Supabase (aprobado). 3) Anti-pérdida remisiones/plan (DELETE+POST sin transacción, consecutivos locales) + cuota localStorage + cap 100. 4) Seguridad (RLS, secretos, contraseñas). 5) Cadena comercial ligada. Menores: WhatsApp real de Mario, tarifas flete, apagar sync-inventory.yml, redesplegar proxy (cap 16000), webhook Teams, informes por definir.

### 🌃 Cierre de la noche (SW v104 → v106)
- **v104**: Excel de Programación Equipo solo para dirección.
- **v105 — CONSECUTIVO DEL SERVIDOR** (Paso A esencial): al guardar cotización nueva se reclama el número en Supabase con INSERT atómico (PK id) — imposible duplicar consecutivos aunque dos personas guarden en el mismo milisegundo; probado con carrera real (201 vs 409). Backfill verificado: las 108 cotizaciones activas + ítems completos en Supabase. **Cutover total (quitar JSON) = opcional NO recomendado** (perdería el offline; el híbrido es el diseño correcto).
- **v106 — Plan de Sostenibilidad, sesión 1**: (1) guardado SIN PÉRDIDA en Remisiones y Plan de Compras (insertar nuevas → borrar viejas; adiós al DELETE+POST que podía borrar todo si fallaba la red); (2) aviso global de cuota del navegador llena (antes fallaba en silencio); (3) **informe 🩺 Salud del Sistema**: workflow `salud-sistema.yml` — día 1 de mes 7AM informe completo a Andrea + gerencia, lunes 7AM chequeo silencioso (solo escribe si hay alertas). Vigila backups (frescura y completitud), tamaños de data/ con tendencia, filas de Supabase, robots fallidos y tamaño del repo. Pendiente plan: sesión 2 (fotos→Supabase Storage + archivado automático) y sesión 3 (copia mensual OneDrive).

---

## 🚑 Sesión 08/jul/2026 (día) — líos urgentes + fixes (SW v107 → v112)

### Accesos del equipo
- **404 al abrir la dirección corta** → `404.html` en la raíz redirige a `/plataforma-eyg/Index.html` (GitHub Pages es case-sensitive, solo existe Index.html). Link SIEMPRE bueno: `.../plataforma-eyg/Index.html`.
- **Contraseña de Lina** es `Lina2026` (sin @). El login normaliza (minúsculas, sin espacios).
- **Base de Datos de cotizaciones vacía para Sandra** (SW v107): el filtro `.includes('comercial')` atrapaba a "Asistente Comercial" (Sandra) y "Coordinador Comercial" (Mario) y les mostraba solo las suyas → BD vacía. FIX: `isComercial=false` → TODO el equipo ve TODAS las cotizaciones (los costos/FACTOR siguen ocultos aparte).

### 🔴 BUG CRÍTICO: se borraban los precios de cotizaciones (SW v108)
- CAUSA: el consecutivo del servidor (v105) marcó cotizaciones con `renumeradaAt`; el reconciliador `_cotizReconciliar` prefería SIEMPRE la versión renumerada (aunque vieja y en $0) → al poner precios, la versión vieja los borraba. FIX: gana la de `updatedAt` más reciente; renumeradaAt solo desempata.
- **Trampa**: el fix solo protege a equipos ya actualizados. Un equipo con CÓDIGO VIEJO (Lina) seguía borrando precios y pisando el server con $0. Por eso TODOS deben recargar.
- Recuperación: la doble escritura a Supabase salvó LM1780 (restaurada blindada con renumeradaAt+updatedAt nuevos, aguanta código viejo). LM1767 (Petroleum Blending) se perdió en todas las copias → REDIGITAR (pendiente).
- **Indicador de versión** (v109): abajo-izquierda de la barra lateral muestra la vNNN real instalada (para detectar equipos desactualizados).

### Otros
- **Blindaje anti-duplicados Remisiones/Plan de Compras** (v110): `_dedupLineas` al leer, por si un corte de red dejó líneas repetidas (el guardado sin pérdida de v106).
- **Auto-vencidas revisado**: correcto, 0 errores (las 33 tienen fechaVenc realmente pasada).
- **Programación Equipo**: nuevo estado 🚗 Visita de campo (v111).
- **Semanario visitas** (v112): 2+ visitas el mismo día ahora se ven separadas (borde completo, numeradas 1/2, hora en badge) — antes se veían pegadas y parecían una sola.

### ⚠️ PENDIENTE al cierre
- **LM1767 (Petroleum Blending)**: redigitar precios (perdidos) — hacerlo en equipo v108+ o pasárselos a Claude para escribir blindados.
- **Confirmar que TODOS los equipos estén en la última versión** (mirar abajo-izq). Un equipo viejo puede volver a borrar precios.
- Regla de oro: dos personas NO deben editar la misma cotización a la vez (se pisan, gana el último que guarda).

## 🔧 Sesión 08/jul (tarde-noche) — Remisiones/Planes/Visitas + varios (SW v113 → v120)
- **Remisiones (v114) y Plan de Compras (v113) anti-borrado**: aviso que consulta el SERVIDOR antes de reemplazar un número/CC que ya existe. Causa: número de remisión se calcula desde memoria → colisiones; el plan usa CC determinístico por cotización → dos personas chocan. Caso real: SERSOLCA (rem 26150) pisó a QCIEN → QCIEN recuperada del backup `backups/remisiones.json`, SERSOLCA reubicada en 26202. (Otras 4 remisiones de junio bajo 26150 —SAR ENERGY/Wiler/PGS/MYM— quedan en el backup si se necesitan.)
- **Visitas comerciales**: borrado con lápida `deleted:true` para que NO revivan al sincronizar (v115); nombres completos en el semanario (Mario Rodríguez/Sandra Sánchez, solo etiqueta vía VIS_NOMBRES); quitado el campo "Cliente" del modal Programar visita → queda "Empresa/Cliente" (v116); FIX visitas que desaparecían al guardar/borrar otra —ahora une remoto+LOCAL+cambio— (v117); **visita con VARIOS comerciales** (casillas, campo `comerciales[]`, sale en la fila de cada uno con 👥, backward-compatible) (v118).
- **Programación Equipo**: agregada Sheila (híbrido) (v119).
- **Cotizaciones IA**: pestaña **Clientes** visible SOLO para gerencia (Alberto) y Andrea (v120).
- **Informe de pagos**: GitHub se saltó la corrida automática de las 6PM (congestión a la hora en punto); se disparó manual y se movió el cron a las 6:17PM (`17 23`) para que no se salte. Si falla un día, el siguiente recoge lo pendiente.

### ⚠️ PENDIENTE al cierre 08/jul
- **LM1767 (Petroleum Blending)**: redigitar precios (se perdieron en todas las copias).
- **Que TODOS los equipos estén en v120** (indicador abajo-izq de la barra lateral). Un equipo con versión vieja puede volver a borrar precios/datos.
- Mejora de fondo pendiente: **consecutivo de remisión del lado del servidor** (como el de cotizaciones) para que nunca se repita el número.
- Contraseña de Lina: `Lina2026`.

*Fin del archivo de memoria. Última actualización: 08/jul/2026 (sesión completa: SW v86 → v120)*
