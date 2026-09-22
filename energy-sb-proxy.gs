/**
 * ===== ENERGY — Proxy SOLO Supabase (carril rápido de escritura) =====
 *
 * POR QUÉ EXISTE (22-sep-2026)
 * El energy-proxy.gs hace tres cosas a la vez: habla con Claude, sube archivos a
 * GitHub y escribe en Supabase. El problema es el segundo: `data/cotizaciones.json`
 * pesa 4,5 MB y se sube entero en cada guardado — 239 veces en cinco horas ese día.
 * Apps Script atiende por turnos, así que la escritura de una OP hacía cola detrás
 * de esos 4,5 MB. Medido sobre 8 llamadas seguidas: 24,4 s de media, 70 s la peor,
 * 5 por encima de 15 s. Y como la plataforma cortaba a los 20 s, la OP no se creaba
 * — pero el consecutivo ya estaba pedido. Así se perdieron OP-2026-0097 a 0101.
 *
 * Este script hace UNA sola cosa: escribir en Supabase. Son peticiones de unos
 * pocos KB, así que nunca tiene que esperar detrás de una subida grande.
 *
 * NO REEMPLAZA AL OTRO. El energy-proxy.gs sigue igual, con su manejador de
 * Supabase intacto: si este carril falla, la plataforma se va por el de siempre
 * (ver `_sbWriteRaw` en Index.html). Por eso desplegarlo no puede romper nada.
 *
 * ─── CÓMO DESPLEGARLO (10 minutos, una sola vez) ────────────────────────────
 * 1. https://script.google.com/home → Nuevo proyecto. Llámalo "ENERGY Supabase".
 *    Tiene que ser un proyecto NUEVO: si lo pegas dentro del que ya existe,
 *    comparte turno con las subidas a GitHub y no sirve de nada.
 * 2. Pega TODO este archivo en Code.gs (borra lo que traiga por defecto).
 * 3. ⚙ Configuración del proyecto → Propiedades del script → Agregar:
 *      SUPABASE_SECRET    = sb_secret_...   (Supabase → Settings → API Keys → "secret")
 *      SHARED_SECRET      = eyg_prx_...     (el MISMO que está en Index.html)
 *      MIN_WRITE_VERSION  = 130             (frena pestañas con código viejo)
 *      RATE_LIMIT_PER_MIN = 600             (aquí solo entran escrituras pequeñas)
 *    La SUPABASE_SECRET y la SHARED_SECRET son las MISMAS del proxy que ya existe:
 *    ábrelo en la otra pestaña y cópialas de ahí. No hay que generar nada nuevo.
 * 4. Implementar → Nueva implementación → Aplicación web
 *      - Ejecutar como: Tu cuenta
 *      - Quién tiene acceso: Cualquier usuario, incluso anónimo
 * 5. Copia la URL que termina en /exec.
 * 6. En la plataforma → Configuración → "⚡ Proxy Supabase (carril rápido)":
 *    pégala y dale Guardar. Se reparte sola a todos los equipos vía
 *    data/config.json — nadie más tiene que hacer nada.
 *
 * Para comprobar que quedó bien: abre la URL /exec en el navegador. Tiene que
 * responder {"ok":true,...,"sbListo":true}. Si dice sbListo:false, falta la
 * propiedad SUPABASE_SECRET del paso 3.
 */

const PROPS = PropertiesService.getScriptProperties();

const SB_URL = 'https://juprjevxkcitqpsnemto.supabase.co/rest/v1';

// Solo las tablas de la plataforma; cualquier otra se rechaza. Es la misma lista
// del energy-proxy.gs: si algún día se agrega una tabla, hay que agregarla en los
// DOS sitios o la escritura funcionará por un carril y por el otro no.
const SB_TABLAS = ['productos','kardex','familias','conteos','conteo_items',
                   'remisiones','cotizaciones','cotizacion_items',
                   'plan_compras','oc_compras','proveedores',
                   'verificacion_despacho',
                   'ops','op_items','op_reservas','op_certificados','op_eventos',
                   'op_consecutivos','proveedor_sedes','proveedor_sede_memoria','zonas_ruta',
                   'plan_gastos','op_recepciones'];

// Funciones del servidor invocables (rpc/<nombre>). Una por una a propósito:
// un comodín aquí deja expuesta cualquier función de la base.
const SB_RPC = ['op_nuevo_numero'];

const RATE_LIMIT_DEFAULT = 600;

/** Token compartido. Sin la propiedad configurada se acepta todo (modo
 *  compatibilidad, para no dejar el carril muerto a mitad del despliegue). */
function _autorizado(body) {
  const secret = PROPS.getProperty('SHARED_SECRET');
  if (!secret) return true;
  return String(body && body.secret || '') === secret;
}

/** Límite por minuto (ventana de 60 s en CacheService). */
function _dentroDelLimite() {
  const cache = CacheService.getScriptCache();
  const limite = parseInt(PROPS.getProperty('RATE_LIMIT_PER_MIN')) || RATE_LIMIT_DEFAULT;
  const ventana = 'rl_' + Math.floor(Date.now() / 60000);
  const actual = parseInt(cache.get(ventana)) || 0;
  if (actual >= limite) return false;
  cache.put(ventana, String(actual + 1), 120);
  return true;
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!_autorizado(body)) return _json({ error: 'No autorizado' });
    if (!_dentroDelLimite()) return _json({ error: 'Límite de peticiones excedido, intenta en un momento' });
    if (body.sb && typeof body.sb === 'object') return _handleSupabaseWrite(body.sb);
    // A propósito no atiende ni `messages` (Claude) ni `file` (GitHub): si este
    // carril aceptara subidas grandes volvería a ser el cuello de botella que vino
    // a resolver. Esas siguen yendo al energy-proxy.gs.
    return _json({ error: 'Este proxy solo atiende escrituras de Supabase — falta sb{}' });
  } catch (err) {
    return _json({ error: 'Proxy error: ' + err.message });
  }
}

function doGet() {
  return _json({
    ok: true,
    service: 'ENERGY proxy Supabase (carril rápido)',
    handles: ['supabase-write'],
    sbListo: !!PROPS.getProperty('SUPABASE_SECRET'),
    minWriteVersion: parseInt(PROPS.getProperty('MIN_WRITE_VERSION')) || 0
  });
}

function _handleSupabaseWrite(sb) {
  const key = PROPS.getProperty('SUPABASE_SECRET');
  if (!key) return _json({ error: 'SUPABASE_SECRET no configurada en Propiedades del script' });

  const method = String(sb.method || '').toLowerCase();
  if (['post','patch','delete'].indexOf(method) < 0) {
    return _json({ error: 'Método no permitido: ' + sb.method });
  }
  const path = String(sb.path || '').replace(/^\/+/, '');
  const tabla = path.split('?')[0].split('/')[0];
  if (tabla === 'rpc') {
    const fn = path.split('?')[0].split('/')[1] || '';
    if (SB_RPC.indexOf(fn) < 0) return _json({ error: 'Función no permitida: ' + fn });
  } else if (SB_TABLAS.indexOf(tabla) < 0) {
    return _json({ error: 'Tabla no permitida: ' + tabla });
  }

  // Gate de versión mínima para ESCRIBIR (frena pestañas con código viejo, que es
  // lo que duplicó LM2137). Solo se exige si la propiedad existe.
  const minV = parseInt(PROPS.getProperty('MIN_WRITE_VERSION')) || 0;
  if (minV && (parseInt(sb.v) || 0) < minV) {
    return _json({ status: 426, body: JSON.stringify({ message:
      'Tu pestaña tiene una versión vieja de la plataforma. Recarga la página (Ctrl+F5) para seguir guardando.' }) });
  }

  const headers = { 'apikey': key, 'Authorization': 'Bearer ' + key };
  if (sb.prefer) headers['Prefer'] = String(sb.prefer);
  const opts = { method: method, headers: headers, muteHttpExceptions: true };
  if (sb.body != null && method !== 'delete') {
    opts.contentType = 'application/json';
    opts.payload = (typeof sb.body === 'string') ? sb.body : JSON.stringify(sb.body);
  }
  const resp = UrlFetchApp.fetch(SB_URL + '/' + path, opts);
  // El status REAL de Supabase viaja DENTRO del JSON, porque Apps Script siempre
  // responde 200. `_sbWriteRaw` lo desempaca: 201/204/409 llegan intactos.
  return _json({ status: resp.getResponseCode(), body: resp.getContentText() });
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
