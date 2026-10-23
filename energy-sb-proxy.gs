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
                   'saldos_proveedor','saldos_movimientos',   // 💳 Saldos a favor (13-oct-2026)
                   'verificacion_despacho',
                   'ops','op_items','op_reservas','op_certificados','op_eventos',
                   'op_consecutivos','proveedor_sedes','proveedor_sede_memoria','zonas_ruta',
                   'plan_gastos','op_recepciones',
                   'contratos_marco','contratos_consumos',
                   'fe_productos','fe_kardex','fe_clientes','fe_pedidos','fe_pedido_items',
                   'fe_remisiones','fe_remision_items',   // 🔧 Ferrenergy (20-oct-2026)   // contratos marco (9-oct-2026)
                   'portal_cuentas','portal_documentos'];   // 🔗 Portal de vinculación (22-oct-2026). portal_credenciales NO: solo por _handlePortal

// Funciones del servidor invocables (rpc/<nombre>). Una por una a propósito:
// un comodín aquí deja expuesta cualquier función de la base.
const SB_RPC = ['op_nuevo_numero', 'fe_nuevo_numero'];   // fe_: consecutivos FE-/PE- de Ferrenergy (20-oct-2026)

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
    if (body.portal && typeof body.portal === 'object') return _handlePortal(body.portal);   // 🔗 Portal de vinculación (22-oct-2026)
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
    handles: ['supabase-write','portal'],
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

// ═══════════════════════════════════════════════════════════════════════════
// 🔗 PORTAL DE VINCULACIÓN — clientes y proveedores suben sus documentos (22-oct-2026)
//
// El portal (portal.html) es una página PÚBLICA: la abre gente sin cuenta en la
// plataforma. Por eso NO lee Supabase con la key publishable (vería los datos de
// los demás) ni escribe por `sb{}` (podría tocar cualquier tabla). Todo lo del
// portal entra por aquí, con `body.portal = {accion, ...}`:
//   · el proxy valida NIT + clave contra `portal_credenciales` (hash con sal) y
//     entrega un pase temporal (6 h, en CacheService);
//   · con el pase, el portal solo ve y toca lo de SU NIT;
//   · los archivos NO pasan por Apps Script: el proxy emite una URL firmada de
//     subida y el navegador sube directo a Supabase Storage (bucket privado
//     `vinculacion`). Igual para ver: URL firmada de 10 minutos.
// Las acciones `invitar` y `ver_url` son de la plataforma (van con el SHARED_SECRET
// como todo lo demás). Este bloque está IGUAL en energy-proxy.gs y en
// energy-sb-proxy.gs: si se cambia, cambiar en los dos.
// ═══════════════════════════════════════════════════════════════════════════
var PORTAL_BUCKET = 'vinculacion';
var PORTAL_MAX_BYTES = 15 * 1024 * 1024;
var PORTAL_MIMES = ['application/pdf', 'image/jpeg', 'image/png'];
var PORTAL_DOCS = {
  rut:      { nombre: 'RUT (Registro Único Tributario)', vigenciaDias: 0 },
  camara:   { nombre: 'Cámara de Comercio',               vigenciaDias: 30 },
  cedula:   { nombre: 'Cédula del representante legal',  vigenciaDias: 0 },
  bancaria: { nombre: 'Certificación bancaria',          vigenciaDias: 0 }
};
var PORTAL_SESION_SEG = 21600;   // 6 h, el máximo de CacheService

function _ptKey() { return PROPS.getProperty('SUPABASE_SECRET'); }
function _ptBase() { return SB_URL.replace(/\/rest\/v1\/?$/, ''); }
function _ptNit(v) { return String(v || '').split('-')[0].replace(/\D/g, ''); }
function _ptHex(bytes) { return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join(''); }
function _ptHash(clave, sal) {
  return _ptHex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + ':' + clave, Utilities.Charset.UTF_8));
}
function _ptCodigo() {   // 6 caracteres sin confusables (0/O, 1/I/L)
  var abc = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 6; i++) s += abc.charAt(Math.floor(Math.random() * abc.length));
  return s;
}
// REST de Supabase con la key secreta. Devuelve {status, data}.
function _ptRest(path, method, body, prefer) {
  var key = _ptKey();
  var h = { 'apikey': key, 'Authorization': 'Bearer ' + key };
  if (prefer) h['Prefer'] = prefer;
  var o = { method: method || 'get', headers: h, muteHttpExceptions: true };
  if (body != null) { o.contentType = 'application/json'; o.payload = JSON.stringify(body); }
  var r = UrlFetchApp.fetch(SB_URL + '/' + path, o);
  var t = r.getContentText(), d = null;
  try { d = t ? JSON.parse(t) : null; } catch (_) { d = t; }
  return { status: r.getResponseCode(), data: d };
}
function _ptStorage(path, method, body) {
  var key = _ptKey();
  var o = { method: method || 'post', headers: { 'apikey': key, 'Authorization': 'Bearer ' + key }, muteHttpExceptions: true };
  if (body != null) { o.contentType = 'application/json'; o.payload = JSON.stringify(body); }
  var r = UrlFetchApp.fetch(_ptBase() + '/storage/v1/' + path, o);
  var t = r.getContentText(), d = null;
  try { d = t ? JSON.parse(t) : null; } catch (_) { d = t; }
  return { status: r.getResponseCode(), data: d };
}
function _ptSesionNueva(nit) {
  var tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('ptk_' + tok, nit, PORTAL_SESION_SEG);
  return tok;
}
function _ptSesion(tok) {
  if (!tok || String(tok).length < 32) return null;
  return CacheService.getScriptCache().get('ptk_' + tok) || null;
}
function _ptCuenta(nit) {
  var r = _ptRest('portal_cuentas?nit=eq.' + nit + '&select=id,nit,razon_social,tipo,correo,contacto,telefono,estado,activado_at,ultimo_ingreso&limit=1');
  return (r.status === 200 && r.data && r.data[0]) || null;
}
function _ptCred(nit) {
  var r = _ptRest('portal_credenciales?nit=eq.' + nit + '&limit=1');
  return (r.status === 200 && r.data && r.data[0]) || null;
}
function _ptDocs(nit) {
  var r = _ptRest('portal_documentos?nit=eq.' + nit + '&select=id,tipo_doc,nombre_archivo,ruta,tamano,mime,fecha_expedicion,vence,estado,observacion,revisado_at,subido_at&order=subido_at.desc&limit=200');
  return (r.status === 200 && Array.isArray(r.data)) ? r.data : [];
}
function _ptEstado(nit) {
  return { ok: true, cuenta: _ptCuenta(nit), documentos: _ptDocs(nit), requeridos: PORTAL_DOCS };
}
function _ptSafeName(n) {
  return String(n || 'archivo').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9._-]+/g, '_').slice(-80);
}

function _handlePortal(p) {
  if (!_ptKey()) return _json({ ok: false, error: 'SUPABASE_SECRET no configurada en Propiedades del script' });
  var accion = String(p.accion || '');
  var nit, cred, cuenta, r, ahora = new Date();

  // ── Primera vez: NIT + código del correo → crea su clave ─────────────────
  if (accion === 'activar') {
    nit = _ptNit(p.nit);
    var codigo = String(p.codigo || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    var clave = String(p.clave || '');
    if (!nit) return _json({ ok: false, error: 'Escribe el NIT (solo números, sin dígito de verificación).' });
    if (clave.length < 6) return _json({ ok: false, error: 'La clave debe tener al menos 6 caracteres.' });
    cred = _ptCred(nit); cuenta = _ptCuenta(nit);
    if (!cred || !cuenta) return _json({ ok: false, error: 'Ese NIT no tiene invitación. Pide a E&G que te invite al portal.' });
    if (cuenta.estado === 'bloqueado') return _json({ ok: false, error: 'La cuenta está bloqueada. Comunícate con E&G.' });
    if (!cred.codigo_invitacion || cred.codigo_invitacion !== codigo) return _json({ ok: false, error: 'El código de invitación no coincide. Revisa el correo que te envió E&G.' });
    var sal = Utilities.getUuid();
    r = _ptRest('portal_credenciales?nit=eq.' + nit, 'patch',
      { clave_hash: _ptHash(clave, sal), sal: sal, codigo_invitacion: null, intentos_fallidos: 0, bloqueado_hasta: null, updated_at: ahora.toISOString() }, 'return=minimal');
    if (r.status >= 300) return _json({ ok: false, error: 'No se pudo guardar la clave (' + r.status + ').' });
    _ptRest('portal_cuentas?nit=eq.' + nit, 'patch',
      { estado: (cuenta.estado === 'invitado' ? 'activo' : cuenta.estado), activado_at: cuenta.activado_at || ahora.toISOString(), ultimo_ingreso: ahora.toISOString(), updated_at: ahora.toISOString() }, 'return=minimal');
    var e1 = _ptEstado(nit); e1.token = _ptSesionNueva(nit); return _json(e1);
  }

  // ── Entrar: NIT + clave ──────────────────────────────────────────────────
  if (accion === 'entrar') {
    nit = _ptNit(p.nit);
    var cl = String(p.clave || '');
    if (!nit || !cl) return _json({ ok: false, error: 'Escribe el NIT y la clave.' });
    cred = _ptCred(nit); cuenta = _ptCuenta(nit);
    if (!cred || !cuenta) return _json({ ok: false, error: 'Ese NIT no está registrado en el portal.' });
    if (cuenta.estado === 'bloqueado') return _json({ ok: false, error: 'La cuenta está bloqueada. Comunícate con E&G.' });
    if (cred.bloqueado_hasta && new Date(cred.bloqueado_hasta) > ahora) return _json({ ok: false, error: 'Demasiados intentos. Espera 15 minutos y vuelve a intentar.' });
    if (!cred.clave_hash) return _json({ ok: false, error: 'Todavía no has creado tu clave. Usa «Primera vez» con el código del correo.', primeraVez: true });
    if (_ptHash(cl, cred.sal || '') !== cred.clave_hash) {
      var n = (cred.intentos_fallidos || 0) + 1;
      var upd = { intentos_fallidos: n, updated_at: ahora.toISOString() };
      if (n >= 5) { upd.bloqueado_hasta = new Date(ahora.getTime() + 15 * 60000).toISOString(); upd.intentos_fallidos = 0; }
      _ptRest('portal_credenciales?nit=eq.' + nit, 'patch', upd, 'return=minimal');
      return _json({ ok: false, error: 'Clave incorrecta.' + (n >= 5 ? ' La cuenta queda bloqueada 15 minutos.' : '') });
    }
    _ptRest('portal_credenciales?nit=eq.' + nit, 'patch', { intentos_fallidos: 0, bloqueado_hasta: null }, 'return=minimal');
    _ptRest('portal_cuentas?nit=eq.' + nit, 'patch', { ultimo_ingreso: ahora.toISOString() }, 'return=minimal');
    var e2 = _ptEstado(nit); e2.token = _ptSesionNueva(nit); return _json(e2);
  }

  // ── Con pase: lo de SU NIT ───────────────────────────────────────────────
  nit = _ptSesion(p.token);
  if (accion === 'estado' || accion === 'subir_url' || accion === 'confirmar' || accion === 'cambiar_clave') {
    if (!nit) return _json({ ok: false, error: 'La sesión venció. Vuelve a entrar.', sesion: false });
  }
  if (accion === 'estado') return _json(_ptEstado(nit));

  if (accion === 'cambiar_clave') {
    cred = _ptCred(nit);
    if (!cred || _ptHash(String(p.actual || ''), cred.sal || '') !== cred.clave_hash) return _json({ ok: false, error: 'La clave actual no coincide.' });
    if (String(p.nueva || '').length < 6) return _json({ ok: false, error: 'La clave nueva debe tener al menos 6 caracteres.' });
    var sal2 = Utilities.getUuid();
    r = _ptRest('portal_credenciales?nit=eq.' + nit, 'patch', { clave_hash: _ptHash(String(p.nueva), sal2), sal: sal2, updated_at: ahora.toISOString() }, 'return=minimal');
    return _json({ ok: r.status < 300, error: r.status < 300 ? null : 'No se pudo cambiar la clave.' });
  }

  if (accion === 'subir_url') {
    var tipo = String(p.tipo || '');
    if (!PORTAL_DOCS[tipo]) return _json({ ok: false, error: 'Tipo de documento no válido.' });
    if (PORTAL_MIMES.indexOf(String(p.mime || '')) < 0) return _json({ ok: false, error: 'Solo se aceptan PDF, JPG o PNG.' });
    if (!(+p.tamano > 0) || +p.tamano > PORTAL_MAX_BYTES) return _json({ ok: false, error: 'El archivo debe pesar entre 1 byte y 15 MB.' });
    var ruta = nit + '/' + tipo + '/' + Date.now() + '_' + _ptSafeName(p.nombre);
    r = _ptStorage('object/upload/sign/' + PORTAL_BUCKET + '/' + ruta, 'post', {});
    if (r.status >= 300 || !r.data || !r.data.url) return _json({ ok: false, error: 'No se pudo preparar la subida (' + r.status + '): ' + JSON.stringify(r.data).slice(0, 160) });
    return _json({ ok: true, ruta: ruta, subirUrl: _ptBase() + '/storage/v1' + r.data.url, bucket: PORTAL_BUCKET });
  }

  if (accion === 'confirmar') {
    var tipo2 = String(p.tipo || ''), ruta2 = String(p.ruta || '');
    if (!PORTAL_DOCS[tipo2]) return _json({ ok: false, error: 'Tipo de documento no válido.' });
    if (ruta2.indexOf(nit + '/' + tipo2 + '/') !== 0) return _json({ ok: false, error: 'La ruta no corresponde a tu NIT.' });
    // ¿El archivo de verdad quedó en el bucket? (lista la carpeta y busca el nombre)
    var carpeta = ruta2.slice(0, ruta2.lastIndexOf('/')), archivo = ruta2.slice(ruta2.lastIndexOf('/') + 1);
    var ls = _ptStorage('object/list/' + PORTAL_BUCKET, 'post', { prefix: carpeta, limit: 100, search: archivo });
    var existe = Array.isArray(ls.data) && ls.data.some(function (o) { return o && o.name === archivo; });
    if (!existe) return _json({ ok: false, error: 'El archivo no llegó al servidor. Intenta subirlo de nuevo.' });
    var fexp = /^\d{4}-\d{2}-\d{2}$/.test(String(p.fecha_expedicion || '')) ? String(p.fecha_expedicion) : null;
    var vd = PORTAL_DOCS[tipo2].vigenciaDias, vence = null;
    if (vd) {
      if (!fexp) return _json({ ok: false, error: 'Indica la fecha de expedición del certificado.' });
      var d = new Date(fexp + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + vd); vence = d.toISOString().slice(0, 10);
      if (new Date(vence + 'T12:00:00Z') < ahora) return _json({ ok: false, error: 'Ese certificado ya venció (más de ' + vd + ' días). Sube uno reciente.' });
    }
    // El anterior del mismo tipo queda como historial
    _ptRest('portal_documentos?nit=eq.' + nit + '&tipo_doc=eq.' + tipo2 + '&estado=in.(pendiente,aprobado,rechazado)', 'patch', { estado: 'reemplazado' }, 'return=minimal');
    r = _ptRest('portal_documentos', 'post', {
      nit: nit, tipo_doc: tipo2, nombre_archivo: String(p.nombre || archivo).slice(0, 200), ruta: ruta2,
      tamano: +p.tamano || null, mime: String(p.mime || ''), fecha_expedicion: fexp, vence: vence, estado: 'pendiente', subido_at: ahora.toISOString()
    }, 'return=minimal');
    if (r.status >= 300) return _json({ ok: false, error: 'No se pudo registrar el documento (' + r.status + ').' });
    // Si la cuenta estaba aprobada y entra un documento nuevo, vuelve a revisión
    cuenta = _ptCuenta(nit);
    if (cuenta && cuenta.estado === 'aprobado') _ptRest('portal_cuentas?nit=eq.' + nit, 'patch', { estado: 'activo', updated_at: ahora.toISOString() }, 'return=minimal');
    return _json(_ptEstado(nit));
  }

  // ── De la plataforma (ya viene con el SHARED_SECRET) ─────────────────────
  if (accion === 'invitar') {
    nit = _ptNit(p.nit);
    var razon = String(p.razon_social || '').trim().toUpperCase(), correo = String(p.correo || '').trim().toLowerCase();
    if (!nit || nit.length < 6) return _json({ ok: false, error: 'NIT inválido.' });
    if (!razon) return _json({ ok: false, error: 'Falta la razón social.' });
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) return _json({ ok: false, error: 'Correo inválido.' });
    var tipoC = (String(p.tipo || '') === 'proveedor') ? 'proveedor' : 'cliente';
    var existente = _ptCuenta(nit);
    var cuerpo = { nit: nit, razon_social: razon, tipo: tipoC, correo: correo, contacto: p.contacto || null, telefono: p.telefono || null,
                   notas: p.notas || null, updated_at: ahora.toISOString() };
    if (!existente) { cuerpo.estado = 'invitado'; cuerpo.invitado_por = p.por || null; cuerpo.invitado_at = ahora.toISOString(); }
    r = _ptRest('portal_cuentas?on_conflict=nit', 'post', cuerpo, 'resolution=merge-duplicates,return=minimal');
    if (r.status >= 300) return _json({ ok: false, error: 'No se pudo guardar la cuenta (' + r.status + '): ' + JSON.stringify(r.data).slice(0, 160) });
    var cod = _ptCodigo();
    // El código nuevo deja crear (o re-crear) la clave; la clave vieja, si había, sigue sirviendo.
    r = _ptRest('portal_credenciales?on_conflict=nit', 'post', { nit: nit, codigo_invitacion: cod, intentos_fallidos: 0, bloqueado_hasta: null, updated_at: ahora.toISOString() }, 'resolution=merge-duplicates,return=minimal');
    if (r.status >= 300) return _json({ ok: false, error: 'No se pudo guardar el código (' + r.status + ').' });
    return _json({ ok: true, nit: nit, codigo: cod, nueva: !existente });
  }

  if (accion === 'ver_url') {
    var ruta3 = String(p.ruta || '');
    if (!ruta3 || ruta3.indexOf('..') >= 0) return _json({ ok: false, error: 'Ruta inválida.' });
    r = _ptStorage('object/sign/' + PORTAL_BUCKET + '/' + ruta3, 'post', { expiresIn: 600 });
    if (r.status >= 300 || !r.data || !r.data.signedURL) return _json({ ok: false, error: 'No se pudo firmar la descarga (' + r.status + ').' });
    return _json({ ok: true, url: _ptBase() + '/storage/v1' + r.data.signedURL });
  }

  return _json({ ok: false, error: 'Acción de portal desconocida: ' + accion });
}
