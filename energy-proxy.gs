/**
 * ===== ENERGY — Proxy UNIFICADO (Claude + GitHub Writes) =====
 *
 * Reemplaza al claude-proxy.gs y al github-proxy.gs anteriores.
 * Un solo Apps Script, una sola URL, que detecta automáticamente
 * qué tipo de petición es:
 *   - Si el body tiene { messages:[...] }  → reenvía a Anthropic Claude
 *   - Si el body tiene { file, content }   → escribe el archivo en GitHub
 *   - Si el body tiene { sb:{...} }        → escribe en Supabase (RLS cerrada)
 *
 * Cómo desplegarlo (una sola vez):
 * 1. https://script.google.com/home → Nuevo proyecto.
 * 2. Pega TODO este archivo en Code.gs.
 * 3. ⚙ Configuración del proyecto → Propiedades del script → Agregar:
 *      CLAUDE_API_KEY = sk-ant-api03-...   (para el chat ENERGY IA)
 *      GH_TOKEN       = ghp_...            (PAT con permiso 'repo')
 *      GH_OWNER       = cami902026-oss
 *      GH_REPO        = plataforma-eyg
 *      GH_BRANCH      = main
 *      SHARED_SECRET  = eyg_prx_...        (token de seguridad; el MISMO que está en Index.html)
 *      SUPABASE_SECRET = sb_secret_...     (key SECRETA de Supabase: Dashboard → Settings →
 *                                           API Keys → "secret". Escribe saltándose la RLS;
 *                                           NUNCA va en el HTML, solo aquí.)
 *      RATE_LIMIT_PER_MIN = 240            (recomendado subirlo: ahora las escrituras de
 *                                           inventario/remisiones/cotizaciones pasan por aquí)
 *
 *   IMPORTANTE: agrega SHARED_SECRET SOLO cuando el Index.html que ya manda el token
 *   esté publicado (Ctrl+F5). Mientras la propiedad no exista, el proxy acepta todo
 *   (modo compatibilidad) — así no se cae el servicio durante el despliegue.
 * 4. Implementar → Nueva implementación → Aplicación web
 *      - Ejecutar como: Tu cuenta
 *      - Acceso: Cualquier usuario, incluso anónimo
 * 5. Copia la URL /exec.
 * 6. En la plataforma → Configuración pega la MISMA URL en:
 *      - "🤖 ENERGY IA — URL del Proxy seguro"
 *      - "📤 Proxy GitHub (escritura compartida sin PAT)"
 *
 * Después de pegar y guardar, todos los dispositivos del equipo la reciben
 * automáticamente vía data/config.json.
 */

const PROPS = PropertiesService.getScriptProperties();
const MODEL_DEFAULT  = 'claude-haiku-4-5-20251001';
const MAX_TOKENS_CAP = 16000;

// ─── Seguridad ────────────────────────────────────────────────────────────────
// Modelos permitidos a través de este proxy (evita que alguien gaste tu cuota en
// modelos caros). Cualquier otro modelo solicitado se degrada a MODEL_DEFAULT.
const MODELOS_PERMITIDOS = ['claude-haiku-4-5-20251001', 'claude-sonnet-4-6'];
// Límite de peticiones por minuto (global, todas las llamadas al proxy juntas).
// Configurable con la propiedad RATE_LIMIT_PER_MIN. Protege la factura ante abuso.
const RATE_LIMIT_DEFAULT = 90;

/**
 * Capa 1 — Token compartido. Devuelve true si la petición está autorizada.
 * Si la propiedad SHARED_SECRET NO está configurada, se permite todo
 * (modo compatibilidad, para no romper nada antes de terminar el despliegue).
 */
function _autorizado(body) {
  const secret = PROPS.getProperty('SHARED_SECRET');
  if (!secret) return true;                 // aún no configurado → no exige token
  return String(body && body.secret || '') === secret;
}

/**
 * Capa 2 — Límite de peticiones por minuto (CacheService, ventana de 60s).
 * Devuelve true si la petición está dentro del límite.
 */
function _dentroDelLimite() {
  const cache = CacheService.getScriptCache();
  const limite = parseInt(PROPS.getProperty('RATE_LIMIT_PER_MIN')) || RATE_LIMIT_DEFAULT;
  const ventana = 'rl_' + Math.floor(Date.now() / 60000);   // clave por minuto
  const actual = parseInt(cache.get(ventana)) || 0;
  if (actual >= limite) return false;
  cache.put(ventana, String(actual + 1), 120);              // expira en 2 min
  return true;
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');

    // Capa 1: token
    if (!_autorizado(body)) {
      return _json({ error: 'No autorizado' });
    }
    // Capa 2: límite por minuto
    if (!_dentroDelLimite()) {
      return _json({ error: 'Límite de peticiones excedido, intenta en un momento' });
    }

    // Detección por contenido
    if (Array.isArray(body.messages)) {
      return _handleClaude(body);
    }
    if (typeof body.file === 'string') {
      return _handleGitHubWrite(body);
    }
    if (body.sb && typeof body.sb === 'object') {
      return _handleSupabaseWrite(body.sb);
    }
    // Disparar un workflow de GitHub (aviso instantaneo de OP pendiente)
    if (body.wf && typeof body.wf === 'object') {
      return _handleWorkflowDispatch(body.wf);
    }
    // 🔗 Portal de vinculación (22-oct-2026): clientes/proveedores sin cuenta
    if (body.portal && typeof body.portal === 'object') {
      return _handlePortal(body.portal);
    }
    return _json({ error: 'Petición inválida — falta messages[], file o sb' });

  } catch (err) {
    return _json({ error: 'Proxy error: ' + err.message });
  }
}

function doGet() {
  return _json({
    ok: true,
    service: 'ENERGY proxy unificado',
    handles: ['claude','github-write','supabase-write','portal'],
    sbListo: !!PROPS.getProperty('SUPABASE_SECRET'),
    model: MODEL_DEFAULT
  });
}

// ─── Supabase write (RLS cerrada: la key publishable del HTML solo LEE; ─────
// ─── las escrituras entran por aquí con la key secreta del servidor) ────────
const SB_URL = 'https://juprjevxkcitqpsnemto.supabase.co/rest/v1';
// Solo las tablas de la plataforma; cualquier otra se rechaza.
const SB_TABLAS = ['productos','kardex','familias','conteos','conteo_items',
                   'remisiones','cotizaciones','cotizacion_items',
                   'plan_compras','oc_compras','proveedores',
                   'saldos_proveedor','saldos_movimientos',   // 💳 Saldos a favor (13-oct-2026)
                   'verificacion_despacho',   // SIG-CAL-FR-009 (6-ago-2026)
                   // OP — Órdenes de Pedido (23-ago-2026)
                   'ops','op_items','op_reservas','op_certificados','op_eventos',
                   'op_consecutivos','proveedor_sedes','proveedor_sede_memoria','zonas_ruta',
                   'plan_gastos',      // centro de costos del plan (24-ago-2026)
                   'op_recepciones',
                   'contratos_marco','contratos_consumos',
                   'fe_productos','fe_kardex','fe_clientes','fe_pedidos','fe_pedido_items',
                   'fe_remisiones','fe_remision_items',   // 🔧 Ferrenergy (20-oct-2026)   // contratos marco (9-oct-2026)
                   'portal_cuentas','portal_documentos'];   // 🔗 Portal de vinculación (22-oct-2026). portal_credenciales NO: solo por _handlePortal

// Funciones del servidor que la plataforma puede invocar (rpc/<nombre>).
// Se listan una por una a propósito: `rpc` no puede ser un comodín, o cualquier
// función de la base quedaría expuesta.
const SB_RPC = ['op_nuevo_numero', 'fe_nuevo_numero'];   // fe_: consecutivos FE-/PE- de Ferrenergy (20-oct-2026)

// ─── DISPARAR UN WORKFLOW DE GITHUB (23-ago-2026) ───────────────────────────
// La plataforma no puede llamar a la API de GitHub: haria falta un token en el
// navegador, y este repo es publico. El proxy si tiene GH_TOKEN en sus
// propiedades, asi que hace de puente.
//
// Se usa para que el aviso de "OP pendiente de aprobacion" salga AL INSTANTE
// en vez de esperar al ciclo de 15 minutos.
//
// Lista blanca por nombre de archivo: nunca un comodin. Un proxy que dispara
// cualquier workflow es un proxy que ejecuta codigo arbitrario en el repo.
var GH_WORKFLOWS = ['aviso-op-pendiente.yml',
                    'portal-aviso.yml'];   // 🔗 correos del portal de vinculación (22-oct-2026)

function _handleWorkflowDispatch(wf) {
  var nombre = String((wf && wf.file) || '');
  if (GH_WORKFLOWS.indexOf(nombre) < 0) {
    return _json({ ok: false, error: 'Workflow no permitido: ' + nombre });
  }
  var tok = PROPS.getProperty('GH_TOKEN');
  if (!tok) return _json({ ok: false, error: 'GH_TOKEN no configurado' });
  var owner = PROPS.getProperty('GH_OWNER') || 'cami902026-oss';
  var repo  = PROPS.getProperty('GH_REPO')  || 'plataforma-eyg';
  var url = 'https://api.github.com/repos/' + owner + '/' + repo
          + '/actions/workflows/' + encodeURIComponent(nombre) + '/dispatches';
  var r = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + tok, 'Accept': 'application/vnd.github+json' },
    // 22-oct-2026: `inputs` (todos string) para los workflows que los declaran; si no vienen, el JSON no los lleva.
    payload: JSON.stringify({ ref: PROPS.getProperty('GH_BRANCH') || 'main', inputs: (wf.inputs && typeof wf.inputs === 'object') ? wf.inputs : undefined }),
    muteHttpExceptions: true
  });
  var code = r.getResponseCode();
  return _json({ ok: code === 204, status: code, body: code === 204 ? '' : r.getContentText().slice(0, 200) });
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
    // Llamada a una función del servidor: se valida el NOMBRE contra su propia
    // lista, no basta con que empiece por rpc/.
    const fn = path.split('?')[0].split('/')[1] || '';
    if (SB_RPC.indexOf(fn) < 0) {
      return _json({ error: 'Función no permitida: ' + fn });
    }
  } else if (SB_TABLAS.indexOf(tabla) < 0) {
    return _json({ error: 'Tabla no permitida: ' + tabla });
  }
  // Gate de versión mínima para ESCRIBIR (protege contra pestañas con código viejo,
  // causa raíz del rebote LM1790). Solo se exige si la propiedad existe.
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
  // El status REAL de Supabase viaja dentro del JSON (Apps Script siempre responde 200);
  // el cliente lo desempaca en _sbWriteRaw. 201/204/409 etc. llegan intactos.
  return _json({ status: resp.getResponseCode(), body: resp.getContentText() });
}

// ─── Claude (chat IA) ────────────────────────────────────────────────────────
function _handleClaude(body) {
  const apiKey = PROPS.getProperty('CLAUDE_API_KEY');
  if (!apiKey) {
    return _json({ error: 'CLAUDE_API_KEY no configurada en Propiedades del script' });
  }

  // Capa 3: solo se permiten modelos de la lista blanca (evita gasto en modelos caros)
  const modeloPedido = body.model || MODEL_DEFAULT;
  const modelo = MODELOS_PERMITIDOS.indexOf(modeloPedido) >= 0 ? modeloPedido : MODEL_DEFAULT;

  const payload = {
    model:      modelo,
    max_tokens: Math.min(parseInt(body.max_tokens) || 1024, MAX_TOKENS_CAP),
    system:     body.system     || '',
    messages:   body.messages
  };
  if (body.tools)        payload.tools        = body.tools;
  if (body.tool_choice)  payload.tool_choice  = body.tool_choice;
  if (body.temperature !== undefined) payload.temperature = body.temperature;

  if (!payload.messages.length) {
    return _json({ error: 'messages[] vacío' });
  }

  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:           'post',
    contentType:      'application/json',
    headers:          { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload:          JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return ContentService
    .createTextOutput(resp.getContentText())
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GitHub write ────────────────────────────────────────────────────────────
function _handleGitHubWrite(body) {
  const tok    = PROPS.getProperty('GH_TOKEN');
  const owner  = PROPS.getProperty('GH_OWNER')  || 'cami902026-oss';
  const repo   = PROPS.getProperty('GH_REPO')   || 'plataforma-eyg';
  const branch = PROPS.getProperty('GH_BRANCH') || 'main';

  if (!tok) return _json({ ok:false, error:'GH_TOKEN no configurado en Propiedades del script' });

  const file = String(body.file || '').replace(/^\/+/, '');
  if (!file) return _json({ ok:false, error:'Falta el campo "file"' });

  let contentStr;
  if (typeof body.content === 'string') contentStr = body.content;
  else contentStr = JSON.stringify(body.content, null, 2);

  const label = body.label || ('Update ' + file);

  if (!_isAllowedFile(file)) {
    return _json({ ok:false, error:'Archivo no permitido: '+file });
  }

  const sha = _ghGetSha(owner, repo, file, branch, tok);
  let result = _ghPutFile(owner, repo, file, branch, tok, contentStr, label, sha);

  if (!result.ok && (result.status === 422 || result.status === 409)) {
    const sha2 = _ghGetSha(owner, repo, file, branch, tok);
    result = _ghPutFile(owner, repo, file, branch, tok, contentStr, label, sha2);
    if (result.ok) return _json({ ok:true, file:file, sha:result.sha, retried:true });
  }
  if (result.ok) return _json({ ok:true, file:file, sha:result.sha });
  return _json({ ok:false, error:'GitHub error '+result.status, body:result.body });
}

function _isAllowedFile(file) {
  const allow = [
    /^ordenes\.json$/,
    /^reuniones\.json$/,
    /^data\/[a-z0-9_\-]+\.json$/i,
  ];
  return allow.some(re => re.test(file));
}

function _ghGetSha(owner, repo, file, branch, tok) {
  const url = 'https://api.github.com/repos/'+owner+'/'+repo+'/contents/'+encodeURI(file)+'?ref='+branch;
  const resp = UrlFetchApp.fetch(url, {
    method:'get',
    headers:{ 'Authorization':'Bearer '+tok, 'Accept':'application/vnd.github+json' },
    muteHttpExceptions:true
  });
  if (resp.getResponseCode() === 200) {
    const j = JSON.parse(resp.getContentText());
    return j.sha;
  }
  return null;
}

function _ghPutFile(owner, repo, file, branch, tok, contentStr, label, sha) {
  const url = 'https://api.github.com/repos/'+owner+'/'+repo+'/contents/'+encodeURI(file);
  const payload = {
    message: label + ' — ' + new Date().toLocaleString('es-CO'),
    content: Utilities.base64Encode(contentStr, Utilities.Charset.UTF_8),
    branch: branch
  };
  if (sha) payload.sha = sha;
  const resp = UrlFetchApp.fetch(url, {
    method:'put',
    contentType:'application/json',
    headers:{ 'Authorization':'Bearer '+tok, 'Accept':'application/vnd.github+json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions:true
  });
  const status = resp.getResponseCode();
  const text = resp.getContentText();
  if (status >= 200 && status < 300) {
    const j = JSON.parse(text);
    return { ok:true, sha: j.content && j.content.sha };
  }
  return { ok:false, status:status, body:text };
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
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
