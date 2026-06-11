/**
 * ===== ENERGY — WhatsApp Bot (Twilio Sandbox + Claude tools) =====
 *
 * Recibe mensajes de WhatsApp via Twilio webhook → identifica al usuario
 * por su número → llama a Claude con las herramientas → devuelve respuesta.
 *
 * Cómo desplegarlo:
 * 1. https://script.google.com/home → Nuevo proyecto.
 * 2. Pega TODO este archivo en Code.gs.
 * 3. ⚙ Configuración del proyecto → Propiedades del script → Agregar:
 *      CLAUDE_API_KEY  = sk-ant-api03-...
 *      GH_TOKEN        = ghp_... (PAT con permiso 'repo')
 *      GH_OWNER        = cami902026-oss
 *      GH_REPO         = plataforma-eyg
 *      GH_BRANCH       = main
 *      TWILIO_SID      = AC...   (tu Account SID de Twilio Console)
 *      TWILIO_TOKEN    = ...     (tu Auth Token de Twilio Console)
 *      TWILIO_NUMBER   = whatsapp:+14155238886
 *      WHATSAPP_WEBHOOK_TOKEN = (un token secreto inventado, p.ej. eyg_wa_a1b2c3...)
 * 4. Implementar → Nueva implementación → Aplicación web
 *      - Ejecutar como: Tu cuenta
 *      - Acceso: Cualquier usuario, incluso anónimo
 * 5. Copia la URL /exec
 * 6. Ve a https://console.twilio.com → Messaging → Try it out → Sandbox settings
 *    En "When a message comes in" pega la URL /exec AÑADIENDO al final
 *    ?token=EL_MISMO_WHATSAPP_WEBHOOK_TOKEN  → Save
 *    (así solo Twilio conoce la URL con el token; sin él, las peticiones se rechazan)
 *
 * Uso:
 * Cualquier usuario autorizado (Alberto, Andrea, Sheila, Alexandra, Lina)
 * envía WhatsApp a +1 415 523 8886. Primera vez: 'join result-experience'.
 * Después: cualquier mensaje natural en español.
 *
 * Bloqueado: Nelsy (usa la web directamente).
 */

const PROPS = PropertiesService.getScriptProperties();
const CACHE = CacheService.getScriptCache();
const MODEL = 'claude-haiku-4-5-20251001';

// Usuarios autorizados (se chequea por número de WhatsApp normalizado)
// Mantener en sincronía con team.json. Nelsy queda EXPLÍCITAMENTE bloqueada.
const AUTHORIZED = {
  '+573113134451': { name:'Alberto',       role:'JEFE',         email:'alberto'   },
  '+573204947227': { name:'Sheila Baron',  role:'COLABORADOR',  email:'sheila'    },
  '+573144858382': { name:'Alexandra',     role:'COLABORADOR',  email:'alexandra' },
  '+573107574110': { name:'Andrea',        role:'ADMIN',        email:'andrea'    },
  // Lina pendiente: agregar su número aquí cuando lo tengas
  // '+573xxxxxxxxx': { name:'Lina Cifientes', role:'COLABORADOR', email:'lina' },
};
const DENIED = {
  '+573125099056': 'Nelsy', // contabilidad — usa la web
};

function doPost(e) {
  try {
    const params = e.parameter || {};

    // Seguridad: token secreto en la URL del webhook (Apps Script no puede leer el
    // header X-Twilio-Signature, así que validamos un token que solo Twilio conoce).
    // La URL en Twilio debe terminar en ?token=...  Si WHATSAPP_WEBHOOK_TOKEN no está
    // configurado, se permite todo (modo compatibilidad para no romper el despliegue).
    const tokenEsperado = PROPS.getProperty('WHATSAPP_WEBHOOK_TOKEN');
    if (tokenEsperado && String(params.token || '') !== tokenEsperado) {
      Logger.log('WhatsApp: token de webhook inválido — petición rechazada');
      return _twiml('');
    }

    const from = (params.From || '').replace('whatsapp:','').trim();
    const body = (params.Body || '').trim();

    // Adjuntos de WhatsApp (Twilio): fotos / PDF de solicitudes de cotización
    const media = [];
    const numMedia = parseInt(params.NumMedia || '0', 10) || 0;
    for (var mi = 0; mi < numMedia; mi++) {
      var murl = params['MediaUrl' + mi];
      var mct  = params['MediaContentType' + mi] || '';
      if (murl) media.push({ url: murl, ctype: mct });
    }

    if (!from) return _twiml('Error: no recibí tu número.');
    if (!body && !media.length) return _twiml('Hola — escríbeme algo (o envía la foto/PDF de la solicitud) y te ayudo.');

    if (DENIED[from]) {
      return _twiml(`Hola ${DENIED[from]}, este bot no está habilitado para tu rol. Usa la plataforma web: https://cami902026-oss.github.io/plataforma-eyg/Index.html`);
    }
    const user = AUTHORIZED[from];
    if (!user) {
      return _twiml(`Tu número (${from}) no está autorizado para usar el bot ENERGY. Si crees que es un error, contacta a Andrea.`);
    }

    // Llamar a Claude con tools (texto + adjuntos)
    const reply = chatWithClaude(body, user, from, media);
    return _twiml(reply);

  } catch (err) {
    Logger.log('doPost error: ' + err);
    return _twiml('Hubo un error. Intenta de nuevo en un momento.');
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    ok: true, service: 'ENERGY WhatsApp bot', authorizedCount: Object.keys(AUTHORIZED).length
  })).setMimeType(ContentService.MimeType.JSON);
}

function _twiml(message) {
  const safe = String(message)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const xml = '<?xml version="1.0" encoding="UTF-8"?><Response><Message>' + safe + '</Message></Response>';
  return ContentService.createTextOutput(xml).setMimeType(ContentService.MimeType.XML);
}

// ─── HERRAMIENTAS DISPONIBLES PARA EL BOT ─────────────────────────────────
const TOOLS = [
  {
    name: 'crear_tarea',
    description: 'Crea una tarea nueva en el módulo Tareas (visible para todo el equipo).',
    input_schema: {
      type:'object',
      properties: {
        titulo:      { type:'string' },
        responsable: { type:'string', description:'Alberto, Sheila Baron, Alexandra, Andrea, Nelsy o Lina Cifientes' },
        fecha:       { type:'string', description:'YYYY-MM-DD' },
        prioridad:   { type:'string', enum:['Alta','Media','Baja'] },
        desc:        { type:'string' }
      },
      required: ['titulo','responsable','fecha']
    }
  },
  {
    name: 'marcar_etapa_oc',
    description: 'Marca una etapa de Orden de Compra como completada. Etapas: compra, entrega, certificado, factura.',
    input_schema: {
      type:'object',
      properties: {
        num:   { type:'string', description:'Número de la O.C. (ej: LM-1389, ODC-4049)' },
        etapa: { type:'string', enum:['compra','entrega','certificado','factura'] },
        fecha: { type:'string', description:'YYYY-MM-DD (default hoy)' },
        nota:  { type:'string' }
      },
      required: ['num','etapa']
    }
  },
  {
    name: 'listar_oc_pendientes',
    description: 'Lista las O.C. activas con su etapa pendiente. Útil para "qué OC están vencidas" o "estado de las OC".',
    input_schema: {
      type:'object',
      properties: {
        cliente: { type:'string', description:'Filtrar por nombre de cliente (opcional)' }
      }
    }
  },
  {
    name: 'listar_tareas_pendientes',
    description: 'Lista tareas pendientes (no completadas).',
    input_schema: {
      type:'object',
      properties: {
        responsable: { type:'string', description:'Filtrar por responsable (opcional)' }
      }
    }
  },
  {
    name: 'crear_oc',
    description: 'Crea una Orden de Pedido nueva en el módulo Procesos OC. Usa el número que vino del cliente (ej. LM1500, ODC-4170, 4600007246). Si la fecha de compra no se da, usa hoy.',
    input_schema: {
      type:'object',
      properties: {
        num:           { type:'string', description:'Número de la OC tal como lo envió el cliente' },
        cliente:       { type:'string', description:'Nombre del cliente (ej. PETRORIOS, SAR ENERGY, CIAM)' },
        desc:          { type:'string', description:'Descripción de lo que se está comprando' },
        fecha_compra:  { type:'string', description:'YYYY-MM-DD (default hoy)' },
        nota:          { type:'string', description:'Nota opcional' }
      },
      required: ['num','cliente']
    }
  },
  {
    name: 'crear_solicitud_cotizacion',
    description: 'Registra una SOLICITUD DE COTIZACIÓN de un cliente para que el equipo de cotización (Alexandra/Lina) la trabaje. Úsala cuando la comercial envíe una lista de productos, una foto de un pedido o un PDF pidiendo precios/cotización. Extrae los productos de la imagen/PDF si viene adjunto.',
    input_schema: {
      type:'object',
      properties: {
        cliente:       { type:'string', description:'Empresa o persona que necesita los materiales' },
        contacto:      { type:'string', description:'Nombre del contacto del cliente si aparece' },
        productos:     { type:'array', description:'Productos solicitados', items:{ type:'object', properties:{ descripcion:{type:'string'}, cantidad:{type:'string'} } } },
        formaPago:     { type:'string', description:'Condición de pago si se menciona' },
        urgencia:      { type:'string', enum:['alta','media','baja'], description:'alta si es urgente/hoy-mañana, baja si más de una semana, media en otro caso' },
        observaciones: { type:'string', description:'Dato adicional: lugar de entrega, especificaciones, etc.' }
      },
      required: ['cliente','productos']
    }
  }
];

// ─── LOOP AGÉNTICO CON CLAUDE ─────────────────────────────────────────────
function chatWithClaude(prompt, user, from, media) {
  const apiKey = PROPS.getProperty('CLAUDE_API_KEY');
  if (!apiKey) return 'Error: CLAUDE_API_KEY no configurada en el script.';

  const today = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd');
  const system = 'Eres ENERGY, asistente vía WhatsApp para E&G Energy Group SAS. Hoy es ' + today + '. ' +
    'Usuario actual: ' + user.name + ' (' + user.role + '). ' +
    'Sé MUY BREVE — máximo 2-3 frases en cada respuesta. Sin listas largas. ' +
    'Cuando el usuario pida hacer algo (crear tarea, marcar O.C., etc.), USA las herramientas. ' +
    'Si te envían una lista de productos, una foto de un pedido o un PDF pidiendo precios/cotización ' +
    '(típicamente de la comercial), usa la herramienta crear_solicitud_cotizacion para registrar la ' +
    'solicitud — la tomarán Alexandra o Lina. Lee bien la imagen/PDF para extraer los productos. ' +
    'Después de ejecutar, confirma brevemente. Sin emojis excesivos.';

  // Primer mensaje: si vienen adjuntos (fotos/PDF), pasarlos a la visión de Claude
  var primerContenido;
  if (media && media.length) {
    primerContenido = [];
    for (var k = 0; k < media.length && k < 5; k++) {
      var bloque = _mediaABloqueClaude(media[k]);
      if (bloque) primerContenido.push(bloque);
    }
    primerContenido.push({ type:'text', text: prompt || 'Procesa la solicitud de cotización de la imagen/PDF adjunto.' });
  } else {
    primerContenido = prompt;
  }
  const messages = [{ role:'user', content: primerContenido }];

  for (let i = 0; i < 5; i++) {
    const payload = {
      model: MODEL,
      max_tokens: 800,
      system: system,
      messages: messages,
      tools: TOOLS
    };
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const status = resp.getResponseCode();
    const text = resp.getContentText();
    if (status >= 400) {
      Logger.log('Claude error ' + status + ': ' + text);
      return 'Error con la IA. Intenta de nuevo.';
    }
    const data = JSON.parse(text);
    if (data.error) return 'Error: ' + (data.error.message || JSON.stringify(data.error));
    if (!data.content) return 'Respuesta vacía.';

    if (data.stop_reason !== 'tool_use') {
      const txt = (data.content || []).filter(c => c.type === 'text').map(c => c.text).join('\n').trim();
      return txt || '(sin respuesta)';
    }

    // Ejecutar tool_use
    const toolUseBlocks = data.content.filter(c => c.type === 'tool_use');
    const toolResults = [];
    for (const block of toolUseBlocks) {
      let result;
      try {
        result = executeTool(block.name, block.input || {}, user);
      } catch (e) {
        result = { ok: false, mensaje: 'Error ejecutando ' + block.name + ': ' + e.message };
      }
      toolResults.push({
        type: 'tool_result',
        tool_use_id: block.id,
        content: JSON.stringify(result)
      });
    }
    messages.push({ role: 'assistant', content: data.content });
    messages.push({ role: 'user', content: toolResults });
  }
  return 'Demasiadas iteraciones. Reformula la pregunta.';
}

// ─── EJECUTOR DE HERRAMIENTAS ─────────────────────────────────────────────
function executeTool(name, args, user) {
  if (name === 'crear_tarea')             return tool_crear_tarea(args, user);
  if (name === 'marcar_etapa_oc')         return tool_marcar_etapa_oc(args, user);
  if (name === 'listar_oc_pendientes')    return tool_listar_oc_pendientes(args);
  if (name === 'listar_tareas_pendientes') return tool_listar_tareas_pendientes(args);
  if (name === 'crear_oc')                return tool_crear_oc(args, user);
  if (name === 'crear_solicitud_cotizacion') return tool_crear_solicitud_cotizacion(args, user);
  return { ok: false, mensaje: 'Herramienta desconocida: ' + name };
}

// Crea una solicitud de cotización en data/solicitudes_cotiz.json (mismo formato que
// el robot de correo email-to-cotiz), para que la tome el equipo de cotización.
function tool_crear_solicitud_cotizacion(args, user) {
  var sols = _ghLoadJSON('data/solicitudes_cotiz.json') || [];
  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'America/Bogota', 'yyyyMMdd');
  var rand = Math.floor(Math.random() * 900 + 100).toString();
  var productos = args.productos || [];
  var descripcion = productos.length
    ? productos.map(function(p){ return (p.cantidad ? p.cantidad + ' — ' : '') + (p.descripcion || ''); }).join(' | ').substring(0, 400)
    : (args.cliente || 'Solicitud');
  var sol = {
    id: 'SOL-' + dateStr + '-' + rand,
    fecha: now.toISOString(),
    cliente: args.cliente || 'Sin identificar',
    contacto: args.contacto || '',
    productos: productos,
    descripcion: descripcion,
    formaPago: args.formaPago || '',
    observaciones: args.observaciones || '',
    urgencia: args.urgencia || 'media',
    correoOrigen: 'WhatsApp: ' + user.name,
    asuntoOrigen: 'Solicitud por WhatsApp',
    estado: 'pendiente',
    cotizacionId: null,
    createdAt: now.toISOString(),
    updatedAt: now.toISOString(),
    createdBy: user.name + ' (vía WhatsApp)'
  };
  sols.push(sol);
  _ghSaveJSON('data/solicitudes_cotiz.json', sols, '📩 Solicitud (WA): ' + sol.cliente);
  return { ok: true, mensaje: 'Solicitud ' + sol.id + ' creada para ' + sol.cliente + '. La verá el equipo de cotización (Alexandra/Lina).' };
}

// Descarga un adjunto de Twilio (requiere auth Basic con SID/TOKEN) y lo convierte
// en un bloque de contenido para la API de Claude (imagen o documento PDF).
function _mediaABloqueClaude(m) {
  try {
    var sid = PROPS.getProperty('TWILIO_SID');
    var tok = PROPS.getProperty('TWILIO_TOKEN');
    var opts = { muteHttpExceptions: true, followRedirects: true };
    if (sid && tok) opts.headers = { Authorization: 'Basic ' + Utilities.base64Encode(sid + ':' + tok) };
    var resp = UrlFetchApp.fetch(m.url, opts);
    if (resp.getResponseCode() >= 300) { Logger.log('Media fetch HTTP ' + resp.getResponseCode()); return null; }
    var b64 = Utilities.base64Encode(resp.getBlob().getBytes());
    var ct = (m.ctype || '').toLowerCase();
    if (ct.indexOf('image/') === 0) {
      return { type:'image', source:{ type:'base64', media_type: ct, data: b64 } };
    }
    if (ct.indexOf('pdf') >= 0) {
      return { type:'document', source:{ type:'base64', media_type:'application/pdf', data: b64 } };
    }
    Logger.log('Tipo de adjunto no soportado: ' + ct);
    return null;
  } catch (e) { Logger.log('Media error: ' + e); return null; }
}

function tool_crear_tarea(args, user) {
  const tasks = _ghLoadJSON('data/tasks.json') || [];
  const newId = (tasks.length ? Math.max.apply(null, tasks.map(function(t){return t.id||0;})) : 0) + 1;
  const now = new Date().toISOString();
  const task = {
    id: newId,
    titulo: args.titulo,
    desc: args.desc || '',
    responsable: args.responsable,
    fecha: args.fecha,
    prioridad: args.prioridad || 'Media',
    estado: 'En Progreso',
    proyecto: '',
    createdAt: now,
    createdBy: user.name + ' (vía WhatsApp)',
    updatedAt: now,
    updatedBy: user.name + ' (vía WhatsApp)'
  };
  tasks.push(task);
  _ghSaveJSON('data/tasks.json', tasks, '✅ Tarea (WA): ' + args.titulo);
  return { ok: true, mensaje: 'Tarea "' + args.titulo + '" creada para ' + args.responsable + ', vence ' + args.fecha + '.' };
}

function tool_marcar_etapa_oc(args, user) {
  const list = _ghLoadJSON('ordenes.json') || [];
  const stageMap = { compra:0, entrega:1, certificado:2, certif:2, factura:3, facturacion:3 };
  const idx = stageMap[(args.etapa || '').toLowerCase()];
  if (idx === undefined) return { ok:false, mensaje:'Etapa inválida. Usa: compra, entrega, certificado o factura.' };
  const target = (args.num || '').toString().toLowerCase();
  const order = list.find(function(o){ return (o.num||'').toString().toLowerCase() === target || o.id === args.num; });
  if (!order) return { ok:false, mensaje:'No encontré la O.C. ' + args.num + '.' };
  const stages = order.stages || [{},{},{},{}];
  while (stages.length < 4) stages.push({});
  const fecha = args.fecha || new Date().toISOString().slice(0,10);
  stages[idx] = Object.assign({}, stages[idx], { s:'done', f:fecha, n: args.nota || stages[idx].n || '' });
  order.stages = stages;
  order.updatedAt = new Date().toISOString();
  order.updatedBy = user.name + ' (vía WhatsApp)';
  if (stages.every(function(s){ return s.s === 'done' || !!s.f; })) order.estado = 'completado';
  _ghSaveJSON('ordenes.json', list, '🚛 Etapa OC (WA)');
  const labels = ['Compra','Entrega','Certificado','Facturación'];
  return { ok:true, mensaje:'Etapa ' + labels[idx] + ' de O.C. ' + args.num + ' marcada el ' + fecha + '.' };
}

function tool_listar_oc_pendientes(args) {
  const list = (_ghLoadJSON('ordenes.json') || []).filter(function(o){ return o.estado === 'activo'; });
  const cliente = args && args.cliente ? args.cliente.toLowerCase() : '';
  const filtered = cliente ? list.filter(function(o){ return (o.cliente||'').toLowerCase().indexOf(cliente) >= 0; }) : list;
  const STAGE_LBL = ['Compra','Entrega','Certif.','Factura'];
  const today = new Date(); today.setHours(0,0,0,0);
  const lines = filtered.slice(0, 10).map(function(o){
    const stages = o.stages || [];
    let pending = '?', estado = '';
    for (let i = 0; i < 4; i++) {
      const s = stages[i] || {};
      const done = i >= 2 ? !!s.f : s.s === 'done';
      if (!done) {
        pending = STAGE_LBL[i];
        if (s.f) {
          const d = new Date(s.f); d.setHours(0,0,0,0);
          const diff = Math.round((d - today) / 86400000);
          estado = diff < 0 ? 'vencida ' + (-diff) + 'd' : (diff === 0 ? 'hoy' : 'en ' + diff + 'd');
        } else {
          estado = 'sin fecha';
        }
        break;
      }
    }
    return '• ' + (o.num || o.id) + ' (' + (o.cliente || '-') + '): ' + pending + ' — ' + estado;
  });
  if (!lines.length) return { ok:true, lista: 'No hay O.C. activas pendientes' + (cliente ? ' para "' + args.cliente + '"' : '') + '.' };
  return { ok:true, total: filtered.length, lista: lines.join('\n') + (filtered.length > 10 ? '\n…y ' + (filtered.length - 10) + ' más' : '') };
}

function tool_listar_tareas_pendientes(args) {
  const tasks = (_ghLoadJSON('data/tasks.json') || []).filter(function(t){ return t.estado !== 'Completado' && t.estado !== 'Cancelado'; });
  const resp = args && args.responsable ? args.responsable.toLowerCase() : '';
  const filtered = resp ? tasks.filter(function(t){ return (t.responsable||'').toLowerCase().indexOf(resp) >= 0; }) : tasks;
  if (!filtered.length) return { ok:true, lista: 'No hay tareas pendientes' + (resp ? ' para "' + args.responsable + '"' : '') + '.' };
  const lines = filtered.slice(0, 10).map(function(t){
    return '• ' + t.titulo + ' (' + (t.responsable||'-') + ', ' + (t.fecha||'sin fecha') + ', ' + (t.prioridad||'') + ')';
  });
  return { ok:true, total: filtered.length, lista: lines.join('\n') + (filtered.length > 10 ? '\n…y ' + (filtered.length - 10) + ' más' : '') };
}

// ─── HELPERS GITHUB ────────────────────────────────────────────────────────
function _ghOwner()  { return PROPS.getProperty('GH_OWNER')  || 'cami902026-oss'; }
function _ghRepo()   { return PROPS.getProperty('GH_REPO')   || 'plataforma-eyg'; }
function _ghBranch() { return PROPS.getProperty('GH_BRANCH') || 'main'; }
function _ghTok()    { return PROPS.getProperty('GH_TOKEN'); }

function _ghLoadJSON(path) {
  // Lectura por raw (sin token, repo público)
  const url = 'https://raw.githubusercontent.com/' + _ghOwner() + '/' + _ghRepo() + '/' + _ghBranch() + '/' + path + '?t=' + Date.now();
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) return JSON.parse(r.getContentText());
  } catch(e) { Logger.log('GH load ' + path + ': ' + e); }
  return null;
}

function _ghSaveJSON(path, data, label) {
  const tok = _ghTok();
  if (!tok) throw new Error('GH_TOKEN no configurado');

  // Get SHA
  let sha = null;
  try {
    const getResp = UrlFetchApp.fetch(
      'https://api.github.com/repos/' + _ghOwner() + '/' + _ghRepo() + '/contents/' + encodeURI(path) + '?ref=' + _ghBranch(),
      { method:'get', headers:{ 'Authorization':'Bearer '+tok, 'Accept':'application/vnd.github+json' }, muteHttpExceptions:true }
    );
    if (getResp.getResponseCode() === 200) sha = JSON.parse(getResp.getContentText()).sha;
  } catch(e) { Logger.log('GH SHA error: ' + e); }

  const payload = {
    message: (label || 'Update ' + path) + ' — ' + new Date().toLocaleString('es-CO'),
    content: Utilities.base64Encode(JSON.stringify(data, null, 2), Utilities.Charset.UTF_8),
    branch: _ghBranch()
  };
  if (sha) payload.sha = sha;

  const r = UrlFetchApp.fetch(
    'https://api.github.com/repos/' + _ghOwner() + '/' + _ghRepo() + '/contents/' + encodeURI(path),
    {
      method: 'put',
      contentType: 'application/json',
      headers: { 'Authorization':'Bearer '+tok, 'Accept':'application/vnd.github+json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );
  if (r.getResponseCode() >= 400) {
    Logger.log('GH save error ' + r.getResponseCode() + ': ' + r.getContentText());
    throw new Error('GitHub PUT ' + r.getResponseCode());
  }
}

// ─── tool_crear_oc ─────────────────────────────────────────────────────────
// Crea OC en ordenes.json. Usa num del cliente como num de OC.
// Stage 0 (Compra) queda con la fecha indicada (default hoy). Resto vacío.
function tool_crear_oc(args, user) {
  if (!args || !args.num || !args.cliente) {
    return { ok:false, mensaje:'Faltan datos. Necesito al menos número y cliente.' };
  }
  const list = _ghLoadJSON('ordenes.json') || [];
  // Evitar duplicados por num
  const numNorm = String(args.num).trim();
  const yaExiste = list.find(function(o){ return String(o.num||'').trim().toLowerCase() === numNorm.toLowerCase(); });
  if (yaExiste) {
    return { ok:false, mensaje:'Ya existe una OC con número ' + numNorm + ' (cliente: ' + (yaExiste.cliente||'-') + ').' };
  }
  const fecha = args.fecha_compra || new Date().toISOString().slice(0,10);
  const now = new Date().toISOString();
  const nuevoId = 'poc_' + Date.now();
  const oc = {
    id: nuevoId,
    num: numNorm,
    cliente: String(args.cliente).trim().toUpperCase(),
    desc: args.desc || '',
    estado: 'activo',
    stages: [
      { s:'done', f: fecha, n: args.nota || '' },
      { s:'pending', f:'', n:'' },
      { s:'pending', f:'', n:'' },
      { s:'pending', f:'', n:'' }
    ],
    createdAt: now,
    createdBy: (user && user.name) ? user.name + ' (vía WhatsApp)' : 'Bot WhatsApp',
    updatedAt: now,
    updatedBy: (user && user.name) ? user.name + ' (vía WhatsApp)' : 'Bot WhatsApp'
  };
  list.push(oc);
  _ghSaveJSON('ordenes.json', list, '🆕 OC ' + numNorm + ' (' + oc.cliente + ') vía bot');
  return { ok:true, mensaje:'OC ' + numNorm + ' creada para ' + oc.cliente + ', compra ' + fecha + '.' };
}
