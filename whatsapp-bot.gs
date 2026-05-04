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
 * 4. Implementar → Nueva implementación → Aplicación web
 *      - Ejecutar como: Tu cuenta
 *      - Acceso: Cualquier usuario, incluso anónimo
 * 5. Copia la URL /exec
 * 6. Ve a https://console.twilio.com → Messaging → Try it out → Sandbox settings
 *    En "When a message comes in" pega la URL /exec → Save
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
    const from = (params.From || '').replace('whatsapp:','').trim();
    const body = (params.Body || '').trim();

    if (!from) return _twiml('Error: no recibí tu número.');
    if (!body) return _twiml('Hola — escríbeme algo y te ayudo.');

    if (DENIED[from]) {
      return _twiml(`Hola ${DENIED[from]}, este bot no está habilitado para tu rol. Usa la plataforma web: https://cami902026-oss.github.io/plataforma-eyg/Index.html`);
    }
    const user = AUTHORIZED[from];
    if (!user) {
      return _twiml(`Tu número (${from}) no está autorizado para usar el bot ENERGY. Si crees que es un error, contacta a Andrea.`);
    }

    // Llamar a Claude con tools
    const reply = chatWithClaude(body, user, from);
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
  }
];

// ─── LOOP AGÉNTICO CON CLAUDE ─────────────────────────────────────────────
function chatWithClaude(prompt, user, from) {
  const apiKey = PROPS.getProperty('CLAUDE_API_KEY');
  if (!apiKey) return 'Error: CLAUDE_API_KEY no configurada en el script.';

  const today = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd');
  const system = 'Eres ENERGY, asistente vía WhatsApp para E&G Energy Group SAS. Hoy es ' + today + '. ' +
    'Usuario actual: ' + user.name + ' (' + user.role + '). ' +
    'Sé MUY BREVE — máximo 2-3 frases en cada respuesta. Sin listas largas. ' +
    'Cuando el usuario pida hacer algo (crear tarea, marcar O.C., etc.), USA las herramientas. ' +
    'Después de ejecutar, confirma brevemente. Sin emojis excesivos.';

  const messages = [{ role:'user', content: prompt }];

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
  return { ok: false, mensaje: 'Herramienta desconocida: ' + name };
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
