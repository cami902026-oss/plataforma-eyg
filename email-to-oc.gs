/**
 * ===== ENERGY — Email → OC Auto-creator (Apps Script independiente) =====
 *
 * Recibe POST de Power Automate cuando llega correo con asunto
 * "ORDEN DE COMPRA ENERGY" a andrea.bernal@eygenergygroup.com.
 * Lee el cuerpo + adjuntos (PDF/imagen) con Claude Vision, extrae los
 * datos de la OC y la crea en ordenes.json en GitHub + Google Sheets.
 *
 * NO depende del bot WhatsApp ni de Twilio. Es un proyecto Apps Script
 * SEPARADO (despliega su propio /exec).
 *
 * Cómo desplegarlo (paso a paso):
 * 1. https://script.google.com/home → Nuevo proyecto → nombrar "ENERGY Email to OC"
 * 2. Pega TODO este archivo en Code.gs
 * 3. ⚙ Configuración → Propiedades del script:
 *      CLAUDE_API_KEY  = sk-ant-api03-...
 *      GH_TOKEN        = ghp_... (PAT con permiso 'repo')
 *      GH_OWNER        = cami902026-oss
 *      GH_REPO         = plataforma-eyg
 *      GH_BRANCH       = main
 *      GS_SCRIPT_URL   = https://script.google.com/macros/s/AKfy.../exec
 *                        (URL del Apps Script de Google Sheets, mismo que usa Index.html)
 * 4. Implementar → Nueva implementación → Aplicación web
 *      - Ejecutar como: Tu cuenta
 *      - Acceso: Cualquier usuario, incluso anónimo
 * 5. Copia la URL /exec y pégala en Power Automate (paso siguiente)
 *
 * Power Automate (Office 365):
 * - Trigger: When a new email arrives (V3)
 *     - Folder: Inbox
 *     - Subject filter: "ORDEN DE COMPRA ENERGY"
 *     - Include Attachments: Yes
 * - Action: HTTP
 *     - Method: POST
 *     - URI: <URL del /exec>
 *     - Headers: Content-Type = application/json
 *     - Body:
 *       {
 *         "subject":  @{triggerOutputs()?['body/subject']},
 *         "from":     @{triggerOutputs()?['body/from']},
 *         "bodyText": @{triggerOutputs()?['body/bodyPreview']},
 *         "attachments": @{triggerOutputs()?['body/attachments']}
 *       }
 *
 * Lo que devuelve este script:
 *   { ok:true, num:"LM1500", cliente:"PETRORIOS" }   — éxito
 *   { ok:false, error:"..." }                         — fallo
 */

const PROPS = PropertiesService.getScriptProperties();
const MODEL = 'claude-sonnet-4-6';   // Sonnet — mejor para extracción visual

function doGet() {
  return _json({
    ok: true,
    service: 'ENERGY Email→OC',
    propsConfigured: ['CLAUDE_API_KEY','GH_TOKEN','GH_OWNER','GH_REPO','GS_SCRIPT_URL']
                     .filter(function(k){ return !!PROPS.getProperty(k); })
  });
}

// ─── POLLING DE GMAIL (corre cada 1 min con trigger de tiempo) ─────────────
// Lee correos no leídos cuyo asunto contenga "ORDEN DE COMPRA ENERGY",
// extrae datos con Claude y crea la OC. Marca el correo como leído + label
// "ENERGY-OC-Procesado" para no procesarlo dos veces.
function procesarCorreosNuevos() {
  var query = 'subject:"ORDEN DE COMPRA ENERGY" is:unread newer_than:7d';
  var threads = GmailApp.search(query, 0, 20);
  if (!threads.length) {
    Logger.log('No hay correos nuevos.');
    return;
  }
  Logger.log('Hilos encontrados: ' + threads.length);

  // Asegurar label de procesados
  var labelName = 'ENERGY-OC-Procesado';
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);

  for (var t = 0; t < threads.length; t++) {
    var thread = threads[t];
    var msgs = thread.getMessages();
    for (var m = 0; m < msgs.length; m++) {
      var msg = msgs[m];
      if (!msg.isUnread()) continue;
      // Verificar de nuevo el asunto (search es laxo)
      var subj = msg.getSubject() || '';
      if (subj.toUpperCase().indexOf('ORDEN DE COMPRA ENERGY') < 0) continue;

      try {
        // Construir el body que normalmente arma Power Automate
        var attsRaw = msg.getAttachments() || [];
        var attachments = [];
        for (var a = 0; a < attsRaw.length && a < 5; a++) {
          var att = attsRaw[a];
          var ct = att.getContentType() || '';
          var name = att.getName() || '';
          // Solo PDFs e imágenes
          if (ct.indexOf('pdf') < 0 && ct.indexOf('image') < 0 &&
              !/\.(pdf|jpe?g|png|gif|webp)$/i.test(name)) continue;
          var bytes = att.getBytes();
          attachments.push({
            name: name,
            contentType: ct,
            contentBytes: Utilities.base64Encode(bytes)
          });
        }
        var body = {
          subject: subj,
          from: msg.getFrom() || '',
          bodyText: (msg.getPlainBody() || '').slice(0, 5000),
          attachments: attachments
        };

        // Reusar el mismo pipeline del doPost
        var fakeEvent = { postData: { contents: JSON.stringify(body) } };
        var resp = doPost(fakeEvent);
        var content = resp.getContent ? resp.getContent() : '';
        Logger.log('Procesado msg "' + subj + '" → ' + content);

        // Marcar como leído + label (se procese OK o no, para no reintentar infinito)
        msg.markRead();
        thread.addLabel(label);
      } catch (err) {
        Logger.log('ERROR procesando msg: ' + err);
        // No marcamos como leído para que reintente la próxima vez
      }
    }
  }
}

// ─── INSTALADOR DEL TRIGGER (correrlo UNA SOLA VEZ a mano) ─────────────────
// Crea el trigger de tiempo que dispara procesarCorreosNuevos() cada 1 min.
// En el editor: dropdown de funciones → "instalarTriggerCorreo" → ▶ Ejecutar.
function instalarTriggerCorreo() {
  // Borrar triggers viejos del mismo handler
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'procesarCorreosNuevos') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Crear trigger nuevo cada 1 min
  ScriptApp.newTrigger('procesarCorreosNuevos')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('✅ Trigger instalado: procesarCorreosNuevos cada 1 min');
}

function doPost(e) {
  try {
    var body = {};
    try { body = JSON.parse(e.postData.contents || '{}'); } catch(_) {}
    var subject     = body.subject     || '';
    var fromAddr    = body.from        || '';
    var bodyText    = body.bodyText    || body.body || '';
    var attachments = body.attachments || [];

    Logger.log('Email recibido — subject: ' + subject + ' | from: ' + fromAddr + ' | adjuntos: ' + attachments.length);

    // 1. Pedir a Claude que extraiga los datos
    var datos = _extraerDatosOC(bodyText, attachments, subject, fromAddr);
    if (!datos || !datos.num) {
      return _json({ ok:false, error:'No se pudo extraer número de OC del correo', subject: subject });
    }

    // 2. Cargar ordenes.json actual
    var list = _ghLoadJSON('ordenes.json') || [];
    var numNorm = String(datos.num).trim();

    // 3. Evitar duplicados
    var existente = list.find(function(o){ return String(o.num||'').trim().toLowerCase() === numNorm.toLowerCase(); });
    if (existente) {
      Logger.log('OC ya existe: ' + numNorm);
      return _json({ ok:false, error:'OC ya existe', num: numNorm, cliente: existente.cliente });
    }

    // 4. Crear OC
    var fecha = datos.fecha || new Date().toISOString().slice(0,10);
    var now = new Date().toISOString();
    var nuevaOC = {
      id: 'poc_' + Date.now(),
      num: numNorm,
      cliente: String(datos.cliente || '').trim().toUpperCase() || 'POR DEFINIR',
      desc: datos.desc || '',
      estado: 'activo',
      stages: [
        { s:'done', f: fecha, n: 'Auto-creada desde correo' },
        { s:'pending', f:'', n:'' },
        { s:'pending', f:'', n:'' },
        { s:'pending', f:'', n:'' }
      ],
      createdAt: now,
      createdBy: 'Auto (correo de ' + fromAddr + ')',
      updatedAt: now,
      updatedBy: 'Auto (correo)',
      origenCorreo: { from: fromAddr, subject: subject, fecha: now }
    };
    list.push(nuevaOC);

    // 5. Guardar en GitHub
    _ghSaveJSON('ordenes.json', list, '📨 OC auto ' + numNorm + ' (' + nuevaOC.cliente + ')');

    // 6. Sincronizar a Google Sheets (para que la app la vea sin tener que recargar 2 veces)
    _saveToGS(list);

    Logger.log('OC creada: ' + numNorm + ' / ' + nuevaOC.cliente);
    return _json({ ok:true, num: numNorm, cliente: nuevaOC.cliente, desc: nuevaOC.desc });

  } catch (err) {
    Logger.log('doPost error: ' + err);
    return _json({ ok:false, error: String(err) });
  }
}

// ─── EXTRACTOR CON CLAUDE VISION ──────────────────────────────────────────
function _extraerDatosOC(bodyText, attachments, subject, fromAddr) {
  var apiKey = PROPS.getProperty('CLAUDE_API_KEY');
  if (!apiKey) throw new Error('CLAUDE_API_KEY no configurada');

  // Construir bloques de contenido para Claude (texto + imágenes/PDFs)
  var content = [];
  var instrucciones =
    'De este correo y/o adjuntos te están enviando una ORDEN DE COMPRA. ' +
    'Debes extraer los siguientes campos y devolverlos en JSON puro (sin texto extra, sin markdown):\n' +
    '{\n' +
    '  "num":     "número de la OC tal como aparece (LM1500, ODC-4170, 4600007246, etc.) — REQUERIDO",\n' +
    '  "cliente": "nombre del cliente / razón social en MAYÚSCULAS",\n' +
    '  "desc":    "descripción corta de qué se está comprando (máximo 80 caracteres)",\n' +
    '  "fecha":   "fecha de la OC en formato YYYY-MM-DD (si no se ve, usa la de hoy)"\n' +
    '}\n' +
    'Si no encuentras el número de la OC, devuelve {"num": null}.\n' +
    'Asunto del correo: ' + subject + '\n' +
    'Remitente: ' + fromAddr + '\n' +
    'Cuerpo del correo (puede estar vacío):\n' + (bodyText || '(sin texto)');
  content.push({ type:'text', text: instrucciones });

  // Adjuntar imágenes y PDFs (Claude soporta image y document)
  for (var i=0; i<attachments.length && i<5; i++) {
    var a = attachments[i];
    var name = (a.name || a.fileName || '').toLowerCase();
    var ct   = (a.contentType || a.contentTypeId || '').toLowerCase();
    var b64  = a.contentBytes || a.content || '';
    if (!b64) continue;

    if (/\.pdf$/.test(name) || ct.indexOf('pdf') >= 0) {
      content.push({
        type: 'document',
        source: { type:'base64', media_type:'application/pdf', data: b64 }
      });
    } else if (/\.(jpe?g|png|gif|webp)$/.test(name) || ct.indexOf('image') >= 0) {
      var media = ct && ct.indexOf('image/') === 0 ? ct
                : /\.png$/.test(name) ? 'image/png'
                : /\.gif$/.test(name) ? 'image/gif'
                : /\.webp$/.test(name) ? 'image/webp'
                : 'image/jpeg';
      content.push({
        type: 'image',
        source: { type:'base64', media_type: media, data: b64 }
      });
    }
  }

  var payload = {
    model: MODEL,
    max_tokens: 1000,
    messages: [{ role:'user', content: content }]
  };

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var status = resp.getResponseCode();
  var text = resp.getContentText();
  if (status >= 400) {
    Logger.log('Claude error ' + status + ': ' + text);
    throw new Error('Claude API ' + status);
  }
  var data = JSON.parse(text);
  var raw = (data.content || []).filter(function(c){ return c.type==='text'; })
                                .map(function(c){ return c.text; }).join('').trim();
  Logger.log('Claude respondió: ' + raw);

  // Parsear JSON (a veces viene envuelto en ```json ... ```)
  var jsonStr = raw.replace(/^```json\s*/i,'').replace(/```\s*$/,'').trim();
  // Buscar el primer { ... } válido
  var mb = jsonStr.indexOf('{');
  var me = jsonStr.lastIndexOf('}');
  if (mb >= 0 && me > mb) jsonStr = jsonStr.substring(mb, me+1);

  try {
    return JSON.parse(jsonStr);
  } catch(e) {
    Logger.log('No se pudo parsear JSON: ' + jsonStr);
    return null;
  }
}

// ─── HELPERS GITHUB ────────────────────────────────────────────────────────
function _ghOwner()  { return PROPS.getProperty('GH_OWNER')  || 'cami902026-oss'; }
function _ghRepo()   { return PROPS.getProperty('GH_REPO')   || 'plataforma-eyg'; }
function _ghBranch() { return PROPS.getProperty('GH_BRANCH') || 'main'; }
function _ghTok()    { return PROPS.getProperty('GH_TOKEN'); }

function _ghLoadJSON(path) {
  var url = 'https://raw.githubusercontent.com/' + _ghOwner() + '/' + _ghRepo() + '/' + _ghBranch() + '/' + path + '?t=' + Date.now();
  try {
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) return JSON.parse(r.getContentText());
  } catch(e) { Logger.log('GH load ' + path + ': ' + e); }
  return null;
}

function _ghSaveJSON(path, data, label) {
  var tok = _ghTok();
  if (!tok) throw new Error('GH_TOKEN no configurado');
  var sha = null;
  try {
    var getResp = UrlFetchApp.fetch(
      'https://api.github.com/repos/' + _ghOwner() + '/' + _ghRepo() + '/contents/' + encodeURI(path) + '?ref=' + _ghBranch(),
      { method:'get', headers:{ 'Authorization':'Bearer '+tok, 'Accept':'application/vnd.github+json' }, muteHttpExceptions:true }
    );
    if (getResp.getResponseCode() === 200) sha = JSON.parse(getResp.getContentText()).sha;
  } catch(e) { Logger.log('GH SHA error: ' + e); }

  var payload = {
    message: (label || 'Update ' + path) + ' — ' + new Date().toLocaleString('es-CO'),
    content: Utilities.base64Encode(JSON.stringify(data, null, 2), Utilities.Charset.UTF_8),
    branch: _ghBranch()
  };
  if (sha) payload.sha = sha;

  var r = UrlFetchApp.fetch(
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

// ─── HELPER GOOGLE SHEETS ──────────────────────────────────────────────────
function _saveToGS(list) {
  var url = PROPS.getProperty('GS_SCRIPT_URL');
  if (!url) { Logger.log('GS_SCRIPT_URL no configurado, skip'); return; }
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: 'action=saveAll&ops=' + encodeURIComponent(JSON.stringify(list)),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log('GS sync error: ' + e); }
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
