/**
 * ===== ENERGY — Solicitudes de Cotización (Gmail polling) =====
 *
 * Sheila envía a cotizacionenergy@gmail.com con texto, imagen o PDF.
 * Claude extrae los datos → crea solicitud en GitHub → aparece en la
 * plataforma para que Alexandra haga la cotización.
 * Cuando la cotización se marca "Enviada" → notifica a gerencia.
 *
 * CONFIGURACIÓN (una sola vez):
 * 1. Crear cuenta Gmail: cotizacionenergy@gmail.com
 * 2. Pegar este código en un nuevo proyecto Apps Script
 *    creado con esa cuenta (cotizacionenergy@gmail.com)
 * 3. Agregar propiedades del script (ver abajo)
 * 4. Implementar como Web App (acceso: Cualquier usuario)
 * 5. Ejecutar crearTrigger() una sola vez para activar el polling
 *
 * Propiedades del script:
 *   CLAUDE_API_KEY  = sk-ant-api03-...
 *   GH_TOKEN        = ghp_...
 *   GH_OWNER        = cami902026-oss
 *   GH_REPO         = plataforma-eyg
 *   GH_BRANCH       = main
 *   NOTIF_EMAIL     = gerenciageneral@eygenergygroup.com
 */

var PROPS        = PropertiesService.getScriptProperties();
var MODEL        = 'claude-haiku-4-5-20251001';
var PROCESSED_KEY = 'cotiz_processed_ids';
var NOTIF_CC     = 'cotizacionenergy@gmail.com'; // copia en notificaciones

// ─── Trigger de 1 minuto (ejecutar una sola vez) ─────────────────────────────

function crearTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkEmails') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('checkEmails').timeBased().everyMinutes(1).create();
  Logger.log('Trigger activado: checkEmails cada 1 minuto');
}

// ─── Función principal — revisa la bandeja cada 1 minuto ─────────────────────

function checkEmails() {
  var processed = new Set();
  try {
    var stored = PROPS.getProperty(PROCESSED_KEY);
    if (stored) JSON.parse(stored).forEach(function(id) { processed.add(id); });
  } catch(_) {}

  // Busca todos los correos no procesados de los últimos 7 días
  var threads = GmailApp.search('newer_than:7d in:inbox -from:me', 0, 20);
  var newProcessed = [];

  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    for (var j = 0; j < msgs.length; j++) {
      var msg = msgs[j];
      var msgId = msg.getId();
      if (processed.has(msgId)) { newProcessed.push(msgId); continue; }

      var from    = msg.getFrom();
      var subject = msg.getSubject();
      var body    = msg.getPlainBody();
      var atts    = msg.getAttachments();

      // Saltar mailer-daemon y notificaciones automáticas
      var fromLower = from.toLowerCase();
      if (fromLower.indexOf('mailer-daemon') !== -1 ||
          fromLower.indexOf('noreply') !== -1 ||
          fromLower.indexOf('no-reply') !== -1) {
        newProcessed.push(msgId);
        continue;
      }

      // Extraer datos con Claude (texto + imágenes adjuntas)
      var sol = _extractData(subject, from, body, atts);
      newProcessed.push(msgId);
      if (!sol) continue;

      // Guardar en GitHub
      var saved = _saveGitHub(sol);
      if (!saved) { Logger.log('Error guardando ' + sol.id); continue; }

      // Notificar a Andrea y Gerencia
      _enviarEmailSolicitud(sol);

      Logger.log('Solicitud creada: ' + sol.id + ' — ' + sol.cliente);
    }
  }

  var todos = Array.from(processed).concat(newProcessed).slice(-500);
  PROPS.setProperty(PROCESSED_KEY, JSON.stringify(todos));
}

// ─── Extraer datos con Claude (texto + visión para imágenes) ─────────────────

function _extractData(subject, from, bodyText, attachments) {
  var apiKey = PROPS.getProperty('CLAUDE_API_KEY');
  if (!apiKey) throw new Error('CLAUDE_API_KEY no configurada');

  var systemPrompt =
    'Eres el asistente de E&G Energy Group SAS (empresa de suministro industrial). ' +
    'La comercial reenvía solicitudes de clientes: pueden venir como texto, imagen o PDF adjunto. ' +
    'Tu tarea es extraer los datos de la solicitud para crear un registro de seguimiento.';

  var promptText =
    'Analiza este mensaje de la comercial. Puede ser un reenvío de un cliente o una solicitud directa.\n\n' +
    'Asunto: ' + subject + '\n' +
    'De: ' + from + '\n' +
    'Cuerpo:\n' + bodyText.substring(0, 3000) + '\n\n' +
    'Responde SOLO con JSON válido (sin markdown):\n' +
    '{\n' +
    '  "esSolicitud": true o false,\n' +
    '  "cliente": "nombre de la empresa o persona que necesita los materiales",\n' +
    '  "descripcion": "qué materiales o equipos piden, máximo 300 caracteres",\n' +
    '  "urgencia": "alta / media / baja",\n' +
    '  "contacto": "nombre del contacto del cliente si aparece, sino cadena vacía"\n' +
    '}\n\n' +
    'Es solicitud si piden precios, disponibilidad o cotización de materiales/equipos industriales.\n' +
    'NO es solicitud si es: spam, factura recibida, confirmación de pago, newsletter, alerta automática.';

  // Construir el contenido del mensaje (texto + imágenes si hay)
  var content = [];

  // Agregar imágenes adjuntas para visión de Claude
  if (attachments && attachments.length > 0) {
    for (var i = 0; i < attachments.length && i < 3; i++) {
      var att = attachments[i];
      var mime = att.getContentType() || '';
      if (mime.indexOf('image/') === 0) {
        try {
          var b64 = Utilities.base64Encode(att.copyBlob().getBytes());
          content.push({
            type: 'image',
            source: { type: 'base64', media_type: mime, data: b64 }
          });
        } catch(e) { Logger.log('Error leyendo imagen: ' + e.message); }
      } else if (mime === 'application/pdf') {
        // Para PDFs, intentar leer como texto (limitado en Apps Script)
        try {
          var pdfText = att.copyBlob().getDataAsString();
          if (pdfText && pdfText.length > 50) {
            promptText += '\n\n[Texto extraído del PDF adjunto]:\n' + pdfText.substring(0, 2000);
          }
        } catch(e) { Logger.log('PDF no legible como texto'); }
      }
    }
  }

  // El prompt de texto siempre al final
  content.push({ type: 'text', text: promptText });

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    payload: JSON.stringify({
      model: MODEL,
      max_tokens: 400,
      system: systemPrompt,
      messages: [{ role: 'user', content: content }]
    }),
    muteHttpExceptions: true
  });

  var data = JSON.parse(resp.getContentText());
  if (!data.content || !data.content[0]) return null;

  var text = data.content[0].text.trim().replace(/```json\n?|\n?```/g, '');
  var parsed;
  try { parsed = JSON.parse(text); } catch(_) { return null; }
  if (!parsed.esSolicitud) return null;

  var now     = new Date();
  var dateStr = Utilities.formatDate(now, 'America/Bogota', 'yyyyMMdd');
  var rand    = Math.floor(Math.random() * 900 + 100).toString();

  return {
    id:           'SOL-' + dateStr + '-' + rand,
    fecha:        now.toISOString(),
    cliente:      parsed.cliente      || 'Sin identificar',
    descripcion:  parsed.descripcion  || subject,
    urgencia:     parsed.urgencia     || 'media',
    contacto:     parsed.contacto     || '',
    correoOrigen: from,
    asuntoOrigen: subject,
    estado:       'pendiente',
    cotizacionId: null,
    createdAt:    now.toISOString(),
    updatedAt:    now.toISOString(),
    createdBy:    'sistema'
  };
}

// ─── Guardar en GitHub ────────────────────────────────────────────────────────

function _saveGitHub(sol) {
  var token  = PROPS.getProperty('GH_TOKEN');
  var owner  = PROPS.getProperty('GH_OWNER');
  var repo   = PROPS.getProperty('GH_REPO');
  var branch = PROPS.getProperty('GH_BRANCH') || 'main';
  var file   = 'data/solicitudes_cotiz.json';
  var url    = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + file;
  var hdrs   = { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json' };

  var existing = [];
  var sha = '';
  try {
    var r = UrlFetchApp.fetch(url, { headers: hdrs, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) {
      var d = JSON.parse(r.getContentText());
      sha = d.sha;
      existing = JSON.parse(
        Utilities.newBlob(Utilities.base64Decode(d.content.replace(/\n/g,''))).getDataAsString()
      );
    }
  } catch(_) {}

  existing.push(sol);

  var body = {
    message: 'feat: nueva solicitud ' + sol.id + ' — ' + sol.cliente,
    content: Utilities.base64Encode(JSON.stringify(existing, null, 2), Utilities.Charset.UTF_8),
    branch:  branch
  };
  if (sha) body.sha = sha;

  var r2 = UrlFetchApp.fetch(url, {
    method: 'PUT',
    headers: Object.assign({}, hdrs, { 'Content-Type': 'application/json' }),
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  return r2.getResponseCode() === 200 || r2.getResponseCode() === 201;
}

// ─── Web App — notificaciones desde la plataforma ────────────────────────────

// Destinatarios de todas las notificaciones
var EMAIL_NOTIF = [
  'andrea.bernal@eygenergygroup.com',
  'gerenciageneral@eygenergygroup.com'
];

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action || '';
    if (action === 'notificar_cotizacion') {
      _enviarEmailCotizacion(params.cotizacion, params.tipo || 'creada');
      return _json({ ok: true });
    }
    return _json({ ok: false, error: 'Accion desconocida' });
  } catch(ex) {
    return _json({ ok: false, error: ex.message });
  }
}

function doGet(e) {
  return _json({ ok: true, msg: 'ENERGY Cotiz Script activo' });
}

function _enviarEmailSolicitud(sol) {
  var urgEmoji = sol.urgencia === 'alta' ? '[ALTA]' : sol.urgencia === 'media' ? '[MEDIA]' : '[BAJA]';
  var asunto = 'Nueva solicitud de cotizacion ' + urgEmoji + ': ' + sol.cliente + ' [' + sol.id + ']';
  var cuerpo =
    'Llego una nueva solicitud de cotizacion.\n\n' +
    'ID: ' + sol.id + '\n' +
    'Cliente: ' + sol.cliente + '\n' +
    'Solicitan: ' + (sol.descripcion || '') + '\n' +
    'Urgencia: ' + (sol.urgencia || 'media') + '\n' +
    (sol.contacto ? 'Contacto: ' + sol.contacto + '\n' : '') +
    (sol.correoOrigen ? 'Correo del remitente: ' + sol.correoOrigen + '\n' : '') +
    '\nVer en la plataforma:\n' +
    'https://cami902026-oss.github.io/plataforma-eyg/Index.html\n' +
    '(Cotizaciones → Solicitudes)';
  for (var i = 0; i < EMAIL_NOTIF.length; i++) {
    try { GmailApp.sendEmail(EMAIL_NOTIF[i], asunto, cuerpo); } catch(e2) {
      Logger.log('Error enviando a ' + EMAIL_NOTIF[i] + ': ' + e2.message);
    }
  }
}

function _enviarEmailCotizacion(cot, tipo) {
  var total = cot.total ? '$' + Number(cot.total).toLocaleString('es-CO') : '';
  var esEnviada = tipo === 'enviada';
  var asunto = esEnviada
    ? 'Cotizacion ENVIADA al cliente: ' + cot.cliente + ' [' + cot.id + ']'
    : 'Nueva cotizacion creada: ' + cot.cliente + ' [' + cot.id + ']';
  var cuerpo =
    (esEnviada
      ? 'La asistente envio una cotizacion al cliente.\n\n'
      : 'Se creo una nueva cotizacion en la plataforma.\n\n') +
    'No.: ' + cot.id + '\n' +
    'Cliente: ' + cot.cliente + '\n' +
    (total ? 'Total: ' + total + '\n' : '') +
    (cot.vendedor ? 'Vendedor: ' + cot.vendedor + '\n' : '') +
    (cot.realizadaPor ? 'Realizada por: ' + cot.realizadaPor + '\n' : '') +
    '\nVer en la plataforma:\n' +
    'https://cami902026-oss.github.io/plataforma-eyg/Index.html\n' +
    '(Cotizaciones → Base de Datos)';
  for (var i = 0; i < EMAIL_NOTIF.length; i++) {
    try { GmailApp.sendEmail(EMAIL_NOTIF[i], asunto, cuerpo); } catch(e2) {
      Logger.log('Error enviando a ' + EMAIL_NOTIF[i] + ': ' + e2.message);
    }
  }
}

// ─── Helper ───────────────────────────────────────────────────────────────────

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
