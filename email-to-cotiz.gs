/**
 * ===== ENERGY — Email → Solicitud Cotización (Gmail polling) =====
 *
 * Revisa Gmail cada 1 minuto buscando correos reenviados desde
 * info@eygenergygroup.com. Claude extrae los datos → crea solicitud
 * en GitHub → envía WhatsApp a Alexandra, Alberto y Andrea.
 *
 * REQUISITO PREVIO: Configurar reenvío automático en Outlook de
 * info@eygenergygroup.com → cami902026@gmail.com
 *
 * Cómo activarlo (una sola vez):
 * 1. Pega este código en Apps Script
 * 2. Agrega las propiedades del script (ver abajo)
 * 3. Selecciona la función "crearTrigger" en el menú desplegable
 * 4. Clic en ▶ Ejecutar — esto activa la revisión cada 1 minuto
 *
 * Propiedades del script necesarias:
 *   CLAUDE_API_KEY  = sk-ant-api03-...
 *   GH_TOKEN        = ghp_...
 *   GH_OWNER        = cami902026-oss
 *   GH_REPO         = plataforma-eyg
 *   GH_BRANCH       = main
 *   TWILIO_SID      = (ver config WhatsApp bot)
 *   TWILIO_TOKEN    = (ver config WhatsApp bot)
 *   TWILIO_NUMBER   = whatsapp:+14155238886
 */

var PROPS = PropertiesService.getScriptProperties();
var MODEL = 'claude-haiku-4-5-20251001';
var PROCESSED_KEY = 'cotiz_processed_ids';

var WA_DESTINOS = [
  { nombre: 'Alexandra', numero: '+573144858382' },
  { nombre: 'Alberto',   numero: '+573113134451' },
  { nombre: 'Andrea',    numero: '+573107574110' }
];

// ─── Función principal — corre cada 1 minuto ─────────────────────────────────

function checkEmails() {
  var processed = new Set();
  try {
    var stored = PROPS.getProperty(PROCESSED_KEY);
    if (stored) JSON.parse(stored).forEach(function(id) { processed.add(id); });
  } catch(_) {}

  var threads = GmailApp.search('newer_than:2d in:inbox', 0, 20);
  var newProcessed = [];

  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    for (var j = 0; j < msgs.length; j++) {
      var msg = msgs[j];
      var msgId = msg.getId();
      if (processed.has(msgId)) { continue; }

      var from    = msg.getFrom();
      var subject = msg.getSubject();
      var body    = msg.getPlainBody();

      // Saltar correos internos
      var fromLower = from.toLowerCase();
      if (fromLower.indexOf('eygenergygroup.com') !== -1 ||
          fromLower.indexOf('cami902026@gmail.com') !== -1 ||
          fromLower.indexOf('mailer-daemon') !== -1) {
        newProcessed.push(msgId);
        continue;
      }

      // Extraer datos con Claude
      var sol = _extractData(subject, from, body);
      newProcessed.push(msgId);
      if (!sol) continue;

      // Guardar en GitHub
      var saved = _saveGitHub(sol);
      if (!saved) { Logger.log('Error guardando ' + sol.id); continue; }

      // Enviar WhatsApp a los 3 destinatarios
      _sendWhatsApp(sol);

      Logger.log('✅ Solicitud creada: ' + sol.id + ' — ' + sol.cliente);
    }
  }

  // Guardar IDs procesados (máximo 500 para no llenar Properties)
  var todos = Array.from(processed).concat(newProcessed).slice(-500);
  PROPS.setProperty(PROCESSED_KEY, JSON.stringify(todos));
}

// ─── Crear trigger de 1 minuto (ejecutar una sola vez) ───────────────────────

function crearTrigger() {
  // Eliminar triggers anteriores de checkEmails para no duplicar
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkEmails') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('checkEmails')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('✅ Trigger activado: checkEmails cada 1 minuto');
}

// ─── Extraer datos con Claude ────────────────────────────────────────────────

function _extractData(subject, from, bodyText) {
  var apiKey = PROPS.getProperty('CLAUDE_API_KEY');
  if (!apiKey) throw new Error('CLAUDE_API_KEY no configurada');

  var prompt =
    'Analiza este correo recibido en info@eygenergygroup.com (empresa de suministro industrial).\n\n' +
    'Asunto: ' + subject + '\n' +
    'De: ' + from + '\n' +
    'Cuerpo:\n' + bodyText.substring(0, 2000) + '\n\n' +
    'Responde SOLO con JSON válido:\n' +
    '{\n' +
    '  "esSolicitudCotizacion": true o false,\n' +
    '  "cliente": "nombre de la empresa o persona que solicita",\n' +
    '  "descripcion": "qué materiales o equipos piden, máx 200 chars",\n' +
    '  "urgencia": "alta / media / baja",\n' +
    '  "contacto": "nombre del contacto si aparece, sino cadena vacía"\n' +
    '}\n\n' +
    'Es solicitud si piden precios, disponibilidad o cotización de materiales/equipos industriales.\n' +
    'NO es solicitud si es: spam, factura, confirmación de pago, saludo, newsletter, notificación automática.';

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    payload: JSON.stringify({
      model: MODEL,
      max_tokens: 300,
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  var data = JSON.parse(resp.getContentText());
  if (!data.content || !data.content[0]) return null;

  var text = data.content[0].text.trim().replace(/```json\n?|\n?```/g, '');
  var parsed;
  try { parsed = JSON.parse(text); } catch(_) { return null; }
  if (!parsed.esSolicitudCotizacion) return null;

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

// ─── Guardar en GitHub ───────────────────────────────────────────────────────

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

  var code = r2.getResponseCode();
  return code === 200 || code === 201;
}

// ─── Enviar WhatsApp ─────────────────────────────────────────────────────────

function _sendWhatsApp(sol) {
  var sid    = PROPS.getProperty('TWILIO_SID');
  var tkn    = PROPS.getProperty('TWILIO_TOKEN');
  var fromWA = PROPS.getProperty('TWILIO_NUMBER');
  if (!sid || !tkn || !fromWA) return;

  var urgEmoji = sol.urgencia === 'alta' ? '🔴' : sol.urgencia === 'media' ? '🟡' : '🟢';
  var fechaFmt = Utilities.formatDate(new Date(sol.fecha), 'America/Bogota', 'dd/MM/yyyy HH:mm');
  var auth     = 'Basic ' + Utilities.base64Encode(sid + ':' + tkn);

  var msg =
    '📋 *Nueva solicitud de cotización*\n\n' +
    '👤 *Cliente:* ' + sol.cliente + '\n' +
    '📝 *Solicitud:* ' + sol.descripcion + '\n' +
    urgEmoji + ' *Urgencia:* ' + sol.urgencia + '\n' +
    '🕐 *Recibida:* ' + fechaFmt + '\n' +
    (sol.contacto ? '📞 *Contacto:* ' + sol.contacto + '\n' : '') +
    '\n⏰ Tiempo límite: *12 horas*\n' +
    '🆔 ' + sol.id;

  for (var i = 0; i < WA_DESTINOS.length; i++) {
    try {
      UrlFetchApp.fetch(
        'https://api.twilio.com/2010-04-01/Accounts/' + sid + '/Messages.json',
        {
          method: 'POST',
          headers: { 'Authorization': auth },
          payload: { From: fromWA, To: 'whatsapp:' + WA_DESTINOS[i].numero, Body: msg },
          muteHttpExceptions: true
        }
      );
    } catch(e) { Logger.log('Error WA ' + WA_DESTINOS[i].nombre + ': ' + e.message); }
  }
}

// ─── Helper ──────────────────────────────────────────────────────────────────

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
