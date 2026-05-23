/**
 * ===== ENERGY — Email → Solicitud Cotización (Apps Script independiente) =====
 *
 * Power Automate monitorea info@eygenergygroup.com. Cuando llega un correo nuevo,
 * hace POST a este script con el asunto, remitente y cuerpo.
 * Claude extrae cliente + descripción + urgencia → crea solicitud en GitHub
 * → envía WhatsApp a Alexandra, Alberto y Andrea.
 *
 * Cómo desplegarlo:
 * 1. https://script.google.com/home → Nuevo proyecto → "ENERGY Email to Cotiz"
 * 2. Pega este archivo en Code.gs
 * 3. ⚙ Propiedades del script:
 *      CLAUDE_API_KEY  = sk-ant-api03-...
 *      GH_TOKEN        = ghp_... (PAT con permiso 'repo')
 *      GH_OWNER        = cami902026-oss
 *      GH_REPO         = plataforma-eyg
 *      GH_BRANCH       = main
 *      TWILIO_SID      = (ver config WhatsApp bot)
 *      TWILIO_TOKEN    = (ver config WhatsApp bot)
 *      TWILIO_NUMBER   = whatsapp:+14155238886
 * 4. Implementar → Nueva implementación → Aplicación web
 *      - Ejecutar como: Tu cuenta
 *      - Acceso: Cualquier usuario, incluso anónimo
 * 5. Copiar URL /exec → pegar en Power Automate
 *
 * Power Automate (Office 365):
 * - Trigger: When a new email arrives (V3) → Folder: Inbox (de info@)
 * - Action: HTTP POST a la URL /exec
 *   Body: {
 *     "subject":  "@{triggerOutputs()?['body/subject']}",
 *     "from":     "@{triggerOutputs()?['body/from']}",
 *     "bodyText": "@{triggerOutputs()?['body/bodyPreview']}"
 *   }
 */

var PROPS = PropertiesService.getScriptProperties();
var MODEL = 'claude-haiku-4-5-20251001';

// Destinatarios WhatsApp (siempre reciben la notificación)
var WA_DESTINOS = [
  { nombre: 'Alexandra', numero: '+573144858382' },
  { nombre: 'Alberto',   numero: '+573113134451' },
  { nombre: 'Andrea',    numero: '+573107574110' }
];

// ─── Entry points ───────────────────────────────────────────────────────────

function doGet() {
  return _json({ ok: true, service: 'ENERGY Email→Cotiz', version: '1.0' });
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var subject  = payload.subject  || '';
    var from     = payload.from     || '';
    var bodyText = payload.bodyText || '';

    // Ignorar correos internos (evitar bucles)
    if (from.toLowerCase().indexOf('eygenergygroup.com') !== -1) {
      return _json({ ok: true, skipped: 'correo interno' });
    }

    // Extraer datos con Claude
    var sol = _extractData(subject, from, bodyText);
    if (!sol) {
      return _json({ ok: true, skipped: 'no es solicitud de cotización' });
    }

    // Guardar en GitHub
    var saved = _saveGitHub(sol);
    if (!saved) {
      return _json({ ok: false, error: 'No se pudo guardar en GitHub' });
    }

    // Enviar WhatsApp a los 3 destinatarios
    _sendWhatsApp(sol);

    return _json({ ok: true, id: sol.id, cliente: sol.cliente });

  } catch (err) {
    return _json({ ok: false, error: err.message });
  }
}

// ─── Extraer datos con Claude ────────────────────────────────────────────────

function _extractData(subject, from, bodyText) {
  var apiKey = PROPS.getProperty('CLAUDE_API_KEY');
  if (!apiKey) throw new Error('CLAUDE_API_KEY no configurada');

  var prompt = 'Analiza este correo recibido en info@eygenergygroup.com (empresa de suministro industrial).\n\n' +
    'Asunto: ' + subject + '\n' +
    'De: ' + from + '\n' +
    'Cuerpo:\n' + bodyText.substring(0, 2000) + '\n\n' +
    'Responde SOLO con JSON:\n' +
    '{\n' +
    '  "esSolicitudCotizacion": true o false,\n' +
    '  "cliente": "nombre de la empresa o persona que solicita",\n' +
    '  "descripcion": "qué materiales o equipos piden, máx 200 chars",\n' +
    '  "urgencia": "alta / media / baja",\n' +
    '  "contacto": "nombre del contacto si aparece, sino vacío"\n' +
    '}\n\n' +
    'Es solicitud de cotización si piden precios, disponibilidad o cotización de materiales/equipos industriales.\n' +
    'NO es solicitud si es: spam, factura, confirmación de pago, saludo, newsletter.';

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
  var parsed = JSON.parse(text);
  if (!parsed.esSolicitudCotizacion) return null;

  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'America/Bogota', 'yyyyMMdd');
  var rand = Math.floor(Math.random() * 900 + 100).toString();
  var id = 'SOL-' + dateStr + '-' + rand;

  return {
    id:            id,
    fecha:         now.toISOString(),
    cliente:       parsed.cliente   || 'Sin identificar',
    descripcion:   parsed.descripcion || subject,
    urgencia:      parsed.urgencia  || 'media',
    contacto:      parsed.contacto  || '',
    correoOrigen:  from,
    asuntoOrigen:  subject,
    estado:        'pendiente',
    cotizacionId:  null,
    createdAt:     now.toISOString(),
    updatedAt:     now.toISOString(),
    createdBy:     'sistema'
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
  var headers = {
    'Authorization': 'Bearer ' + token,
    'Accept': 'application/vnd.github+json'
  };

  var existing = [];
  var sha = '';
  try {
    var r = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) {
      var d = JSON.parse(r.getContentText());
      sha = d.sha;
      existing = JSON.parse(
        Utilities.newBlob(
          Utilities.base64Decode(d.content.replace(/\n/g, ''))
        ).getDataAsString()
      );
    }
  } catch (_) {}

  existing.push(sol);

  var content = Utilities.base64Encode(
    JSON.stringify(existing, null, 2),
    Utilities.Charset.UTF_8
  );
  var body = {
    message: 'feat: nueva solicitud cotiz ' + sol.id + ' — ' + sol.cliente,
    content: content,
    branch:  branch
  };
  if (sha) body.sha = sha;

  var r2 = UrlFetchApp.fetch(url, {
    method: 'PUT',
    headers: Object.assign({}, headers, { 'Content-Type': 'application/json' }),
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = r2.getResponseCode();
  return code === 200 || code === 201;
}

// ─── Enviar WhatsApp ─────────────────────────────────────────────────────────

function _sendWhatsApp(sol) {
  var sid    = PROPS.getProperty('TWILIO_SID');
  var token  = PROPS.getProperty('TWILIO_TOKEN');
  var fromWA = PROPS.getProperty('TWILIO_NUMBER');
  if (!sid || !token || !fromWA) return;

  var urgEmoji = sol.urgencia === 'alta' ? '🔴' : sol.urgencia === 'media' ? '🟡' : '🟢';
  var fechaFmt = Utilities.formatDate(
    new Date(sol.fecha), 'America/Bogota', 'dd/MM/yyyy HH:mm'
  );

  var msg =
    '📋 *Nueva solicitud de cotización*\n\n' +
    '👤 *Cliente:* ' + sol.cliente + '\n' +
    '📝 *Solicitud:* ' + sol.descripcion + '\n' +
    urgEmoji + ' *Urgencia:* ' + sol.urgencia + '\n' +
    '🕐 *Recibida:* ' + fechaFmt + '\n' +
    (sol.contacto ? '📞 *Contacto:* ' + sol.contacto + '\n' : '') +
    '\n⏰ Tiempo límite de respuesta: *12 horas*\n' +
    '🆔 ' + sol.id;

  var auth = 'Basic ' + Utilities.base64Encode(sid + ':' + token);

  for (var i = 0; i < WA_DESTINOS.length; i++) {
    var dest = WA_DESTINOS[i];
    try {
      UrlFetchApp.fetch(
        'https://api.twilio.com/2010-04-01/Accounts/' + sid + '/Messages.json',
        {
          method: 'POST',
          headers: { 'Authorization': auth },
          payload: {
            From: fromWA,
            To:   'whatsapp:' + dest.numero,
            Body: msg
          },
          muteHttpExceptions: true
        }
      );
    } catch (e) {
      Logger.log('Error WA ' + dest.nombre + ': ' + e.message);
    }
  }
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
