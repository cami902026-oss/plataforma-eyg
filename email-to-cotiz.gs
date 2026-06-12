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
      var atts    = msg.getAttachments({ includeInlineImages: true, includeAttachments: true });

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
      // No es una solicitud de cotización → marcar como procesado (no reintentar)
      if (!sol) { newProcessed.push(msgId); continue; }

      // Guardar en GitHub
      var saved = _saveGitHub(sol);
      if (!saved) {
        // Fallo real de guardado (red/SHA): NO marcar como procesado para que el
        // próximo ciclo lo reintente y NO se pierda la solicitud del cliente.
        Logger.log('Error guardando ' + sol.id + ' — se reintentará en el próximo ciclo');
        continue;
      }

      // Éxito → notificar y marcar como procesado
      _enviarEmailSolicitud(sol);
      newProcessed.push(msgId);
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
    'La comercial reenvía solicitudes de clientes que pueden venir como texto, imagen o PDF. ' +
    'Si hay una imagen adjunta, SIEMPRE léela — puede contener una lista de productos escrita a mano, ' +
    'una foto de un pedido, una captura de pantalla de WhatsApp o un listado impreso. ' +
    'Si el cuerpo del correo está vacío pero hay imagen, asume que la imagen contiene la solicitud.\n\n' +
    'REGLA CRÍTICA — TRANSCRIPCIÓN LITERAL: copia los productos, cantidades y descripciones ' +
    'EXACTAMENTE como aparecen en el mensaje original (texto, imagen o PDF). NO traduzcas, NO corrijas ' +
    'ortografía, NO cambies ni conviertas unidades, NO normalices ni reformules, NO abrevies ni completes ' +
    'palabras. Si el cliente escribe en inglés, déjalo en inglés tal cual. Respeta mayúsculas, minúsculas, ' +
    'símbolos, números, códigos y referencias EXACTAMENTE como están escritos.';

  var promptText =
    'Analiza este mensaje. Puede ser un reenvío de un cliente o una solicitud directa de cotización.\n\n' +
    'Asunto: ' + subject + '\n' +
    'De: ' + from + '\n' +
    'Cuerpo:\n' + bodyText.substring(0, 3000) + '\n\n' +
    'Responde SOLO con JSON válido (sin markdown, sin explicaciones):\n' +
    '{\n' +
    '  "esSolicitud": true o false,\n' +
    '  "cliente": "nombre de la empresa o persona que necesita los materiales",\n' +
    '  "contacto": "nombre del contacto del cliente si aparece, sino cadena vacía",\n' +
    '  "productos": [\n' +
    '    { "descripcion": "el producto o material EXACTAMENTE como está escrito en el original, sin traducir ni modificar nada", "cantidad": "la cantidad y unidad EXACTAMENTE como aparece, sin convertir ni cambiar, sino cadena vacía" }\n' +
    '  ],\n' +
    '  "formaPago": "condición de pago si se menciona (contado, crédito 30 días, etc.), sino cadena vacía",\n' +
    '  "urgencia": "alta si lo piden urgente o para hoy/mañana, baja si tienen más de una semana, media en cualquier otro caso",\n' +
    '  "observaciones": "cualquier dato relevante adicional: lugar de entrega, especificaciones técnicas, etc., sino cadena vacía"\n' +
    '}\n\n' +
    'Es solicitud si piden precios, disponibilidad o cotización de materiales o equipos industriales.\n' +
    'NO es solicitud si es: spam, factura recibida, confirmación de pago, newsletter, notificación automática.\n' +
    'Si hay imagen adjunta, léela para extraer los productos listados en ella.\n' +
    'RECUERDA: transcribe los productos, cantidades y descripciones TAL CUAL aparecen, sin traducir ni cambiar nada.';

  // Construir el contenido del mensaje (texto + imágenes si hay)
  var content = [];

  // Agregar imágenes adjuntas para visión de Claude
  var adjuntosAgregados = 0;
  if (attachments && attachments.length > 0) {
    Logger.log('Adjuntos encontrados: ' + attachments.length);
    for (var i = 0; i < attachments.length && adjuntosAgregados < 5; i++) {
      var att  = attachments[i];
      var mime = att.getContentType() || '';
      var nom  = att.getName() || '';
      Logger.log('Adjunto ' + i + ': ' + mime + ' — ' + nom);

      // IMÁGENES → visión de Claude
      if (mime.indexOf('image/') === 0) {
        try {
          var b64 = Utilities.base64Encode(att.copyBlob().getBytes());
          content.push({ type: 'image', source: { type: 'base64', media_type: mime, data: b64 } });
          adjuntosAgregados++;
          Logger.log('Imagen agregada al contexto de Claude');
        } catch(e) { Logger.log('Error leyendo imagen: ' + e.message); }

      // PDF → documento NATIVO de Claude (lee el PDF real, no texto plano)
      } else if (mime === 'application/pdf' || /\.pdf$/i.test(nom)) {
        try {
          var pbytes = att.copyBlob().getBytes();
          if (pbytes.length > 9 * 1024 * 1024) {
            Logger.log('PDF muy grande (>9MB), ignorado');
          } else {
            content.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: Utilities.base64Encode(pbytes) } });
            adjuntosAgregados++;
            Logger.log('PDF agregado como documento');
          }
        } catch(e) { Logger.log('Error leyendo PDF: ' + e.message); }

      // EXCEL / CSV → convertir a texto y agregarlo al prompt
      } else if (/\.(xlsx|xlsm|xls|csv)$/i.test(nom) ||
                 mime.indexOf('spreadsheet') >= 0 || mime.indexOf('excel') >= 0 || mime === 'text/csv') {
        try {
          var tablaTxt = _archivoTablaATexto(att);
          if (tablaTxt && tablaTxt.length > 0) {
            promptText += '\n\n[Datos del archivo "' + nom + '"]:\n' + tablaTxt.substring(0, 5000);
            adjuntosAgregados++;
            Logger.log('Excel/CSV agregado como texto (' + tablaTxt.length + ' chars)');
          }
        } catch(e) { Logger.log('Error leyendo Excel/CSV: ' + e.message + ' (¿falta habilitar el servicio Drive API?)'); }
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
      max_tokens: 2000,
      system: systemPrompt,
      messages: [{ role: 'user', content: content }]
    }),
    muteHttpExceptions: true
  });

  var data = JSON.parse(resp.getContentText());
  if (!data.content || !data.content[0]) {
    Logger.log('Claude no devolvio contenido. Respuesta: ' + resp.getContentText().substring(0, 300));
    return null;
  }

  var text = data.content[0].text.trim().replace(/```json\n?|\n?```/g, '');
  Logger.log('Claude respondio: ' + text.substring(0, 500));
  var parsed;
  try { parsed = JSON.parse(text); } catch(e) {
    Logger.log('Error parseando JSON: ' + e.message);
    return null;
  }
  if (!parsed.esSolicitud) {
    Logger.log('Claude determino que NO es solicitud');
    return null;
  }

  var now     = new Date();
  var dateStr = Utilities.formatDate(now, 'America/Bogota', 'yyyyMMdd');
  var rand    = Math.floor(Math.random() * 900 + 100).toString();

  // Construir descripción resumida a partir de los productos extraídos
  var productos = parsed.productos || [];
  var descripcion = productos.length > 0
    ? productos.map(function(p) {
        return (p.cantidad ? p.cantidad + ' — ' : '') + (p.descripcion || '');
      }).join(' | ').substring(0, 400)
    : subject;

  return {
    id:            'SOL-' + dateStr + '-' + rand,
    fecha:         now.toISOString(),
    cliente:       parsed.cliente       || 'Sin identificar',
    contacto:      parsed.contacto      || '',
    productos:     productos,
    descripcion:   descripcion,
    formaPago:     parsed.formaPago     || '',
    observaciones: parsed.observaciones || '',
    urgencia:      parsed.urgencia      || 'media',
    correoOrigen:  from,
    asuntoOrigen:  subject,
    estado:        'pendiente',
    cotizacionId:  null,
    createdAt:     now.toISOString(),
    updatedAt:     now.toISOString(),
    createdBy:     'sistema'
  };
}

// ─── Convertir Excel/CSV adjunto a texto (SIN depender de la Drive API) ──────
// CSV: lectura directa. Excel moderno (.xlsx/.xlsm) = archivo ZIP con XML
// adentro → se lee con Utilities.unzip + XmlService (servicios estándar,
// SIEMPRE disponibles, no hay que habilitar nada). El .xls antiguo (binario)
// intenta Drive API solo como último recurso.
function _archivoTablaATexto(att) {
  var nom  = att.getName() || '';
  var mime = att.getContentType() || '';

  // CSV → directo
  if (/\.csv$/i.test(nom) || mime === 'text/csv') {
    return att.copyBlob().getDataAsString();
  }

  // Excel moderno (.xlsx / .xlsm) → leer el ZIP interno (sin Drive API)
  if (/\.(xlsx|xlsm)$/i.test(nom) ||
      mime.indexOf('openxmlformats') >= 0 || mime.indexOf('spreadsheetml') >= 0) {
    try {
      var txt = _xlsxATexto(att.copyBlob());
      if (txt) return txt;
    } catch (e) {
      Logger.log('Error leyendo xlsx por ZIP: ' + e.message);
    }
  }

  // .xls antiguo (binario) o fallback → Drive API si estuviera disponible
  try {
    var tmp = Drive.Files.insert(
      { title: 'tmp_cotiz_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
      att.copyBlob(),
      { convert: true }
    );
    var out = [];
    try {
      var sheets = SpreadsheetApp.openById(tmp.id).getSheets();
      for (var s = 0; s < sheets.length; s++) {
        var vals = sheets[s].getDataRange().getValues();
        for (var r = 0; r < vals.length; r++) {
          var fila = vals[r].join('\t').trim();
          if (fila) out.push(fila);
        }
      }
    } finally {
      try { Drive.Files.remove(tmp.id); } catch(_) {}
    }
    return out.join('\n');
  } catch (e) {
    Logger.log('Drive API no disponible para .xls: ' + e.message);
    return '';
  }
}

// ─── Lee un .xlsx/.xlsm (ZIP de XML) y lo vuelve texto, sin servicios avanzados.
function _xlsxATexto(blob) {
  blob.setContentType('application/zip');
  var files = Utilities.unzip(blob);
  var map = {};
  for (var i = 0; i < files.length; i++) { map[files[i].getName()] = files[i]; }

  // 1) Tabla de textos compartidos (sharedStrings.xml)
  var shared = [];
  if (map['xl/sharedStrings.xml']) {
    var ssRoot = XmlService.parse(map['xl/sharedStrings.xml'].getDataAsString()).getRootElement();
    var ssNs   = ssRoot.getNamespace();
    var sis    = ssRoot.getChildren('si', ssNs);
    for (var s = 0; s < sis.length; s++) { shared.push(_xlsxTextoNodo(sis[s], ssNs)); }
  }

  // 2) Recorrer cada hoja (worksheets/sheetN.xml)
  var out = [];
  var hojas = Object.keys(map)
    .filter(function (n) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(n); })
    .sort();
  for (var h = 0; h < hojas.length; h++) {
    var wsRoot = XmlService.parse(map[hojas[h]].getDataAsString()).getRootElement();
    var wns    = wsRoot.getNamespace();
    var data   = wsRoot.getChild('sheetData', wns);
    if (!data) continue;
    var rows = data.getChildren('row', wns);
    for (var r = 0; r < rows.length; r++) {
      var cells = rows[r].getChildren('c', wns);
      var fila = [];
      for (var c = 0; c < cells.length; c++) {
        var cell  = cells[c];
        var tAttr = cell.getAttribute('t');
        var tipo  = tAttr ? tAttr.getValue() : '';
        var valor = '';
        if (tipo === 's') {                       // texto compartido
          var v = cell.getChild('v', wns);
          if (v) { var idx = parseInt(v.getText(), 10); valor = shared[idx] || ''; }
        } else if (tipo === 'inlineStr') {        // texto en línea
          var is = cell.getChild('is', wns);
          if (is) valor = _xlsxTextoNodo(is, wns);
        } else {                                  // número / fecha / fórmula
          var v2 = cell.getChild('v', wns);
          if (v2) valor = v2.getText();
        }
        if (valor !== '') fila.push(valor);
      }
      var filaTxt = fila.join('\t').trim();
      if (filaTxt) out.push(filaTxt);
    }
  }
  return out.join('\n');
}

// Extrae el texto de un nodo <si>/<is> (puede ser <t> directo o varios <r><t>)
function _xlsxTextoNodo(node, ns) {
  var t = node.getChild('t', ns);
  if (t) return t.getText();
  var partes = [];
  var rs = node.getChildren('r', ns);
  for (var i = 0; i < rs.length; i++) {
    var rt = rs[i].getChild('t', ns);
    if (rt) partes.push(rt.getText());
  }
  return partes.join('');
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
