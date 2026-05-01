// =====================================================
// E&G ENERGY — Google Apps Script para Órdenes de Pedido
// Pegar este código en: script.google.com → Nuevo proyecto
// =====================================================

const SHEET_ID   = '1QjeJiCQ8fND57r0MNw1vRNdcnukqpBNHDfxjbeW7Lq0';
const SHEET_NAME = 'Pedido_Seguimiento';

// Columnas de la hoja (en orden)
const HEADERS = [
  'ID',
  'Número OP',
  'Cliente',
  'Descripción',
  'Estado',
  // Compra
  '🛒 Compra - Estado',
  '🛒 Compra - Fecha',
  '🛒 Compra - Proveedor/Nota',
  // Entrega
  '🚚 Entrega - Estado',
  '🚚 Entrega - Fecha',
  '🚚 Entrega - Tipo Envío',
  '🚚 Entrega - N° Guía',
  '🚚 Entrega - Destino',
  '🚚 Entrega - Responsable',
  // Certificado
  '📋 Certificado - Estado',
  '📋 Certificado - Fecha',
  '📋 Certificado - Nota',
  // Facturación
  '💰 Facturación - Estado',
  '💰 Facturación - Fecha',
  '💰 Facturación - Nota',
  // Extra
  '¿Facturado?',
  'Última Actualización',
  'Etapas (JSON)'
];

const ENVIO_LABELS = {
  'camion_energy': 'Camión E&G',
  'transportadora': 'Transportadora',
  'mensajero':      'Mensajero',
  'cliente_recoge': 'Cliente recoge',
  'otro':           'Otro'
};

// ─── Lectura (GET) ────────────────────────────────────────────────────────────
function doGet(e) {
  const callback = e.parameter.callback || '';
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let sheet   = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = _initSheet(ss);

    const lr  = sheet.getLastRow();
    let ops   = [];
    if (lr > 1) {
      const data = sheet.getRange(2, 1, lr - 1, HEADERS.length).getValues();
      ops = data
        .filter(r => r[0])
        .map(r => {
          // Reconstruir stages desde columnas individuales (col 22 = JSON completo)
          const stagesJSON = r[HEADERS.length - 1];
          const stages = _parseJSON(stagesJSON, _buildStagesFromRow(r));
          return {
            id:        r[0],
            num:       r[1],
            cliente:   r[2],
            desc:      r[3],
            estado:    r[4],
            stages:    stages,
            updatedAt: r[21]   // FIX: r[20] era la columna "¿Facturado?" (etiqueta), no el timestamp
          };
        });
    }

    const json = JSON.stringify({ ok: true, data: ops });
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const json = JSON.stringify({ ok: false, error: String(err) });
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Escritura (POST) ─────────────────────────────────────────────────────────
function doPost(e) {
  try {
    let payload;
    if (e.parameter && e.parameter.ops) {
      payload = {
        action: e.parameter.action || 'saveAll',
        ops:    _parseJSON(e.parameter.ops, [])
      };
    } else {
      payload = JSON.parse(e.postData.contents);
    }

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let sheet   = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = _initSheet(ss);

    // ── Archivar completadas del mes ──────────────────────────────────────────
    if (payload.action === 'archiveCompleted') {
      const ops = payload.ops || [];
      let archSheet = ss.getSheetByName('Historial');
      if (!archSheet) {
        archSheet = ss.insertSheet('Historial');
        archSheet.appendRow([...HEADERS, 'Mes Archivado']);
        archSheet.getRange(1, 1, 1, HEADERS.length + 1)
          .setFontWeight('bold')
          .setBackground('#1a3a5c')
          .setFontColor('#ffffff');
      }
      const mesLabel = payload.mesLabel ||
        new Date().toLocaleDateString('es-CO', { month: 'long', year: 'numeric' });
      ops.forEach(op => {
        archSheet.appendRow([..._opToRow(op), mesLabel]);
      });
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, archived: ops.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Guardar todas las órdenes ─────────────────────────────────────────────
    if (payload.action === 'saveAll') {
      const ops = payload.ops || [];
      const lr  = sheet.getLastRow();
      // FIX: usar clearContent en vez de deleteRows para evitar
      // "No se pueden eliminar todas las filas que no están inmovilizadas"
      if (lr > 1) {
        sheet.getRange(2, 1, lr - 1, HEADERS.length).clearContent();
      }
      // Usar setValues en bloque (mucho más rápido que appendRow uno a uno)
      if (ops.length > 0) {
        const rows = ops.map(_opToRow);
        sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
      }
      _formatSheet(sheet, ops.length);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Convertir una orden a fila de hoja ──────────────────────────────────────
function _opToRow(op) {
  const s = op.stages || [];
  const s0 = s[0] || {};  // Compra
  const s1 = s[1] || {};  // Entrega
  const s2 = s[2] || {};  // Certificado
  const s3 = s[3] || {};  // Facturación

  const estadoLabel = { done: '✅ Completado', active: '🔄 En proceso', pending: '⏳ Pendiente' };

  return [
    op.id    || '',
    op.num   || '',
    op.cliente || '',
    op.desc  || '',
    op.estado || '',
    // Compra
    estadoLabel[s0.s] || '⏳ Pendiente',
    s0.f || '',
    s0.n || '',
    // Entrega
    estadoLabel[s1.s] || '⏳ Pendiente',
    s1.f || '',
    ENVIO_LABELS[s1.envio] || s1.envio || '',
    s1.guia || '',
    s1.dest || '',
    s1.resp || '',
    // Certificado
    estadoLabel[s2.s] || '⏳ Pendiente',
    s2.f || '',
    s2.n || '',
    // Facturación
    estadoLabel[s3.s] || '⏳ Pendiente',
    s3.f || '',
    s3.n || '',
    // Extra
    (s3.s === 'done' || s3.f) ? '✅ Facturado' : '❌ Pendiente',
    op.updatedAt || new Date().toISOString(),
    JSON.stringify(op.stages || [])
  ];
}

// ─── Reconstruir stages desde columnas (fallback) ────────────────────────────
function _buildStagesFromRow(r) {
  const estadoInv = { '✅ Completado': 'done', '🔄 En proceso': 'active', '⏳ Pendiente': 'pending' };
  return [
    { s: estadoInv[r[5]]  || 'pending', f: r[6],  n: r[7]  },
    { s: estadoInv[r[8]]  || 'pending', f: r[9],  envio: r[10], guia: r[11], dest: r[12], resp: r[13] },
    { s: estadoInv[r[14]] || 'pending', f: r[15], n: r[16] },
    { s: estadoInv[r[17]] || 'pending', f: r[18], n: r[19] }
  ];
}

// ─── Inicializar hoja con encabezados ────────────────────────────────────────
function _initSheet(ss) {
  const sheet = ss.insertSheet(SHEET_NAME);
  sheet.appendRow(HEADERS);
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setFontWeight('bold').setBackground('#0f2b5b').setFontColor('#ffffff');
  headerRange.setWrap(true);

  // Anchos de columna
  const widths = [130,120,180,280,90, 120,100,200, 120,100,130,120,160,160, 120,100,200, 120,100,200, 120,160,300];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setRowHeight(1, 50);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  return sheet;
}

// ─── Colorear filas por estado ───────────────────────────────────────────────
function _formatSheet(sheet, numRows) {
  if (numRows < 1) return;

  // Colores de encabezado por sección
  const sectionColors = [
    [1,5,'#0f2b5b'],      // ID → Estado: azul oscuro
    [6,8,'#1a3a1a'],      // Compra: verde oscuro
    [9,14,'#1a2a3a'],     // Entrega: azul acero
    [15,17,'#2a1a3a'],    // Certificado: morado oscuro
    [18,20,'#3a1a1a'],    // Facturación: rojo oscuro
    [21,21,'#1a3a1a'],    // ¿Facturado?: verde oscuro
    [22,23,'#111111']     // Extra: negro
  ];
  sectionColors.forEach(([from, to, color]) => {
    sheet.getRange(1, from, 1, to - from + 1)
      .setBackground(color)
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  });

  // Color de filas por estado
  const estados = sheet.getRange(2, 5, numRows, 1).getValues();
  estados.forEach((row, i) => {
    const color = row[0] === 'activo'     ? '#f0f7f0' :
                  row[0] === 'completado' ? '#e8f0fe' :
                  row[0] === 'cancelado'  ? '#fce8e6' : '#ffffff';
    sheet.getRange(i + 2, 1, 1, HEADERS.length).setBackground(color);
  });
}

function _parseJSON(str, def) {
  try { return JSON.parse(str); } catch (e) { return def; }
}
