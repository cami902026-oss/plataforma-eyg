/**
 * ===== ENERGY — Claude API Proxy (Google Apps Script) =====
 *
 * Cómo desplegarlo:
 * 1. Ve a https://script.google.com/home → Nuevo proyecto.
 * 2. Pega TODO este archivo en Code.gs.
 * 3. Menú "Configuración del proyecto" (icono ⚙) → Propiedades del script
 *    → Agregar propiedad: CLAUDE_API_KEY = sk-ant-api03-...
 * 4. Implementar → Nueva implementación → tipo "Aplicación web"
 *    - Ejecutar como: Tu cuenta
 *    - Quién tiene acceso: Cualquier usuario (incluso anónimo)
 *    - Implementar
 * 5. Copia la URL que termina en /exec
 * 6. Pégala en la plataforma ENERGY → Configuración → "URL del Proxy ENERGY IA"
 *
 * IMPORTANTE: cada vez que cambies este código debes hacer
 * Implementar → Administrar implementaciones → ✏️ → Versión: Nueva versión.
 */

const MODEL_DEFAULT  = 'claude-haiku-4-5-20251001';
const MAX_TOKENS_CAP = 4096;

function doPost(e) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey) {
      return _json({ error: 'CLAUDE_API_KEY no configurada en Propiedades del script' });
    }

    const body = JSON.parse(e.postData.contents);
    const payload = {
      model:      body.model      || MODEL_DEFAULT,
      max_tokens: Math.min(parseInt(body.max_tokens) || 1024, MAX_TOKENS_CAP),
      system:     body.system     || '',
      messages:   Array.isArray(body.messages) ? body.messages : []
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

  } catch (err) {
    return _json({ error: 'Proxy error: ' + err.message });
  }
}

// Health check rápido en navegador
function doGet() {
  return _json({ ok: true, service: 'ENERGY Claude proxy', model: MODEL_DEFAULT });
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
