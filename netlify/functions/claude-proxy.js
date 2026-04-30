// ===== ENERGY — Claude API Proxy =====
// Recibe peticiones del frontend (Index.html) y las reenvía a Anthropic
// usando la API Key almacenada en variable de entorno de Netlify.
// La clave NUNCA viaja al navegador.
//
// Configuración requerida en Netlify → Site settings → Environment variables:
//   CLAUDE_API_KEY = sk-ant-...
//   ENERGY_ALLOWED_ORIGINS = https://cami902026-oss.github.io,https://eygenergygroup.com  (opcional, separar con coma)

const CLAUDE_API_KEY = process.env.CLAUDE_API_KEY;
const ALLOWED_ORIGINS = (process.env.ENERGY_ALLOWED_ORIGINS || '').split(',').map(s => s.trim()).filter(Boolean);
const MODEL_DEFAULT  = 'claude-haiku-4-5-20251001';
const MAX_TOKENS_CAP = 4096;

function corsHeaders(origin) {
  // Si hay lista blanca, validar; si no, permitir cualquier origen (modo desarrollo)
  const allow = ALLOWED_ORIGINS.length === 0 || ALLOWED_ORIGINS.includes(origin)
    ? (origin || '*')
    : ALLOWED_ORIGINS[0];
  return {
    'Access-Control-Allow-Origin':  allow,
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Vary': 'Origin',
  };
}

exports.handler = async (event) => {
  const origin = event.headers.origin || event.headers.Origin || '';
  const cors = corsHeaders(origin);

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: cors, body: '' };
  }
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers: cors, body: 'Method not allowed' };
  }
  if (!CLAUDE_API_KEY) {
    return {
      statusCode: 500,
      headers: { ...cors, 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: 'CLAUDE_API_KEY no configurada en Netlify' })
    };
  }

  try {
    const body = JSON.parse(event.body || '{}');
    // Whitelist de campos para evitar que el cliente abuse de la API
    const payload = {
      model:      body.model      || MODEL_DEFAULT,
      max_tokens: Math.min(parseInt(body.max_tokens) || 1024, MAX_TOKENS_CAP),
      system:     body.system     || '',
      messages:   Array.isArray(body.messages) ? body.messages : [],
    };
    if (body.tools)        payload.tools        = body.tools;
    if (body.tool_choice)  payload.tool_choice  = body.tool_choice;
    if (body.temperature !== undefined) payload.temperature = body.temperature;

    if (!payload.messages.length) {
      return {
        statusCode: 400,
        headers: { ...cors, 'Content-Type': 'application/json' },
        body: JSON.stringify({ error: 'messages[] vacío' })
      };
    }

    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method:  'POST',
      headers: {
        'Content-Type':      'application/json',
        'x-api-key':         CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(payload),
    });

    const data = await resp.json();
    return {
      statusCode: resp.status,
      headers:    { ...cors, 'Content-Type': 'application/json' },
      body:       JSON.stringify(data),
    };
  } catch (err) {
    console.error('claude-proxy error:', err.message);
    return {
      statusCode: 500,
      headers:    { ...cors, 'Content-Type': 'application/json' },
      body:       JSON.stringify({ error: err.message }),
    };
  }
};
