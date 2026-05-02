/**
 * ===== ENERGY — Proxy UNIFICADO (Claude + GitHub Writes) =====
 *
 * Reemplaza al claude-proxy.gs y al github-proxy.gs anteriores.
 * Un solo Apps Script, una sola URL, que detecta automáticamente
 * qué tipo de petición es:
 *   - Si el body tiene { messages:[...] }  → reenvía a Anthropic Claude
 *   - Si el body tiene { file, content }   → escribe el archivo en GitHub
 *
 * Cómo desplegarlo (una sola vez):
 * 1. https://script.google.com/home → Nuevo proyecto.
 * 2. Pega TODO este archivo en Code.gs.
 * 3. ⚙ Configuración del proyecto → Propiedades del script → Agregar:
 *      CLAUDE_API_KEY = sk-ant-api03-...   (para el chat ENERGY IA)
 *      GH_TOKEN       = ghp_...            (PAT con permiso 'repo')
 *      GH_OWNER       = cami902026-oss
 *      GH_REPO        = plataforma-eyg
 *      GH_BRANCH      = main
 * 4. Implementar → Nueva implementación → Aplicación web
 *      - Ejecutar como: Tu cuenta
 *      - Acceso: Cualquier usuario, incluso anónimo
 * 5. Copia la URL /exec.
 * 6. En la plataforma → Configuración pega la MISMA URL en:
 *      - "🤖 ENERGY IA — URL del Proxy seguro"
 *      - "📤 Proxy GitHub (escritura compartida sin PAT)"
 *
 * Después de pegar y guardar, todos los dispositivos del equipo la reciben
 * automáticamente vía data/config.json.
 */

const PROPS = PropertiesService.getScriptProperties();
const MODEL_DEFAULT  = 'claude-haiku-4-5-20251001';
const MAX_TOKENS_CAP = 4096;

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');

    // Detección por contenido
    if (Array.isArray(body.messages)) {
      return _handleClaude(body);
    }
    if (typeof body.file === 'string') {
      return _handleGitHubWrite(body);
    }
    return _json({ error: 'Petición inválida — falta messages[] o file' });

  } catch (err) {
    return _json({ error: 'Proxy error: ' + err.message });
  }
}

function doGet() {
  return _json({
    ok: true,
    service: 'ENERGY proxy unificado',
    handles: ['claude','github-write'],
    model: MODEL_DEFAULT
  });
}

// ─── Claude (chat IA) ────────────────────────────────────────────────────────
function _handleClaude(body) {
  const apiKey = PROPS.getProperty('CLAUDE_API_KEY');
  if (!apiKey) {
    return _json({ error: 'CLAUDE_API_KEY no configurada en Propiedades del script' });
  }

  const payload = {
    model:      body.model      || MODEL_DEFAULT,
    max_tokens: Math.min(parseInt(body.max_tokens) || 1024, MAX_TOKENS_CAP),
    system:     body.system     || '',
    messages:   body.messages
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
}

// ─── GitHub write ────────────────────────────────────────────────────────────
function _handleGitHubWrite(body) {
  const tok    = PROPS.getProperty('GH_TOKEN');
  const owner  = PROPS.getProperty('GH_OWNER')  || 'cami902026-oss';
  const repo   = PROPS.getProperty('GH_REPO')   || 'plataforma-eyg';
  const branch = PROPS.getProperty('GH_BRANCH') || 'main';

  if (!tok) return _json({ ok:false, error:'GH_TOKEN no configurado en Propiedades del script' });

  const file = String(body.file || '').replace(/^\/+/, '');
  if (!file) return _json({ ok:false, error:'Falta el campo "file"' });

  let contentStr;
  if (typeof body.content === 'string') contentStr = body.content;
  else contentStr = JSON.stringify(body.content, null, 2);

  const label = body.label || ('Update ' + file);

  if (!_isAllowedFile(file)) {
    return _json({ ok:false, error:'Archivo no permitido: '+file });
  }

  const sha = _ghGetSha(owner, repo, file, branch, tok);
  let result = _ghPutFile(owner, repo, file, branch, tok, contentStr, label, sha);

  if (!result.ok && (result.status === 422 || result.status === 409)) {
    const sha2 = _ghGetSha(owner, repo, file, branch, tok);
    result = _ghPutFile(owner, repo, file, branch, tok, contentStr, label, sha2);
    if (result.ok) return _json({ ok:true, file:file, sha:result.sha, retried:true });
  }
  if (result.ok) return _json({ ok:true, file:file, sha:result.sha });
  return _json({ ok:false, error:'GitHub error '+result.status, body:result.body });
}

function _isAllowedFile(file) {
  const allow = [
    /^ordenes\.json$/,
    /^reuniones\.json$/,
    /^data\/[a-z0-9_\-]+\.json$/i,
  ];
  return allow.some(re => re.test(file));
}

function _ghGetSha(owner, repo, file, branch, tok) {
  const url = 'https://api.github.com/repos/'+owner+'/'+repo+'/contents/'+encodeURI(file)+'?ref='+branch;
  const resp = UrlFetchApp.fetch(url, {
    method:'get',
    headers:{ 'Authorization':'Bearer '+tok, 'Accept':'application/vnd.github+json' },
    muteHttpExceptions:true
  });
  if (resp.getResponseCode() === 200) {
    const j = JSON.parse(resp.getContentText());
    return j.sha;
  }
  return null;
}

function _ghPutFile(owner, repo, file, branch, tok, contentStr, label, sha) {
  const url = 'https://api.github.com/repos/'+owner+'/'+repo+'/contents/'+encodeURI(file);
  const payload = {
    message: label + ' — ' + new Date().toLocaleString('es-CO'),
    content: Utilities.base64Encode(contentStr, Utilities.Charset.UTF_8),
    branch: branch
  };
  if (sha) payload.sha = sha;
  const resp = UrlFetchApp.fetch(url, {
    method:'put',
    contentType:'application/json',
    headers:{ 'Authorization':'Bearer '+tok, 'Accept':'application/vnd.github+json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions:true
  });
  const status = resp.getResponseCode();
  const text = resp.getContentText();
  if (status >= 200 && status < 300) {
    const j = JSON.parse(text);
    return { ok:true, sha: j.content && j.content.sha };
  }
  return { ok:false, status:status, body:text };
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
