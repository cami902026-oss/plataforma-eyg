/**
 * ===== ENERGY — GitHub Write Proxy (Google Apps Script) =====
 *
 * Permite que CUALQUIER dispositivo del equipo escriba al repo sin tener
 * que configurar un Personal Access Token (PAT) propio. El PAT vive sólo
 * en este Apps Script (Propiedades del script).
 *
 * Cómo desplegarlo:
 * 1. https://script.google.com/home → Nuevo proyecto.
 * 2. Pega TODO este archivo en Code.gs.
 * 3. Configuración del proyecto (⚙) → Propiedades del script → Agregar:
 *    - GH_TOKEN  = ghp_...   (PAT con permiso 'repo')
 *    - GH_OWNER  = cami902026-oss
 *    - GH_REPO   = plataforma-eyg
 *    - GH_BRANCH = main
 * 4. Implementar → Nueva implementación → Aplicación web
 *    - Ejecutar como: Tu cuenta
 *    - Acceso: Cualquier usuario, incluso anónimo
 * 5. Copia la URL /exec y pégala en la plataforma → Configuración →
 *    "Proxy GitHub (escritura compartida sin PAT)".
 *
 * El frontend envía POST con {file, content, label}. La función:
 *   - Obtiene el SHA actual del archivo (si existe)
 *   - Hace PUT al endpoint de contenido de GitHub con base64
 *   - Reintenta una vez si hay 422/409 (race condition con commits paralelos)
 *
 * NUNCA expone el token al navegador.
 */

const PROPS = PropertiesService.getScriptProperties();

function _cors() {
  return ContentService.createTextOutput()
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const tok    = PROPS.getProperty('GH_TOKEN');
    const owner  = PROPS.getProperty('GH_OWNER')  || 'cami902026-oss';
    const repo   = PROPS.getProperty('GH_REPO')   || 'plataforma-eyg';
    const branch = PROPS.getProperty('GH_BRANCH') || 'main';

    if (!tok) {
      return _json({ ok:false, error:'GH_TOKEN no configurado en Propiedades del script' });
    }

    const body = JSON.parse(e.postData.contents || '{}');
    const file = String(body.file || '').replace(/^\/+/, '');
    if (!file) return _json({ ok:false, error:'Falta el campo "file"' });

    // El contenido puede venir como objeto/array (lo serializamos) o como string
    let contentStr;
    if (typeof body.content === 'string') contentStr = body.content;
    else contentStr = JSON.stringify(body.content, null, 2);

    const label = body.label || ('Update ' + file);

    // Whitelist de archivos: sólo aceptamos paths conocidos del repo para evitar abuso
    if (!_isAllowedFile(file)) {
      return _json({ ok:false, error:'Archivo no permitido: '+file });
    }

    const sha = _getSha(owner, repo, file, branch, tok);
    const result = _putFile(owner, repo, file, branch, tok, contentStr, label, sha);

    if (result.ok) return _json({ ok:true, file:file, sha:result.sha });
    // Reintento ante race condition
    if (result.status === 422 || result.status === 409) {
      const sha2 = _getSha(owner, repo, file, branch, tok);
      const result2 = _putFile(owner, repo, file, branch, tok, contentStr, label, sha2);
      if (result2.ok) return _json({ ok:true, file:file, sha:result2.sha, retried:true });
      return _json({ ok:false, error:'GitHub error '+result2.status, body:result2.body });
    }
    return _json({ ok:false, error:'GitHub error '+result.status, body:result.body });
  } catch (err) {
    return _json({ ok:false, error:String(err) });
  }
}

function doGet() {
  return _json({ ok:true, service:'ENERGY GitHub write proxy' });
}

// Sólo se permiten estos archivos (todo lo demás se rechaza)
function _isAllowedFile(file) {
  const allow = [
    /^ordenes\.json$/,
    /^reuniones\.json$/,
    /^data\/[a-z0-9_\-]+\.json$/i,
  ];
  return allow.some(re => re.test(file));
}

function _getSha(owner, repo, file, branch, tok) {
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

function _putFile(owner, repo, file, branch, tok, contentStr, label, sha) {
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
