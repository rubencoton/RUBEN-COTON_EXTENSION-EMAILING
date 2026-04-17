/**
 * hubConnector.js — Conexion con RUBEN-COTON_API-GOOGLE hub
 * Reutiliza el oauth-manager.js del hub central para acceder a todo el ecosistema Google.
 *
 * Uso:
 *   const { getHub, listSheetsInFolder, deployToSheet } = require('./hubConnector');
 *   const hub = await getHub();
 *   const sheets = await listSheetsInFolder(hub.drive, 'FOLDER_ID');
 */

const path = require('path');

// Ruta al hub central
const HUB_PATH = path.join(__dirname, '..', '..', 'RUBEN-COTON_API-GOOGLE');
const CODEX_HUB_PATH = path.join(__dirname, '..', '..', 'codex-google-hub');

/**
 * Obtiene las instancias de APIs del hub central
 */
async function getHub() {
  try {
    const { getServices, getAuthClient } = require(path.join(HUB_PATH, 'src', 'auth', 'oauth-manager'));
    const services = await getServices();
    console.log('[hubConnector] Conectado via RUBEN-COTON_API-GOOGLE');
    return services;
  } catch (err) {
    console.warn('[hubConnector] Hub RUBEN-COTON_API-GOOGLE no disponible:', err.message);
    console.log('[hubConnector] Intentando via codex-google-hub...');
    return getHubFallback_();
  }
}

/**
 * Fallback: usar codex-google-hub directamente
 */
async function getHubFallback_() {
  const { google } = require('googleapis');
  const fs = require('fs');

  // Buscar token y client secret en codex-google-hub
  const tokenPath = path.join(CODEX_HUB_PATH, 'secrets', 'oauth_token_manager_rubencoton_com.json');
  const clientPath = path.join(CODEX_HUB_PATH, 'secrets', 'client_secret.json');

  if (!fs.existsSync(tokenPath) || !fs.existsSync(clientPath)) {
    throw new Error('No se encontraron credenciales en codex-google-hub');
  }

  const token = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
  const clientData = JSON.parse(fs.readFileSync(clientPath, 'utf8'));
  const { client_id, client_secret } = clientData.installed || clientData.web || clientData;

  const auth = new google.auth.OAuth2(client_id, client_secret);
  auth.setCredentials(token);

  console.log('[hubConnector] Conectado via codex-google-hub (fallback)');

  return {
    auth,
    gmail: google.gmail({ version: 'v1', auth }),
    drive: google.drive({ version: 'v3', auth }),
    sheets: google.sheets({ version: 'v4', auth }),
    calendar: google.calendar({ version: 'v3', auth }),
    people: google.people({ version: 'v1', auth }),
  };
}

/**
 * Lista todas las hojas de calculo en una carpeta de Drive (recursivo)
 */
async function listSheetsInFolder(drive, folderId, depth) {
  depth = depth || 0;
  const results = [];

  const res = await drive.files.list({
    q: `'${folderId}' in parents and trashed=false`,
    fields: 'files(id,name,mimeType)',
    pageSize: 100
  });

  for (const f of (res.data.files || [])) {
    if (f.mimeType === 'application/vnd.google-apps.spreadsheet') {
      results.push({ id: f.id, name: f.name });
    } else if (f.mimeType === 'application/vnd.google-apps.folder' && depth < 5) {
      const sub = await listSheetsInFolder(drive, f.id, depth + 1);
      results.push(...sub);
    }
  }

  return results;
}

/**
 * Crea un proyecto Apps Script vinculado a una hoja de calculo
 * y sube el codigo de EMAILING RUBEN COTON
 */
async function deployToSheet(auth, spreadsheetId, spreadsheetName, extensionCode) {
  const { google } = require('googleapis');
  const script = google.script({ version: 'v1', auth });

  // 1. Crear proyecto vinculado
  const createRes = await script.projects.create({
    requestBody: {
      title: 'EMAILING RUBEN COTON — ' + spreadsheetName,
      parentId: spreadsheetId
    }
  });

  const scriptId = createRes.data.scriptId;
  console.log(`  Proyecto creado: ${scriptId}`);

  // 2. Subir contenido
  await script.projects.updateContent({
    scriptId: scriptId,
    requestBody: {
      files: extensionCode
    }
  });

  console.log(`  Codigo subido OK`);
  return scriptId;
}

module.exports = {
  getHub,
  listSheetsInFolder,
  deployToSheet,
};
