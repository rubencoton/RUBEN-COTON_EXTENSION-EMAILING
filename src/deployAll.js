/**
 * deployAll.js — Despliega EMAILING RUBEN COTON en todas las hojas de una carpeta
 *
 * Uso: node src/deployAll.js
 *
 * 1. Conecta con el hub de APIs (RUBEN-COTON_API-GOOGLE o codex-google-hub)
 * 2. Lista todas las hojas de calculo en la carpeta objetivo
 * 3. Crea un proyecto Apps Script vinculado en cada hoja
 * 4. Sube el codigo de la extension (Main + MailEngine + HTML)
 */

require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { getHub, listSheetsInFolder, deployToSheet } = require('./hubConnector');

// Carpeta objetivo: 🚀 CRM - RUBEN COTON ➕ ARTES BUHO
const TARGET_FOLDER = '1A7agk072QZIS_6HVB3SmX8j2pERh4fW3';

// Leer el HTML de los sidebars
const MAIL_MERGE_HTML = fs.readFileSync(
  path.join(__dirname, '..', 'src-addon', 'MailMerge.html'), 'utf8'
);
const TRACKING_HTML = fs.readFileSync(
  path.join(__dirname, '..', 'src-addon', 'TrackingReport.html'), 'utf8'
);

// Codigo GS compacto del Main
const MAIN_CODE = `
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Iniciar Mail Merge", "showMailMergeSidebar")
    .addSeparator()
    .addItem("Ver informe de seguimiento", "showTrackingReportSidebar")
    .addItem("Estado de cuota", "showQuotaAlert")
    .addToUi();
}
function onInstall(e) { onOpen(e); }
function showMailMergeSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("MailMerge")
    .setTitle("EMAILING RUBEN COTON").setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}
function showTrackingReportSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("TrackingReport")
    .setTitle("Informe de Seguimiento").setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}
function showQuotaAlert() {
  var q = MailApp.getRemainingDailyQuota();
  SpreadsheetApp.getUi().alert("Cuota de envio", "Emails restantes hoy: " + q, SpreadsheetApp.getUi().ButtonSet.OK);
}
`;

// Codigo del motor de mail merge (compacto)
const ENGINE_CODE = fs.readFileSync(
  path.join(__dirname, '..', 'src-addon', 'MailEngine.js'), 'utf8'
);

// Codigo de tracking
const TRACKING_CODE = fs.readFileSync(
  path.join(__dirname, '..', 'src-addon', 'Tracking.js'), 'utf8'
);

// Manifest
const MANIFEST = JSON.stringify({
  timeZone: "Europe/Madrid",
  exceptionLogging: "STACKDRIVER",
  runtimeVersion: "V8",
  oauthScopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.send_mail",
    "https://www.googleapis.com/auth/userinfo.email"
  ]
});

// Files payload para la Apps Script API
const EXTENSION_FILES = [
  { name: "appsscript", type: "JSON", source: MANIFEST },
  { name: "Main", type: "SERVER_JS", source: MAIN_CODE },
  { name: "MailEngine", type: "SERVER_JS", source: ENGINE_CODE },
  { name: "Tracking", type: "SERVER_JS", source: TRACKING_CODE },
  { name: "MailMerge", type: "HTML", source: MAIL_MERGE_HTML },
  { name: "TrackingReport", type: "HTML", source: TRACKING_HTML },
];

async function main() {
  console.log('=== EMAILING RUBEN COTON — Despliegue masivo ===\n');

  // 1. Conectar al hub
  console.log('[1] Conectando al hub de APIs...');
  const hub = await getHub();
  console.log('    Conectado.\n');

  // 2. Listar hojas
  console.log('[2] Explorando carpeta: ' + TARGET_FOLDER);
  const sheets = await listSheetsInFolder(hub.drive, TARGET_FOLDER);
  console.log('    Encontradas: ' + sheets.length + ' hojas de calculo\n');

  if (sheets.length === 0) {
    console.log('    No se encontraron hojas. Verifica el ID de la carpeta.');
    return;
  }

  for (const s of sheets) {
    console.log('    - ' + s.name + ' (' + s.id + ')');
  }

  // 3. Desplegar en cada hoja
  console.log('\n[3] Desplegando extension...\n');
  const results = [];

  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i];
    console.log(`  [${i+1}/${sheets.length}] ${s.name}`);

    try {
      const scriptId = await deployToSheet(hub.auth, s.id, s.name, EXTENSION_FILES);
      results.push({ name: s.name, id: s.id, status: 'OK', scriptId });
    } catch (err) {
      console.log('    ERROR: ' + err.message);
      results.push({ name: s.name, id: s.id, status: 'ERROR', error: err.message });
    }

    // Pausa entre creaciones
    if (i < sheets.length - 1) {
      await new Promise(r => setTimeout(r, 1500));
    }
  }

  // 4. Resumen
  console.log('\n=== RESUMEN ===');
  const ok = results.filter(r => r.status === 'OK').length;
  const fail = results.filter(r => r.status === 'ERROR').length;
  console.log(`Desplegados: ${ok} | Errores: ${fail} | Total: ${sheets.length}`);

  // Guardar resultado
  const resultPath = path.join(__dirname, '..', 'deploy-results.json');
  fs.writeFileSync(resultPath, JSON.stringify(results, null, 2));
  console.log('Resultados guardados en: deploy-results.json');
}

main().catch(err => {
  console.error('Error fatal:', err.message);
  process.exit(1);
});
