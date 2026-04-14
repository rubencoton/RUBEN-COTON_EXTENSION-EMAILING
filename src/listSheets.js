/**
 * listSheets.js — Listar hojas conectadas
 * Uso: npm run list-sheets
 */

const { listAllSheets, getSheetsClient, SHEET_IDS } = require("./sheetsConnector");

(async () => {
  const client = getSheetsClient();
  if (!client) {
    console.error("Sin credenciales. Copia .env.example a .env y configura.");
    process.exit(1);
  }

  console.log(`\nModo auth: ${client.mode}`);
  console.log(`Hojas configuradas: ${SHEET_IDS.length}\n`);

  const sheets = await listAllSheets();
  console.log("\n--- Resumen ---");
  for (const s of sheets) {
    if (s.error) {
      console.log(`  [ERROR] ${s.id}: ${s.error}`);
    } else {
      console.log(`  ${s.title} — ${s.tabs.length} tabs — ${s.url}`);
    }
  }
})();
