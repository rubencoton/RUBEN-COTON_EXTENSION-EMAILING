/**
 * syncOnce.js — Sync manual de todas las hojas
 * Uso: npm run sync
 */

const { getSheetsClient, getSheetMeta, readTab, SHEET_IDS } = require("./sheetsConnector");

const EMAIL_HEADERS = ["email", "correo", "correo electrónico", "correo electronico", "e-mail"];

const findEmailCol = (headers) => {
  const lower = headers.map((h) => (h || "").toLowerCase().trim());
  for (const candidate of EMAIL_HEADERS) {
    const idx = lower.indexOf(candidate);
    if (idx !== -1) return idx;
  }
  return -1;
};

(async () => {
  const client = getSheetsClient();
  if (!client) {
    console.error("Sin credenciales. Copia .env.example a .env y configura.");
    process.exit(1);
  }

  console.log(`Modo auth: ${client.mode}`);
  console.log(`Hojas: ${SHEET_IDS.length}\n`);

  let totalEmails = 0;

  for (const id of SHEET_IDS) {
    try {
      const meta = await getSheetMeta(client.sheets, id);
      console.log(`\n📋 ${meta.title}`);

      for (const tab of meta.tabs) {
        if (tab.toLowerCase() === "ccaa") continue;

        const rows = await readTab(client.sheets, id, tab);
        if (rows.length < 2) {
          console.log(`  ${tab}: vacia`);
          continue;
        }

        const emailIdx = findEmailCol(rows[0]);
        if (emailIdx === -1) {
          console.log(`  ${tab}: sin columna email`);
          continue;
        }

        const emails = rows
          .slice(1)
          .map((r) => (r[emailIdx] || "").trim())
          .filter((e) => e.includes("@"));

        totalEmails += emails.length;
        console.log(`  ${tab}: ${emails.length} emails`);
      }
    } catch (err) {
      console.error(`  ERROR ${id}: ${err.message}`);
    }
  }

  console.log(`\n--- Total: ${totalEmails} emails ---`);
})();
