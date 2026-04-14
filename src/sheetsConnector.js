/**
 * sheetsConnector.js — Conector Google Sheets para manager@rubencoton.com
 *
 * Soporta 3 modos de autenticacion:
 *   1. API Key (hojas publicas)
 *   2. OAuth refresh_token (hojas privadas, cuenta manager@rubencoton.com)
 *   3. Service Account JSON
 *
 * Patron reutilizado de APP_ARTES-BUHO_EMAILING/src/sheetsSync.js
 */

require("dotenv").config();
const { google } = require("googleapis");

/* ─── Config ─── */
const API_KEY = process.env.GOOGLE_SHEETS_API_KEY || "";
const CLIENT_ID = process.env.GOOGLE_CLIENT_ID || "";
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || "";
const REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN || "";
const CREDENTIALS_JSON = process.env.GOOGLE_SHEETS_CREDENTIALS || "";
const SHEET_IDS = (process.env.SHEETS_SYNC_IDS || "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);

/* ─── Auth ─── */
const getAuth = () => {
  // Modo 1: API Key
  if (API_KEY) return { mode: "api_key", auth: API_KEY };

  // Modo 2: OAuth con refresh_token (manager@rubencoton.com)
  if (CLIENT_ID && CLIENT_SECRET && REFRESH_TOKEN) {
    const oauth2 = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET);
    oauth2.setCredentials({ refresh_token: REFRESH_TOKEN });
    return { mode: "oauth", auth: oauth2 };
  }

  // Modo 3: Service Account
  if (CREDENTIALS_JSON) {
    try {
      const creds = JSON.parse(CREDENTIALS_JSON);
      const auth = new google.auth.GoogleAuth({
        credentials: creds,
        scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
      });
      return { mode: "service_account", auth };
    } catch (err) {
      console.error("[sheetsConnector] Error parsing credentials:", err.message);
    }
  }

  return null;
};

/* ─── Sheets client ─── */
const getSheetsClient = () => {
  const result = getAuth();
  if (!result) return null;

  if (result.mode === "api_key") {
    return { sheets: google.sheets({ version: "v4", key: result.auth }), mode: result.mode };
  }
  return { sheets: google.sheets({ version: "v4", auth: result.auth }), mode: result.mode };
};

/* ─── Leer metadata de una hoja ─── */
const getSheetMeta = async (sheets, spreadsheetId) => {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  return {
    id: spreadsheetId,
    title: meta.data.properties?.title || spreadsheetId,
    tabs: (meta.data.sheets || []).map((s) => s.properties.title),
    locale: meta.data.properties?.locale,
    url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
  };
};

/* ─── Leer datos de un tab ─── */
const readTab = async (sheets, spreadsheetId, tabName, range) => {
  const fullRange = range || `'${tabName}'!A1:Z`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: fullRange });
  return res.data.values || [];
};

/* ─── Listar todas las hojas configuradas ─── */
const listAllSheets = async () => {
  const client = getSheetsClient();
  if (!client) {
    console.error("[sheetsConnector] Sin credenciales configuradas");
    return [];
  }

  const results = [];
  for (const id of SHEET_IDS) {
    try {
      const meta = await getSheetMeta(client.sheets, id);
      results.push(meta);
      console.log(`  OK: ${meta.title} (${meta.tabs.length} tabs)`);
    } catch (err) {
      results.push({ id, title: "?", error: err.message });
      console.error(`  ERROR: ${id} — ${err.message}`);
    }
  }
  return results;
};

module.exports = {
  getAuth,
  getSheetsClient,
  getSheetMeta,
  readTab,
  listAllSheets,
  SHEET_IDS,
};
