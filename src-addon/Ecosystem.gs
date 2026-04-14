/**
 * Ecosystem.gs — Conexion con todo el ecosistema Google de manager@rubencoton.com
 */

/* ═══════════════════════════════════════════
   GOOGLE SHEETS — CRMs
   ═══════════════════════════════════════════ */

/** IDs de los 6 CRMs principales */
var CRM_SHEETS = {
  "VENTA-BOOKING":         "1_VK6eXqsBuxzvkEJ-uoCIOrsgz-ivpzYKiwwKWzxCtA",
  "OTROS":                 "1WrO-GTi5kvEWnZ_aL4R_PolTIPP3x-lGoAtmTYHTuVo",
  "MUNDO-DISCOGRAFICO":    "1az7PQVf6KlNqvbZh4A6B0_mqLbb2g2Rmq2Y4Febpnas",
  "MARKETING-PROMOCION":   "1Uq2wm2tMKxv11C_ULNszsHzlxvQuZ0EAwzX7AVfXqic",
  "FESTIVALES":            "1H5haxYS4a36srL_aBlOo9W-2AEMqcUP_Bm4kZBbEg_Q",
  "BELLA-BESTIA":          "1OJIIqh60OAE1589Brql9e_RjWvXrCGYttz5g5pb72so"
};

/** Hojas adicionales del ecosistema */
var ECOSYSTEM_SHEETS = {
  "CALENDAR-HUB":          "12uliCcyuRKKfP0C2qTJpsByKC1EwIK_X-eD1CaHP4B0",
  "GMAIL-CRM":             "1ix5hlJmDkdvrzAHzXKLu5JJUjqnuvc0KpHSdntmvKys",
  "BUSCA-CONTACTOS":       "10HvM6MRq2s8nhg-ZTssAWFxuw6VhBxe7RVu006NDUC8",
  "CONTABILIDAD":          "1f1JTbbf1IL7FABJRdrl-rWurDka8VIFQIo7W8Z3eVGg",
  "CRM-CENTRAL":           "1gNg71eDZefbx_8leQ8ij7pOaBXdsWIFnNK_QhXjv8W0"
};

/** Headers conocidos para email */
var EMAIL_HEADERS = ["email", "correo", "correo electrónico", "correo electronico", "e-mail"];
var NAME_HEADERS = ["nombre contacto", "nombre completo", "nombre de contacto"];
var COMPANY_HEADERS = ["nombre", "nombre festival", "nombre discográfica", "nombre discografica", "medio o empresa"];

/**
 * Abrir un CRM por nombre
 */
function openCrmSheet(crmName) {
  var id = CRM_SHEETS[crmName];
  if (!id) throw new Error("CRM no encontrado: " + crmName);
  return SpreadsheetApp.openById(id);
}

/**
 * Listar todos los CRMs con metadata
 */
function listAllCrms() {
  var results = [];
  var keys = Object.keys(CRM_SHEETS);
  for (var i = 0; i < keys.length; i++) {
    try {
      var ss = SpreadsheetApp.openById(CRM_SHEETS[keys[i]]);
      results.push({
        name: keys[i],
        id: CRM_SHEETS[keys[i]],
        title: ss.getName(),
        tabs: ss.getSheets().length,
        url: ss.getUrl()
      });
    } catch (err) {
      results.push({ name: keys[i], id: CRM_SHEETS[keys[i]], error: err.message });
    }
  }
  return results;
}

/**
 * Contar emails en todas las hojas CRM
 */
function countAllCrmEmails() {
  var total = 0;
  var keys = Object.keys(CRM_SHEETS);
  for (var i = 0; i < keys.length; i++) {
    try {
      var ss = SpreadsheetApp.openById(CRM_SHEETS[keys[i]]);
      var sheets = ss.getSheets();
      for (var j = 0; j < sheets.length; j++) {
        if (sheets[j].getName().toLowerCase() === "ccaa") continue;
        var data = sheets[j].getDataRange().getValues();
        if (data.length < 2) continue;
        var emailIdx = findHeaderIndex_(data[0], EMAIL_HEADERS);
        if (emailIdx === -1) continue;
        for (var r = 1; r < data.length; r++) {
          var v = String(data[r][emailIdx] || "").trim();
          if (v.indexOf("@") > -1) total++;
        }
      }
    } catch (e) { /* skip errores de acceso */ }
  }
  return total;
}

/* ═══════════════════════════════════════════
   GOOGLE DRIVE
   ═══════════════════════════════════════════ */

/**
 * Estadisticas basicas de Drive
 */
function getDriveStats() {
  var files = DriveApp.getFiles();
  var count = 0;
  // Contar hasta 1000 para no agotar tiempo
  while (files.hasNext() && count < 1000) {
    files.next();
    count++;
  }
  return {
    fileCount: count,
    hasMore: files.hasNext(),
    storageQuota: DriveApp.getStorageLimit(),
    storageUsed: DriveApp.getStorageUsed()
  };
}

/**
 * Buscar archivos en Drive
 */
function searchDrive(query, maxResults) {
  maxResults = maxResults || 20;
  var files = DriveApp.searchFiles(query);
  var results = [];
  while (files.hasNext() && results.length < maxResults) {
    var f = files.next();
    results.push({
      id: f.getId(),
      name: f.getName(),
      type: f.getMimeType(),
      url: f.getUrl(),
      lastUpdated: f.getLastUpdated()
    });
  }
  return results;
}

/* ═══════════════════════════════════════════
   GMAIL
   ═══════════════════════════════════════════ */

/**
 * Contar emails no leidos
 */
function getUnreadCount() {
  return GmailApp.getInboxUnreadCount();
}

/**
 * Buscar hilos de email
 */
function searchGmail(query, maxResults) {
  maxResults = maxResults || 20;
  var threads = GmailApp.search(query, 0, maxResults);
  return threads.map(function(t) {
    return {
      id: t.getId(),
      subject: t.getFirstMessageSubject(),
      from: t.getMessages()[0].getFrom(),
      date: t.getLastMessageDate(),
      unread: t.isUnread(),
      messageCount: t.getMessageCount()
    };
  });
}

/**
 * Enviar email
 */
function sendEmail(to, subject, htmlBody, options) {
  options = options || {};
  GmailApp.sendEmail(to, subject, "", {
    htmlBody: htmlBody,
    name: options.senderName || "RUBEN COTON",
    replyTo: options.replyTo || "manager@rubencoton.com"
  });
  return { sent: true, to: to, subject: subject };
}

/* ═══════════════════════════════════════════
   GOOGLE CALENDAR
   ═══════════════════════════════════════════ */

/**
 * Eventos de hoy
 */
function getTodayEvents() {
  var cal = CalendarApp.getDefaultCalendar();
  var now = new Date();
  var start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var end = new Date(start.getTime() + 24 * 60 * 60 * 1000);
  var events = cal.getEvents(start, end);
  return events.map(function(ev) {
    return {
      title: ev.getTitle(),
      start: ev.getStartTime(),
      end: ev.getEndTime(),
      location: ev.getLocation(),
      guests: ev.getGuestList().map(function(g) { return g.getEmail(); })
    };
  });
}

/**
 * Crear evento
 */
function createCalendarEvent(title, startDate, endDate, options) {
  var cal = CalendarApp.getDefaultCalendar();
  var event = cal.createEvent(title, new Date(startDate), new Date(endDate));
  if (options && options.description) event.setDescription(options.description);
  if (options && options.location) event.setLocation(options.location);
  return { created: true, id: event.getId(), title: title };
}

/* ═══════════════════════════════════════════
   CONTACTS (People API via UrlFetchApp)
   ═══════════════════════════════════════════ */

/**
 * Obtener contactos via People API
 */
function getGoogleContacts(maxResults) {
  maxResults = maxResults || 100;
  var url = "https://people.googleapis.com/v1/people/me/connections"
    + "?personFields=names,emailAddresses,phoneNumbers,organizations"
    + "&pageSize=" + maxResults;
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  });
  var data = JSON.parse(response.getContentText());
  return (data.connections || []).map(function(p) {
    return {
      name: p.names ? p.names[0].displayName : "",
      email: p.emailAddresses ? p.emailAddresses[0].value : "",
      phone: p.phoneNumbers ? p.phoneNumbers[0].value : "",
      org: p.organizations ? p.organizations[0].name : ""
    };
  });
}

/* ═══════════════════════════════════════════
   STATS AGREGADAS
   ═══════════════════════════════════════════ */

function getEcosystemStats_() {
  var stats = { sheets: 0, driveFiles: 0, unreadEmails: 0, todayEvents: 0 };
  try { stats.sheets = Object.keys(CRM_SHEETS).length + Object.keys(ECOSYSTEM_SHEETS).length; } catch(e) {}
  try {
    var files = DriveApp.getFiles();
    var c = 0;
    while (files.hasNext() && c < 500) { files.next(); c++; }
    stats.driveFiles = c;
  } catch(e) {}
  try { stats.unreadEmails = GmailApp.getInboxUnreadCount(); } catch(e) {}
  try { stats.todayEvents = getTodayEvents().length; } catch(e) {}
  return stats;
}

/* ═══════════════════════════════════════════
   UTILIDADES
   ═══════════════════════════════════════════ */

function findHeaderIndex_(headers, candidates) {
  var lower = headers.map(function(h) { return String(h || "").toLowerCase().trim(); });
  for (var i = 0; i < candidates.length; i++) {
    var idx = lower.indexOf(candidates[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

function countEmailsInActiveSheet_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return 0;
    var sheet = ss.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return 0;
    var idx = findHeaderIndex_(data[0], EMAIL_HEADERS);
    if (idx === -1) return 0;
    var count = 0;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idx] || "").indexOf("@") > -1) count++;
    }
    return count;
  } catch(e) { return 0; }
}

/**
 * Info de la cuenta activa
 */
function getAccountInfo() {
  return {
    email: Session.getActiveUser().getEmail(),
    effectiveEmail: Session.getEffectiveUser().getEmail(),
    timeZone: Session.getScriptTimeZone(),
    locale: Session.getActiveUserLocale()
  };
}
