/**
 * Tracking.js — Seguimiento de campanas
 * Lee la columna "Merge status" y genera estadisticas
 */

/**
 * Obtiene estadisticas de la hoja activa basado en la columna Merge status
 * @returns {Object} stats
 */
function getCampaignStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return { total: 0, sent: 0, opened: 0, clicked: 0, responded: 0, bounced: 0, errors: 0, pending: 0 };
  }

  var headers = data[0].map(function(h) { return String(h || "").toLowerCase().trim(); });
  var msIdx = headers.indexOf("merge status");

  if (msIdx === -1) {
    return { total: data.length - 1, sent: 0, opened: 0, clicked: 0, responded: 0, bounced: 0, errors: 0, pending: data.length - 1, noColumn: true };
  }

  // Detectar columna email para contar solo filas con email
  var emailIdx = -1;
  var emailCandidates = ["email", "correo", "correo electrónico", "correo electronico", "e-mail", "mail"];
  for (var e = 0; e < emailCandidates.length; e++) {
    var i = headers.indexOf(emailCandidates[e]);
    if (i !== -1) { emailIdx = i; break; }
  }

  var stats = {
    total: 0,
    sent: 0,
    opened: 0,
    clicked: 0,
    responded: 0,
    bounced: 0,
    unsubscribed: 0,
    errors: 0,
    pending: 0,
    byStatus: {}
  };

  for (var r = 1; r < data.length; r++) {
    // Solo contar filas con email valido
    if (emailIdx !== -1) {
      var email = String(data[r][emailIdx] || "").trim();
      if (!email || email.indexOf("@") === -1) continue;
    }

    stats.total++;
    var status = String(data[r][msIdx] || "").trim().toUpperCase();

    if (!status) {
      stats.pending++;
      continue;
    }

    // Conteo por categoria
    if (status === "EMAIL_SENT") stats.sent++;
    else if (status === "OPENED") stats.opened++;
    else if (status === "CLICKED") stats.clicked++;
    else if (status === "RESPONDED") stats.responded++;
    else if (status === "BOUNCED") stats.bounced++;
    else if (status === "UNSUBSCRIBED") stats.unsubscribed++;
    else if (status.indexOf("ERROR") === 0) stats.errors++;
    else stats.sent++; // cualquier otro estado se cuenta como enviado

    // Conteo detallado
    stats.byStatus[status] = (stats.byStatus[status] || 0) + 1;
  }

  // Tasas
  var delivered = stats.sent + stats.opened + stats.clicked + stats.responded;
  stats.deliveryRate = stats.total > 0 ? Math.round((delivered / stats.total) * 100) : 0;
  stats.openRate = delivered > 0 ? Math.round(((stats.opened + stats.clicked + stats.responded) / delivered) * 100) : 0;
  stats.clickRate = delivered > 0 ? Math.round(((stats.clicked + stats.responded) / delivered) * 100) : 0;
  stats.responseRate = delivered > 0 ? Math.round((stats.responded / delivered) * 100) : 0;
  stats.bounceRate = stats.total > 0 ? Math.round((stats.bounced / stats.total) * 100) : 0;

  return stats;
}

/**
 * Obtiene lista detallada de destinatarios con su estado
 * @param {string} filterStatus - filtrar por estado (opcional)
 * @returns {Array<{email, status, row}>}
 */
function getDetailedTracking(filterStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) return [];

  var headers = data[0].map(function(h) { return String(h || "").toLowerCase().trim(); });
  var msIdx = headers.indexOf("merge status");
  if (msIdx === -1) return [];

  var emailIdx = -1;
  var emailCandidates = ["email", "correo", "correo electrónico", "correo electronico", "e-mail"];
  for (var e = 0; e < emailCandidates.length; e++) {
    var i = headers.indexOf(emailCandidates[e]);
    if (i !== -1) { emailIdx = i; break; }
  }
  if (emailIdx === -1) return [];

  var nameIdx = -1;
  var nameCandidates = ["nombre contacto", "nombre completo", "nombre de contacto", "first name", "nombre"];
  for (var n = 0; n < nameCandidates.length; n++) {
    var ni = headers.indexOf(nameCandidates[n]);
    if (ni !== -1) { nameIdx = ni; break; }
  }

  var results = [];
  for (var r = 1; r < data.length; r++) {
    var email = String(data[r][emailIdx] || "").trim();
    if (!email || email.indexOf("@") === -1) continue;

    var status = String(data[r][msIdx] || "").trim();

    if (filterStatus && status.toUpperCase() !== filterStatus.toUpperCase()) continue;

    results.push({
      row: r + 1,
      email: email,
      name: nameIdx !== -1 ? String(data[r][nameIdx] || "").trim() : "",
      status: status || "(pendiente)"
    });
  }

  return results;
}

/**
 * Comprueba replies en Gmail para emails enviados y actualiza Merge status
 * Busca respuestas a los emails que tienen estado EMAIL_SENT
 * @returns {Object} {checked, updated}
 */
function checkForReplies() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) return { checked: 0, updated: 0 };

  var headers = data[0].map(function(h) { return String(h || "").toLowerCase().trim(); });
  var msIdx = headers.indexOf("merge status");
  var emailIdx = -1;
  var emailCandidates = ["email", "correo", "correo electrónico", "correo electronico", "e-mail"];
  for (var e = 0; e < emailCandidates.length; e++) {
    var i = headers.indexOf(emailCandidates[e]);
    if (i !== -1) { emailIdx = i; break; }
  }

  if (msIdx === -1 || emailIdx === -1) return { checked: 0, updated: 0 };

  var mergeStatusCol = msIdx + 1; // 1-indexed
  var checked = 0;
  var updated = 0;

  for (var r = 1; r < data.length; r++) {
    var status = String(data[r][msIdx] || "").trim().toUpperCase();
    if (status !== "EMAIL_SENT" && status !== "OPENED") continue;

    var email = String(data[r][emailIdx] || "").trim();
    if (!email) continue;

    checked++;

    // Buscar en Gmail si hay reply de este destinatario
    try {
      var threads = GmailApp.search("from:" + email + " is:inbox", 0, 1);
      if (threads.length > 0) {
        var lastMsg = threads[0].getLastMessageDate();
        var daysDiff = (new Date() - lastMsg) / (1000 * 60 * 60 * 24);
        if (daysDiff <= 10) {
          sheet.getRange(r + 1, mergeStatusCol).setValue("RESPONDED");
          updated++;
        }
      }
    } catch (err) {
      // Skip search errors
    }

    // Limitar busquedas para no agotar tiempo de ejecucion
    if (checked >= 50) break;
  }

  SpreadsheetApp.flush();
  return { checked: checked, updated: updated };
}

/**
 * Detectar bounces buscando en Gmail mensajes de "delivery failure"
 * @returns {Object} {checked, updated}
 */
function checkForBounces() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) return { checked: 0, updated: 0 };

  var headers = data[0].map(function(h) { return String(h || "").toLowerCase().trim(); });
  var msIdx = headers.indexOf("merge status");
  var emailIdx = -1;
  var emailCandidates = ["email", "correo", "correo electrónico", "correo electronico", "e-mail"];
  for (var e = 0; e < emailCandidates.length; e++) {
    var i = headers.indexOf(emailCandidates[e]);
    if (i !== -1) { emailIdx = i; break; }
  }

  if (msIdx === -1 || emailIdx === -1) return { checked: 0, updated: 0 };

  // Buscar bounces en los ultimos 10 dias
  var bounceThreads = GmailApp.search("from:mailer-daemon OR from:postmaster subject:(delivery OR undeliverable OR failure)", 0, 50);
  var bouncedEmails = {};

  for (var t = 0; t < bounceThreads.length; t++) {
    var msgs = bounceThreads[t].getMessages();
    for (var m = 0; m < msgs.length; m++) {
      var body = msgs[m].getPlainBody().toLowerCase();
      // Extraer emails del cuerpo del bounce
      var emailMatches = body.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/g);
      if (emailMatches) {
        for (var em = 0; em < emailMatches.length; em++) {
          bouncedEmails[emailMatches[em]] = true;
        }
      }
    }
  }

  var mergeStatusCol = msIdx + 1;
  var updated = 0;

  for (var r = 1; r < data.length; r++) {
    var status = String(data[r][msIdx] || "").trim().toUpperCase();
    if (status !== "EMAIL_SENT") continue;

    var email = String(data[r][emailIdx] || "").trim().toLowerCase();
    if (bouncedEmails[email]) {
      sheet.getRange(r + 1, mergeStatusCol).setValue("BOUNCED");
      updated++;
    }
  }

  SpreadsheetApp.flush();
  return { checked: Object.keys(bouncedEmails).length, updated: updated };
}

/**
 * Ejecutar todas las comprobaciones de tracking
 */
function refreshTracking() {
  var replies = checkForReplies();
  var bounces = checkForBounces();
  var stats = getCampaignStats();

  return {
    replies: replies,
    bounces: bounces,
    stats: stats
  };
}
