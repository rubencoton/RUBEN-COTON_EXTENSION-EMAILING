/**
 * MailEngine.js — Motor principal de Mail Merge
 * Clon de YAMM para "EMAILING RUBEN COTON"
 * Cuenta: manager@rubencoton.com
 */

/* ═══════════════════════════════════════════
   BORRADORES DE GMAIL
   ═══════════════════════════════════════════ */

/**
 * Lista todos los borradores de Gmail del usuario
 * @returns {Array<{id, subject, snippet}>}
 */
function getDrafts() {
  var drafts = GmailApp.getDrafts();
  return drafts.map(function(d) {
    var msg = d.getMessage();
    return {
      id: d.getId(),
      messageId: msg.getId(),
      subject: msg.getSubject() || "(sin asunto)",
      snippet: msg.getPlainBody().substring(0, 80),
      hasAttachments: msg.getAttachments().length > 0,
      attachmentCount: msg.getAttachments().length
    };
  });
}

/**
 * Obtiene el contenido completo de un borrador
 * @param {string} draftId
 * @returns {Object} {subject, htmlBody, plainBody, from, attachments}
 */
function getDraftById(draftId) {
  var drafts = GmailApp.getDrafts();
  var draft = null;
  for (var i = 0; i < drafts.length; i++) {
    if (drafts[i].getId() === draftId) {
      draft = drafts[i];
      break;
    }
  }
  if (!draft) throw new Error("Borrador no encontrado: " + draftId);

  var msg = draft.getMessage();
  var attachments = msg.getAttachments();

  return {
    id: draftId,
    subject: msg.getSubject(),
    htmlBody: msg.getBody(),
    plainBody: msg.getPlainBody(),
    from: msg.getFrom(),
    attachments: attachments.map(function(a) {
      return { name: a.getName(), size: a.getSize(), type: a.getContentType() };
    })
  };
}

/* ═══════════════════════════════════════════
   ALIASES DE GMAIL
   ═══════════════════════════════════════════ */

/**
 * Obtiene los aliases configurados en Gmail via API
 * @returns {Array<{email, displayName, isDefault}>}
 */
function getAliases() {
  var url = "https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs";
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  });
  var data = JSON.parse(response.getContentText());
  return (data.sendAs || []).map(function(a) {
    return {
      email: a.sendAsEmail,
      displayName: a.displayName || "",
      isDefault: a.isDefault || false,
      isPrimary: a.isPrimary || false
    };
  });
}

/* ═══════════════════════════════════════════
   LEER DESTINATARIOS DE LA HOJA ACTIVA
   ═══════════════════════════════════════════ */

/** Headers conocidos para detectar email */
var MERGE_EMAIL_HEADERS = ["email", "correo", "correo electrónico", "correo electronico",
  "e-mail", "email address", "mail", "direccion de correo"];

/**
 * Lee los destinatarios y datos de la hoja activa
 * @param {Object} options - {filterColumn, filterValue, skipSent}
 * @returns {Object} {headers, recipients, emailColIndex, mergeStatusColIndex, totalRows}
 */
function getRecipientsFromSheet(options) {
  options = options || {};
  var ss, sheet;
  // Fallback para triggers programados (no hay spreadsheet activo)
  if (options._sheetId) {
    ss = SpreadsheetApp.openById(options._sheetId);
    sheet = options._sheetName ? ss.getSheetByName(options._sheetName) : ss.getSheets()[0];
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.getActiveSheet();
  }
  var data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return { headers: [], recipients: [], emailColIndex: -1, mergeStatusColIndex: -1, totalRows: 0 };
  }

  var headers = data[0].map(function(h) { return String(h || "").trim(); });
  var headersLower = headers.map(function(h) { return h.toLowerCase(); });

  // Detectar columna de email
  var emailColIndex = -1;
  for (var e = 0; e < MERGE_EMAIL_HEADERS.length; e++) {
    var idx = headersLower.indexOf(MERGE_EMAIL_HEADERS[e]);
    if (idx !== -1) { emailColIndex = idx; break; }
  }
  if (emailColIndex === -1) {
    throw new Error("No se encontro columna de email en la hoja. Headers: " + headers.join(", "));
  }

  // Detectar columna Merge status
  var mergeStatusColIndex = headersLower.indexOf("merge status");

  // Filtro por columna custom
  var filterColIndex = -1;
  if (options.filterColumn) {
    filterColIndex = headersLower.indexOf(options.filterColumn.toLowerCase());
  }

  var recipients = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var email = String(row[emailColIndex] || "").trim();
    if (!email || email.indexOf("@") === -1) continue;

    // Skip si ya fue enviado y skipSent esta activo
    if (options.skipSent && mergeStatusColIndex !== -1) {
      var status = String(row[mergeStatusColIndex] || "").trim().toUpperCase();
      if (status === "EMAIL_SENT" || status === "OPENED" || status === "CLICKED" || status === "RESPONDED") {
        continue;
      }
    }

    // Filtro por columna
    if (filterColIndex !== -1 && options.filterValue) {
      var cellVal = String(row[filterColIndex] || "").trim();
      if (cellVal.toLowerCase() !== options.filterValue.toLowerCase()) continue;
    }

    // Construir objeto de merge vars
    var mergeVars = {};
    for (var c = 0; c < headers.length; c++) {
      mergeVars[headers[c]] = String(row[c] || "");
    }

    recipients.push({
      rowIndex: r + 1, // 1-indexed para Sheets
      email: email,
      mergeVars: mergeVars
    });
  }

  return {
    headers: headers,
    recipients: recipients,
    emailColIndex: emailColIndex,
    mergeStatusColIndex: mergeStatusColIndex,
    totalRows: data.length - 1,
    sheetName: sheet.getName()
  };
}

/**
 * Devuelve las columnas de la hoja activa (para el UI de filtro)
 */
function getSheetColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 1) return [];
  return data[0].map(function(h) { return String(h || "").trim(); }).filter(Boolean);
}

/* ═══════════════════════════════════════════
   SUSTITUCION DE VARIABLES {{var}}
   ═══════════════════════════════════════════ */

/**
 * Reemplaza {{variable}} en un texto con los valores del row
 * @param {string} template
 * @param {Object} mergeVars - {headerName: value}
 * @returns {string}
 */
function replaceVariables(template, mergeVars) {
  if (!template) return "";
  return template.replace(/\{\{([^}]+)\}\}/g, function(match, varName) {
    var key = varName.trim();
    return mergeVars.hasOwnProperty(key) ? mergeVars[key] : "";
  });
}

/* ═══════════════════════════════════════════
   COLUMNA MERGE STATUS
   ═══════════════════════════════════════════ */

/**
 * Asegura que existe la columna "Merge status" en la hoja activa
 * @returns {number} indice de la columna (1-indexed)
 */
function ensureMergeStatusColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headersLower = headers.map(function(h) { return String(h || "").toLowerCase().trim(); });

  var idx = headersLower.indexOf("merge status");
  if (idx !== -1) return idx + 1; // ya existe, devolver 1-indexed

  // Crear la columna al final
  var newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue("Merge status");
  return newCol;
}

/**
 * Escribe el estado de merge en una fila
 * @param {number} rowIndex - fila (1-indexed)
 * @param {number} mergeStatusCol - columna (1-indexed)
 * @param {string} status - EMAIL_SENT, ERROR, BOUNCED, etc.
 */
function updateMergeStatus(rowIndex, mergeStatusCol, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.getRange(rowIndex, mergeStatusCol).setValue(status);
  SpreadsheetApp.flush();
}

/* ═══════════════════════════════════════════
   CUOTA DE ENVIO
   ═══════════════════════════════════════════ */

/**
 * Obtiene la cuota restante de email del dia
 */
function getRemainingQuota() {
  return MailApp.getRemainingDailyQuota();
}

/**
 * Funcion auxiliar para autorizar permisos desde el editor.
 * Ejecutar esta funcion UNA VEZ desde el editor para activar el popup de autorizacion.
 */
function autorizar() {
  var email = Session.getActiveUser().getEmail();
  var quota = MailApp.getRemainingDailyQuota();
  Logger.log("Autorizado como: " + email + " | Cuota restante: " + quota);
}

/* ═══════════════════════════════════════════
   ENVIO DE PRUEBA
   ═══════════════════════════════════════════ */

/**
 * Envia un email de prueba al propio usuario con datos de la primera fila
 * @param {Object} config - {draftId, senderName, alias}
 * @returns {Object} resultado
 */
function sendTestEmail(config) {
  var draft = getDraftById(config.draftId);
  var recipientData = getRecipientsFromSheet({});

  if (recipientData.recipients.length === 0) {
    throw new Error("No hay destinatarios en la hoja");
  }

  var firstRow = recipientData.recipients[0];
  var subject = replaceVariables(draft.subject, firstRow.mergeVars);
  var htmlBody = replaceVariables(draft.htmlBody, firstRow.mergeVars);
  var myEmail = Session.getActiveUser().getEmail();

  var emailOptions = {
    htmlBody: htmlBody,
    name: config.senderName || "EMAILING RUBEN COTON"
  };

  if (config.alias) {
    emailOptions.from = config.alias;
  }

  // Adjuntos del borrador
  var draftMsg = GmailApp.getDrafts().filter(function(d) { return d.getId() === config.draftId; })[0];
  if (draftMsg) {
    var attachments = draftMsg.getMessage().getAttachments();
    if (attachments.length > 0) {
      emailOptions.attachments = attachments;
    }
  }

  GmailApp.sendEmail(myEmail, "[TEST] " + subject, "", emailOptions);

  return {
    sent: true,
    to: myEmail,
    subject: "[TEST] " + subject,
    preview: "Datos de fila: " + firstRow.email
  };
}

/* ═══════════════════════════════════════════
   MAIL MERGE PRINCIPAL
   ═══════════════════════════════════════════ */

/**
 * Ejecuta el mail merge completo
 * @param {Object} config
 *   - draftId: ID del borrador
 *   - senderName: nombre del remitente
 *   - alias: email alias (opcional)
 *   - cc: direccion CC (opcional)
 *   - bcc: direccion BCC (opcional)
 *   - filterColumn: columna para filtrar (opcional)
 *   - filterValue: valor del filtro (opcional)
 *   - skipSent: boolean, saltar filas ya enviadas
 * @returns {Object} resultado del merge
 */
function sendMailMerge(config) {
  var draft = getDraftById(config.draftId);
  var recipientData = getRecipientsFromSheet({
    filterColumn: config.filterColumn,
    filterValue: config.filterValue,
    skipSent: config.skipSent !== false
  });

  if (recipientData.recipients.length === 0) {
    return { sent: 0, errors: 0, total: 0, message: "No hay destinatarios para enviar" };
  }

  // Verificar cuota
  var quota = getRemainingQuota();
  var toSend = recipientData.recipients.length;
  if (toSend > quota) {
    // Marcar los que exceden como SCHEDULED
    var mergeStatusColQ = ensureMergeStatusColumn();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    for (var q = quota; q < recipientData.recipients.length; q++) {
      sheet.getRange(recipientData.recipients[q].rowIndex, mergeStatusColQ).setValue("SCHEDULED");
    }
    SpreadsheetApp.flush();
    // Solo enviar los que caben en la cuota
    recipientData.recipients = recipientData.recipients.slice(0, quota);
    toSend = recipientData.recipients.length;
  }

  // Asegurar columna Merge status
  var mergeStatusCol = ensureMergeStatusColumn();

  // Obtener adjuntos del borrador
  var draftObj = GmailApp.getDrafts().filter(function(d) { return d.getId() === config.draftId; })[0];
  var attachments = draftObj ? draftObj.getMessage().getAttachments() : [];

  var sent = 0;
  var errors = 0;
  var bounced = 0;
  var cache = CacheService.getScriptCache();

  // Progreso inicial
  cache.put("mergeProgress", JSON.stringify({
    sent: 0, errors: 0, bounced: 0, total: toSend, running: true, startTime: new Date().toISOString()
  }), 600);

  for (var i = 0; i < recipientData.recipients.length; i++) {
    var recipient = recipientData.recipients[i];

    try {
      var subject = replaceVariables(draft.subject, recipient.mergeVars);
      var htmlBody = replaceVariables(draft.htmlBody, recipient.mergeVars);

      var emailOptions = {
        htmlBody: htmlBody,
        name: config.senderName || "EMAILING RUBEN COTON"
      };

      if (config.alias) emailOptions.from = config.alias;
      if (config.cc) emailOptions.cc = config.cc;
      if (config.bcc) emailOptions.bcc = config.bcc;
      if (config.replyTo) emailOptions.replyTo = config.replyTo;
      if (attachments.length > 0) emailOptions.attachments = attachments;

      GmailApp.sendEmail(recipient.email, subject, "", emailOptions);

      updateMergeStatus(recipient.rowIndex, mergeStatusCol, "EMAIL_SENT");
      sent++;

    } catch (err) {
      var errMsg = err.message || "";
      // Detectar bounces (invalid email, recipient not found)
      if (errMsg.indexOf("Invalid") > -1 || errMsg.indexOf("not found") > -1 || errMsg.indexOf("invalid") > -1) {
        updateMergeStatus(recipient.rowIndex, mergeStatusCol, "BOUNCED");
        bounced++;
      } else {
        updateMergeStatus(recipient.rowIndex, mergeStatusCol, "ERROR");
        errors++;
      }
    }

    // Actualizar progreso cada 5 emails (para polling en tiempo real)
    if (i % 5 === 0 || i === recipientData.recipients.length - 1) {
      cache.put("mergeProgress", JSON.stringify({
        sent: sent, errors: errors, bounced: bounced, total: toSend,
        current: i + 1, running: (i < recipientData.recipients.length - 1),
        lastEmail: recipient.email
      }), 600);
    }

    // Pausa entre envios (100ms para no saturar Gmail)
    if (i < recipientData.recipients.length - 1) {
      Utilities.sleep(100);
    }
  }

  // Progreso final
  cache.put("mergeProgress", JSON.stringify({
    sent: sent, errors: errors, bounced: bounced, total: toSend,
    current: toSend, running: false
  }), 600);

  // Guardar campana
  var campaignId = "campaign_" + new Date().getTime();
  var campaignData = {
    id: campaignId,
    date: new Date().toISOString(),
    draftSubject: draft.subject,
    senderName: config.senderName,
    sent: sent,
    errors: errors,
    bounced: bounced,
    total: toSend,
    sheetName: recipientData.sheetName
  };

  var props = PropertiesService.getScriptProperties();
  props.setProperty(campaignId, JSON.stringify(campaignData));
  props.setProperty("lastCampaignId", campaignId);

  SpreadsheetApp.flush();

  return {
    campaignId: campaignId,
    sent: sent,
    errors: errors,
    bounced: bounced,
    total: toSend,
    quota: getRemainingQuota()
  };
}

/**
 * Progreso en tiempo real del merge (polling desde el sidebar via CacheService)
 */
function getMergeProgress() {
  var cache = CacheService.getScriptCache();
  var data = cache.get("mergeProgress");
  return data ? JSON.parse(data) : null;
}

/**
 * Lista campanas anteriores
 */
function listCampaigns() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var campaigns = [];
  var keys = Object.keys(all);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf("campaign_") === 0) {
      try {
        campaigns.push(JSON.parse(all[keys[i]]));
      } catch (e) { /* skip corrupted */ }
    }
  }
  // Ordenar por fecha desc
  campaigns.sort(function(a, b) { return b.date > a.date ? 1 : -1; });
  return campaigns;
}
