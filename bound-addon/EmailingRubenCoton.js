/**
 * EmailingRubenCoton.js — BUNDLE COMPLETO (bound add-on)
 * Concatenacion de: Main.js + MailEngine.js + Tracking.js + Contacts.js + Scheduler.js + Ecosystem.js
 * Cuenta: manager@rubencoton.com
 * Generado: 2026-04-14
 */

// ═══════════════════════════════════════════════════════════════════════════════
// SECCION 1 — Main.js
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Main.js — EMAILING RUBEN COTON
 * Clon de YAMM (Yet Another Mail Merge)
 * Cuenta: manager@rubencoton.com
 */

/* ═══════════════════════════════════════════
   MENU (al abrir una hoja de calculo)
   ═══════════════════════════════════════════ */

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Iniciar Mail Merge", "showMailMergeSidebar")
    .addSeparator()
    .addItem("Importar Google Contacts", "showImportContactsSidebar")
    .addItem("Ver informe de seguimiento", "showTrackingReportSidebar")
    .addSeparator()
    .addItem("Campanas anteriores", "showCampaignHistory")
    .addItem("Estado de cuota", "showQuotaStatus")
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

/* ═══════════════════════════════════════════
   SIDEBARS
   ═══════════════════════════════════════════ */

/** Sidebar principal: configurar y enviar mail merge */
function showMailMergeSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("MailMerge")
    .setTitle("EMAILING RUBEN COTON")
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Sidebar: importar contactos de Google */
function showImportContactsSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("ImportContacts")
    .setTitle("Importar Contactos")
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Sidebar: informe de tracking */
function showTrackingReportSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("TrackingReport")
    .setTitle("Informe de Seguimiento")
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ═══════════════════════════════════════════
   DIALOGS
   ═══════════════════════════════════════════ */

/** Mostrar historial de campanas */
function showCampaignHistory() {
  var campaigns = listCampaigns();
  var ui = SpreadsheetApp.getUi();

  if (campaigns.length === 0) {
    ui.alert("Campanas", "No hay campanas registradas.", ui.ButtonSet.OK);
    return;
  }

  var msg = campaigns.slice(0, 10).map(function(c) {
    return c.date.substring(0, 10) + " — " + c.draftSubject +
      " (" + c.sent + " enviados, " + c.errors + " errores)";
  }).join("\n");

  ui.alert("Ultimas campanas", msg, ui.ButtonSet.OK);
}

/** Mostrar cuota restante */
function showQuotaStatus() {
  var quota = getRemainingQuota();
  var ui = SpreadsheetApp.getUi();
  ui.alert("Cuota de envio", "Emails restantes hoy: " + quota, ui.ButtonSet.OK);
}

/* ═══════════════════════════════════════════
   CARD UI (Add-on v2 — homepage triggers)
   ═══════════════════════════════════════════ */

function onHomepage(e) {
  return buildHomeCard_();
}

function onSheetsHomepage(e) {
  return buildMailMergeCard_();
}

/** Card principal */
function buildHomeCard_() {
  var card = CardService.newCardBuilder();
  card.setHeader(
    CardService.newCardHeader()
      .setTitle("EMAILING RUBEN COTON")
      .setSubtitle("Mail Merge para Google Sheets")
  );

  // Stats
  var section = CardService.newCardSection().setHeader("Estado");
  try {
    var quota = getRemainingQuota();
    section.addWidget(
      CardService.newDecoratedText().setText("Cuota restante: " + quota + " emails")
    );
  } catch(e) {
    section.addWidget(
      CardService.newDecoratedText().setText("Cuota: no disponible")
    );
  }

  var campaigns = listCampaigns();
  section.addWidget(
    CardService.newDecoratedText().setText("Campanas enviadas: " + campaigns.length)
  );

  if (campaigns.length > 0) {
    var last = campaigns[0];
    section.addWidget(
      CardService.newDecoratedText()
        .setText("Ultima: " + last.draftSubject)
        .setBottomLabel(last.sent + " enviados — " + last.date.substring(0, 10))
    );
  }

  card.addSection(section);

  // Acciones
  var actions = CardService.newCardSection().setHeader("Acciones");
  actions.addWidget(
    CardService.newTextButton()
      .setText("Iniciar Mail Merge")
      .setOnClickAction(CardService.newAction().setFunctionName("cardOpenMailMerge"))
  );
  actions.addWidget(
    CardService.newTextButton()
      .setText("Ver informe de seguimiento")
      .setOnClickAction(CardService.newAction().setFunctionName("cardOpenTracking"))
  );
  card.addSection(actions);

  return card.build();
}

/** Card para Sheets: muestra info de la hoja activa */
function buildMailMergeCard_() {
  var card = CardService.newCardBuilder();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ss ? ss.getName() : "(sin hoja activa)";

  card.setHeader(
    CardService.newCardHeader()
      .setTitle("EMAILING RUBEN COTON")
      .setSubtitle(sheetName)
  );

  var section = CardService.newCardSection().setHeader("Hoja actual");

  if (ss) {
    var emailCount = countEmailsInActiveSheet_();
    section.addWidget(
      CardService.newDecoratedText().setText("Emails detectados: " + emailCount)
    );
    section.addWidget(
      CardService.newDecoratedText().setText("Tabs: " + ss.getSheets().length)
    );
  }

  section.addWidget(
    CardService.newTextButton()
      .setText("Iniciar Mail Merge")
      .setOnClickAction(CardService.newAction().setFunctionName("cardOpenMailMerge"))
  );

  card.addSection(section);
  return card.build();
}

/* ═══════════════════════════════════════════
   CARD ACTIONS
   ═══════════════════════════════════════════ */

function cardOpenMailMerge(e) {
  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("Abre desde: Extensiones > EMAILING RUBEN COTON > Iniciar Mail Merge")
    )
    .build();
}

function cardOpenTracking(e) {
  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("Abre desde: Extensiones > EMAILING RUBEN COTON > Ver informe")
    )
    .build();
}


// ═══════════════════════════════════════════════════════════════════════════════
// SECCION 2 — MailEngine.js
// ═══════════════════════════════════════════════════════════════════════════════

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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
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
    throw new Error("Cuota insuficiente. Necesitas " + toSend + " pero quedan " + quota + " emails hoy.");
  }

  // Asegurar columna Merge status
  var mergeStatusCol = ensureMergeStatusColumn();

  // Obtener adjuntos del borrador
  var draftObj = GmailApp.getDrafts().filter(function(d) { return d.getId() === config.draftId; })[0];
  var attachments = draftObj ? draftObj.getMessage().getAttachments() : [];

  var sent = 0;
  var errors = 0;
  var results = [];

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
      results.push({ email: recipient.email, status: "EMAIL_SENT", row: recipient.rowIndex });

      // Pausa entre envios para no saturar (50ms)
      if (i < recipientData.recipients.length - 1) {
        Utilities.sleep(50);
      }

    } catch (err) {
      updateMergeStatus(recipient.rowIndex, mergeStatusCol, "ERROR: " + err.message.substring(0, 50));
      errors++;
      results.push({ email: recipient.email, status: "ERROR", error: err.message, row: recipient.rowIndex });
    }
  }

  // Guardar metadata de campana en PropertiesService
  var campaignId = "campaign_" + new Date().getTime();
  var campaignData = {
    id: campaignId,
    date: new Date().toISOString(),
    draftSubject: draft.subject,
    senderName: config.senderName,
    sent: sent,
    errors: errors,
    total: toSend,
    sheetName: recipientData.sheetName
  };

  var props = PropertiesService.getScriptProperties();
  props.setProperty(campaignId, JSON.stringify(campaignData));

  // Guardar ultima campana
  props.setProperty("lastCampaignId", campaignId);

  return {
    campaignId: campaignId,
    sent: sent,
    errors: errors,
    total: toSend,
    quota: getRemainingQuota(),
    results: results
  };
}

/**
 * Obtiene el progreso del merge (para polling desde el sidebar)
 */
function getMergeProgress() {
  var props = PropertiesService.getScriptProperties();
  var lastId = props.getProperty("lastCampaignId");
  if (!lastId) return null;
  var data = props.getProperty(lastId);
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


// ═══════════════════════════════════════════════════════════════════════════════
// SECCION 3 — Tracking.js
// ═══════════════════════════════════════════════════════════════════════════════

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


// ═══════════════════════════════════════════════════════════════════════════════
// SECCION 4 — Contacts.js
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Contacts.js — Importar contactos desde Google Contacts
 * Usa People API via UrlFetchApp
 */

/**
 * Obtiene contactos de Google Contacts via People API
 * @param {number} maxResults - maximo de contactos (default 200)
 * @returns {Array<Object>} contactos
 */
function fetchGoogleContacts(maxResults) {
  maxResults = maxResults || 200;
  var contacts = [];
  var pageToken = "";

  while (contacts.length < maxResults) {
    var pageSize = Math.min(200, maxResults - contacts.length);
    var url = "https://people.googleapis.com/v1/people/me/connections"
      + "?personFields=names,emailAddresses,phoneNumbers,organizations"
      + "&pageSize=" + pageSize
      + "&sortOrder=LAST_MODIFIED_DESCENDING";

    if (pageToken) url += "&pageToken=" + pageToken;

    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code !== 200) {
      throw new Error("Error People API (" + code + "): " + response.getContentText().substring(0, 100));
    }

    var data = JSON.parse(response.getContentText());
    var connections = data.connections || [];

    for (var i = 0; i < connections.length; i++) {
      var p = connections[i];
      var email = "";
      if (p.emailAddresses && p.emailAddresses.length > 0) {
        email = p.emailAddresses[0].value;
      }
      if (!email) continue; // Solo contactos con email

      contacts.push({
        name: p.names ? p.names[0].displayName : "",
        firstName: p.names ? (p.names[0].givenName || "") : "",
        lastName: p.names ? (p.names[0].familyName || "") : "",
        email: email,
        phone: p.phoneNumbers ? p.phoneNumbers[0].value : "",
        organization: p.organizations ? p.organizations[0].name : "",
        title: p.organizations ? (p.organizations[0].title || "") : ""
      });
    }

    pageToken = data.nextPageToken;
    if (!pageToken) break;
  }

  return contacts;
}

/**
 * Obtiene grupos de contactos (etiquetas)
 * @returns {Array<{id, name, count}>}
 */
function getContactGroups() {
  var url = "https://people.googleapis.com/v1/contactGroups"
    + "?groupFields=name,groupType,memberCount";

  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  var data = JSON.parse(response.getContentText());
  var groups = (data.contactGroups || []).filter(function(g) {
    return g.groupType === "USER_CONTACT_GROUP" || g.name === "myContacts";
  });

  return groups.map(function(g) {
    return {
      id: g.resourceName,
      name: g.formattedName || g.name,
      count: g.memberCount || 0
    };
  });
}

/**
 * Importa contactos a la hoja activa
 * @param {Object} options - {maxResults, overwrite}
 * @returns {Object} {imported, skipped}
 */
function importContactsToSheet(options) {
  options = options || {};
  var contacts = fetchGoogleContacts(options.maxResults || 200);

  if (contacts.length === 0) {
    return { imported: 0, skipped: 0, message: "No se encontraron contactos con email" };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet;

  if (options.overwrite) {
    // Escribir en la hoja activa
    sheet = ss.getActiveSheet();
  } else {
    // Crear nueva hoja "Contactos Google"
    var sheetName = "Contactos Google";
    sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
  }

  // Headers
  var headers = ["Email", "Nombre", "Apellido", "Telefono", "Organizacion", "Cargo"];
  var startRow = 1;

  if (options.overwrite) {
    // Buscar si ya hay headers
    var existing = sheet.getDataRange().getValues();
    if (existing.length > 0) {
      startRow = existing.length + 1;
    } else {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      startRow = 2;
    }
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    startRow = 2;
  }

  // Datos
  var rows = contacts.map(function(c) {
    return [c.email, c.firstName, c.lastName, c.phone, c.organization, c.title];
  });

  if (rows.length > 0) {
    sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
  }

  // Formato
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.autoResizeColumns(1, headers.length);

  return {
    imported: contacts.length,
    skipped: 0,
    sheetName: sheet.getName(),
    message: contacts.length + " contactos importados"
  };
}

/**
 * Preview de cuantos contactos hay disponibles (sin importar)
 */
function getContactsCount() {
  try {
    var url = "https://people.googleapis.com/v1/people/me/connections"
      + "?personFields=emailAddresses&pageSize=1";
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());
    return data.totalPeople || data.totalItems || 0;
  } catch (e) {
    return -1;
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// SECCION 5 — Scheduler.js
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Scheduler.js — Envio programado con time-based triggers
 */

/**
 * Programa una campana para enviar en una fecha/hora futura
 * @param {Object} config - misma config que sendMailMerge
 * @param {string} datetime - fecha ISO string
 * @returns {Object} resultado
 */
function scheduleCampaign(config, datetime) {
  var targetDate = new Date(datetime);
  var now = new Date();

  if (targetDate <= now) {
    throw new Error("La fecha debe ser futura");
  }

  // Guardar config en PropertiesService
  var scheduleId = "scheduled_" + new Date().getTime();
  var props = PropertiesService.getScriptProperties();
  props.setProperty(scheduleId, JSON.stringify({
    id: scheduleId,
    config: config,
    targetDate: targetDate.toISOString(),
    createdAt: now.toISOString(),
    status: "SCHEDULED",
    sheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
    sheetName: SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()
  }));

  // Crear trigger
  var trigger = ScriptApp.newTrigger("executeScheduledCampaign")
    .timeBased()
    .at(targetDate)
    .create();

  // Asociar trigger ID al schedule
  var data = JSON.parse(props.getProperty(scheduleId));
  data.triggerId = trigger.getUniqueId();
  props.setProperty(scheduleId, JSON.stringify(data));

  // Guardar mapping trigger → schedule
  props.setProperty("trigger_" + trigger.getUniqueId(), scheduleId);

  return {
    scheduleId: scheduleId,
    triggerId: trigger.getUniqueId(),
    targetDate: targetDate.toISOString(),
    status: "SCHEDULED"
  };
}

/**
 * Callback ejecutado por el trigger programado
 * @param {Object} e - evento del trigger
 */
function executeScheduledCampaign(e) {
  var props = PropertiesService.getScriptProperties();

  // Buscar el schedule asociado a este trigger
  var triggerId = e.triggerUid;
  var scheduleId = props.getProperty("trigger_" + triggerId);

  if (!scheduleId) {
    console.error("[Scheduler] No se encontro schedule para trigger " + triggerId);
    return;
  }

  var scheduleData = JSON.parse(props.getProperty(scheduleId) || "null");
  if (!scheduleData) {
    console.error("[Scheduler] Datos no encontrados para " + scheduleId);
    return;
  }

  // Actualizar estado
  scheduleData.status = "EXECUTING";
  props.setProperty(scheduleId, JSON.stringify(scheduleData));

  try {
    // Abrir la hoja guardada
    var ss = SpreadsheetApp.openById(scheduleData.sheetId);
    var sheet = ss.getSheetByName(scheduleData.sheetName);

    if (!sheet) {
      throw new Error("Hoja no encontrada: " + scheduleData.sheetName);
    }

    // Activar la hoja (necesario para que getActiveSheet funcione en MailEngine)
    SpreadsheetApp.setActiveSpreadsheet(ss);
    sheet.activate();

    // Ejecutar el merge
    var result = sendMailMerge(scheduleData.config);

    // Actualizar estado con resultado
    scheduleData.status = "COMPLETED";
    scheduleData.result = result;
    scheduleData.completedAt = new Date().toISOString();
    props.setProperty(scheduleId, JSON.stringify(scheduleData));

    console.log("[Scheduler] Campana completada: " + result.sent + " enviados");

  } catch (err) {
    scheduleData.status = "ERROR";
    scheduleData.error = err.message;
    props.setProperty(scheduleId, JSON.stringify(scheduleData));
    console.error("[Scheduler] Error: " + err.message);
  }

  // Limpiar trigger
  deleteTriggerById_(triggerId);
  props.deleteProperty("trigger_" + triggerId);
}

/**
 * Cancela un envio programado
 * @param {string} scheduleId
 * @returns {Object}
 */
function cancelScheduled(scheduleId) {
  var props = PropertiesService.getScriptProperties();
  var data = JSON.parse(props.getProperty(scheduleId) || "null");

  if (!data) throw new Error("Schedule no encontrado");
  if (data.status !== "SCHEDULED") throw new Error("Solo se pueden cancelar envios pendientes");

  // Borrar trigger
  if (data.triggerId) {
    deleteTriggerById_(data.triggerId);
    props.deleteProperty("trigger_" + data.triggerId);
  }

  // Actualizar estado
  data.status = "CANCELLED";
  data.cancelledAt = new Date().toISOString();
  props.setProperty(scheduleId, JSON.stringify(data));

  return { cancelled: true, scheduleId: scheduleId };
}

/**
 * Lista envios programados
 * @returns {Array}
 */
function listScheduledCampaigns() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var scheduled = [];

  var keys = Object.keys(all);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf("scheduled_") === 0) {
      try {
        scheduled.push(JSON.parse(all[keys[i]]));
      } catch (e) { /* skip */ }
    }
  }

  scheduled.sort(function(a, b) { return a.targetDate > b.targetDate ? 1 : -1; });
  return scheduled;
}

/**
 * Elimina un trigger por su unique ID
 */
function deleteTriggerById_(triggerId) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(triggers[i]);
      return;
    }
  }
}

/**
 * Reanuda campanas grandes que se cortaron por limite de 6 min
 * Se usa con un trigger periodico (cada 5 min)
 */
function resumeUnfinishedCampaigns() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();

  var keys = Object.keys(all);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf("scheduled_") === 0) {
      try {
        var data = JSON.parse(all[keys[i]]);
        if (data.status === "EXECUTING") {
          // Re-ejecutar (las filas ya enviadas se skipean por Merge status)
          data.config.skipSent = true;
          executeScheduledCampaign({ triggerUid: data.triggerId });
        }
      } catch (e) { /* skip */ }
    }
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// SECCION 6 — Ecosystem.js
// ═══════════════════════════════════════════════════════════════════════════════

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
