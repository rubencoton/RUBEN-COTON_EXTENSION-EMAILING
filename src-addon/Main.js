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
