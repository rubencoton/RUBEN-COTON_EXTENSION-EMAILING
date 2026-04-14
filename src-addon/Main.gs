/**
 * RUBEN COTON Emailing — Extension Google Workspace
 * Cuenta: manager@rubencoton.com
 *
 * Add-on que conecta con todo el ecosistema Google:
 * - Sheets: lectura/escritura de CRMs
 * - Drive: acceso a archivos y carpetas
 * - Gmail: envio y lectura de correos
 * - Calendar: eventos y disponibilidad
 * - Contacts: directorio de contactos
 */

/* ═══════════════════════════════════════════
   TRIGGERS DEL ADD-ON
   ═══════════════════════════════════════════ */

/**
 * Homepage comun del add-on (aparece en cualquier app Google)
 */
function onHomepage(e) {
  return buildHomeCard_();
}

/**
 * Homepage especifica para Google Sheets
 */
function onSheetsHomepage(e) {
  return buildSheetsCard_();
}

/**
 * Menu contextual al abrir una hoja (legacy, para editor add-on)
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Abrir panel", "showSidebar")
    .addSeparator()
    .addItem("Sync contactos", "syncContactsFromSheet")
    .addItem("Ver ecosistema", "showEcosystemInfo")
    .addItem("Estado cuenta", "showAccountStatus")
    .addToUi();
}

/**
 * Al instalar el add-on
 */
function onInstall(e) {
  onOpen(e);
}

/* ═══════════════════════════════════════════
   SIDEBAR
   ═══════════════════════════════════════════ */

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("RUBEN COTON Emailing")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ═══════════════════════════════════════════
   CARDS UI (Add-on v2)
   ═══════════════════════════════════════════ */

function buildHomeCard_() {
  var card = CardService.newCardBuilder();
  card.setHeader(
    CardService.newCardHeader()
      .setTitle("RUBEN COTON Emailing")
      .setSubtitle("manager@rubencoton.com")
  );

  var section = CardService.newCardSection()
    .setHeader("Ecosistema Google");

  var stats = getEcosystemStats_();

  section.addWidget(
    CardService.newDecoratedText()
      .setText("Hojas CRM: " + stats.sheets)
      .setStartIcon(CardService.newIconImage().setIconUrl("https://www.gstatic.com/images/branding/product/2x/sheets_2020q4_48dp.png"))
  );
  section.addWidget(
    CardService.newDecoratedText()
      .setText("Archivos Drive: " + stats.driveFiles)
      .setStartIcon(CardService.newIconImage().setIconUrl("https://www.gstatic.com/images/branding/product/2x/drive_2020q4_48dp.png"))
  );
  section.addWidget(
    CardService.newDecoratedText()
      .setText("Emails no leidos: " + stats.unreadEmails)
      .setStartIcon(CardService.newIconImage().setIconUrl("https://www.gstatic.com/images/branding/product/2x/gmail_2020q4_48dp.png"))
  );
  section.addWidget(
    CardService.newDecoratedText()
      .setText("Eventos hoy: " + stats.todayEvents)
      .setStartIcon(CardService.newIconImage().setIconUrl("https://www.gstatic.com/images/branding/product/2x/calendar_2020q4_48dp.png"))
  );

  card.addSection(section);

  // Acciones
  var actions = CardService.newCardSection().setHeader("Acciones");
  actions.addWidget(
    CardService.newTextButton()
      .setText("Sync contactos")
      .setOnClickAction(CardService.newAction().setFunctionName("cardSyncContacts"))
  );
  actions.addWidget(
    CardService.newTextButton()
      .setText("Ver todas las hojas")
      .setOnClickAction(CardService.newAction().setFunctionName("cardListSheets"))
  );
  card.addSection(actions);

  return card.build();
}

function buildSheetsCard_() {
  var card = CardService.newCardBuilder();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ss ? ss.getName() : "(sin hoja activa)";

  card.setHeader(
    CardService.newCardHeader()
      .setTitle("RUBEN COTON Emailing")
      .setSubtitle(sheetName)
  );

  var section = CardService.newCardSection().setHeader("Hoja actual");

  if (ss) {
    var sheets = ss.getSheets();
    section.addWidget(
      CardService.newDecoratedText().setText("Tabs: " + sheets.length)
    );

    // Contar emails en la hoja activa
    var emailCount = countEmailsInActiveSheet_();
    section.addWidget(
      CardService.newDecoratedText().setText("Emails detectados: " + emailCount)
    );
  }

  section.addWidget(
    CardService.newTextButton()
      .setText("Sync esta hoja")
      .setOnClickAction(CardService.newAction().setFunctionName("cardSyncCurrentSheet"))
  );

  card.addSection(section);
  return card.build();
}
