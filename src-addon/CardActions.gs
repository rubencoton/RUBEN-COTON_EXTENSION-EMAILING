/**
 * CardActions.gs — Acciones de los botones del Card UI
 */

function cardSyncContacts(e) {
  var total = countAllCrmEmails();
  var notification = CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification()
        .setText("Total emails en CRMs: " + total)
    )
    .build();
  return notification;
}

function cardListSheets(e) {
  var card = CardService.newCardBuilder();
  card.setHeader(
    CardService.newCardHeader()
      .setTitle("Hojas CRM")
      .setSubtitle("Ecosistema manager@rubencoton.com")
  );

  // CRMs
  var crmSection = CardService.newCardSection().setHeader("CRMs (" + Object.keys(CRM_SHEETS).length + ")");
  var keys = Object.keys(CRM_SHEETS);
  for (var i = 0; i < keys.length; i++) {
    crmSection.addWidget(
      CardService.newDecoratedText()
        .setText(keys[i])
        .setBottomLabel(CRM_SHEETS[keys[i]].substring(0, 20) + "...")
        .setOpenLink(
          CardService.newOpenLink()
            .setUrl("https://docs.google.com/spreadsheets/d/" + CRM_SHEETS[keys[i]])
        )
    );
  }
  card.addSection(crmSection);

  // Ecosistema
  var ecoSection = CardService.newCardSection().setHeader("Otras hojas (" + Object.keys(ECOSYSTEM_SHEETS).length + ")");
  var ecoKeys = Object.keys(ECOSYSTEM_SHEETS);
  for (var j = 0; j < ecoKeys.length; j++) {
    ecoSection.addWidget(
      CardService.newDecoratedText()
        .setText(ecoKeys[j])
        .setOpenLink(
          CardService.newOpenLink()
            .setUrl("https://docs.google.com/spreadsheets/d/" + ECOSYSTEM_SHEETS[ecoKeys[j]])
        )
    );
  }
  card.addSection(ecoSection);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

function cardSyncCurrentSheet(e) {
  var count = countEmailsInActiveSheet_();
  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification()
        .setText("Emails en hoja activa: " + count)
    )
    .build();
}
