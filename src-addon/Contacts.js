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
