/**
 * Deployer.js — Despliegue masivo de EMAILING RUBEN COTON
 * Recorre una carpeta de Drive, encuentra hojas de calculo,
 * y crea un proyecto Apps Script vinculado con el codigo de la extension en cada una.
 */

/** ID de la carpeta objetivo */
var TARGET_FOLDER_ID = "1A7agk072QZIS_6HVB3SmX8j2pERh4fW3";

/**
 * Lista todas las hojas de calculo en la carpeta (y subcarpetas)
 * @returns {Array<{id, name, url}>}
 */
function listSpreadsheetsInFolder() {
  var folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  var results = [];
  scanFolder_(folder, results);
  Logger.log("Total hojas encontradas: " + results.length);
  results.forEach(function(s) {
    Logger.log("  " + s.name + " — " + s.id);
  });
  return results;
}

function scanFolder_(folder, results) {
  // Buscar hojas de calculo en esta carpeta
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    var file = files.next();
    results.push({
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl()
    });
  }

  // Recursion en subcarpetas
  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    scanFolder_(subfolders.next(), results);
  }
}

/**
 * Codigo GS de la extension (compactado) que se inyecta en cada hoja.
 * Contiene: onOpen menu + showMailMergeSidebar + funciones esenciales del motor.
 */
function getExtensionCode_() {
  return '/**\n'
    + ' * EMAILING RUBEN COTON — Extension Mail Merge\n'
    + ' * Instalado automaticamente desde el proyecto central.\n'
    + ' */\n\n'
    + 'function onOpen(e) {\n'
    + '  SpreadsheetApp.getUi()\n'
    + '    .createAddonMenu()\n'
    + '    .addItem("Iniciar Mail Merge", "showMailMergeSidebar")\n'
    + '    .addSeparator()\n'
    + '    .addItem("Ver informe de seguimiento", "showTrackingReportSidebar")\n'
    + '    .addItem("Estado de cuota", "showQuotaAlert")\n'
    + '    .addToUi();\n'
    + '}\n\n'
    + 'function onInstall(e) { onOpen(e); }\n\n'
    + 'function showMailMergeSidebar() {\n'
    + '  var html = HtmlService.createHtmlOutputFromFile("MailMerge")\n'
    + '    .setTitle("EMAILING RUBEN COTON").setWidth(340);\n'
    + '  SpreadsheetApp.getUi().showSidebar(html);\n'
    + '}\n\n'
    + 'function showTrackingReportSidebar() {\n'
    + '  var html = HtmlService.createHtmlOutputFromFile("TrackingReport")\n'
    + '    .setTitle("Informe de Seguimiento").setWidth(340);\n'
    + '  SpreadsheetApp.getUi().showSidebar(html);\n'
    + '}\n\n'
    + 'function showQuotaAlert() {\n'
    + '  var q = MailApp.getRemainingDailyQuota();\n'
    + '  SpreadsheetApp.getUi().alert("Cuota de envio", "Emails restantes hoy: " + q, SpreadsheetApp.getUi().ButtonSet.OK);\n'
    + '}\n';
}

/**
 * Despliega la extension en TODAS las hojas de la carpeta
 * Usa la Apps Script API para crear proyectos vinculados
 */
function deployToAllSheets() {
  var sheets = listSpreadsheetsInFolder();
  var token = ScriptApp.getOAuthToken();
  var results = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    Logger.log("\n--- Procesando: " + sheet.name + " ---");

    try {
      // 1. Verificar si ya tiene un proyecto Apps Script vinculado con nuestra extension
      var existing = checkExistingProject_(sheet.id, token);
      if (existing) {
        Logger.log("  Ya tiene extension instalada, actualizando...");
        updateProjectContent_(existing, token);
        results.push({ name: sheet.name, id: sheet.id, status: "UPDATED", scriptId: existing });
        continue;
      }

      // 2. Crear nuevo proyecto vinculado
      var scriptId = createBoundProject_(sheet.id, sheet.name, token);
      if (scriptId) {
        // 3. Subir codigo
        updateProjectContent_(scriptId, token);
        results.push({ name: sheet.name, id: sheet.id, status: "CREATED", scriptId: scriptId });
        Logger.log("  OK: Proyecto creado " + scriptId);
      }
    } catch (err) {
      Logger.log("  ERROR: " + err.message);
      results.push({ name: sheet.name, id: sheet.id, status: "ERROR", error: err.message });
    }

    // Pausa entre operaciones
    Utilities.sleep(1000);
  }

  // Resumen
  Logger.log("\n=== RESUMEN ===");
  var created = results.filter(function(r) { return r.status === "CREATED"; }).length;
  var updated = results.filter(function(r) { return r.status === "UPDATED"; }).length;
  var errors = results.filter(function(r) { return r.status === "ERROR"; }).length;
  Logger.log("Creados: " + created + " | Actualizados: " + updated + " | Errores: " + errors);

  return results;
}

/**
 * Crea un proyecto Apps Script vinculado a una hoja de calculo
 */
function createBoundProject_(spreadsheetId, title, token) {
  var url = "https://script.googleapis.com/v1/projects";
  var payload = {
    title: "EMAILING RUBEN COTON — " + title,
    parentId: spreadsheetId
  };

  var response = UrlFetchApp.fetch(url, {
    method: "POST",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = JSON.parse(response.getContentText());

  if (code !== 200) {
    throw new Error("API error " + code + ": " + (body.error ? body.error.message : response.getContentText().substring(0, 100)));
  }

  return body.scriptId;
}

/**
 * Sube el codigo de la extension a un proyecto
 */
function updateProjectContent_(scriptId, token) {
  var url = "https://script.googleapis.com/v1/projects/" + scriptId + "/content";

  // Leer los HTML del proyecto central
  var mailMergeHtml = getMailMergeHtml_();
  var trackingReportHtml = getTrackingReportHtml_();

  var payload = {
    files: [
      {
        name: "appsscript",
        type: "JSON",
        source: JSON.stringify({
          timeZone: "Europe/Madrid",
          exceptionLogging: "STACKDRIVER",
          runtimeVersion: "V8",
          oauthScopes: [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/gmail.readonly",
            "https://www.googleapis.com/auth/gmail.send",
            "https://www.googleapis.com/auth/gmail.compose",
            "https://www.googleapis.com/auth/script.external_request",
            "https://www.googleapis.com/auth/script.send_mail",
            "https://www.googleapis.com/auth/userinfo.email"
          ]
        })
      },
      {
        name: "Main",
        type: "SERVER_JS",
        source: getExtensionCode_()
      },
      {
        name: "MailEngine",
        type: "SERVER_JS",
        source: getMailEngineCode_()
      },
      {
        name: "Tracking",
        type: "SERVER_JS",
        source: getTrackingCode_()
      },
      {
        name: "MailMerge",
        type: "HTML",
        source: mailMergeHtml
      },
      {
        name: "TrackingReport",
        type: "HTML",
        source: trackingReportHtml
      }
    ]
  };

  var response = UrlFetchApp.fetch(url, {
    method: "PUT",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    var body = JSON.parse(response.getContentText());
    throw new Error("Update error " + code + ": " + (body.error ? body.error.message : "unknown"));
  }

  return true;
}

/**
 * Busca si una hoja ya tiene un proyecto con nuestra extension
 */
function checkExistingProject_(spreadsheetId, token) {
  // No hay forma directa de listar proyectos vinculados a un archivo via API.
  // Devolvemos null para crear siempre nuevo.
  return null;
}

/**
 * Codigo del motor de mail merge (version compacta para bound scripts)
 */
function getMailEngineCode_() {
  // Motor esencial: getDrafts, getAliases, getRecipientsFromSheet, replaceVariables,
  // ensureMergeStatusColumn, sendMailMerge, sendTestEmail, getRemainingQuota
  var code = '';
  code += 'var MERGE_EMAIL_HEADERS = ["email","correo","correo electrónico","correo electronico","e-mail","email address","mail"];\n\n';

  code += 'function getDrafts() {\n';
  code += '  var drafts = GmailApp.getDrafts();\n';
  code += '  return drafts.map(function(d) {\n';
  code += '    var msg = d.getMessage();\n';
  code += '    return { id: d.getId(), subject: msg.getSubject() || "(sin asunto)",\n';
  code += '      snippet: msg.getPlainBody().substring(0,80),\n';
  code += '      hasAttachments: msg.getAttachments().length > 0 };\n';
  code += '  });\n}\n\n';

  code += 'function getDraftById(draftId) {\n';
  code += '  var drafts = GmailApp.getDrafts();\n';
  code += '  for (var i=0;i<drafts.length;i++) { if (drafts[i].getId()===draftId) {\n';
  code += '    var msg=drafts[i].getMessage();\n';
  code += '    return {id:draftId,subject:msg.getSubject(),htmlBody:msg.getBody(),\n';
  code += '      plainBody:msg.getPlainBody(),attachments:msg.getAttachments()};\n';
  code += '  }} throw new Error("Borrador no encontrado");\n}\n\n';

  code += 'function getAliases() {\n';
  code += '  try { var r=UrlFetchApp.fetch("https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs",\n';
  code += '    {headers:{Authorization:"Bearer "+ScriptApp.getOAuthToken()}});\n';
  code += '    var d=JSON.parse(r.getContentText());\n';
  code += '    return (d.sendAs||[]).map(function(a){return{email:a.sendAsEmail,\n';
  code += '      displayName:a.displayName||"",isDefault:a.isDefault||false};});\n';
  code += '  } catch(e){return [{email:Session.getActiveUser().getEmail(),displayName:"",isDefault:true}];}\n}\n\n';

  code += 'function getRecipientsFromSheet(options) {\n';
  code += '  options=options||{};\n';
  code += '  var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();\n';
  code += '  var data=sheet.getDataRange().getValues();\n';
  code += '  if(data.length<2) return {headers:[],recipients:[],emailColIndex:-1,mergeStatusColIndex:-1,totalRows:0};\n';
  code += '  var headers=data[0].map(function(h){return String(h||"").trim();});\n';
  code += '  var headersLower=headers.map(function(h){return h.toLowerCase();});\n';
  code += '  var emailColIndex=-1;\n';
  code += '  for(var e=0;e<MERGE_EMAIL_HEADERS.length;e++){\n';
  code += '    var idx=headersLower.indexOf(MERGE_EMAIL_HEADERS[e]);\n';
  code += '    if(idx!==-1){emailColIndex=idx;break;}}\n';
  code += '  if(emailColIndex===-1) throw new Error("No se encontro columna de email");\n';
  code += '  var msIdx=headersLower.indexOf("merge status");\n';
  code += '  var recipients=[];\n';
  code += '  for(var r=1;r<data.length;r++){\n';
  code += '    var email=String(data[r][emailColIndex]||"").trim();\n';
  code += '    if(!email||email.indexOf("@")===-1) continue;\n';
  code += '    if(options.skipSent&&msIdx!==-1){\n';
  code += '      var st=String(data[r][msIdx]||"").trim().toUpperCase();\n';
  code += '      if(st==="EMAIL_SENT"||st==="OPENED"||st==="CLICKED"||st==="RESPONDED") continue;}\n';
  code += '    var mv={}; for(var c=0;c<headers.length;c++) mv[headers[c]]=String(data[r][c]||"");\n';
  code += '    recipients.push({rowIndex:r+1,email:email,mergeVars:mv});}\n';
  code += '  return {headers:headers,recipients:recipients,emailColIndex:emailColIndex,\n';
  code += '    mergeStatusColIndex:msIdx,totalRows:data.length-1,sheetName:sheet.getName()};\n}\n\n';

  code += 'function getSheetColumns() {\n';
  code += '  var data=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();\n';
  code += '  if(data.length<1) return [];\n';
  code += '  return data[0].map(function(h){return String(h||"").trim();}).filter(Boolean);\n}\n\n';

  code += 'function replaceVariables(template,mergeVars) {\n';
  code += '  if(!template) return "";\n';
  code += '  return template.replace(/\\{\\{([^}]+)\\}\\}/g,function(m,v){\n';
  code += '    var k=v.trim(); return mergeVars.hasOwnProperty(k)?mergeVars[k]:"";});\n}\n\n';

  code += 'function ensureMergeStatusColumn() {\n';
  code += '  var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();\n';
  code += '  var headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];\n';
  code += '  var idx=headers.map(function(h){return String(h||"").toLowerCase().trim();}).indexOf("merge status");\n';
  code += '  if(idx!==-1) return idx+1;\n';
  code += '  var newCol=sheet.getLastColumn()+1;\n';
  code += '  sheet.getRange(1,newCol).setValue("Merge status"); return newCol;\n}\n\n';

  code += 'function getRemainingQuota(){return MailApp.getRemainingDailyQuota();}\n\n';

  code += 'function sendTestEmail(config) {\n';
  code += '  var draft=getDraftById(config.draftId);\n';
  code += '  var rd=getRecipientsFromSheet({});\n';
  code += '  if(rd.recipients.length===0) throw new Error("No hay destinatarios");\n';
  code += '  var fr=rd.recipients[0];\n';
  code += '  var subj=replaceVariables(draft.subject,fr.mergeVars);\n';
  code += '  var html=replaceVariables(draft.htmlBody,fr.mergeVars);\n';
  code += '  var myEmail=Session.getActiveUser().getEmail();\n';
  code += '  var opts={htmlBody:html,name:config.senderName||"EMAILING RUBEN COTON"};\n';
  code += '  if(config.alias) opts.from=config.alias;\n';
  code += '  if(draft.attachments&&draft.attachments.length>0) opts.attachments=draft.attachments;\n';
  code += '  GmailApp.sendEmail(myEmail,"[TEST] "+subj,"",opts);\n';
  code += '  return {sent:true,to:myEmail,subject:"[TEST] "+subj};\n}\n\n';

  code += 'function sendMailMerge(config) {\n';
  code += '  var draft=getDraftById(config.draftId);\n';
  code += '  var rd=getRecipientsFromSheet({filterColumn:config.filterColumn,\n';
  code += '    filterValue:config.filterValue,skipSent:config.skipSent!==false});\n';
  code += '  if(rd.recipients.length===0) return {sent:0,errors:0,total:0,message:"No hay destinatarios"};\n';
  code += '  var quota=getRemainingQuota();\n';
  code += '  if(rd.recipients.length>quota) throw new Error("Cuota insuficiente: "+rd.recipients.length+" necesarios, "+quota+" disponibles");\n';
  code += '  var msCol=ensureMergeStatusColumn();\n';
  code += '  var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();\n';
  code += '  var sent=0,errors=0;\n';
  code += '  for(var i=0;i<rd.recipients.length;i++){\n';
  code += '    var r=rd.recipients[i];\n';
  code += '    try{\n';
  code += '      var subj=replaceVariables(draft.subject,r.mergeVars);\n';
  code += '      var html=replaceVariables(draft.htmlBody,r.mergeVars);\n';
  code += '      var opts={htmlBody:html,name:config.senderName||"EMAILING RUBEN COTON"};\n';
  code += '      if(config.alias) opts.from=config.alias;\n';
  code += '      if(config.cc) opts.cc=config.cc;\n';
  code += '      if(config.bcc) opts.bcc=config.bcc;\n';
  code += '      if(draft.attachments&&draft.attachments.length>0) opts.attachments=draft.attachments;\n';
  code += '      GmailApp.sendEmail(r.email,subj,"",opts);\n';
  code += '      sheet.getRange(r.rowIndex,msCol).setValue("EMAIL_SENT");\n';
  code += '      sent++; if(i<rd.recipients.length-1) Utilities.sleep(50);\n';
  code += '    }catch(err){\n';
  code += '      sheet.getRange(r.rowIndex,msCol).setValue("ERROR: "+err.message.substring(0,50));\n';
  code += '      errors++;}}\n';
  code += '  SpreadsheetApp.flush();\n';
  code += '  return {sent:sent,errors:errors,total:rd.recipients.length,quota:getRemainingQuota()};\n}\n\n';

  code += 'function getCampaignStats() {\n';
  code += '  var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();\n';
  code += '  var data=sheet.getDataRange().getValues();\n';
  code += '  if(data.length<2) return {total:0,sent:0,opened:0,clicked:0,responded:0,bounced:0,errors:0,pending:0};\n';
  code += '  var headers=data[0].map(function(h){return String(h||"").toLowerCase().trim();});\n';
  code += '  var msIdx=headers.indexOf("merge status");\n';
  code += '  if(msIdx===-1) return {total:data.length-1,sent:0,pending:data.length-1,noColumn:true,\n';
  code += '    deliveryRate:0,openRate:0,clickRate:0,responseRate:0,bounceRate:0};\n';
  code += '  var s={total:0,sent:0,opened:0,clicked:0,responded:0,bounced:0,unsubscribed:0,errors:0,pending:0};\n';
  code += '  var eIdx=-1; var ec=["email","correo","e-mail"];\n';
  code += '  for(var e=0;e<ec.length;e++){var ei=headers.indexOf(ec[e]);if(ei!==-1){eIdx=ei;break;}}\n';
  code += '  for(var r=1;r<data.length;r++){\n';
  code += '    if(eIdx!==-1&&String(data[r][eIdx]||"").indexOf("@")===-1) continue;\n';
  code += '    s.total++; var st=String(data[r][msIdx]||"").trim().toUpperCase();\n';
  code += '    if(!st){s.pending++;continue;}\n';
  code += '    if(st==="EMAIL_SENT")s.sent++;else if(st==="OPENED")s.opened++;else if(st==="CLICKED")s.clicked++;\n';
  code += '    else if(st==="RESPONDED")s.responded++;else if(st==="BOUNCED")s.bounced++;\n';
  code += '    else if(st.indexOf("ERROR")===0)s.errors++;else s.sent++;}\n';
  code += '  var del=s.sent+s.opened+s.clicked+s.responded;\n';
  code += '  s.deliveryRate=s.total>0?Math.round(del/s.total*100):0;\n';
  code += '  s.openRate=del>0?Math.round((s.opened+s.clicked+s.responded)/del*100):0;\n';
  code += '  s.clickRate=del>0?Math.round((s.clicked+s.responded)/del*100):0;\n';
  code += '  s.responseRate=del>0?Math.round(s.responded/del*100):0;\n';
  code += '  s.bounceRate=s.total>0?Math.round(s.bounced/s.total*100):0;\n';
  code += '  return s;\n}\n\n';

  code += 'function getDetailedTracking(filterStatus) {\n';
  code += '  var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();\n';
  code += '  var data=sheet.getDataRange().getValues();\n';
  code += '  if(data.length<2) return [];\n';
  code += '  var headers=data[0].map(function(h){return String(h||"").toLowerCase().trim();});\n';
  code += '  var msIdx=headers.indexOf("merge status"); if(msIdx===-1) return [];\n';
  code += '  var eIdx=-1; var ec=["email","correo","e-mail"];\n';
  code += '  for(var e=0;e<ec.length;e++){var ei=headers.indexOf(ec[e]);if(ei!==-1){eIdx=ei;break;}}\n';
  code += '  if(eIdx===-1) return [];\n';
  code += '  var results=[];\n';
  code += '  for(var r=1;r<data.length;r++){\n';
  code += '    var email=String(data[r][eIdx]||"").trim();\n';
  code += '    if(!email||email.indexOf("@")===-1) continue;\n';
  code += '    var status=String(data[r][msIdx]||"").trim();\n';
  code += '    if(filterStatus&&status.toUpperCase()!==filterStatus.toUpperCase()) continue;\n';
  code += '    results.push({row:r+1,email:email,name:"",status:status||"(pendiente)"});}\n';
  code += '  return results;\n}\n';

  return code;
}

/**
 * HTML del sidebar MailMerge (leido del archivo del proyecto)
 */
function getMailMergeHtml_() {
  try {
    return HtmlService.createHtmlOutputFromFile("MailMerge").getContent();
  } catch(e) {
    // Fallback: HTML minimo
    return getMailMergeHtmlFallback_();
  }
}

function getTrackingReportHtml_() {
  try {
    return HtmlService.createHtmlOutputFromFile("TrackingReport").getContent();
  } catch(e) {
    return getTrackingReportHtmlFallback_();
  }
}

function getMailMergeHtmlFallback_() {
  return '<!DOCTYPE html><html><head><base target="_top"><style>'
    + 'body{font-family:Arial,sans-serif;padding:16px;font-size:13px;}'
    + '.section{background:#fff;border-radius:8px;padding:12px;margin-bottom:12px;box-shadow:0 1px 2px rgba(0,0,0,0.08);}'
    + 'label{display:block;font-size:12px;color:#5f6368;margin:8px 0 4px;}'
    + 'select,input{width:100%;padding:8px;border:1px solid #dadce0;border-radius:6px;font-size:13px;}'
    + '.btn{display:block;width:100%;padding:10px;border:none;border-radius:6px;font-size:13px;font-weight:600;cursor:pointer;margin-top:8px;}'
    + '.btn-primary{background:#f60;color:#fff;}.btn-secondary{background:#e8eaed;color:#202124;}'
    + '.stat-box{background:#f1f3f4;border-radius:6px;padding:10px;text-align:center;margin:8px 0;}'
    + '.stat-number{font-size:24px;font-weight:700;color:#f60;}'
    + '</style></head><body>'
    + '<h2 style="text-align:center;">EMAILING RUBEN COTON</h2>'
    + '<div class="section"><label>Nombre remitente</label>'
    + '<input type="text" id="senderName" value="RUBEN COTON">'
    + '<label>Borrador Gmail</label>'
    + '<select id="draftSelect"><option value="">Cargando...</option></select>'
    + '<button class="btn btn-secondary" onclick="loadDrafts()" style="margin-top:4px;">Refrescar</button></div>'
    + '<div class="section"><div class="stat-box"><div class="stat-number" id="recipientCount">-</div>'
    + '<div style="font-size:11px;color:#5f6368;">emails encontrados</div></div>'
    + '<label><input type="checkbox" id="skipSent" checked> Saltar ya enviados</label></div>'
    + '<div class="section">'
    + '<button class="btn btn-secondary" onclick="sendTest()" id="btnTest">Enviar prueba</button>'
    + '<button class="btn btn-primary" onclick="sendMerge()" id="btnSend" disabled>Enviar emails</button></div>'
    + '<div id="statusMsg" style="text-align:center;font-size:11px;color:#5f6368;padding:8px;"></div>'
    + '<script>'
    + 'function loadDrafts(){google.script.run.withSuccessHandler(function(d){'
    + 'var s=document.getElementById("draftSelect");s.innerHTML="<option value=\\"\\">-- Selecciona --</option>";'
    + 'd.forEach(function(x){var o=document.createElement("option");o.value=x.id;o.textContent=x.subject;s.appendChild(o);});'
    + '}).getDrafts();}'
    + 'function loadRecipients(){google.script.run.withSuccessHandler(function(d){'
    + 'document.getElementById("recipientCount").textContent=d.recipients.length;'
    + 'document.getElementById("btnSend").disabled=!document.getElementById("draftSelect").value||d.recipients.length===0;'
    + '}).getRecipientsFromSheet({skipSent:document.getElementById("skipSent").checked});}'
    + 'document.getElementById("draftSelect").addEventListener("change",function(){loadRecipients();});'
    + 'document.getElementById("skipSent").addEventListener("change",loadRecipients);'
    + 'function sendTest(){var did=document.getElementById("draftSelect").value;if(!did)return;'
    + 'document.getElementById("statusMsg").textContent="Enviando prueba...";'
    + 'google.script.run.withSuccessHandler(function(r){document.getElementById("statusMsg").textContent="Prueba enviada a "+r.to;'
    + '}).withFailureHandler(function(e){document.getElementById("statusMsg").textContent="Error: "+e.message;'
    + '}).sendTestEmail({draftId:did,senderName:document.getElementById("senderName").value});}'
    + 'function sendMerge(){var did=document.getElementById("draftSelect").value;if(!did)return;'
    + 'if(!confirm("Enviar emails?"))return;document.getElementById("statusMsg").textContent="Enviando...";'
    + 'document.getElementById("btnSend").disabled=true;'
    + 'google.script.run.withSuccessHandler(function(r){document.getElementById("statusMsg").textContent=r.sent+" enviados, "+r.errors+" errores";'
    + 'document.getElementById("btnSend").disabled=false;loadRecipients();'
    + '}).withFailureHandler(function(e){document.getElementById("statusMsg").textContent="Error: "+e.message;'
    + 'document.getElementById("btnSend").disabled=false;}).sendMailMerge({draftId:did,'
    + 'senderName:document.getElementById("senderName").value,skipSent:document.getElementById("skipSent").checked});}'
    + 'window.addEventListener("load",function(){loadDrafts();loadRecipients();});'
    + '</script></body></html>';
}

function getTrackingReportHtmlFallback_() {
  return '<!DOCTYPE html><html><head><base target="_top"><style>'
    + 'body{font-family:Arial,sans-serif;padding:16px;font-size:13px;}'
    + '.stat{display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #eee;}'
    + '.btn{display:block;width:100%;padding:10px;border:none;border-radius:6px;font-size:13px;cursor:pointer;margin-top:8px;background:#e8eaed;}'
    + '</style></head><body>'
    + '<h2 style="text-align:center;">Informe de Seguimiento</h2>'
    + '<div id="stats">Cargando...</div>'
    + '<button class="btn" onclick="load()">Actualizar</button>'
    + '<script>'
    + 'function load(){google.script.run.withSuccessHandler(function(s){'
    + 'document.getElementById("stats").innerHTML='
    + '"<div class=\\"stat\\"><span>Total</span><b>"+s.total+"</b></div>"'
    + '+"<div class=\\"stat\\"><span>Enviados</span><b>"+(s.sent+s.opened+s.clicked+s.responded)+"</b></div>"'
    + '+"<div class=\\"stat\\"><span>Pendientes</span><b>"+s.pending+"</b></div>"'
    + '+"<div class=\\"stat\\"><span>Errores</span><b>"+s.errors+"</b></div>"'
    + '+"<div class=\\"stat\\"><span>Rebotados</span><b>"+s.bounced+"</b></div>"'
    + '+"<div class=\\"stat\\"><span>Tasa entrega</span><b>"+s.deliveryRate+"%</b></div>";'
    + '}).getCampaignStats();}window.addEventListener("load",load);'
    + '</script></body></html>';
}
