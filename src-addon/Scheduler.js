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

    // Nota: en time-based triggers no hay spreadsheet activo.
    // sendMailMerge usa getActiveSpreadsheet() internamente, asi que
    // inyectamos _sheetId y _sheetName en la config como fallback.
    scheduleData.config._sheetId = scheduleData.sheetId;
    scheduleData.config._sheetName = scheduleData.sheetName;
    scheduleData.config.skipSent = true;

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
