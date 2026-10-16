// Hale's Tools usage log.
// Paste this into the Apps Script editor of a Google Sheet (Extensions > Apps Script),
// then Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
// Copy the web app URL into USAGE_LOG_URL in index.html.
//
// Writes two tabs:
//   Usage Log - one row per event, with a Flag column for anything suspicious
//   Devices   - every name + device pair ever seen, with first/last seen times

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action !== 'logUsage') return jsonOut({ status: 'ignored' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var log = getOrCreateSheet(ss, 'Usage Log', ['Time', 'Name', 'Action', 'Details', 'Device', 'Device ID', 'Flag']);
    var devices = getOrCreateSheet(ss, 'Devices', ['Name', 'Device ID', 'Device', 'First Seen', 'Last Seen']);

    var now = new Date();
    var name = String(data.name || '').trim();
    var nameKey = name.toLowerCase();
    var deviceId = String(data.deviceId || '');
    var device = String(data.device || '');

    var rows = devices.getLastRow() > 1 ? devices.getRange(2, 1, devices.getLastRow() - 1, 2).getValues() : [];
    var pairRow = -1;
    var nameSeenElsewhere = false;
    var otherNamesOnDevice = [];
    for (var i = 0; i < rows.length; i++) {
      var rowName = String(rows[i][0]);
      var rowDevice = String(rows[i][1]);
      if (rowName.toLowerCase() === nameKey && rowDevice === deviceId) pairRow = i + 2;
      else if (rowName.toLowerCase() === nameKey) nameSeenElsewhere = true;
      else if (rowDevice === deviceId && otherNamesOnDevice.indexOf(rowName) === -1) otherNamesOnDevice.push(rowName);
    }

    var flags = [];
    if (pairRow === -1 && nameSeenElsewhere) flags.push('NEW DEVICE for this name');
    if (otherNamesOnDevice.length) flags.push('Device also used as: ' + otherNamesOnDevice.join(', '));

    if (pairRow === -1) {
      devices.appendRow([name, deviceId, device, now, now]);
    } else {
      devices.getRange(pairRow, 5).setValue(now);
    }

    log.appendRow([now, name, data.event || '', data.details || '', device, deviceId, flags.join('; ')]);
    if (flags.length) log.getRange(log.getLastRow(), 1, 1, 7).setBackground('#fde2e2');
  } finally {
    lock.releaseLock();
  }

  return jsonOut({ status: 'ok' });
}

function getOrCreateSheet(ss, title, headers) {
  var sheet = ss.getSheetByName(title);
  if (!sheet) {
    sheet = ss.insertSheet(title);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
