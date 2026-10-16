// Hale's Tools usage log.
// Paste this into the Apps Script editor of a Google Sheet (Extensions > Apps Script),
// then Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
// Copy the web app URL into USAGE_LOG_URL in index.html.
//
// Writes two tabs:
//   Devices   - one row per phone/computer. Type the person's name in the Name
//               column; every later event from that device is logged under it.
//   Usage Log - one row per event. Devices with no name yet are flagged red.

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action !== 'logUsage') return jsonOut({ status: 'ignored' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var devices = getOrCreateSheet(ss, 'Devices', ['Device ID', 'Name (fill in)', 'Device', 'First Seen', 'Last Seen', 'Times Used']);
    var log = getOrCreateSheet(ss, 'Usage Log', ['Time', 'Name', 'Action', 'Details', 'Device', 'Device ID', 'Flag']);

    var now = new Date();
    var deviceId = String(data.deviceId || '');
    var device = String(data.device || '');

    var ids = devices.getLastRow() > 1 ? devices.getRange(2, 1, devices.getLastRow() - 1, 1).getValues() : [];
    var row = -1;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === deviceId) { row = i + 2; break; }
    }

    var name = '';
    var flag = '';
    if (row === -1) {
      devices.appendRow([deviceId, '', device, now, now, 1]);
      devices.getRange(devices.getLastRow(), 1, 1, 6).setBackground('#fde2e2');
      flag = 'NEW DEVICE';
    } else {
      name = String(devices.getRange(row, 2).getValue()).trim();
      var count = Number(devices.getRange(row, 6).getValue()) || 0;
      devices.getRange(row, 5, 1, 2).setValues([[now, count + 1]]);
      devices.getRange(row, 1, 1, 6).setBackground(name ? null : '#fde2e2');
      if (!name) flag = 'Unknown device';
    }

    log.appendRow([now, name, data.event || '', data.details || '', device, deviceId, flag]);
    if (flag) log.getRange(log.getLastRow(), 1, 1, 7).setBackground('#fde2e2');
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
