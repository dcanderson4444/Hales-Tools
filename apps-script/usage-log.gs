// Hale's Tools usage log.
// Paste this into the Apps Script editor of a Google Sheet (Extensions > Apps Script),
// then Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
// Copy the web app URL into USAGE_LOG_URL in index.html.
//
// Writes two tabs:
//   Devices   - one row per phone/computer. Type the person's name in the Name
//               column; every later event from that device is logged under it.
//   Usage Log - one row per event. Devices with no name yet are flagged red.
//               Copied service notes include the full note text.
//
// Typing (or changing) a name in Devices also fills that name into every past
// Usage Log row for that device. "Usage Log > Refresh names" re-applies all names.

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action !== 'logUsage') return jsonOut({ status: 'ignored' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var devices = getOrCreateSheet(ss, 'Devices', ['Device ID', 'Name (fill in)', 'Device', 'First Seen', 'Last Seen', 'Times Used']);
    var log = getOrCreateSheet(ss, 'Usage Log', ['Time', 'Name', 'Action', 'Details', 'Device', 'Device ID', 'Flag', 'Note']);

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

    log.appendRow([now, name, data.event || '', data.details || '', device, deviceId, flag, data.note || '']);
    if (flag) log.getRange(log.getLastRow(), 1, 1, 8).setBackground('#fde2e2');
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

// Runs automatically whenever someone edits the sheet by hand.
function onEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== 'Devices' || e.range.getLastColumn() < 2 || e.range.getColumn() > 2) return;
  syncNames();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Usage Log')
    .addItem('Refresh names', 'syncNames')
    .addToUi();
}

// Copies each device's name from Devices onto all of its Usage Log rows, and
// sets the red highlight on rows whose device still has no name.
function syncNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var devices = ss.getSheetByName('Devices');
  var log = ss.getSheetByName('Usage Log');
  if (!devices || !log || log.getLastRow() < 2) return;

  var names = {};
  if (devices.getLastRow() > 1) {
    var deviceRows = devices.getRange(2, 1, devices.getLastRow() - 1, 2).getValues();
    var deviceColors = [];
    deviceRows.forEach(function(r) {
      var name = String(r[1]).trim();
      names[String(r[0])] = name;
      deviceColors.push(new Array(6).fill(name ? null : '#fde2e2'));
    });
    devices.getRange(2, 1, deviceRows.length, 6).setBackgrounds(deviceColors);
  }

  var count = log.getLastRow() - 1;
  var rows = log.getRange(2, 1, count, 8).getValues();
  var nameCol = [];
  var flagCol = [];
  var colors = [];
  rows.forEach(function(r) {
    var name = names[String(r[5])] || '';
    var flag = String(r[6]);
    if (name && (flag === 'NEW DEVICE' || flag === 'Unknown device')) flag = '';
    if (!name && !flag) flag = 'Unknown device';
    nameCol.push([name]);
    flagCol.push([flag]);
    colors.push(new Array(8).fill(flag ? '#fde2e2' : null));
  });
  log.getRange(2, 2, count, 1).setValues(nameCol);
  log.getRange(2, 7, count, 1).setValues(flagCol);
  log.getRange(2, 1, count, 8).setBackgrounds(colors);
}
