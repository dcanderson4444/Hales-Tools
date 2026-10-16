// Hale's Tools usage log.
// Paste this into the Apps Script editor of a Google Sheet (Extensions > Apps Script),
// then Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
// Copy the web app URL into USAGE_LOG_URL in index.html.

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action !== 'logUsage') {
    return ContentService.createTextOutput(JSON.stringify({ status: 'ignored' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usage Log');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Usage Log');
    sheet.appendRow(['Time', 'Name', 'Action', 'Details']);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([new Date(), data.name || '', data.event || '', data.details || '']);

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
