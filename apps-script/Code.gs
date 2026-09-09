/**
 * Shared backend for every form under forms.baithak.org.
 * One deployment serves all forms: each submission carries a `formId`,
 * which maps 1:1 to a tab in this bound Spreadsheet. New forms need no
 * script changes — the first submission for a new formId creates its tab
 * and header row from the field names sent by the page.
 *
 * Deploy: Extensions > Apps Script (from the Sheet) > paste this file >
 * Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 */

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var formId = String(payload.formId || 'untitled-form').substring(0, 100);
  var fields = payload.fields || {};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(formId);
  if (!sheet) {
    sheet = ss.insertSheet(formId);
    sheet.appendRow(['Timestamp'].concat(Object.keys(fields)));
  }

  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : ['Timestamp'];
  var newKeys = Object.keys(fields).filter(function (k) {
    return headers.indexOf(k) === -1;
  });
  if (newKeys.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, newKeys.length).setValues([newKeys]);
    headers = headers.concat(newKeys);
  }

  var row = headers.map(function (h) {
    if (h === 'Timestamp') return new Date();
    return fields[h] !== undefined ? fields[h] : '';
  });
  sheet.appendRow(row);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success', formId: formId }))
    .setMimeType(ContentService.MimeType.JSON);
}
