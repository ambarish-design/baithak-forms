// Midpoint (in lac INR) for each bracket offered by the income question.
var INCOME_BRACKET_MIDPOINTS = {
  'Less than 1 lac': 0.5,
  '1 lac - 2 lac': 1.5,
  '2 lac - 4 lac': 3,
  '4 lac - 6 lac': 5,
  '6 lac - 8 lac': 7,
  'Above 8 lac': 9
};

// Read-only stats for a form, fetched cross-origin via JSONP
// (Apps Script Web Apps can't set CORS headers, so plain fetch() can't read
// the response from a different origin — a <script src="...&callback=x">
// tag isn't subject to CORS, so we wrap the JSON in that callback instead).
// Usage: GET ?formId=<id>&callback=<jsFnName>
function doGet(e) {
  var formId = String((e.parameter.formId || '')).substring(0, 100);
  var callback = e.parameter.callback;
  var result = {
    formId: formId,
    count: 0,
    averageIncomeLac: null,
    minIncomeLac: null,
    maxIncomeLac: null
  };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(formId);
  if (sheet && sheet.getLastRow() > 1) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var incomeCol = headers.indexOf('Annual Income (INR)');
    if (incomeCol > -1) {
      var values = sheet.getRange(2, incomeCol + 1, sheet.getLastRow() - 1, 1).getValues();
      var sum = 0, n = 0, min = null, max = null;
      values.forEach(function (row) {
        var midpoint = INCOME_BRACKET_MIDPOINTS[row[0]];
        if (midpoint !== undefined) {
          sum += midpoint;
          n += 1;
          if (min === null || midpoint < min) min = midpoint;
          if (max === null || midpoint > max) max = midpoint;
        }
      });
      if (n > 0) {
        result.count = n;
        result.averageIncomeLac = Math.round((sum / n) * 10) / 10;
        result.minIncomeLac = min;
        result.maxIncomeLac = max;
      }
    }
  }

  var json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
