// 5 Minute Dog – waitlist receiver. Paste into Extensions → Apps Script of the
// "5 Minute Dog – Waitlist" sheet, then Deploy → New deployment → Web app
// (Execute as: Me, Who has access: Anyone). Copy the /exec URL into WEBAPP_URL in index.html.
var FIELDS = ['email', 'whatsapp', 'delivery', 'dogName', 'age', 'problem'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Signups');
    var p = (e && e.parameter) || {};
    var row = [new Date()].concat(FIELDS.map(function (k) {
      var v = String(p[k] || '').slice(0, 200);
      return /^[=+\-@]/.test(v) ? "'" + v : v; // block formula injection
    }));
    sheet.appendRow(row);
    return ContentService.createTextOutput('OK');
  } finally {
    lock.releaseLock();
  }
}
