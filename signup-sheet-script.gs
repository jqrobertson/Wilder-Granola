// Google Apps Script that receives email sign-ups from the website
// and adds them to the "Wilder Granola Email Signups" Google Sheet.
//
// Setup (one time):
// 1. Open the sheet, then Extensions > Apps Script.
// 2. Replace everything in Code.gs with this file and click Save.
// 3. Deploy > New deployment > type "Web app".
//      Execute as: Me    Who has access: Anyone
// 4. Authorize when asked, then copy the Web app URL (ends in /exec)
//    into data-endpoint="" on the signup form in index.html.

function doPost(e) {
  var p = (e && e.parameter) || {};
  var email = String(p.email || '').trim().toLowerCase();

  // Ignore bots (hidden field filled in) and anything that isn't an email.
  if (p.website || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return ContentService.createTextOutput('ignored');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var existing = sheet.getRange('B:B').getValues().flat();
    if (existing.indexOf(email) === -1) {
      sheet.appendRow([new Date(), email, p.page || '']);
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}
