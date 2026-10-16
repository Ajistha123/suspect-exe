// suspect.exe registration backend (Google Apps Script) for Spark Club, IIIT Bhopal
//
// Setup (about 3 minutes):
// 1. Create a new Google Sheet. Open Extensions > Apps Script.
// 2. Delete any code there, paste this whole file, and click Save.
// 3. Click Deploy > New deployment > gear icon > Web app.
//      Execute as: Me
//      Who has access: Anyone
//    Click Deploy and approve the permission prompts. Copy the Web app URL.
// 4. In index.html, paste that URL into FORM_ENDPOINT, then re-upload the file.
//
// After you edit this script later, use Deploy > Manage deployments > Edit > New version,
// otherwise the old code keeps running. The sheet gets a "Registrations" tab automatically.

var CAPACITY = 100;            // maximum number of registrations
var SHEET_NAME = 'Registrations';
var HEADERS = ['Time', 'Code', 'Name', 'Roll number', 'Program', 'Year', 'WhatsApp', 'Email'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange('G:G').setNumberFormat('@');   // keep WhatsApp numbers as text
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// The page calls this to show how many seats are left.
function doGet() {
  var taken = Math.max(getSheet_().getLastRow() - 1, 0);
  return json_({ ok: true, capacity: CAPACITY, left: Math.max(CAPACITY - taken, 0) });
}

// The page calls this to register someone, or to update a registration when edit is true.
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var sh = getSheet_();
    var last = sh.getLastRow();
    var rows = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];

    var roll = String(d.roll || '').trim().toUpperCase();
    var email = String(d.email || '').trim().toLowerCase();
    if (!roll || !email || !d.name) return json_({ ok: false, error: 'Missing fields' });

    var own = -1;      // the row being edited, if any
    var clashRoll = false, clashEmail = false;
    for (var i = 0; i < rows.length; i++) {
      var sameCode = String(rows[i][1]) === String(d.code);
      if (d.edit && sameCode) { own = i; continue; }
      if (String(rows[i][3]).toUpperCase() === roll) clashRoll = true;
      if (String(rows[i][7]).toLowerCase() === email) clashEmail = true;
    }
    if (clashRoll) return json_({ ok: false, dup: 'roll' });
    if (clashEmail) return json_({ ok: false, dup: 'email' });

    var row = [new Date(), d.code, d.name, roll, d.prog, d.year, d.wa, email];
    if (d.edit && own >= 0) {
      sh.getRange(own + 2, 1, 1, HEADERS.length).setValues([row]);
      return json_({ ok: true, code: rows[own][1] });
    }
    if (rows.length >= CAPACITY) return json_({ ok: false, full: true });
    sh.appendRow(row);
    return json_({ ok: true, code: d.code });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
