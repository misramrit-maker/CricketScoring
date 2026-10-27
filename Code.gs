/**
 * CricketScorer - Google Apps Script receiver
 *
 * Setup:
 *  1. Create a Google Sheet, then open Extensions > Apps Script.
 *  2. Replace the default code with this file and save.
 *  3. Deploy > New deployment > type "Web app".
 *     Execute as: Me.   Who has access: Anyone.
 *  4. Copy the web app URL (ends in /exec) and paste it into the app's
 *     Sync URL box (Match history > Sync), or into SYNC_URL in the HTML file.
 *  5. If you change this code later, use Deploy > Manage deployments >
 *     Edit > New version so the same URL keeps working.
 *
 * The app sends everything as JSON in a POST body. Each sync rewrites four tabs:
 *   Players, Teams, Matches, Live
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse(e.postData.contents);
    if (data.app !== 'CricketScorer') {
      return out_({ ok: false, error: 'Not a CricketScorer payload' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    writeTab_(ss, 'Players', ['Id', 'Name'], data.players || []);
    writeTab_(ss, 'Teams', ['Id', 'Team', 'Players'], data.teams || []);
    writeTab_(ss, 'Matches', data.head || [], data.rows || []);
    writeTab_(ss, 'Live', data.liveHead || [], data.live || []);

    var meta = ss.getSheetByName('Players');
    if (meta) {
      PropertiesService.getScriptProperties().setProperty('lastSync', data.sentAt || new Date().toISOString());
    }

    return out_({ ok: true, sentAt: data.sentAt || null });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() {
  var last = PropertiesService.getScriptProperties().getProperty('lastSync') || 'never';
  return out_({ ok: true, app: 'CricketScorer', lastSync: last });
}

function writeTab_(ss, name, header, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();

  var width = header.length;
  rows.forEach(function (r) { if (r.length > width) width = r.length; });
  if (width === 0) return;

  var values = [];
  if (header.length) values.push(pad_(header, width));
  rows.forEach(function (r) { values.push(pad_(r, width)); });

  sh.getRange(1, 1, values.length, width).setValues(values);

  if (header.length) {
    sh.getRange(1, 1, 1, width).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function pad_(row, width) {
  var r = row.slice();
  while (r.length < width) r.push('');
  return r;
}

function out_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
