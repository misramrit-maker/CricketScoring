/**
 * CricketScorer - Google Apps Script receiver
 *
 * Every time anything changes in the app, it sends ALL of its data to this
 * script and the script rewrites one tab per kind of data:
 *
 *   Players        every saved player and the teams they are in
 *   Teams          saved teams with their players, captain and wicketkeeper
 *   Matches        every finished match: date, teams, toss, scores, result,
 *                  category (Tournament match / Regular match) and tournament
 *   Innings        one row per innings (the match in progress is included): total, wickets, overs, extras, fall of wickets
 *   Batting        every batter in every innings: runs, balls, 4s, 6s, strike rate, how out
 *   Bowling        every bowler in every innings: overs, maidens, runs, wickets, wides, no balls, economy
 *   Lineups        every player in every match: team, captain, wicketkeeper, playing or substitute
 *   Ball by ball   every ball and event of every match, finished or in progress
 *   Live           ball by ball record of the match being scored
 *   Current match  the match in progress: teams, captains, keepers, toss, score
 *   Tournaments    every tournament with its status, its people and its score or result
 *   People         each person in a tournament: teams played as, match wins
 *   Fixtures       every fixture: stage, teams, who played as each, status, result, score after the match
 *   Series         the running score between each pair of teams, e.g. 3-1
 *   Activity       dated log of what was done in each tournament
 *
 * Setup:
 *  1. Create a Google Sheet, then open Extensions > Apps Script.
 *  2. Replace the default code with this file and save.
 *  3. Deploy > New deployment > type "Web app".
 *     Execute as: Me.   Who has access: Anyone.
 *  4. Copy the web app URL (ends in /exec) and paste it into the app:
 *     menu (three dots) > Match history > Sync URL.
 *  5. If you change this code later, use Deploy > Manage deployments >
 *     Edit > New version so the same URL keeps working.
 *
 * Open the web app URL in a browser to check it is live. It shows when the
 * last sync arrived and how many rows each tab holds.
 */

var MAX_CELL = 45000; // Google Sheets allows 50,000 characters in a cell.

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);
    if (data.app !== 'CricketScorer') {
      return out_({ ok: false, error: 'Not a CricketScorer payload' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tabs = data.tabs;

    // Older versions of the app sent fixed fields instead of a list of tabs.
    if (!tabs || !tabs.length) {
      tabs = [
        { name: 'Players', head: ['Id', 'Name'], rows: data.players || [] },
        { name: 'Teams', head: ['Id', 'Team', 'Players'], rows: data.teams || [] },
        { name: 'Matches', head: data.head || [], rows: data.rows || [] },
        { name: 'Live', head: data.liveHead || [], rows: data.live || [] }
      ];
    }

    var counts = {};
    for (var i = 0; i < tabs.length; i++) {
      var t = tabs[i];
      if (!t || !t.name) continue;
      writeTab_(ss, String(t.name), t.head || [], t.rows || []);
      counts[t.name] = (t.rows || []).length;
    }

    var props = PropertiesService.getScriptProperties();
    props.setProperty('lastSync', data.sentAt || new Date().toISOString());
    props.setProperty('counts', JSON.stringify(counts));

    return out_({ ok: true, sentAt: data.sentAt || null, rows: counts });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() {
  var props = PropertiesService.getScriptProperties();
  var counts = {};
  try { counts = JSON.parse(props.getProperty('counts') || '{}'); } catch (x) {}
  return out_({
    ok: true,
    app: 'CricketScorer',
    lastSync: props.getProperty('lastSync') || 'never',
    rows: counts
  });
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
  var r = [];
  for (var i = 0; i < width; i++) {
    var v = i < row.length ? row[i] : '';
    if (v === null || v === undefined) v = '';
    if (typeof v === 'string' && v.length > MAX_CELL) v = v.slice(0, MAX_CELL);
    r.push(v);
  }
  return r;
}

function out_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
