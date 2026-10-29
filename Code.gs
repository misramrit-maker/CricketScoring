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
 * To add every tab and its column headings straight away, choose the function
 * setupSheets in the Apps Script editor and press Run (once). The headings are
 * also added automatically whenever the app sends data, and whenever the web app
 * URL is opened in a browser.
 *
 * Players and teams typed or saved in the Players and Teams tabs are also read back
 * by the app, so they appear in the team and player pickers when starting a match.
 * In the Teams tab, put the team name in column B and the player names, separated
 * by commas, in column C (captain in D and wicketkeeper in E are optional).
 *
 * Open the web app URL in a browser to check it is live. It shows when the
 * last sync arrived and how many rows each tab holds.
 */

/* Leave this empty if you made the script from the sheet itself (Extensions > Apps Script).
 * If the script was created on its own at script.google.com, paste the sheet's ID here
 * (the long text in the sheet's address between /d/ and /edit). If you leave it empty in
 * that case, the script creates a new sheet called "CricketScorer data" and keeps using it. */
var SPREADSHEET_ID = '';

var MAX_CELL = 45000; // Google Sheets allows 50,000 characters in a cell.

/* The column headings of every tab, so they are always written even when a tab has no rows yet. */
var HEADINGS = {
  "Players": ["Id", "Name", "Teams"],
  "Teams": ["Id", "Team", "Players", "Captain", "Wicketkeeper"],
  "Matches": ["Match ID", "Date", "Team A", "Team B", "Format", "Overs", "Toss", "Scores", "Result", "Category", "Tournament", "Toss called by", "Call", "Bat flip landed", "Toss won by", "Chose to"],
  "Innings": ["Match ID", "Date", "Match", "Innings", "Batting team", "Runs", "Wickets", "Overs", "Extras", "Fall of wickets"],
  "Batting": ["Match ID", "Date", "Match", "Innings", "Team", "Batter", "How out", "Runs", "Balls", "4s", "6s", "Strike rate"],
  "Bowling": ["Match ID", "Date", "Match", "Innings", "Team", "Bowler", "Overs", "Maidens", "Runs", "Wickets", "Wides", "No balls", "Economy"],
  "Lineups": ["Match ID", "Date", "Match", "Team", "Player", "Role", "Status"],
  "Ball by ball": ["Match ID", "Date", "Match", "Innings", "Event #", "Type", "Ball kind", "Runs", "Bowler", "Wicket", "Batter out", "Fielder"],
  "Live": ["Match", "Innings", "Event #", "Type", "Ball kind", "Runs", "Bowler", "Wicket", "Batter out", "Fielder", "Bowler id"],
  "Current match": ["Match", "Format", "Overs", "Toss", "Tournament", "Score", "Status", "Team A", "Captain A", "Wicketkeeper A", "Team B", "Captain B", "Wicketkeeper B", "Toss called by", "Call", "Bat flip landed", "Toss won by", "Chose to"],
  "Tournaments": ["Id", "Tournament", "Type", "Overs", "Status", "Teams", "Fixtures", "Played", "Created", "People", "Score or result"],
  "People": ["Tournament", "Person", "Teams played as", "Match wins", "Status"],
  "Fixtures": ["Tournament", "#", "Stage", "Team A", "Team B", "Status", "Result", "Scores", "Date", "Match ID", "Stage", "Person A", "Person B", "Score after match"],
  "Series": ["Tournament", "Team A", "Wins A", "Wins B", "Team B", "Ties", "Draws", "Played", "Total", "Status"],
  "Activity": ["Tournament", "When", "What happened"]
};
var TAB_ORDER = ["Players", "Teams", "Matches", "Innings", "Batting", "Bowling", "Lineups", "Ball by ball", "Live", "Current match", "Tournaments", "People", "Fixtures", "Series", "Activity"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);
    if (data.app !== 'CricketScorer') {
      return out_({ ok: false, error: 'Not a CricketScorer payload' });
    }

    var ss = getSS_();
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
      var head = (t.head && t.head.length) ? t.head : (HEADINGS[t.name] || []);
      try {
        writeTab_(ss, String(t.name), head, t.rows || []);
        counts[t.name] = (t.rows || []).length;
      } catch (tabErr) {
        // One bad tab must never stop the others from being written.
        counts[t.name] = 'error: ' + String(tabErr);
      }
    }
    // Tabs the app did not send this time still get their headings.
    ensureHeadings_(ss);

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

function doGet(e) {
  // The app asks for the saved players and teams so they can be picked for a match on any device.
  if (e && e.parameter && e.parameter.action === 'lib') {
    try {
      var ss0 = getSS_();
      return out_({ ok: true, players: readTab_(ss0, 'Players'), teams: readTab_(ss0, 'Teams') });
    } catch (err) {
      return out_({ ok: false, error: String(err) });
    }
  }
  var info = { ok: true, app: 'CricketScorer' };
  try {
    var ss = getSS_();
    ensureHeadings_(ss);
    info.sheet = ss.getName();
    info.sheetUrl = ss.getUrl();
    info.tabs = ss.getSheets().map(function (s) { return s.getName(); });
  } catch (e) {
    info.ok = false;
    info.problem = 'Could not open a spreadsheet: ' + String(e);
  }
  var props = PropertiesService.getScriptProperties();
  var counts = {};
  try { counts = JSON.parse(props.getProperty('counts') || '{}'); } catch (x) {}
  info.lastSync = props.getProperty('lastSync') || 'never';
  info.rows = counts;
  return out_(info);
}

/* Reads a tab's rows without its heading row. */
function readTab_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var width = Math.max(sh.getLastColumn(), 1);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  return vals.filter(function (r) { return String(r[1] === undefined ? '' : r[1]) !== '' || String(r[0]) !== ''; })
    .map(function (r) { return r.map(function (v) { return v === null || v === undefined ? '' : String(v); }); });
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
  if (header.length) styleHeading_(sh, width);
}

function styleHeading_(sh, width) {
  var r = sh.getRange(1, 1, 1, width);
  r.setFontWeight('bold');
  r.setBackground('#d9efe9');
  sh.setFrozenRows(1);
}

/* Creates any missing tab and writes the headings into any tab whose first row is empty. */
function ensureHeadings_(ss) {
  TAB_ORDER.forEach(function (name) {
    var head = HEADINGS[name];
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var first = sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, head.length).getValues()[0] : [];
    var empty = true;
    for (var i = 0; i < first.length; i++) { if (String(first[i]) !== '') { empty = false; break; } }
    if (empty) {
      sh.getRange(1, 1, 1, head.length).setValues([head]);
      styleHeading_(sh, head.length);
    }
  });
  // The default empty "Sheet1" is not needed once the real tabs exist.
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) { try { ss.deleteSheet(def); } catch (e) {} }
}

/* Finds the spreadsheet to write to. */
function getSS_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var props = PropertiesService.getScriptProperties();
  var id = SPREADSHEET_ID || props.getProperty('ssId');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('CricketScorer data');
  props.setProperty('ssId', ss.getId());
  return ss;
}

/* Adds a CricketScorer menu to the sheet when the script is bound to it. */
function onOpen() {
  try { SpreadsheetApp.getUi().createMenu('CricketScorer').addItem('Add tabs and headings', 'setupSheets').addToUi(); } catch (e) {}
}

/* Run this once from the editor to create every tab with its headings. */
function setupSheets() {
  var ss = getSS_();
  ensureHeadings_(ss);
  Logger.log('Tabs and headings are ready: ' + ss.getUrl());
  return 'Tabs and headings are ready: ' + ss.getUrl();
}

function pad_(row, width) {
  var r = [];
  for (var i = 0; i < width; i++) {
    var v = i < row.length ? row[i] : '';
    if (v === null || v === undefined) v = '';
    if (typeof v === 'object') v = JSON.stringify(v);
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
