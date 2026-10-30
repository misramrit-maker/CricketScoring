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
 *   Toss           the toss of every match: who called, the call, the bat flip outcome,
 *                  who won the toss and what they chose
 *   Lineups        every player in every match: team, captain, wicketkeeper, playing or substitute
 *   Ball by ball   every ball and event of every match, finished or in progress
 *   Live           ball by ball record of the match being scored
 *   Current match  the match in progress: teams, captains, keepers, toss, score
 *   Tournaments    every tournament with its status, current stage, its people and its score or result
 *   Tournament teams  each team in each tournament: who plays as it, its players and its record
 *   People         each person in a tournament: teams played as, match wins
 *   Fixtures       every fixture: stage, teams, who played as each, status, result, winning team, point to, score after the match
 *   Series         the running score between each pair of teams, e.g. 3-1
 *   Activity       dated log of what was done in each tournament
 *   App data       the whole app's data, saved so every device using this Sync URL shows the same thing (do not edit)
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
 * by commas, in column C (captain in D and wicketkeeper in E are optional; column F is filled in by the app).
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
  "Teams": ["Id", "Team", "Players", "Captain", "Wicketkeeper", "Player ids"],
  "Matches": ["Match ID", "Date", "Team A", "Team B", "Format", "Overs", "Toss", "Scores", "Result", "Category", "Tournament", "Toss called by", "Call", "Bat flip landed", "Toss won by", "Chose to"],
  "Innings": ["Match ID", "Date", "Match", "Innings", "Batting team", "Runs", "Wickets", "Overs", "Extras", "Fall of wickets"],
  "Batting": ["Match ID", "Date", "Match", "Innings", "Team", "Batter", "How out", "Runs", "Balls", "4s", "6s", "Strike rate"],
  "Bowling": ["Match ID", "Date", "Match", "Innings", "Team", "Bowler", "Overs", "Maidens", "Runs", "Wickets", "Wides", "No balls", "Economy"],
  "Toss": ["Match ID", "Date", "Match", "Format", "Toss type", "Called by", "Call", "Bat flip outcome (side facing up)", "Toss won by", "Chose to", "Toss summary", "Result"],
  "Lineups": ["Match ID", "Date", "Match", "Team", "Player", "Role", "Status"],
  "Ball by ball": ["Match ID", "Date", "Match", "Innings", "Event #", "Type", "Ball kind", "Runs", "Bowler", "Wicket", "Batter out", "Fielder"],
  "Live": ["Match", "Innings", "Event #", "Type", "Ball kind", "Runs", "Bowler", "Wicket", "Batter out", "Fielder", "Bowler id"],
  "Current match": ["Match", "Format", "Overs", "Toss", "Tournament", "Score", "Status", "Team A", "Captain A", "Wicketkeeper A", "Team B", "Captain B", "Wicketkeeper B", "Toss called by", "Call", "Bat flip landed", "Toss won by", "Chose to"],
  "Tournaments": ["Id", "Tournament", "Type", "Overs", "Status", "Teams", "Fixtures", "Played", "Created", "People", "Score or result", "Current stage"],
  "Tournament teams": ["Tournament", "Team", "Played by", "Players", "Played", "Won", "Lost", "Tied", "Drawn", "Tournament ID"],
  "People": ["Tournament", "Person", "Teams played as", "Match wins", "Status", "Tournament ID"],
  "Fixtures": ["Tournament", "#", "Stage", "Team A", "Team B", "Status", "Result", "Scores", "Date", "Match ID", "Phase", "Person A", "Person B", "Score after match", "Winning team", "Point to", "Tournament ID"],
  "Series": ["Tournament", "Team A", "Wins A", "Wins B", "Team B", "Ties", "Draws", "Played", "Total", "Status", "Tournament ID"],
  "Activity": ["Tournament", "When", "What happened", "Tournament ID"],
  "App data": ["Part", "Shared app data (do not edit or delete)"]
};
var TAB_ORDER = ["Players", "Teams", "Matches", "Innings", "Batting", "Bowling", "Toss", "Lineups", "Ball by ball", "Live", "Current match", "Tournaments", "Tournament teams", "People", "Fixtures", "Series", "Activity", "App data"];

/* ---- merging, so what two devices enter at the same time is never lost ---- */
function cmp_(a, b) { return a < b ? -1 : a > b ? 1 : 0; }
function canon_(v) {
  if (v === null || v === undefined) return 'null';
  if (Object.prototype.toString.call(v) === '[object Array]') return '[' + v.map(canon_).join(',') + ']';
  if (typeof v === 'object') {
    return '{' + Object.keys(v).filter(function (k) { return v[k] !== undefined; }).sort().map(function (k) { return JSON.stringify(k) + ':' + canon_(v[k]); }).join(',') + '}';
  }
  return JSON.stringify(v);
}
/* Puts the shared data into a fixed order, the same order the app uses, so two copies with the same data compare equal. */
function normCore_(s) {
  s = s || {};
  var lib = s.lib || {};
  var byNm = function (a, b) { return cmp_(String(a.name || '').toLowerCase(), String(b.name || '').toLowerCase()) || cmp_(String(a.id), String(b.id)); };
  return {
    liveAt: String(s.liveAt !== undefined ? s.liveAt : (s.at || '')),
    lib: { players: (lib.players || []).slice().sort(byNm), teams: (lib.teams || []).slice().sort(byNm) },
    hist: (s.hist || []).slice().sort(function (a, b) { return cmp_(String(b.at), String(a.at)) || cmp_(String(a.id), String(b.id)); }),
    tours: (s.tours || []).slice().sort(function (a, b) { return cmp_(String(a.id), String(b.id)); }),
    live: s.live || null,
    del: s.del || {}
  };
}
/* Every record is kept; when both sides have the same record the one updated last wins. A dated delete marker removes a record unless it was edited after the delete. */
function mergeRecs_(x, y, pre, del) {
  var map = {};
  [x || [], y || []].forEach(function (list) {
    list.forEach(function (r) {
      if (!r || !r.id) return;
      var c = map[r.id];
      if (!c || String(r.u || '') >= String(c.u || '')) map[r.id] = r;
    });
  });
  var out = [];
  for (var id in map) {
    var r = map[id], t = del[pre + id];
    if (t && String(t) >= String(r.u || '')) continue;
    out.push(r);
  }
  return out;
}
function mergeStates_(a, b) {
  a = normCore_(a); b = normCore_(b);
  var del = {}, k;
  [a.del, b.del].forEach(function (d) { for (k in d) { if (!del[k] || String(d[k]) > String(del[k])) del[k] = d[k]; } });
  var live = b.liveAt >= a.liveAt ? b : a;
  return normCore_({
    liveAt: live.liveAt, live: live.live, del: del,
    lib: {
      players: mergeRecs_(a.lib.players, b.lib.players, 'p:', del).filter(function (p) { return p.name; }),
      teams: mergeRecs_(a.lib.teams, b.lib.teams, 't:', del)
    },
    hist: mergeRecs_(a.hist, b.hist, 'm:', del).filter(function (r) { return r.t1; }),
    tours: mergeRecs_(a.tours, b.tours, 'o:', del)
  });
}
function loadStateText_(ss) {
  var rows = readTab_(ss, 'App data');
  rows.sort(function (a, b) { return Number(a[0]) - Number(b[0]); });
  return rows.map(function (r) { return String(r[1]).substring(1); }).join('');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);
    if (data.app !== 'CricketScorer') {
      return out_({ ok: false, error: 'Not a CricketScorer payload' });
    }
    var ss = getSS_();
    var props = PropertiesService.getScriptProperties();
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

    // What arrives is MERGED with what is already saved, never written over it. Whatever any device enters is kept.
    var writeTabs = true;
    if (data.state) {
      var incoming = null;
      try { incoming = JSON.parse(data.state); } catch (pe) {}
      if (incoming) {
        var stored = null, text = loadStateText_(ss);
        if (text) { try { stored = JSON.parse(text); } catch (se) {} }
        var inN = normCore_(incoming);
        var merged = stored ? mergeStates_(stored, incoming) : inN;
        var mC = canon_(merged), iC = canon_(inN), sC = stored ? canon_(normCore_(stored)) : '';
        // The readable tabs are only written from a copy that already contains everything saved. A copy that is missing something waits until that device has loaded it and sends again.
        writeTabs = (mC === iC);
        var newAt = new Date().toISOString();
        if (mC !== sC) {
          merged.at = newAt;
          var st = JSON.stringify(merged), chunks = [];
          for (var c = 0; c < st.length; c += 30000) chunks.push([chunks.length + 1, '~' + st.substr(c, 30000)]);
          writeTab_(ss, 'App data', HEADINGS['App data'], chunks);
          props.setProperty('stateAt', newAt);
          // stateBy names the device whose copy already had everything, so that device need not download it again.
          props.setProperty('stateBy', writeTabs ? String(data.dev || '') : '');
        } else if (!writeTabs) {
          // Nothing new was saved, but the sending device is missing something, so tell every device to reload.
          props.setProperty('stateAt', newAt);
          props.setProperty('stateBy', '');
        }
      }
    }

    var counts = {};
    if (writeTabs) {
      for (var i = 0; i < tabs.length; i++) {
        var t = tabs[i];
        if (!t || !t.name || t.name === 'App data') continue;
        var head = (t.head && t.head.length) ? t.head : (HEADINGS[t.name] || []);
        try {
          writeTab_(ss, String(t.name), head, t.rows || []);
          counts[t.name] = (t.rows || []).length;
        } catch (tabErr) {
          // One bad tab must never stop the others from being written.
          counts[t.name] = 'error: ' + String(tabErr);
        }
      }
      props.setProperty('counts', JSON.stringify(counts));
    }
    // Tabs the app did not send this time still get their headings.
    ensureHeadings_(ss);

    var sentAt = data.sentAt || new Date().toISOString();
    var recent = [];
    try { recent = JSON.parse(props.getProperty('recent') || '[]'); } catch (rx) {}
    recent.push(sentAt);
    if (recent.length > 30) recent = recent.slice(-30);
    props.setProperty('recent', JSON.stringify(recent));
    props.setProperty('lastSync', sentAt);

    return out_({ ok: true, sentAt: sentAt, rows: counts, tabsWritten: writeTabs });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : '';
  var props = PropertiesService.getScriptProperties();
  // A very small reply that only says when the shared data last changed, and which device changed it. The app checks this often.
  if (action === 'stateat') {
    return out_({ ok: true, stateAt: props.getProperty('stateAt') || '', stateBy: props.getProperty('stateBy') || '' });
  }
  // A quick reply the app uses to confirm that what it sent has been received.
  if (action === 'check') {
    var chk = { ok: true, lastSync: props.getProperty('lastSync') || 'never', stateAt: props.getProperty('stateAt') || '', stateBy: props.getProperty('stateBy') || '', recent: [] };
    try { chk.recent = JSON.parse(props.getProperty('recent') || '[]'); } catch (xr) {}
    try { getSS_(); } catch (ec) { chk.ok = false; chk.problem = 'Could not open a spreadsheet: ' + String(ec); }
    return out_(chk);
  }
  // The app asks for the shared state to load what was entered on another device.
  if (action === 'state') {
    var lk = LockService.getScriptLock(), got = false;
    try {
      got = lk.tryLock(15000);
      var text = loadStateText_(getSS_());
      return out_({ ok: true, state: text, stateAt: props.getProperty('stateAt') || '', stateBy: props.getProperty('stateBy') || '' });
    } catch (errS) {
      return out_({ ok: false, error: String(errS) });
    } finally {
      if (got) { try { lk.releaseLock(); } catch (xl) {} }
    }
  }
  // Every visit makes sure all tabs exist with their headings.
  try { ensureHeadings_(getSS_()); } catch (x0) {}
  // The app asks for the saved players and teams so they can be picked for a match on any device.
  if (action === 'lib') {
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
  } catch (e2) {
    info.ok = false;
    info.problem = 'Could not open a spreadsheet: ' + String(e2);
  }
  var counts = {};
  try { counts = JSON.parse(props.getProperty('counts') || '{}'); } catch (x) {}
  info.lastSync = props.getProperty('lastSync') || 'never';
  info.stateAt = props.getProperty('stateAt') || '';
  info.stateBy = props.getProperty('stateBy') || '';
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

  var width = header.length;
  rows.forEach(function (r) { if (r.length > width) width = r.length; });
  if (width === 0) return;

  var values = [];
  if (header.length) values.push(pad_(header, width));
  rows.forEach(function (r) { values.push(pad_(r, width)); });

  // A tab whose content has not changed is left alone, which keeps every save fast.
  var props = PropertiesService.getScriptProperties();
  var key = 'h_' + name;
  var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(values)));
  if (props.getProperty(key) === hash && sh.getLastRow() >= values.length && sh.getRange(1, 1).getValue() !== '') return;
  sh.clearContents();

  sh.getRange(1, 1, values.length, width).setValues(values);
  if (header.length) styleHeading_(sh, width);
  props.setProperty(key, hash);
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
