/**
 * RoboWarriors REBUILT Scout -> Google Sheets
 *
 * Pulls match and pit reports from the team's Firebase (Firestore) database
 * every 5 minutes and builds three tabs:
 *   Team Averages  - one row per team, sorted by the column you pick in the dropdown
 *   Match Data     - every match report (raw)
 *   Pit Data       - every pit report (raw)
 *
 * Setup (once): see sheets/README.md. In short, paste this file and appsscript.json
 * into Extensions > Apps Script, then run setup().
 *
 * The script signs in to Firebase as the Google account that runs it, so that account
 * must be an owner or editor of the Firebase project.
 */

const PROJECT_ID = "rebuilt-scouting-f33c1";
const SYNC_MINUTES = 5;

const AVG_SHEET = "Team Averages";
const MATCH_SHEET = "Match Data";
const PIT_SHEET = "Pit Data";
const ACC_SHEET = "Fuel Accuracy";
const SCOUT_SHEET = "Scouter Accuracy";
const PREVIEW_SHEET = "Match Preview";
const SCHED_SHEET = "Schedule";
const MIN_CHECKS = 5;   // fewer checked alliances than this is flagged as a rough number

// Point values, kept in step with index.html
const CLIMB_PTS = { none: 0, fail: 0, L1: 10, L2: 20, L3: 30 };
const AUTO_CLIMB_PTS = 15;

const MATCH_COLS = [
  ["event", "Event"], ["match", "Match"], ["team", "Team"], ["alliance", "Alliance"],
  ["wonAuto", "Won auto"], ["autoFuel", "Auto fuel"], ["autoClimb", "Auto climb L1"],
  ["transFuel", "Transition fuel"], ["s1Fuel", "Shift 1 fuel"], ["s2Fuel", "Shift 2 fuel"],
  ["s3Fuel", "Shift 3 fuel"], ["s4Fuel", "Shift 4 fuel"], ["teleFuel", "Teleop fuel"],
  ["endFuel", "End game fuel"], ["climb", "Climb"], ["estPts", "Est. points"],
  ["s1Role", "Shift 1 inactive role"], ["s2Role", "Shift 2 inactive role"],
  ["s3Role", "Shift 3 inactive role"], ["s4Role", "Shift 4 inactive role"],
  ["defense", "Defense (0-2)"], ["driver", "Driver (1-5)"], ["robot", "Robot status"],
  ["minor", "Minor fouls"], ["major", "Major fouls"], ["scout", "Scout"],
  ["notes", "Notes"], ["saved", "Saved"]
];

const PIT_COLS = [
  ["team", "Team"], ["drive", "Drivetrain"], ["intake", "Intake"], ["shooter", "Shooter"],
  ["capacity", "Fuel capacity"], ["autoFuel", "Auto fuel (claimed)"], ["maxClimb", "Highest climb"],
  ["autoClimbCap", "Climbs in auto"], ["weight", "Weight (lb)"], ["scout", "Scout"],
  ["notes", "Notes"], ["saved", "Saved"]
];

const AVG_COLS = [
  "Team", "Matches", "Avg est. points", "Best est. points", "Avg auto fuel", "Auto climb %",
  "Avg transition fuel", "Avg teleop fuel", "Avg end game fuel", "Avg total fuel",
  "Climb %", "L2+ climb %", "L3 climb %", "Top climb", "Defense %", "Avg foul points",
  "Avg driver", "Robot issues", "Drivetrain", "Shooter", "Fuel capacity"
];

/* ---------------- menu + triggers ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Scouting")
    .addItem("Sync now", "syncFromFirebase")
    .addItem("Set up / repair", "setup")
    .addToUi();
}

/** Re-sorts / re-filters Team Averages when a dropdown at the top changes. */
function onEdit(e) {
  const name = e.range.getSheet().getName();
  const r = e.range.getRow(), c = e.range.getColumn();
  if (name === AVG_SHEET) {
    if (r <= 3 && c === 2) buildAverages_();
  } else if (name === PREVIEW_SHEET) {
    if (c === 2 && r <= 2) buildPreview_(true);                                 // new event or match: pull its teams
    else if (c === 2 && r === 3) buildPreview_(false);                          // stats from this event / all events
    else if (r === PV_TEAM_ROW && c >= 2 && c <= 9) buildPreview_(false);       // a team number was typed in
  }
}

/** Run once: creates the tabs, the 5-minute sync, and does the first sync. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  [AVG_SHEET, PREVIEW_SHEET, ACC_SHEET, SCOUT_SHEET, MATCH_SHEET, PIT_SHEET, SCHED_SHEET].forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  const avg = ss.getSheetByName(AVG_SHEET);
  if (!avg.getRange("A1").getValue()) {
    avg.getRange("A1:A3").setValues([["Event"], ["Sort by"], ["Order"]]).setFontWeight("bold");
    avg.getRange("B1:B3").setValues([["All events"], ["Avg est. points"], ["Highest first"]]);
  }
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "syncFromFirebase")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("syncFromFirebase").timeBased().everyMinutes(SYNC_MINUTES).create();
  const def = ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 3 && def.getLastRow() === 0) ss.deleteSheet(def);
  syncFromFirebase();
  const pvw = ss.getSheetByName(PREVIEW_SHEET);
  if (pvw) { ss.setActiveSheet(pvw); ss.moveActiveSheet(2); }
  ss.setActiveSheet(avg);
}

/* ---------------- sync ---------------- */

function syncFromFirebase() {
  const matches = fetchCollection_("matches").map(normalizeMatch_);
  const pits = fetchCollection_("pit");
  writeTable_(MATCH_SHEET, MATCH_COLS, matches.sort((a, b) =>
    String(a.event).localeCompare(String(b.event)) || a.match - b.match || a.team - b.team));
  writeTable_(PIT_SHEET, PIT_COLS, pits.map(p => Object.assign({}, p, { saved: p.ts ? new Date(p.ts) : "" }))
    .sort((a, b) => a.team - b.team));
  buildAverages_();
  const sched = {};
  try { addAppSchedules_(sched, fetchCollection_("schedules")); } catch (e) { /* no schedules loaded in the app yet */ }
  try { buildAccuracy_(matches, sched); }
  catch (e) { writeAccuracyMessage_("Couldn't check accuracy: " + e.message); }
  writeTable_(SCHED_SHEET, SCHED_COLS, Object.keys(sched).map(k => sched[k])
    .sort((a, b) => String(a.event).localeCompare(String(b.event)) || Number(a.match) - Number(b.match)));
  buildPreview_(false);
  const avg = SpreadsheetApp.getActive().getSheetByName(AVG_SHEET);
  avg.getRange("D1").setValue("Last synced " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, h:mm a"))
    .setFontColor("#6d5f62");
}

/** Reads every document in a Firestore collection using the runner's Google sign-in. */
function fetchCollection_(name) {
  const out = [];
  let pageToken = "";
  do {
    const url = "https://firestore.googleapis.com/v1/projects/" + PROJECT_ID +
      "/databases/(default)/documents/" + name + "?pageSize=300" +
      (pageToken ? "&pageToken=" + encodeURIComponent(pageToken) : "");
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken(), "X-Goog-User-Project": PROJECT_ID },
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code === 403 || code === 401) {
      let why = "";
      try { why = JSON.parse(res.getContentText()).error.message; } catch (e) { why = res.getContentText().slice(0, 300); }
      throw new Error("Firebase refused access (" + code + "). Google says: " + why +
        " | Checks: appsscript.json must include the datastore scope, and the account running this (" +
        Session.getEffectiveUser().getEmail() + ") must be an owner or editor of the Firebase project.");
    }
    if (code !== 200) throw new Error("Firebase returned " + code + ": " + res.getContentText().slice(0, 200));
    const body = JSON.parse(res.getContentText());
    (body.documents || []).forEach(d => out.push(fromFirestoreFields_(d.fields || {})));
    pageToken = body.nextPageToken || "";
  } while (pageToken);
  return out;
}

/** Firestore REST typed values -> plain JS values. */
function fromFirestoreValue_(v) {
  if (v == null) return null;
  if ("stringValue" in v) return v.stringValue;
  if ("integerValue" in v) return Number(v.integerValue);
  if ("doubleValue" in v) return Number(v.doubleValue);
  if ("booleanValue" in v) return v.booleanValue;
  if ("nullValue" in v) return null;
  if ("timestampValue" in v) return new Date(v.timestampValue);
  if ("arrayValue" in v) return (v.arrayValue.values || []).map(fromFirestoreValue_);
  if ("mapValue" in v) return fromFirestoreFields_(v.mapValue.fields || {});
  return null;
}
function fromFirestoreFields_(fields) {
  const o = {};
  Object.keys(fields).forEach(k => { o[k] = fromFirestoreValue_(fields[k]); });
  return o;
}

/** Flattens one match report and adds the estimated points. */
function normalizeMatch_(m) {
  const sf = m.shiftFuel || [];
  const sr = m.shiftRole || [];
  const n = x => Number(x) || 0;
  const row = {
    event: m.event || "", match: n(m.match), team: n(m.team), alliance: m.alliance || "",
    wonAuto: m.wonAuto || "", autoFuel: n(m.autoFuel), autoClimb: m.autoClimb ? "Yes" : "No",
    transFuel: m.transFuel == null ? "" : n(m.transFuel),
    s1Fuel: sf.length ? n(sf[0]) : "", s2Fuel: sf.length ? n(sf[1]) : "",
    s3Fuel: sf.length ? n(sf[2]) : "", s4Fuel: sf.length ? n(sf[3]) : "",
    teleFuel: n(m.teleFuel), endFuel: n(m.endFuel), climb: m.climb || "none",
    s1Role: sr[0] || m.inactiveRole || "", s2Role: sr[1] || "", s3Role: sr[2] || "", s4Role: sr[3] || "",
    defense: n(m.defense), driver: n(m.driver), robot: m.robot || "",
    minor: n(m.minor), major: n(m.major), scout: m.scout || "", notes: m.notes || "",
    saved: m.ts ? new Date(m.ts) : ""
  };
  row.estPts = estPoints_(row, m.autoClimb);
  return row;
}
function estPoints_(r, autoClimb) {
  return r.autoFuel + (autoClimb ? AUTO_CLIMB_PTS : 0) + r.teleFuel + r.endFuel + (CLIMB_PTS[r.climb] || 0);
}

function writeTable_(sheetName, cols, rows) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName) || SpreadsheetApp.getActive().insertSheet(sheetName);
  sh.clearContents();
  const data = [cols.map(c => c[1])].concat(rows.map(r => cols.map(c => r[c[0]] == null ? "" : r[c[0]])));
  sh.getRange(1, 1, data.length, cols.length).setValues(data);
  sh.getRange(1, 1, 1, cols.length).setFontWeight("bold").setBackground("#810f27").setFontColor("#ffffff");
  sh.setFrozenRows(1);
}

/* ---------------- team averages ---------------- */

function buildAverages_() {
  const ss = SpreadsheetApp.getActive();
  const avg = ss.getSheetByName(AVG_SHEET);
  const ms = ss.getSheetByName(MATCH_SHEET);
  const ps = ss.getSheetByName(PIT_SHEET);
  if (!avg || !ms) return;

  const matches = readTable_(ms, MATCH_COLS);
  const pits = ps ? readTable_(ps, PIT_COLS) : [];

  // dropdowns
  const events = Array.from(new Set(matches.map(m => m.event).filter(String))).sort();
  setDropdown_(avg.getRange("B1"), ["All events"].concat(events));
  setDropdown_(avg.getRange("B2"), AVG_COLS);
  setDropdown_(avg.getRange("B3"), ["Highest first", "Lowest first"]);

  const event = avg.getRange("B1").getValue() || "All events";
  const sortBy = AVG_COLS.indexOf(avg.getRange("B2").getValue());
  const desc = avg.getRange("B3").getValue() !== "Lowest first";

  const rows = computeAverages_(matches.filter(m => event === "All events" || m.event === event), pits);
  const col = sortBy < 0 ? 2 : sortBy;
  const climbRank = { None: 0, L1: 1, L2: 2, L3: 3 };
  const key = r => col === 13 ? (climbRank[r[13]] || 0) : r[col];
  rows.sort((a, b) => {
    const x = key(a), y = key(b);
    if (x === y) return a[0] - b[0];
    if (x === "" || x == null) return 1;
    if (y === "" || y == null) return -1;
    return (x > y ? 1 : -1) * (desc ? -1 : 1);
  });

  const start = 5;
  const last = avg.getMaxRows();
  if (last >= start) avg.getRange(start, 1, last - start + 1, avg.getMaxColumns()).clearContent();
  avg.getRange(start, 1, 1, AVG_COLS.length).setValues([AVG_COLS])
    .setFontWeight("bold").setBackground("#810f27").setFontColor("#ffffff").setWrap(true);
  if (rows.length) {
    avg.getRange(start + 1, 1, rows.length, AVG_COLS.length).setValues(rows);
    avg.getRange(start + 1, 3, rows.length, 15).setNumberFormat("0.0");
    avg.getRange(start + 1, 6, rows.length, 1).setNumberFormat("0%");
    avg.getRange(start + 1, 11, rows.length, 3).setNumberFormat("0%");
    avg.getRange(start + 1, 15, rows.length, 1).setNumberFormat("0%");
    avg.getRange(start + 1, 2, rows.length, 1).setNumberFormat("0");
    avg.getRange(start + 1, 4, rows.length, 1).setNumberFormat("0");
    avg.getRange(start + 1, 18, rows.length, 1).setNumberFormat("0");
  } else {
    avg.getRange(start + 1, 1).setValue("No match reports yet" + (event !== "All events" ? " for " + event : "") + ".");
  }
  avg.setFrozenRows(start);
  avg.setFrozenColumns(1);
  avg.getRange("A1:B3").setFontSize(11);
  avg.getRange("B1:B3").setBackground("#f6efd9");
}

/** Pure: match rows (sheet objects) + pit rows -> averages table rows (no header). */
function computeAverages_(matches, pits) {
  const byTeam = {};
  matches.forEach(m => { if (m.team) (byTeam[m.team] = byTeam[m.team] || []).push(m); });
  const pitBy = {};
  pits.forEach(p => { pitBy[p.team] = p; });
  const avg = a => a.length ? a.reduce((s, x) => s + (Number(x) || 0), 0) / a.length : 0;
  const pct = (list, f) => list.filter(f).length / list.length;
  return Object.keys(byTeam).map(t => {
    const L = byTeam[t];
    const climbs = L.map(m => m.climb);
    const top = ["L3", "L2", "L1"].find(c => climbs.indexOf(c) >= 0) || "None";
    const trans = L.filter(m => m.transFuel !== "" && m.transFuel != null);
    const p = pitBy[t] || {};
    return [
      Number(t), L.length,
      avg(L.map(m => m.estPts)), Math.max.apply(null, L.map(m => Number(m.estPts) || 0)),
      avg(L.map(m => m.autoFuel)), pct(L, m => m.autoClimb === "Yes"),
      trans.length ? avg(trans.map(m => m.transFuel)) : "",
      avg(L.map(m => m.teleFuel)), avg(L.map(m => m.endFuel)),
      avg(L.map(m => (Number(m.autoFuel) || 0) + (Number(m.teleFuel) || 0) + (Number(m.endFuel) || 0))),
      pct(L, m => ["L1", "L2", "L3"].indexOf(m.climb) >= 0),
      pct(L, m => ["L2", "L3"].indexOf(m.climb) >= 0),
      pct(L, m => m.climb === "L3"),
      top,
      pct(L, m => Number(m.defense) > 0),
      avg(L.map(m => (Number(m.minor) || 0) * 5 + (Number(m.major) || 0) * 15)),
      avg(L.map(m => m.driver)),
      L.filter(m => m.robot && m.robot !== "ok").length,
      p.drive || "", p.shooter || "", p.capacity === "" || p.capacity == null ? "" : p.capacity
    ];
  });
}

function readTable_(sh, cols) {
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const idx = {};
  v[0].forEach((h, i) => { idx[h] = i; });
  return v.slice(1).map(r => {
    const o = {};
    cols.forEach(c => { o[c[0]] = idx[c[1]] == null ? "" : r[idx[c[1]]]; });
    return o;
  });
}

function setDropdown_(range, values) {
  const current = range.getValue();
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build());
  if (values.indexOf(current) < 0) range.setValue(values[0]);
}


/* ---------------- fuel accuracy vs The Blue Alliance ---------------- */

const ACC_COLS = [
  "Event", "Match", "Alliance", "Teams", "Robots scouted", "Scouts",
  "Scouted auto", "Official auto", "Auto % error",
  "Scouted teleop", "Official teleop", "Teleop % error",
  "Scouted end game", "Official end game", "End game % error",
  "Scouted total", "Official total", "Difference", "Total % error", "Status"
];

/** Reads the Blue Alliance key the app saved in Firestore (settings/tba). */
function getTbaKey_() {
  const url = "https://firestore.googleapis.com/v1/projects/" + PROJECT_ID + "/databases/(default)/documents/settings/tba";
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken(), "X-Goog-User-Project": PROJECT_ID },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return "";
  const f = JSON.parse(res.getContentText()).fields || {};
  return f.key ? fromFirestoreValue_(f.key) : "";
}

function fetchTbaMatches_(eventKey, tbaKey) {
  const res = UrlFetchApp.fetch("https://www.thebluealliance.com/api/v3/event/" + eventKey + "/matches", {
    headers: { "X-TBA-Auth-Key": tbaKey }, muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code === 404) return null;            // event code isn't a Blue Alliance event
  if (code === 401) throw new Error("The Blue Alliance rejected the saved key. Re-save it in the app (Teams tab > Blue Alliance key).");
  if (code !== 200) throw new Error("The Blue Alliance returned " + code + " for " + eventKey + ".");
  return JSON.parse(res.getContentText());
}

function buildAccuracy_(scoutRows, sched) {
  const events = Array.from(new Set(scoutRows.map(r => r.event).filter(e => /^\d{4}[A-Za-z0-9]+$/.test(e))));
  if (!events.length) { writeAccuracyMessage_("No events with a Blue Alliance event code yet. Use the event's Blue Alliance key (like 2026CABL) as the event code in the app."); return; }
  const tbaKey = getTbaKey_();
  if (!tbaKey) { writeAccuracyMessage_("No Blue Alliance key saved yet. Save it in the app: Teams tab > Blue Alliance key."); return; }
  let rows = [], skipped = [];
  events.forEach(ev => {
    const tba = fetchTbaMatches_(ev.toLowerCase(), tbaKey);
    if (!tba) { skipped.push(ev); return; }
    if (sched) addTbaSchedule_(sched, ev, tba);
    rows = rows.concat(computeAccuracy_(tba, scoutRows.filter(r => r.event === ev), ev));
  });
  writeAccuracy_(rows, skipped);
  writeScoutAccuracy_(computeScoutAccuracy_(rows));
}

/**
 * Pure: TBA matches + scouting rows for one event -> one row per played qualification
 * match alliance that has at least one scouted robot. Official fuel uses hubScore points
 * (1 point per FUEL scored in an active HUB), which is what scouts count.
 */
function computeAccuracy_(tbaMatches, scoutRows, eventCode) {
  const byMatchTeam = {};
  scoutRows.forEach(r => { byMatchTeam[r.match + "_" + r.team] = r; });
  const pctErr = (s, o) => o > 0 ? (s - o) / o : (s === 0 ? 0 : "");
  const out = [];
  tbaMatches
    .filter(m => m.comp_level === "qm" && m.score_breakdown)
    .sort((a, b) => a.match_number - b.match_number)
    .forEach(m => {
      ["red", "blue"].forEach(color => {
        const bd = m.score_breakdown[color]; if (!bd || !bd.hubScore) return;
        const h = bd.hubScore;
        const n = k => Number(h[k]) || 0;
        const offAuto = n("autoPoints");
        const offTele = n("transitionPoints") + n("shift1Points") + n("shift2Points") + n("shift3Points") + n("shift4Points");
        const offEnd = n("endgamePoints");
        const teams = (m.alliances[color].team_keys || []).map(k => parseInt(String(k).replace(/^frc/, ""), 10));
        const reps = teams.map(t => byMatchTeam[m.match_number + "_" + t]).filter(Boolean);
        if (!reps.length) return;
        const sum = f => reps.reduce((t, r) => t + (Number(r[f]) || 0), 0);
        const sAuto = sum("autoFuel"), sTele = sum("teleFuel"), sEnd = sum("endFuel");
        const sTot = sAuto + sTele + sEnd, oTot = offAuto + offTele + offEnd;
        const missing = teams.length - reps.length;
        const row = [
          eventCode, m.match_number, color === "red" ? "Red" : "Blue", teams.join(", "),
          reps.length + " of " + teams.length,
          Array.from(new Set(reps.map(r => r.scout).filter(String))).join(", "),
          sAuto, offAuto, pctErr(sAuto, offAuto),
          sTele, offTele, pctErr(sTele, offTele),
          sEnd, offEnd, pctErr(sEnd, offEnd),
          sTot, oTot, sTot - oTot, pctErr(sTot, oTot),
          missing ? "Missing " + missing + " robot" + (missing > 1 ? "s" : "") + " (error will read low)" : "Complete"
        ];
        row.scouts = Array.from(new Set(reps.map(r => String(r.scout || "").trim()).filter(String)));
        out.push(row);
      });
    });
  return out;
}

function accuracySheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(ACC_SHEET) || ss.insertSheet(ACC_SHEET, 1);
}

function writeAccuracyMessage_(msg) {
  const sh = accuracySheet_();
  sh.clear();
  sh.getRange("A1").setValue("Fuel Accuracy").setFontWeight("bold").setFontSize(14);
  sh.getRange("A2").setValue(msg);
  const sc = scoutSheet_();
  sc.clear();
  sc.getRange("A1").setValue("Scouter Accuracy").setFontWeight("bold").setFontSize(14);
  sc.getRange("A2").setValue(msg);
}

function writeAccuracy_(rows, skipped) {
  const sh = accuracySheet_();
  sh.clear();
  const complete = rows.filter(r => r[19] === "Complete" && r[18] !== "");
  const avgAbs = complete.length ? complete.reduce((t, r) => t + Math.abs(r[18]), 0) / complete.length : "";
  sh.getRange("A1").setValue("Fuel Accuracy").setFontWeight("bold").setFontSize(14);
  sh.getRange("A2:B4").setValues([
    ["Alliances checked (all 3 robots scouted)", complete.length],
    ["Average total % error (absolute)", avgAbs],
    ["Color key", "Green within 10%, yellow within 25%, red over 25%. Positive % = scouts counted too many."]
  ]);
  sh.getRange("A2:A4").setFontWeight("bold");
  sh.getRange("B3").setNumberFormat("0.0%");
  if (skipped.length) sh.getRange("A5").setValue("Not on The Blue Alliance (skipped): " + skipped.join(", ")).setFontColor("#6d5f62");

  const start = 7;
  sh.getRange(start, 1, 1, ACC_COLS.length).setValues([ACC_COLS])
    .setFontWeight("bold").setBackground("#810f27").setFontColor("#ffffff").setWrap(true);
  sh.setFrozenRows(start);
  if (!rows.length) { sh.getRange(start + 1, 1).setValue("No played matches with scouting reports yet."); return; }

  sh.getRange(start + 1, 1, rows.length, ACC_COLS.length).setValues(rows);
  [9, 12, 15, 19].forEach(c => sh.getRange(start + 1, c, rows.length, 1).setNumberFormat("+0.0%;-0.0%;0.0%"));
  const color = v => {
    if (v === "" || v == null) return null;
    const a = Math.abs(v);
    return a <= 0.10 ? "#d9f2e3" : a <= 0.25 ? "#fbf0c9" : "#f8d4d7";
  };
  [9, 12, 15, 19].forEach(c => {
    const vals = sh.getRange(start + 1, c, rows.length, 1).getValues();
    sh.getRange(start + 1, c, rows.length, 1).setBackgrounds(vals.map(r => [color(r[0])]));
  });
  sh.getRange(start + 1, 20, rows.length, 1).setFontColor("#6d5f62");
  sh.autoResizeColumns(1, 6);
}


/* ---------------- scouter accuracy ---------------- */

const SCOUT_COLS = [
  "Scouter", "Alliances checked", "Avg total % error", "Avg auto % error", "Avg teleop % error",
  "Avg end game % error", "Worst total % error", "Worst match", "Tends to count", "Note"
];

/**
 * Pure: Fuel Accuracy rows -> one row per scouter, most accurate first.
 * Every error is taken as an absolute value before averaging, so +11% and -11% average to 11%, not 0.
 * Only alliances with all 3 robots scouted count. Each scouter on that alliance is given the
 * alliance's error, because official fuel is only reported per alliance.
 */
function computeScoutAccuracy_(accRows) {
  const by = {};
  const has = v => v !== "" && v != null;
  accRows.filter(r => r[19] === "Complete").forEach(r => {
    (r.scouts || []).forEach(name => {
      const key = name.toLowerCase();
      const s = by[key] || (by[key] = { name: name, total: [], auto: [], tele: [], end: [], signed: [], worst: null });
      if (has(r[8])) s.auto.push(Math.abs(r[8]));
      if (has(r[11])) s.tele.push(Math.abs(r[11]));
      if (has(r[14])) s.end.push(Math.abs(r[14]));
      if (has(r[18])) {
        s.total.push(Math.abs(r[18]));
        s.signed.push(r[18]);
        if (!s.worst || Math.abs(r[18]) > Math.abs(s.worst[18])) s.worst = r;
      }
    });
  });
  const avg = a => a.length ? a.reduce((t, x) => t + x, 0) / a.length : "";
  return Object.keys(by).map(k => {
    const s = by[k];
    const bias = avg(s.signed);
    const tends = s.signed.length < 2 ? "" : bias > 0.05 ? "Too many" : bias < -0.05 ? "Too few" : "About even";
    return [
      s.name, s.total.length, avg(s.total), avg(s.auto), avg(s.tele), avg(s.end),
      s.worst ? Math.abs(s.worst[18]) : "",
      s.worst ? s.worst[0] + " match " + s.worst[1] + " " + s.worst[2] : "",
      tends,
      s.total.length < MIN_CHECKS ? "Fewer than " + MIN_CHECKS + " checks, treat as rough" : ""
    ];
  }).filter(r => r[1] > 0).sort((a, b) => a[2] - b[2] || b[1] - a[1]);
}

function scoutSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SCOUT_SHEET) || ss.insertSheet(SCOUT_SHEET, 2);
}

function writeScoutAccuracy_(rows) {
  const sh = scoutSheet_();
  sh.clear();
  sh.getRange("A1").setValue("Scouter Accuracy").setFontWeight("bold").setFontSize(14);
  sh.getRange("A2:A4").setValues([
    ["Average fuel error for the alliances each scouter helped scout, compared with The Blue Alliance. Most accurate first."],
    ["Errors are absolute values, so over and under counts don't cancel out. Only alliances with all 3 robots scouted are counted."],
    ["Official fuel is per alliance, so each number includes the other two scouters' mistakes. It gets fairer with more matches and when scouters rotate partners."]
  ]).setFontColor("#6d5f62");

  const start = 6;
  sh.getRange(start, 1, 1, SCOUT_COLS.length).setValues([SCOUT_COLS])
    .setFontWeight("bold").setBackground("#810f27").setFontColor("#ffffff").setWrap(true);
  sh.setFrozenRows(start);
  if (!rows.length) { sh.getRange(start + 1, 1).setValue("No alliances with all 3 robots scouted and an official score yet."); return; }

  sh.getRange(start + 1, 1, rows.length, SCOUT_COLS.length).setValues(rows);
  sh.getRange(start + 1, 3, rows.length, 5).setNumberFormat("0.0%");
  const color = v => {
    if (v === "" || v == null) return null;
    return v <= 0.10 ? "#d9f2e3" : v <= 0.25 ? "#fbf0c9" : "#f8d4d7";
  };
  sh.getRange(start + 1, 3, rows.length, 1).setBackgrounds(rows.map(r => [color(r[2])]));
  sh.getRange(start + 1, 10, rows.length, 1).setFontColor("#6d5f62");
  sh.autoResizeColumns(1, 1);
  sh.autoResizeColumns(8, 3);
}


/* ---------------- match schedule (for Match Preview) ---------------- */

const SCHED_COLS = [
  ["event", "Event"], ["match", "Match"], ["red1", "Red 1"], ["red2", "Red 2"], ["red3", "Red 3"],
  ["blue1", "Blue 1"], ["blue2", "Blue 2"], ["blue3", "Blue 3"]
];

function schedRow_(event, match, red, blue) {
  const t = (a, i) => (a && a[i] != null && a[i] !== "") ? Number(a[i]) : "";
  return { event: event, match: Number(match), red1: t(red, 0), red2: t(red, 1), red3: t(red, 2),
           blue1: t(blue, 0), blue2: t(blue, 1), blue3: t(blue, 2) };
}

/** Schedules the app saved in Firestore (Teams tab > Load schedule). */
function addAppSchedules_(sched, docs) {
  docs.forEach(d => {
    const ev = String(d.eventKey || "").toUpperCase();
    if (!ev) return;
    (d.matches || []).forEach(m => { sched[ev + "|" + m.n] = schedRow_(ev, m.n, m.red, m.blue); });
  });
}

/** Qualification schedule straight from The Blue Alliance match list. */
function addTbaSchedule_(sched, ev, tbaMatches) {
  const nums = keys => (keys || []).map(k => parseInt(String(k).replace(/^frc/, ""), 10));
  tbaMatches.filter(m => m.comp_level === "qm").forEach(m => {
    sched[ev + "|" + m.match_number] = schedRow_(ev, m.match_number, nums(m.alliances.red.team_keys), nums(m.alliances.blue.team_keys));
  });
}

/* ---------------- match preview ---------------- */

const PV_TEAM_ROW = 6;
const PV_FIRST_STAT_ROW = 7;
const PV_LIST_COL = 12;         // hidden column L: match numbers that feed the Match # dropdown
const RP_ENERGIZED = 100, RP_SUPERCHARGED = 360, RP_TRAVERSAL = 50;   // regional thresholds from the game manual
const ROLE_LABELS = { defense: "Defense", collect: "Collects fuel", feedMid: "Feeds from midzone",
                      feedOpp: "Feeds from opponent zone", feed: "Feeds partners", idle: "Idle" };

// agg: how the alliance column is built. better: which alliance value gets highlighted.
const PREVIEW_STATS = [
  { label: "Matches scouted",               key: "n",            fmt: "0" },
  { label: "Avg est. points",               key: "avgPts",       fmt: "0.0", agg: "sum", better: "high" },
  { label: "Best est. points",              key: "bestPts",      fmt: "0",   agg: "sum", better: "high" },
  { label: "Auto fuel, average",            key: "autoAvg",      fmt: "0.0", agg: "sum", better: "high" },
  { label: "Auto fuel, best",               key: "autoBest",     fmt: "0",   agg: "sum", better: "high" },
  { label: "Teleop fuel, average",          key: "teleAvg",      fmt: "0.0", agg: "sum", better: "high" },
  { label: "Teleop fuel, best",             key: "teleBest",     fmt: "0",   agg: "sum", better: "high" },
  { label: "End game fuel, average",        key: "endAvg",       fmt: "0.0", agg: "sum", better: "high" },
  { label: "End game fuel, best",           key: "endBest",      fmt: "0",   agg: "sum", better: "high" },
  { label: "Total fuel, average",           key: "totAvg",       fmt: "0.0", agg: "sum", better: "high" },
  { label: "Total fuel, best",              key: "totBest",      fmt: "0",   agg: "sum", better: "high" },
  { label: "Avg foul points given up",      key: "foulAvg",      fmt: "0.0", agg: "sum", better: "low" },
  { label: "Avg driver rating (1 to 5)",    key: "driverAvg",    fmt: "0.0", agg: "avg", better: "high" },
  { label: "Matches with robot issues",     key: "issues",       fmt: "0",   agg: "sum", better: "low" }
];

/** Pure: one team's preview numbers from Match Data rows (+ its pit row). */
function teamPreviewStats_(team, rows, pit) {
  const L = rows.filter(r => Number(r.team) === Number(team));
  const p = pit || {};
  const base = { team: team, n: L.length, drive: p.drive || "", shooter: p.shooter || "",
                 capacity: (p.capacity === "" || p.capacity == null) ? "" : p.capacity };
  if (!L.length) return base;
  const num = (r, f) => Number(r[f]) || 0;
  const avg = f => L.reduce((t, r) => t + f(r), 0) / L.length;
  const max = f => Math.max.apply(null, L.map(f));
  const tot = r => num(r, "autoFuel") + num(r, "teleFuel") + num(r, "endFuel");
  const tower = r => (r.autoClimb === "Yes" ? AUTO_CLIMB_PTS : 0) + (CLIMB_PTS[r.climb] || 0);
  const roles = {};
  L.forEach(r => ["s1Role", "s2Role", "s3Role", "s4Role"].forEach(k => { if (r[k]) roles[r[k]] = (roles[r[k]] || 0) + 1; }));
  const topRole = Object.keys(roles).sort((a, b) => roles[b] - roles[a])[0];
  const climbs = L.map(r => r.climb);
  return Object.assign(base, {
    avgPts: avg(r => num(r, "estPts")), bestPts: max(r => num(r, "estPts")),
    autoAvg: avg(r => num(r, "autoFuel")), autoBest: max(r => num(r, "autoFuel")),
    teleAvg: avg(r => num(r, "teleFuel")), teleBest: max(r => num(r, "teleFuel")),
    endAvg: avg(r => num(r, "endFuel")), endBest: max(r => num(r, "endFuel")),
    totAvg: avg(tot), totBest: max(tot),
    autoClimbPct: L.filter(r => r.autoClimb === "Yes").length / L.length,
    climbPct: L.filter(r => ["L1", "L2", "L3"].indexOf(r.climb) >= 0).length / L.length,
    topClimb: ["L3", "L2", "L1"].find(c => climbs.indexOf(c) >= 0) || "None",
    towerAvg: avg(tower),
    defensePct: L.filter(r => num(r, "defense") > 0).length / L.length,
    foulAvg: avg(r => num(r, "minor") * 5 + num(r, "major") * 15),
    driverAvg: avg(r => num(r, "driver")),
    issues: L.filter(r => r.robot && r.robot !== "ok").length,
    role: topRole ? (ROLE_LABELS[topRole] || topRole) : ""
  });
}

/**
 * Pure: red and blue team lists + Match Data rows + Pit rows -> the preview table.
 * Row shape: [label, red1, red2, red3, red alliance, blue alliance, blue1, blue2, blue3].
 */
function computePreview_(redTeams, blueTeams, rows, pits) {
  const pitBy = {};
  pits.forEach(p => { pitBy[Number(p.team)] = p; });
  const stats = t => (t === "" || t == null) ? null : teamPreviewStats_(t, rows, pitBy[Number(t)]);
  const R = redTeams.map(stats), B = blueTeams.map(stats);
  const val = (st, key) => (!st || st[key] == null) ? "" : st[key];
  const agg = (side, def) => {
    if (!def.agg) return "";
    const vals = side.filter(st => st && st.n > 0).map(st => Number(st[def.key]) || 0);
    if (!vals.length) return "";
    const sum = vals.reduce((t, x) => t + x, 0);
    return def.agg === "avg" ? sum / vals.length : sum;
  };
  const table = [], formats = [], better = [];
  PREVIEW_STATS.forEach(def => {
    const ra = agg(R, def), ba = agg(B, def);
    table.push([def.label, val(R[0], def.key), val(R[1], def.key), val(R[2], def.key), ra, ba,
                val(B[0], def.key), val(B[1], def.key), val(B[2], def.key)]);
    formats.push(def.fmt);
    better.push(!def.better || ra === "" || ba === "" || ra === ba ? null
      : ((def.better === "high") === (ra > ba) ? "red" : "blue"));
  });
  const total = (side, key) => side.filter(st => st && st.n > 0).reduce((t, st) => t + (Number(st[key]) || 0), 0);
  const sum = { redPts: total(R, "avgPts"), bluePts: total(B, "avgPts"), redFuel: total(R, "totAvg"), blueFuel: total(B, "totAvg"),
                redTower: total(R, "towerAvg"), blueTower: total(B, "towerAvg") };
  const hasData = side => side.some(st => st && st.n > 0);
  const yn = (side, ok) => !hasData(side) ? "" : ok ? "Yes" : "No";
  const rp = (label, r, b) => { table.push([label, "", "", "", r, b, "", "", ""]); formats.push("@"); better.push(null); };
  rp("On pace for Energized RP (" + RP_ENERGIZED + " fuel)", yn(R, sum.redFuel >= RP_ENERGIZED), yn(B, sum.blueFuel >= RP_ENERGIZED));
  rp("On pace for Supercharged RP (" + RP_SUPERCHARGED + " fuel)", yn(R, sum.redFuel >= RP_SUPERCHARGED), yn(B, sum.blueFuel >= RP_SUPERCHARGED));
  rp("On pace for Traversal RP (" + RP_TRAVERSAL + " tower points)", yn(R, sum.redTower >= RP_TRAVERSAL), yn(B, sum.blueTower >= RP_TRAVERSAL));

  const all = R.concat(B).filter(Boolean);
  const noData = all.filter(st => st.n === 0).map(st => st.team);
  let headline;
  if (!all.length) headline = "Pick a match, or type team numbers in row " + PV_TEAM_ROW + ".";
  else if (all.every(st => st.n === 0)) headline = "No scouting data yet for these teams.";
  else {
    const d = sum.redPts - sum.bluePts;
    headline = "Projected score from scouting averages: Red " + sum.redPts.toFixed(1) + ", Blue " + sum.bluePts.toFixed(1) + ". " +
      (Math.abs(d) < 0.05 ? "Even." : (d > 0 ? "Red" : "Blue") + " favored by " + Math.abs(d).toFixed(1) + ".");
    if (noData.length) headline += " No data for " + noData.join(", ") + ", so that alliance reads low.";
  }
  return { table: table, formats: formats, better: better, headline: headline, summary: sum };
}

function previewSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(PREVIEW_SHEET) || ss.insertSheet(PREVIEW_SHEET, 1);
}

/**
 * Draws the Match Preview tab from the Schedule, Match Data and Pit Data tabs (no internet needed,
 * so it also runs from the dropdowns). fillTeams = true replaces the team row with the picked match.
 */
function buildPreview_(fillTeams) {
  const ss = SpreadsheetApp.getActive();
  const sh = previewSheet_();
  const read = (name, cols) => { const t = ss.getSheetByName(name); return t ? readTable_(t, cols) : []; };
  const matches = read(MATCH_SHEET, MATCH_COLS), pits = read(PIT_SHEET, PIT_COLS), sched = read(SCHED_SHEET, SCHED_COLS);

  // controls
  sh.getRange("A1:A3").setValues([["Event"], ["Match #"], ["Stats from"]]).setFontWeight("bold");
  const events = Array.from(new Set(sched.map(r => r.event).concat(matches.map(r => r.event)).filter(String))).sort();
  if (events.length) setDropdown_(sh.getRange("B1"), events);
  setDropdown_(sh.getRange("B3"), ["This event", "All events"]);
  sh.getRange("B1:B3").setBackground("#f6efd9").setHorizontalAlignment("left");
  const event = String(sh.getRange("B1").getValue() || "");
  const evSched = sched.filter(r => r.event === event && !isNaN(Number(r.match)))
    .sort((a, b) => Number(a.match) - Number(b.match));
  const mCell = sh.getRange("B2");
  // The dropdown reads its choices from a hidden column of real numbers, already in numeric order,
  // so it lists 1, 2, 3 ... 10, 11 instead of 1, 10, 100.
  sh.getRange(1, PV_LIST_COL, sh.getMaxRows(), 1).clearContent();
  if (evSched.length) {
    const listRange = sh.getRange(1, PV_LIST_COL, evSched.length, 1);
    listRange.setValues(evSched.map(r => [Number(r.match)])).setNumberFormat("0");
    mCell.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(listRange, true).setAllowInvalid(true).build());
    if (mCell.getValue() === "") mCell.setValue(Number(evSched[0].match));
  } else {
    mCell.clearDataValidations();
  }
  sh.hideColumns(PV_LIST_COL);
  const matchNo = Number(mCell.getValue()) || "";

  // teams: from the schedule when a match is picked, otherwise whatever is typed in the team row
  const cur = sh.getRange(PV_TEAM_ROW, 2, 1, 8).getValues()[0];
  let red = [cur[0], cur[1], cur[2]], blue = [cur[5], cur[6], cur[7]];
  const blank = red.concat(blue).every(v => v === "" || v == null);
  const row = evSched.filter(r => Number(r.match) === matchNo)[0];
  if ((fillTeams || blank) && row) {
    red = [row.red1, row.red2, row.red3]; blue = [row.blue1, row.blue2, row.blue3];
  } else if (fillTeams && evSched.length && !row) {
    red = ["", "", ""]; blue = ["", "", ""];
  }

  const statsRows = sh.getRange("B3").getValue() === "All events" ? matches : matches.filter(r => r.event === event);
  const pv = computePreview_(red, blue, statsRows, pits);

  let note = pv.headline;
  if (!evSched.length) note += " No schedule for " + (event || "this event") + " yet: type the six team numbers in row " + PV_TEAM_ROW + ".";
  else if (!row && matchNo) note += " Match " + matchNo + " isn't in the schedule.";
  sh.getRange("A4").setValue(note).setFontWeight("bold");
  sh.getRange("D1").setValue("Type over a team number in row " + PV_TEAM_ROW + " to try a different lineup.").setFontColor("#6d5f62");

  // header rows
  const RED = "#d6333d", BLUE = "#2c63d6";
  sh.getRange(PV_TEAM_ROW - 1, 1, 2, 9).setValues([
    ["", "Red 1", "Red 2", "Red 3", "Red alliance", "Blue alliance", "Blue 1", "Blue 2", "Blue 3"],
    ["Team #", red[0], red[1], red[2], "RED", "BLUE", blue[0], blue[1], blue[2]]
  ]).setFontWeight("bold").setHorizontalAlignment("center");
  sh.getRange(PV_TEAM_ROW - 1, 2, 1, 4).setBackground(RED).setFontColor("#ffffff");
  sh.getRange(PV_TEAM_ROW - 1, 6, 1, 4).setBackground(BLUE).setFontColor("#ffffff");
  sh.getRange(PV_TEAM_ROW, 2, 1, 4).setBackground("#f5cfd2").setFontSize(12).setNumberFormat("0");
  sh.getRange(PV_TEAM_ROW, 6, 1, 4).setBackground("#cfdcf7").setFontSize(12).setNumberFormat("0");
  sh.getRange(PV_TEAM_ROW, 1).setHorizontalAlignment("left");

  // stats table
  const n = pv.table.length;
  sh.getRange(PV_FIRST_STAT_ROW, 1, Math.max(1, sh.getMaxRows() - PV_FIRST_STAT_ROW + 1), 9).clear();
  const body = sh.getRange(PV_FIRST_STAT_ROW, 1, n, 9);
  body.setValues(pv.table);
  body.setNumberFormats(pv.formats.map(f => ["@", f, f, f, f, f, f, f, f]));
  body.setBackgrounds(pv.better.map(b => [
    "#ffffff", "#fbe9ea", "#fbe9ea", "#fbe9ea", b === "red" ? "#f6d77a" : "#f5cfd2",
    b === "blue" ? "#f6d77a" : "#cfdcf7", "#e8eefb", "#e8eefb", "#e8eefb"
  ]));
  body.setFontWeights(pv.better.map(b => [
    "bold", "normal", "normal", "normal", b === "red" ? "bold" : "normal", b === "blue" ? "bold" : "normal", "normal", "normal", "normal"
  ]));
  sh.getRange(PV_FIRST_STAT_ROW, 2, n, 8).setHorizontalAlignment("center");
  sh.getRange(PV_FIRST_STAT_ROW + n + 1, 1).setValue(
    "Gold marks the alliance with the edge on that stat. Alliance columns add up the three robots (driver rating is averaged). " +
    "Est. points = fuel + tower points. RP rows use regional thresholds and may not apply at offseason events.").setFontColor("#6d5f62");
  sh.setColumnWidth(1, 250);
  sh.setFrozenRows(PV_TEAM_ROW);
}
