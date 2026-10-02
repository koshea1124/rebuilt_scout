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
  const sh = e.range.getSheet();
  if (sh.getName() !== AVG_SHEET) return;
  if (e.range.getRow() <= 3 && e.range.getColumn() === 2) buildAverages_();
}

/** Run once: creates the tabs, the 5-minute sync, and does the first sync. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  [AVG_SHEET, ACC_SHEET, MATCH_SHEET, PIT_SHEET].forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
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
  try { buildAccuracy_(matches); }
  catch (e) { writeAccuracyMessage_("Couldn't check accuracy: " + e.message); }
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

function buildAccuracy_(scoutRows) {
  const events = Array.from(new Set(scoutRows.map(r => r.event).filter(e => /^\d{4}[A-Za-z0-9]+$/.test(e))));
  if (!events.length) { writeAccuracyMessage_("No events with a Blue Alliance event code yet. Use the event's Blue Alliance key (like 2026CABL) as the event code in the app."); return; }
  const tbaKey = getTbaKey_();
  if (!tbaKey) { writeAccuracyMessage_("No Blue Alliance key saved yet. Save it in the app: Teams tab > Blue Alliance key."); return; }
  let rows = [], skipped = [];
  events.forEach(ev => {
    const tba = fetchTbaMatches_(ev.toLowerCase(), tbaKey);
    if (!tba) { skipped.push(ev); return; }
    rows = rows.concat(computeAccuracy_(tba, scoutRows.filter(r => r.event === ev), ev));
  });
  writeAccuracy_(rows, skipped);
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
        out.push([
          eventCode, m.match_number, color === "red" ? "Red" : "Blue", teams.join(", "),
          reps.length + " of " + teams.length,
          Array.from(new Set(reps.map(r => r.scout).filter(String))).join(", "),
          sAuto, offAuto, pctErr(sAuto, offAuto),
          sTele, offTele, pctErr(sTele, offTele),
          sEnd, offEnd, pctErr(sEnd, offEnd),
          sTot, oTot, sTot - oTot, pctErr(sTot, oTot),
          missing ? "Missing " + missing + " robot" + (missing > 1 ? "s" : "") + " (error will read low)" : "Complete"
        ]);
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
