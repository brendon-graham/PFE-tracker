// ═══════════════════════════════════════════════════════════════════
// PFE FARM TRACKER — Google Apps Script Backend v5.0  (app v14.0.0)
// v14.0.0: real per-id merge (mergeArraySectionRows/mergeMobsWithTombs) for
//          paddocks/pastureMobs/scenarios/backlogJobs/toolboxMinutesList/
//          barnSchedule + pastureBlocks mobs, instead of blind clear+rewrite
//          on every push. Fixes mobs/paddocks/jobs resurrecting after a
//          stale phone pushed — see applySection.
// v13.2.0: weeklyJobs structure (template + weekStart) guarded by structTs so a
//          stale phone can no longer revert a freshly-rolled week — writeWeeklyMerged.
// ═══════════════════════════════════════════════════════════════════
// Deploy → New deployment → Web app
// Execute as: Me | Who has access: Anyone
// After any Code.gs change: Deploy → New version → re-authorise
// ═══════════════════════════════════════════════════════════════════

// ── Section → sheet tab mapping ───────────────────────────────────────
const SHEETS = {
  // Array sections — one JSON row per record
  paddocks:    "Paddocks",
  silages:     "Silages",
  pastureMobs: "PastureMobs",
  scenarios:   "Scenarios",
  dailyLogs:   "DailyLogs",
  backlogJobs: "BacklogJobs",
  weeklyArchive: "WeeklyArchive",
  toolbox:     "ToolboxMinutes",
  barnSchedule:  "BarnSchedule",
  // Object sections — single JSON row
  pastureBlocks: "PastureBlocks",
  barnCalc:    "BarnCalc",
  checks:      "Checks",
  weeklyJobs:  "WeeklyJobs",
  stockRec:    "StockRec",
  animalHealthOrders: "AnimalHealthOrders",
  // Infrastructure
  lastModified: "LastModified",
  syncLog:      "SyncLog",
  errorLog:     "ErrorLog",
};

// Delta section name (from client) → SHEETS key
const DELTA_SECTION_MAP = {
  paddock:    "paddocks",
  silage:     "silages",
  pasturemob: "pastureMobs",
  barncalc:   "barnCalc",
  weeklyjobs: "weeklyJobs",
  backlog:    "backlogJobs",
  checks:     "checks",
  dailylogs:  "dailyLogs",
  toolbox:    "toolbox",
};

// Sections that store a single object (not an array of records)
const OBJECT_SECTIONS = new Set(["barncalc", "checks"]);

// ── Sync v13 — section-level push (client section key → sheet tab) ─────
// Authoritative list of server-persisted sections. Client sends only the
// sections it changed; the server writes only those, never touching the rest.
const SECTION_SHEETS = {
  paddocks:               "Paddocks",
  silages:                "Silages",
  pastureMobs:            "PastureMobs",
  scenarios:              "Scenarios",
  dailyLogs:              "DailyLogs",
  backlogJobs:            "BacklogJobs",
  weeklyCompletedArchive: "WeeklyArchive",
  toolboxMinutesList:     "ToolboxMinutes",
  barnSchedule:           "BarnSchedule",
  pastureBlocks:          "PastureBlocks",
  barnCalc:               "BarnCalc",
  checks:                 "Checks",
  weeklyJobs:             "WeeklyJobs",
  stockRec:               "StockRec",
  animalHealthOrders:     "AnimalHealthOrders",
};
// Section keys that store an array of records (rest are single objects).
const SECTION_ARRAY = new Set([
  "paddocks", "silages", "pastureMobs", "scenarios", "dailyLogs",
  "backlogJobs", "weeklyCompletedArchive", "toolboxMinutesList", "barnSchedule",
]);
// Array sections that get a real per-id merge on push instead of a blind
// clear+rewrite (v14.0). The client sends each record's `_ts` and encodes
// deletes as `{id,_tomb:true,_ts}` markers in the same array (see
// encodeTombs/decodeTombs in index.html); the server merges those against
// what's already stored so ANY push — even from a phone that hasn't synced
// in a while — can only ever win on a genuinely newer edit, never on being
// "whoever happened to push last". Closes the gap where handleSectionPush
// applied every pushed section unconditionally regardless of client
// staleness (silages/dailyLogs/weeklyCompletedArchive keep their existing,
// separate merge behaviour and are not in this set).
const MERGE_ARRAY_SECTIONS = new Set([
  "paddocks", "pastureMobs", "scenarios", "backlogJobs", "toolboxMinutesList", "barnSchedule",
]);

// SyncLog retention — drop entries older than this
const SYNCLOG_RETAIN_DAYS = 30;

// ══════════════════════════════════════════════════════════════════════
// ROUTING
// ══════════════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const action = (e.parameter.action || "ping").toLowerCase();

    if (action === "ping") {
      return respond({ ok: true, message: "PFE Tracker API v4 running", time: new Date().toISOString() });
    }
    if (action === "pull") {
      return respond({ ok: true, ...handlePull(ss) });
    }
    if (action === "pulldelta") {
      const since = Number(e.parameter.since) || 0;
      return respond({ ok: true, ...handlePullDelta(ss, since) });
    }
    return respond({ ok: true, message: "Unknown GET action: " + action });
  } catch (err) {
    logError("doGet", err);
    return respond({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      return respond({ ok: false, error: "No post body received" });
    }
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const body = JSON.parse(e.postData.contents);

    switch (body.action) {
      case "push":
        // Sync v13: section-level push (only changed sections). Legacy full-blob
        // pushes (body.data, no body.sections) still route to handlePush unchanged.
        if (body.sections) return respond({ ok: true, ...handleSectionPush(ss, body) });
        return respond({ ok: true, ...handlePush(ss, body.data, body.user) });
      case "pushDelta": return respond({ ok: true, ...handlePushDelta(ss, body.changes, body.user) });
      case "pull":      return respond({ ok: true, ...handlePull(ss) });
      case "init":      return respond({ ok: true, ...handleInit(ss) });
      case "claude":    return respond({ ok: true, ...claudeProxy(body) });
      default:          return respond({ ok: false, error: "Unknown action: " + body.action });
    }
  } catch (err) {
    logError("doPost", err);
    return respond({ ok: false, error: err.message });
  }
}

// ══════════════════════════════════════════════════════════════════════
// ACTION HANDLERS
// ══════════════════════════════════════════════════════════════════════

// ── Full pull — returns entire state (used on startup and manual sync) ─
function handlePull(ss) {
  const lmRow = readObjectSection(ss, SHEETS.lastModified);

  return {
    data: {
      paddocks:    readArraySection(ss, SHEETS.paddocks),
      silages:     readArraySection(ss, SHEETS.silages),
      pastureMobs: readArraySection(ss, SHEETS.pastureMobs),
      scenarios:   readArraySection(ss, SHEETS.scenarios),
      dailyLogs:   readArraySection(ss, SHEETS.dailyLogs),
      backlogJobs: readArraySection(ss, SHEETS.backlogJobs),
      weeklyCompletedArchive: readArraySection(ss, SHEETS.weeklyArchive),
      toolboxMinutesList:     readArraySection(ss, SHEETS.toolbox),
      barnSchedule:  readArraySection(ss, SHEETS.barnSchedule),
      pastureBlocks: readObjectSection(ss, SHEETS.pastureBlocks),
      barnCalc:    readObjectSection(ss, SHEETS.barnCalc),
      checks:      readObjectSection(ss, SHEETS.checks),
      weeklyJobs:  readObjectSection(ss, SHEETS.weeklyJobs),
      stockRec:           readObjectSection(ss, SHEETS.stockRec),
      animalHealthOrders: readObjectSection(ss, SHEETS.animalHealthOrders),
      lastModified:   lmRow ? (lmRow.lastModified   || 0) : 0,
      version:        lmRow ? (Number(lmRow.version) || 0) : 0,
      barnScheduleTs: lmRow ? (lmRow.barnScheduleTs || 0) : 0,
      altDayShift:    lmRow ? (Number(lmRow.altDayShift) % 2 || 0) : 0,
      syncTime: new Date().toISOString(),
    }
  };
}

// ── Full push — writes complete state (fallback for uninstrumented changes) ─
function handlePush(ss, data, user) {
  if (!data) throw new Error("Push received empty data payload");
  const userName = user || "Staff";

  // Serialise the read-modify-write so two pushes can't interleave.
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return { skipped: "locked" }; }
  try {

  // Stale-push guard (optimistic concurrency).
  // baseTs = the Sheet version the client was editing from. Reject when the
  // Sheet has moved on since then (currentTs > baseTs) — a stale device must
  // not overwrite a newer write (the 20 Jul 2026 WeeklyJobs corruption).
  // Legacy clients that don't send baseTs fall back to the weaker
  // lastModified-vs-current check.
  const lmRow      = readObjectSection(ss, SHEETS.lastModified);
  const currentTs  = lmRow ? Number(lmRow.lastModified || 0) : 0;
  const baseTs     = Number(data.baseTs || 0);
  const incomingTs = Number(data.lastModified || 0);
  if (baseTs > 0) {
    if (currentTs > baseTs) {
      return { skipped: "stale", serverTs: currentTs };
    }
  } else if (incomingTs > 0 && currentTs > 0 && incomingTs <= currentTs) {
    return { skipped: "stale", serverTs: currentTs };
  }

  const ts = new Date().toISOString();

  if (data.paddocks)    writeArraySection(ss, SHEETS.paddocks,    data.paddocks);
  if (data.silages)     writeArraySection(ss, SHEETS.silages,     data.silages);
  if (data.pastureMobs) writeArraySection(ss, SHEETS.pastureMobs, data.pastureMobs);
  if (data.barnCalc)    writeObjectSection(ss, SHEETS.barnCalc,   data.barnCalc);
  if (data.checks)      writeChecksMerged(ss, SHEETS.checks,     data.checks);
  if (data.scenarios)   writeArraySection(ss, SHEETS.scenarios,   data.scenarios);
  if (data.weeklyJobs)  writeWeeklyMerged(ss, SHEETS.weeklyJobs, data.weeklyJobs);
  if (data.backlogJobs) writeArraySection(ss, SHEETS.backlogJobs, data.backlogJobs);
  if (data.weeklyCompletedArchive) writeArraySection(ss, SHEETS.weeklyArchive, data.weeklyCompletedArchive);
  if (data.toolboxMinutesList)     writeArraySection(ss, SHEETS.toolbox, data.toolboxMinutesList);
  // barnSchedule — only write if this device actually modified it (timestamp guard)
  const curBsTs = Number((lmRow||{}).barnScheduleTs || 0);
  const inBsTs  = Number(data.barnScheduleTs || 0);
  if (data.barnSchedule && !(inBsTs > 0 && curBsTs > 0 && inBsTs <= curBsTs)) {
    writeArraySection(ss, SHEETS.barnSchedule, data.barnSchedule);
  }
  // pastureBlocks — only write if this device actually modified it (timestamp guard)
  if (data.pastureBlocks) {
    const curPbTs = Number((readObjectSection(ss, SHEETS.pastureBlocks) || {})._ts || 0);
    const inPbTs  = Number(data.pastureBlocks._ts || 0);
    if (!(inPbTs > 0 && curPbTs > 0 && inPbTs <= curPbTs)) {
      writeObjectSection(ss, SHEETS.pastureBlocks, data.pastureBlocks);
    }
  }
  if (data.dailyLogs)   mergeDailyLogs(ss, data.dailyLogs);
  if (data.stockRec)           writeObjectSection(ss, SHEETS.stockRec,           data.stockRec);
  if (data.animalHealthOrders) writeObjectSection(ss, SHEETS.animalHealthOrders, data.animalHealthOrders);

  // lastModified stays echoed from the client so v12.8.0 clients' post-push
  // confirmation still matches. version is server-authoritative and advances
  // on every accepted write so v13 clients see the change.
  const lastModified = data.lastModified || Date.now();
  const newVersion   = Number((lmRow || {}).version || 0) + 1;
  const newBsTs = (inBsTs > curBsTs) ? inBsTs : curBsTs;
  const curAltShift = Number((lmRow || {}).altDayShift) % 2 || 0;
  const newAltShift = (data.altDayShift != null) ? (Number(data.altDayShift) % 2 || 0) : curAltShift;
  writeObjectSection(ss, SHEETS.lastModified, { lastModified, version: newVersion, pushedBy: userName, ts, barnScheduleTs: newBsTs, altDayShift: newAltShift });
  appendSyncLog(ss, { ts: lastModified, user: userName, action: "push", section: "full", key: "" });

  return { pushed: ts, version: newVersion };
  } finally {
    lock.releaseLock();
  }
}

// ── Sync v13 — section-level push ─────────────────────────────────────
// Writes ONLY the sections the client changed, so a phone editing one tab can
// never clobber another phone's untouched tab. version is server-authoritative
// (no device clock), and the lock makes the read-modify-write atomic.
function handleSectionPush(ss, body) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return { accepted: false, reason: "locked" }; }
  try {
    const userName    = body.user || "Staff";
    const sections    = body.sections || {};
    const baseVersion = Number(body.baseVersion || 0);
    const lmRow          = readObjectSection(ss, SHEETS.lastModified) || {};
    const currentVersion = Number(lmRow.version || 0);

    // Apply each provided section. Sections not present are left untouched.
    const applied = [];
    Object.keys(sections).forEach(function (key) {
      if (!SECTION_SHEETS[key]) { logError("handleSectionPush", new Error("Unknown section: " + key)); return; }
      applySection(ss, key, sections[key]);
      applied.push(key);
    });

    // Scalars that ride alongside sections (monotonic guards preserved).
    let newBsTs = Number(lmRow.barnScheduleTs || 0);
    if (body.barnScheduleTs != null && Number(body.barnScheduleTs) > newBsTs) newBsTs = Number(body.barnScheduleTs);
    let newAltShift = Number(lmRow.altDayShift) % 2 || 0;
    if (body.altDayShift != null) newAltShift = Number(body.altDayShift) % 2 || 0;

    const newVersion = currentVersion + 1;
    const nowMs      = Date.now();
    writeObjectSection(ss, SHEETS.lastModified, {
      lastModified: nowMs, version: newVersion, pushedBy: userName,
      ts: new Date().toISOString(), barnScheduleTs: newBsTs, altDayShift: newAltShift,
    });
    appendSyncLog(ss, { ts: nowMs, user: userName, action: "sectionPush", section: applied.join("+"), key: "" });

    const resp = { accepted: true, version: newVersion, lastModified: nowMs, pushedBy: userName, applied: applied };
    // Client was behind — hand back the full current state so it can merge in
    // whatever else changed. Its own just-pushed sections stay local-authoritative.
    if (baseVersion > 0 && baseVersion < currentVersion) {
      resp.behind    = true;
      resp.serverData = handlePull(ss).data;
    }
    return resp;
  } finally {
    lock.releaseLock();
  }
}

// Per-id merge for a MERGE_ARRAY_SECTIONS push. Both live records and
// `{id,_tomb:true,_ts}` delete-markers are mixed in one flat array (the
// row-per-item Sheets storage doesn't change shape). Newest `_ts` per id
// wins; a tie prefers a live record over a tombstone so a same-millisecond
// recreate always resolves. Mirrors mergeRecordArray in index.html.
function mergeArraySectionRows(existingArr, incomingArr) {
  const byId = {};
  (existingArr || []).forEach(function(r) { if (r && r.id != null) byId[r.id] = r; });
  (incomingArr || []).forEach(function(inc) {
    if (!inc || inc.id == null) return;
    const cur = byId[inc.id];
    const curTs = cur ? Number(cur._ts || 0) : -1;
    const incTs = Number(inc._ts || 0);
    const curIsTomb = !!(cur && cur._tomb);
    if (!cur || incTs > curTs || (incTs === curTs && !inc._tomb && curIsTomb)) {
      byId[inc.id] = inc;
    }
  });
  return Object.keys(byId).map(function(k) { return byId[k]; });
}

// Same merge, but for a {mobs, mobsTombs} pair (pastureBlocks nests mobs
// inside two named blocks rather than a flat Sheets-row array, so tombs are
// tracked as a separate map instead of inline `_tomb` markers).
function mergeMobsWithTombs(localMobs, localTombs, incomingMobs, incomingTombs) {
  const byId = {};
  function consider(id, ts, isTomb, rec) {
    const cur = byId[id];
    if (!cur || ts > cur.ts || (ts === cur.ts && !isTomb && cur.isTomb)) {
      byId[id] = { ts: ts, isTomb: isTomb, rec: rec };
    }
  }
  (localMobs || []).forEach(function(m) { if (m && m.id != null) consider(m.id, Number(m._ts || 0), false, m); });
  Object.keys(localTombs || {}).forEach(function(id) { consider(id, Number(localTombs[id] || 0), true, null); });
  (incomingMobs || []).forEach(function(m) { if (m && m.id != null) consider(m.id, Number(m._ts || 0), false, m); });
  Object.keys(incomingTombs || {}).forEach(function(id) { consider(id, Number(incomingTombs[id] || 0), true, null); });
  const mobs = [], tombs = {};
  Object.keys(byId).forEach(function(id) {
    const w = byId[id];
    if (w.isTomb) tombs[id] = w.ts; else mobs.push(w.rec);
  });
  return { mobs: mobs, tombs: tombs };
}

// pastureBlocks nests a per-mob list inside two named blocks (winterGraze /
// autumnSaved). Merge each block's mobs per-mob by _ts; everything else in a
// block (area, cover, growth rates, coverUpdates) stays on the pre-existing
// whole-object _ts guard, since those fields are edited by one person at a
// time in practice. This closes the actual reported bug ("Cattle Yards 4
// mob keeps coming back") — the object-level _ts guard never protected the
// mobs list on its own, because every save re-stamps it to "now" regardless
// of what changed, so a stale phone's full push always looked newest.
function writePastureBlocksMerged(ss, sheetName, incoming) {
  const current = readObjectSection(ss, sheetName) || {};
  const curTs = Number(current._ts || 0);
  const incTs = Number((incoming || {})._ts || 0);
  const base = (incTs >= curTs) ? (incoming || {}) : current;
  const out = {};
  Object.keys(base).forEach(function(k) { out[k] = base[k]; });
  ["winterGraze", "autumnSaved"].forEach(function(key) {
    const cb = current[key] || {}, ib = (incoming || {})[key] || {};
    const m = mergeMobsWithTombs(cb.mobs, cb.mobsTombs, ib.mobs, ib.mobsTombs);
    out[key] = Object.assign({}, base[key] || {}, { mobs: m.mobs, mobsTombs: m.tombs });
  });
  out._ts = Math.max(curTs, incTs);
  writeObjectSection(ss, sheetName, out);
}

// Write one section by its client key. Array vs object vs merge-append is
// decided by config so the sectioned path matches the legacy full-push exactly.
function applySection(ss, key, value) {
  const sheetName = SECTION_SHEETS[key];
  if (!sheetName) return;
  if (key === "dailyLogs")     { mergeDailyLogs(ss, value); return; }
  if (key === "checks")        { writeChecksMerged(ss, sheetName, value); return; }
  if (key === "weeklyJobs")    { writeWeeklyMerged(ss, sheetName, value); return; }
  if (key === "pastureBlocks") { writePastureBlocksMerged(ss, sheetName, value); return; }
  if (MERGE_ARRAY_SECTIONS.has(key)) {
    const existing = readArraySection(ss, sheetName);
    writeArraySection(ss, sheetName, mergeArraySectionRows(existing, value));
    return;
  }
  if (SECTION_ARRAY.has(key)) writeArraySection(ss, sheetName, value);
  else                        writeObjectSection(ss, sheetName, value);
}

// ── Delta push — applies only changed fields ──────────────────────────
function handlePushDelta(ss, changes, user) {
  if (!changes || !changes.length) return { skipped: "empty" };
  const userName = user || "Staff";

  let maxTs  = 0;
  const logRows = [];

  changes.forEach(function(change) {
    const ts      = Number(change.ts) || Date.now();
    const section = String(change.section || "");
    const key     = String(change.key     || "");
    const value   = change.value !== undefined ? change.value : null;

    if (ts > maxTs) maxTs = ts;

    applyDelta(ss, section, key, value);

    logRows.push([ts, userName, "pushDelta", section, key]);
  });

  // Batch-write log rows
  appendSyncLogBatch(ss, logRows);

  // Update LastModified to the newest change in this batch
  const lmRow = readObjectSection(ss, SHEETS.lastModified) || {};
  if (maxTs > Number(lmRow.lastModified || 0)) {
    writeObjectSection(ss, SHEETS.lastModified, {
      lastModified: maxTs,
      pushedBy: userName,
      ts: new Date().toISOString(),
    });
  }

  return { ts: maxTs, applied: changes.length };
}

// ── Delta pull — returns change events since a timestamp ──────────────
function handlePullDelta(ss, since) {
  const changes = readSyncLogSince(ss, since);
  return { changes };
}

// ── Init — create all sheet tabs if absent ────────────────────────────
function handleInit(ss) {
  Object.values(SHEETS).forEach(name => ensureSheet(ss, name));
  return { initialized: true };
}

// ══════════════════════════════════════════════════════════════════════
// DELTA APPLY
// ══════════════════════════════════════════════════════════════════════

function applyDelta(ss, section, key, value) {
  const sheetsKey = DELTA_SECTION_MAP[section];
  if (!sheetsKey) {
    logError("applyDelta", new Error("Unknown section: " + section));
    return;
  }
  const sheetName = SHEETS[sheetsKey];

  // ── Single-object sections (barnCalc, checks) ─────────────────────
  if (OBJECT_SECTIONS.has(section)) {
    writeObjectSection(ss, sheetName, value);
    return;
  }

  // ── WeeklyJobs — object with template + ticks, partial update ─────
  if (section === "weeklyjobs") {
    const current = readObjectSection(ss, sheetName) || {};
    if (key === "ticks") {
      current.ticks = value;
    } else {
      if (!current.template) current.template = {};
      current.template[key] = value;
    }
    writeObjectSection(ss, sheetName, current);
    return;
  }

  // ── DailyLogs — prepend-newest, 90-entry cap ──────────────────────
  if (section === "dailylogs") {
    if (value === null) {
      deleteRecord(ss, sheetName, "dateISO", key);
    } else {
      upsertRecord(ss, sheetName, "dateISO", key, value, true /* prepend */);
      trimArraySection(ss, sheetName, 90);
    }
    return;
  }

  // ── All other array sections ──────────────────────────────────────
  if (value === null) {
    deleteRecord(ss, sheetName, "id", key);
  } else {
    upsertRecord(ss, sheetName, "id", key, value, false);
  }
}

// ══════════════════════════════════════════════════════════════════════
// DATA ACCESS LAYER
// ══════════════════════════════════════════════════════════════════════
// All sheets have a header row (row 1) and data rows starting at row 2.
// Column A: JSON string of the record.
// Column B: ISO timestamp of last write.

function ensureSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange("A1:B1").setValues([["json_data", "updated_at"]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Read all rows from an array section — returns array of parsed objects
function readArraySection(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const result = [];
  values.forEach(function(row) {
    if (row[0]) {
      try { result.push(JSON.parse(row[0])); } catch(e) {}
    }
  });
  return result;
}

// Write an array section — clears existing data rows and rewrites
function writeArraySection(ss, sheetName, arr) {
  if (!arr) return;
  const sheet = ensureSheet(ss, sheetName);
  const ts = new Date().toISOString();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  if (arr.length > 0) {
    const rows = arr.map(function(item) { return [JSON.stringify(item), ts]; });
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

// Read a single-object section — returns the parsed object or null
function readObjectSection(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const raw = sheet.getRange(2, 1).getValue();
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

// Write a single-object section — overwrites row 2
function writeObjectSection(ss, sheetName, obj) {
  if (!obj) return;
  const sheet = ensureSheet(ss, sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  sheet.getRange(2, 1).setValue(JSON.stringify(obj));
  sheet.getRange(2, 2).setValue(new Date().toISOString());
}

// Find a row in an array section by field+key — returns { rowIndex, item } or null
// rowIndex is 1-based sheet row number
function findRecord(ss, sheetName, field, key) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const lastRow = sheet.getLastRow();
  const values  = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) continue;
    try {
      const item = JSON.parse(values[i][0]);
      if (item[field] === key) return { rowIndex: i + 2, item };
    } catch(e) {}
  }
  return null;
}

// Upsert a record in an array section
function upsertRecord(ss, sheetName, field, key, value, prepend) {
  const sheet  = ensureSheet(ss, sheetName);
  const ts     = new Date().toISOString();
  const found  = findRecord(ss, sheetName, field, key);
  if (found) {
    // Update in place
    sheet.getRange(found.rowIndex, 1).setValue(JSON.stringify(value));
    sheet.getRange(found.rowIndex, 2).setValue(ts);
  } else if (prepend) {
    // Insert at row 2, shifting existing data down
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1).setValue(JSON.stringify(value));
    sheet.getRange(2, 2).setValue(ts);
  } else {
    // Append after last row
    const nextRow = Math.max(2, sheet.getLastRow() + 1);
    sheet.getRange(nextRow, 1).setValue(JSON.stringify(value));
    sheet.getRange(nextRow, 2).setValue(ts);
  }
}

// Delete a record from an array section
function deleteRecord(ss, sheetName, field, key) {
  const found = findRecord(ss, sheetName, field, key);
  if (found) {
    ss.getSheetByName(sheetName).deleteRow(found.rowIndex);
  }
}

// Trim an array section to a maximum number of data rows (keeps newest = rows near top)
function trimArraySection(ss, sheetName, maxRows) {
  const sheet   = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const dataRows = lastRow - 1;
  if (dataRows > maxRows) {
    sheet.deleteRows(maxRows + 2, dataRows - maxRows);
  }
}

// DailyLogs: merge incoming entries — only add dates not already present
function mergeDailyLogs(ss, logs) {
  if (!logs || !logs.length) return;
  const sheet = ensureSheet(ss, SHEETS.dailyLogs);
  const existing = readArraySection(ss, SHEETS.dailyLogs);
  const existingDates = new Set(existing.map(function(e) { return e.dateISO; }));
  const ts = new Date().toISOString();
  const newRows = logs
    .filter(function(log) { return log.dateISO && !existingDates.has(log.dateISO); })
    .map(function(log) { return [JSON.stringify(log), ts]; });
  if (newRows.length > 0) {
    const nextRow = Math.max(2, sheet.getLastRow() + 1);
    sheet.getRange(nextRow, 1, newRows.length, 2).setValues(newRows);
  }
}

// ══════════════════════════════════════════════════════════════════════
// TICK-MAP MERGE (v13.1) — mirrors the client mergeTickMaps
// checks (Feed Out + Breaks + barn) and weeklyJobs.ticks are shared per-day
// maps several phones tick at once. Merge per key by ms timestamp (newest action
// wins) so a whole-object push never wipes another phone's ticks. A key present
// in `times` but absent from `vals` is a tombstone (delete) so un-ticks carry.
// ══════════════════════════════════════════════════════════════════════

function mergeTickMapsGS(valsA, timesA, valsB, timesB) {
  valsA = valsA || {}; timesA = timesA || {}; valsB = valsB || {}; timesB = timesB || {};
  var vals = {}, times = {}, seen = {};
  var all = [].concat(Object.keys(timesA), Object.keys(timesB), Object.keys(valsA), Object.keys(valsB));
  all.forEach(function (k) {
    if (seen[k]) return; seen[k] = 1;
    var ta = Number(timesA[k] || 0), tb = Number(timesB[k] || 0);
    var useB = tb > ta ? true : ta > tb ? false : (k in valsB && !(k in valsA));
    var sv = useB ? valsB : valsA, st = useB ? timesB : timesA;
    var t = Number(st[k] || 0); if (t) times[k] = t;
    if (k in sv) vals[k] = sv[k];   // else tombstone — key stays deleted
  });
  return { vals: vals, times: times };
}

// Merge incoming checks {checks,date,times} into the stored copy, per day.
function writeChecksMerged(ss, sheetName, incoming) {
  if (!incoming) return;
  var cur = readObjectSection(ss, sheetName) || {};
  var inDate = incoming.date || "", curDate = cur.date || "";
  if (inDate && curDate && inDate !== curDate) {
    // Different day — newer date wins wholesale (the daily tick reset).
    if (inDate > curDate) writeObjectSection(ss, sheetName, { checks: incoming.checks || {}, date: inDate, times: incoming.times || {} });
    return;
  }
  var m = mergeTickMapsGS(cur.checks || {}, cur.times || {}, incoming.checks || {}, incoming.times || {});
  writeObjectSection(ss, sheetName, { checks: m.vals, date: inDate || curDate, times: m.times });
}

// Merge incoming weeklyJobs — ticks merge per key (tombstones). The STRUCTURE
// (template + weekStart) is guarded by structTs: the newest week-roll or template
// edit wins, and a stale phone re-pushing an old board (older or absent structTs)
// can no longer revert a freshly-rolled week. All-legacy clients (no structTs on
// either side) fall back to last-writer-wins, so there is no regression.
function writeWeeklyMerged(ss, sheetName, incoming) {
  if (!incoming) return;
  var cur = readObjectSection(ss, sheetName) || {};
  var m = mergeTickMapsGS(cur.ticks || {}, cur.tickTimes || {}, incoming.ticks || {}, incoming.tickTimes || {});
  var curStruct = Number(cur.structTs) || 0;
  var inStruct  = Number(incoming.structTs) || 0;
  var out = {};
  // Start from the stored copy — this preserves cur.template/weekStart/structTs.
  Object.keys(cur).forEach(function (k) { out[k] = cur[k]; });
  if (inStruct >= curStruct) {
    // Incoming structure is newest — adopt all its fields (template, weekStart, etc.).
    Object.keys(incoming).forEach(function (k) { out[k] = incoming[k]; });
    out.structTs = inStruct;
  }
  // else: stored structure is newer — keep it; incoming template/weekStart are ignored.
  out.ticks = m.vals; out.tickTimes = m.times;   // ticks always merge, either way
  writeObjectSection(ss, sheetName, out);
}

// ══════════════════════════════════════════════════════════════════════
// SYNC LOG
// SyncLog columns: A=ts(ms), B=user, C=action, D=section, E=key
// ══════════════════════════════════════════════════════════════════════

function appendSyncLog(ss, entry) {
  appendSyncLogBatch(ss, [[entry.ts, entry.user, entry.action, entry.section, entry.key]]);
}

function appendSyncLogBatch(ss, rows) {
  if (!rows.length) return;
  const sheet = ensureSheet(ss, SHEETS.syncLog);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, 5).setValues(rows);
  pruneSyncLog(ss, sheet);
}

// Remove entries older than SYNCLOG_RETAIN_DAYS
function pruneSyncLog(ss, sheet) {
  const cutoff  = Date.now() - (SYNCLOG_RETAIN_DAYS * 24 * 60 * 60 * 1000);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const tsValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  // Find contiguous old rows from the top of data (oldest entries)
  let deleteCount = 0;
  for (let i = 0; i < tsValues.length; i++) {
    const ts = Number(tsValues[i][0]) || 0;
    if (ts < cutoff && ts > 0) deleteCount++; else break;
  }
  if (deleteCount > 0) sheet.deleteRows(2, deleteCount);
}

// Read delta change events from SyncLog since a given timestamp
function readSyncLogSince(ss, since) {
  const sheet = ss.getSheetByName(SHEETS.syncLog);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const sinceTs  = Number(since) || 0;
  const lastRow  = sheet.getLastRow();
  const values   = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const changes  = [];

  values.forEach(function(row) {
    const ts     = Number(row[0]) || 0;
    const action = String(row[2] || "");
    const section = String(row[3] || "");
    const key    = String(row[4] || "");

    if (ts <= sinceTs) return;
    if (action !== "pushDelta") return; // only return delta events
    if (!section || !key) return;

    // Re-read the current value from the data sheet (source of truth)
    // so pullDelta always returns the authoritative state, not a stale snapshot
    const sheetsKey = DELTA_SECTION_MAP[section];
    if (!sheetsKey) return;
    const sheetName = SHEETS[sheetsKey];
    let value = null;

    if (OBJECT_SECTIONS.has(section)) {
      value = readObjectSection(ss, sheetName);
    } else if (section === "weeklyjobs") {
      const wj = readObjectSection(ss, sheetName) || {};
      value = (key === "ticks") ? (wj.ticks || {}) : ((wj.template || {})[key] || []);
    } else {
      const idField = (section === "dailylogs") ? "dateISO" : "id";
      const found = findRecord(ss, sheetName, idField, key);
      value = found ? found.item : null;
    }

    changes.push({ ts, user: String(row[1] || ""), section, key, value });
  });

  return changes;
}

// ══════════════════════════════════════════════════════════════════════
// ERROR LOG
// ErrorLog columns: A=timestamp(ISO), B=function, C=error message
// ══════════════════════════════════════════════════════════════════════

function logError(context, err) {
  try {
    const sheet = ensureSheet(SpreadsheetApp.getActiveSpreadsheet(), SHEETS.errorLog);
    sheet.appendRow([new Date().toISOString(), context, err.message || String(err)]);
    // Keep last 500 error rows
    const lastRow = sheet.getLastRow();
    if (lastRow > 501) sheet.deleteRows(2, lastRow - 501);
  } catch(e) {
    // If error logging itself fails, swallow silently
  }
}

// ══════════════════════════════════════════════════════════════════════
// RESPONSE HELPER
// ══════════════════════════════════════════════════════════════════════

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════════════
// CLAUDE PROXY
// ══════════════════════════════════════════════════════════════════════

// Run this once from the editor after deploying to create missing sheet tabs
function runInit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  handleInit(ss);
  Logger.log("Init complete");
}

function claudeProxy(data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) return { error: "ANTHROPIC_API_KEY not set in Script Properties" };

  const payload = {
    model:      data.model      || "claude-haiku-4-5-20251001",
    max_tokens: data.max_tokens || 1024,
    messages:   data.messages,
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method:      "post",
    contentType: "application/json",
    headers: {
      "x-api-key":         apiKey,
      "anthropic-version": "2023-06-01",
    },
    payload: JSON.stringify(payload),
  });

  const result = JSON.parse(response.getContentText());
  return { content: result.content[0].text };
}
