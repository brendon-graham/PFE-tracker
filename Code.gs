// ═══════════════════════════════════════════════════════════════════
// PFE FARM TRACKER — Google Apps Script Backend v4.0
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
  if (data.checks)      writeObjectSection(ss, SHEETS.checks,     data.checks);
  if (data.scenarios)   writeArraySection(ss, SHEETS.scenarios,   data.scenarios);
  if (data.weeklyJobs)  writeObjectSection(ss, SHEETS.weeklyJobs, data.weeklyJobs);
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

// Write one section by its client key. Array vs object vs merge-append is
// decided by config so the sectioned path matches the legacy full-push exactly.
function applySection(ss, key, value) {
  const sheetName = SECTION_SHEETS[key];
  if (!sheetName) return;
  if (key === "dailyLogs")     { mergeDailyLogs(ss, value); return; }
  if (key === "pastureBlocks") {
    // Stale-guard on _ts so an out-of-date phone can't roll it back.
    const curTs = Number((readObjectSection(ss, sheetName) || {})._ts || 0);
    const inTs  = Number((value || {})._ts || 0);
    if (!(inTs > 0 && curTs > 0 && inTs <= curTs)) writeObjectSection(ss, sheetName, value);
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
