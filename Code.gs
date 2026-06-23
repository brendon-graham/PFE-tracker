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
  barnSchedule: "BarnSchedule",
  // Object sections — single JSON row
  barnCalc:    "BarnCalc",
  checks:      "Checks",
  weeklyJobs:  "WeeklyJobs",
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
      case "push":      return respond({ ok: true, ...handlePush(ss, body.data, body.user) });
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
      barnSchedule: readArraySection(ss, SHEETS.barnSchedule),
      barnCalc:    readObjectSection(ss, SHEETS.barnCalc),
      checks:      readObjectSection(ss, SHEETS.checks),
      weeklyJobs:  readObjectSection(ss, SHEETS.weeklyJobs),
      lastModified: lmRow ? (lmRow.lastModified || 0) : 0,
      syncTime: new Date().toISOString(),
    }
  };
}

// ── Full push — writes complete state (fallback for uninstrumented changes) ─
function handlePush(ss, data, user) {
  if (!data) throw new Error("Push received empty data payload");
  const userName = user || "Staff";

  // Stale-push guard — reject if incoming lastModified ≤ server's current
  const lmRow      = readObjectSection(ss, SHEETS.lastModified);
  const currentTs  = lmRow ? Number(lmRow.lastModified || 0) : 0;
  const incomingTs = Number(data.lastModified || 0);
  if (incomingTs > 0 && currentTs > 0 && incomingTs <= currentTs) {
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
  if (data.barnSchedule) writeArraySection(ss, SHEETS.barnSchedule, data.barnSchedule);
  if (data.dailyLogs)   mergeDailyLogs(ss, data.dailyLogs);

  const lastModified = data.lastModified || Date.now();
  writeObjectSection(ss, SHEETS.lastModified, { lastModified, pushedBy: userName, ts });
  appendSyncLog(ss, { ts: lastModified, user: userName, action: "push", section: "full", key: "" });

  return { pushed: ts };
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
