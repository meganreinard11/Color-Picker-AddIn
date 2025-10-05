/* eslint-disable no-console */
"use strict";

/**
 * Env-driven error handling for Excel add-ins.
 * Decision source: _Settings!B4 (TRUE => dev/advanced; FALSE => production).
 *
 * Exports:
 *   - getEnvFlag()
 *   - handle(error, context?)
 *   - tryWrap(action, workFn(env), context?)
 */

var SETTINGS_SHEET = "_Settings";
var SETTINGS_FLAG_RANGE = "B4";
var DIAG_SHEET = "_Diagnostics";
var DIAG_TABLE_NAME = "ErrorLog";
var DIAG_HEADERS = ["Timestamp","Action","Message","Code","Location","Statement","Stack"];

/**
 * Main handler. Chooses dev/prod behavior based on _Settings!B4.
 * context: { action?, data?, userMessage?, logToSheet?, forceLogToSheet? }
 * NOTE: This function calls Excel.run internally to read the env (and to log diagnostics),
 * so it's best to call it OUTSIDE any active Excel.run block if possible.
 */
async function handle(err, context) {
  context = context || {};
  var offErr = toOfficeError(err);
  var safeMessage = typeof context.userMessage === "string" ? context.userMessage : "Something went wrong. Please try again.";
  console.error("[" + (context.action || "Operation") + "] " + (offErr.message || "Error"), offErr);
  var env = await getEnvFlag();
  if (env && env.debug) {
    logVerbose(offErr, context); // Verbose console diagnostics.
    var doSheetLog = (typeof context.logToSheet === "boolean") ? context.logToSheet : true; // In dev, default to also logging in a hidden sheet.
    if (doSheetLog) {
      try { await appendDiagnostics(offErr, context); }
      catch (e) { console.warn("[diagnostics] Unable to write to diagnostics sheet.", e); }
    }
    try { alert(buildDevAlert(offErr, context)); } catch (_) {} // Quick visible alert (details live in console/diagnostics).
  } else {
    try { alert(safeMessage); } catch (_) {} // Production: show friendly message only.
    if (context.forceLogToSheet) { try { await appendDiagnostics(offErr, context); } catch (_) {} }
  }
  return false;
}

/** Reads _Settings!B4 and returns { debug: boolean }. Missing sheet/cell => { debug:false }. */
async function getEnvFlag() {
  try {
    return await Excel.run(async function (ctx) {
      try {
        var ws = ctx.workbook.worksheets.getItem(SETTINGS_SHEET);
        var rng = ws.getRange(SETTINGS_FLAG_RANGE);
        rng.load("values");
        await ctx.sync();
        var raw = rng.values && rng.values[0] ? rng.values[0][0] : undefined;
        var debug = coerceBoolean(raw);
        return { debug: debug };
      } catch (err) {
        if (isItemNotFound(err)) {
          console.warn('[env] "_Settings" not found; defaulting to production.');
          return { debug: false };
        }
        console.warn("[env] Could not read _Settings!B4; defaulting to production.", err);
        return { debug: false };
      }
    });
  } catch (outer) {
    console.warn("[env] Excel.run failed while determining environment; defaulting to production.", outer);
    return { debug: false };
  }
}

/**
 * Convenience wrapper: runs an async block and applies the handler on failure.
 * Still reads env up-front so your workFn can branch behavior if needed.
 */
async function tryWrap(action, workFn, context) {
  var env = await getEnvFlag();
  try {
    return await workFn(env); // you can still use env here if helpful
  } catch (err) {
    await handle(err, merge({ action: action }, context || {}));  // fixed: call the correct handler
    return undefined;
  }
}

/* ------------------------ internals ------------------------ */
function coerceBoolean(v) {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (typeof v === "string") {
    var s = v.trim().toLowerCase();
    if (["true", "yes", "y", "1"].indexOf(s) >= 0) return true;
    if (["false", "no", "n", "0", ""].indexOf(s) >= 0) return false;
  }
  return false;
}

function isItemNotFound(e) {
  try {
    var code = e && (e.code || (e.error && e.error.code));
    var apiCode = Excel && Excel.ErrorCodes ? Excel.ErrorCodes.itemNotFound : undefined;
    return code === apiCode || code === "ItemNotFound";
  } catch (_) {
    return false;
  }
}

function toOfficeError(err) {
  var anyErr = err || {};
  return {
    name: anyErr.name || "Error",
    message: anyErr.message || String(anyErr || "Unknown error"),
    stack: anyErr.stack,
    code: anyErr.code,
    debugInfo: anyErr.debugInfo
  };
}

function logVerbose(e, context) {
  var action = context.action || "Operation";
  console.groupCollapsed("%c[DEV] " + action + " failed", "font-weight:bold;color:#c00");
  console.log("Message:", e.message);
  if (e.code) console.log("Code:", e.code);
  if (e.debugInfo && e.debugInfo.errorLocation) console.log("Location:", e.debugInfo.errorLocation);
  if (e.debugInfo && e.debugInfo.statement) console.log("Statement:", e.debugInfo.statement);
  if (e.debugInfo && e.debugInfo.surroundingStatements && e.debugInfo.surroundingStatements.length) console.log("Surrounding:", e.debugInfo.surroundingStatements.join("\n"));
  if (e.debugInfo && e.debugInfo.traceMessages && e.debugInfo.traceMessages.length) console.log("Trace:", e.debugInfo.traceMessages);
  if (typeof context.data !== "undefined") console.log("Context data:", context.data);
  if (e.stack) console.log("Stack:", e.stack);
  console.groupEnd();
}

function buildDevAlert(e, context) {
  var parts = [];
  parts.push("[DEV] " + (context.action || "Operation") + " failed");
  if (e.message) parts.push("• " + e.message);
  if (e.code) parts.push("• Code: " + e.code);
  if (e.debugInfo && e.debugInfo.errorLocation) parts.push("• Where: " + e.debugInfo.errorLocation);
  return parts.join("\n");
}

/** Append a row to a hidden _Diagnostics sheet with an ErrorLog table. */
async function appendDiagnostics(e, context) {
  await Excel.run(async function (excelCtx) {
    var wb = excelCtx.workbook;
    // Ensure sheet exists.
    var ws = wb.worksheets.getItemOrNullObject(DIAG_SHEET);
    ws.load("name,isNullObject,visibility");
    await excelCtx.sync();
    if (ws.isNullObject) {
      ws = wb.worksheets.add(DIAG_SHEET);
      // Create headers
      var headerRange = ws.getRange("A1:" + colLetter(DIAG_HEADERS.length) + "1");
      headerRange.values = [DIAG_HEADERS.slice(0)];
      headerRange.format.font.bold = true;
      // Add as a table
      var t = wb.tables.add(ws.name + "!A1:" + colLetter(DIAG_HEADERS.length) + "1", true);
      t.name = DIAG_TABLE_NAME;
      // Hide the sheet (VeryHidden if supported).
      try {
        if (Excel && Excel.SheetVisibility && Excel.SheetVisibility.veryHidden) {
          ws.visibility = Excel.SheetVisibility.veryHidden;
        } else {
          ws.visibility = "VeryHidden";
        }
      } catch (_) {
        try { ws.visibility = "VeryHidden"; } catch (_) {}
      }
      await excelCtx.sync();
    }
    // Ensure table exists (user might have deleted it).
    var table = wb.tables.getItemOrNullObject(DIAG_TABLE_NAME);
    table.load("name,isNullObject,range");
    await excelCtx.sync();
    if (table.isNullObject) {
      table = wb.tables.add(ws.name + "!A1:" + colLetter(DIAG_HEADERS.length) + "1", true);
      table.name = DIAG_TABLE_NAME;
      await excelCtx.sync();
    }
    // Build row payload
    var row = [
      new Date().toISOString(),
      context.action || "",
      firstLine(e.message || ""),
      e.code || "",
      e.debugInfo && e.debugInfo.errorLocation ? e.debugInfo.errorLocation : "",
      e.debugInfo && e.debugInfo.statement ? e.debugInfo.statement : "",
      truncate(e.stack || "", 4000)
    ];
    table.rows.add(null, [row]);
    await excelCtx.sync();
  });
}

/* ------------------------ tiny helpers ------------------------ */
function colLetter(n) {
  var s = "";
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function firstLine(s) {
  var i = s.indexOf("\n");
  return i >= 0 ? s.slice(0, i) : s;
}

function truncate(s, max) { return s.length > max ? s.slice(0, max - 1) + "…" : s; }

function merge(a, b) {
  var out = {};
  Object.keys(a || {}).forEach(function (k) { out[k] = a[k]; });
  Object.keys(b || {}).forEach(function (k) { out[k] = b[k]; });
  return out;
}

/* ------------------------ exports ------------------------ */

// ESM / CommonJS / UMD-lite export
var ErrorHandler = { getEnvFlag: getEnvFlag, handle: handle, tryWrap: tryWrap };
// Back-compat alias so callers using handleError(...) keep working.
ErrorHandler.handleError = handle;

if (typeof module !== "undefined" && module.exports) {
  module.exports = ErrorHandler;            // CommonJS
} else if (typeof window !== "undefined") {
  window.ErrorHandler = ErrorHandler;       // Browser global (script tag)
} else if (typeof self !== "undefined") {
  self.ErrorHandler = ErrorHandler;         // WebWorker/globalThis
}
