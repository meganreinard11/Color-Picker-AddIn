// utils.js — global helpers for the Excel color picker add-in
// Exposes window.Utils (no module system required).

(function () {
  'use strict';

  const SETTINGS_SHEET = '_Settings';
  const VALUE_COLUMN = 'B';
  const ENVIRONMENT_START = 4;
  
  async function debugMode() {
    return Excel.run(async (context) => {
      const sheet = await ensureStoreSheet(context);
      const rng = sheet.getRange(VALUE_COLUMN + ENVIRONMENT_START);
      rng.load("values");
      await context.sync();
      const val = rng.values?.[0]?.[0];
    }
  }

  function showError(e) {
    console.error(e);
    const msg = (e && e.message) ? e.message : String(e);
    showResult(`⚠️ ${msg}`);
    alert("Excel API error: " + e));
  }

  async function ensureStoreSheet(context) {
    let sheet = context.workbook.worksheets.getItemOrNullObject(SETTINGS_SHEET);
    sheet.load("name,isNullObject,visibility");
    await context.sync();
    if (sheet.isNullObject) {
      sheet = context.workbook.worksheets.add(STORE_SHEET);
      sheet.visibility = "Hidden";
      const r = sheet.getRange(`A1:A${RECENT_LIMIT}`);
      r.numberFormat = "@";
    }
    return sheet;
  }

  window.Utils = {
    SETTINGS_SHEET, VALUE_COLUMN, currentEnvironment
  };
})();
