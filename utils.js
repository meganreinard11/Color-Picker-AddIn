// utils.js — global helpers for the Excel color picker add-in
// Exposes window.Utils (no module system required).

(function () {
  'use strict';

  const SETTINGS_SHEET = '_Settings';

  window.Utils = {
    ensureStoreSheet, readRecentColors, writeRecentColors, pushRecentColor,
    getSelectedRange, readFillColor, readFontColor, applyColorToSelection,
    DEFAULT_STORE_SHEET
  };
})();
