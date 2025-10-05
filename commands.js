(function () {
  /** Safely delegate to our error handler (if present) */
  async function safeHandle(err, context) {
    try {
      const EH =
        (typeof window !== "undefined" && window.ErrorHandler)
          ? window.ErrorHandler
          : (typeof ErrorHandler !== "undefined" ? ErrorHandler : null);
      if (EH && typeof EH.handle === "function") {
        return EH.handle(err, context || {});
      }
    } catch (_) {}
    console.error("[commands] error:", err);
    try { alert((context && context.userMessage) || "Something went wrong."); } catch (_) {}
    return false;
  }

  /* ================= Auto-open task pane when a specific cell is selected ================
     Configure either a named range (recommended) or a fallback sheet+address.
     Create a Named Range in the workbook called "OpenColorPickerCell" that refers to the target cell.
     Or change TRIGGER.fallback below.
  ========================================================================================= */
  const TRIGGER = {
    name: "OpenColorPickerCell",                       // named range to watch (preferred)
    fallback: { sheet: "Sheet1", address: "B2" }       // used only if the name doesn't exist
  };

  let _autoOpenDebounceAt = 0;

  function addressesMatch(selectionAddr, targetAddr) {
    if (!selectionAddr || !targetAddr) return false;
    const norm = s => s.toUpperCase().replace(/\$+/g, "");
    const s = norm(selectionAddr);
    const t = norm(targetAddr);
    // Normalize single-cell forms: "Sheet!B2" vs "Sheet!B2:B2"
    const tSingle = t.includes(":") ? t : (t + ":" + t);
    const sSingle = s.includes(":") ? s : (s + ":" + s);
    return s === t || s === tSingle || sSingle === t || sSingle === tSingle;
  }

  async function resolveTriggerAddress(ctx) {
    // Try named item first
    if (TRIGGER.name) {
      const ni = ctx.workbook.names.getItemOrNullObject(TRIGGER.name);
      ni.load(["name", "referenceAddress", "isNullObject"]);
      await ctx.sync();
      if (!ni.isNullObject && ni.referenceAddress) {
        return ni.referenceAddress;
      }
    }
    // Fallback to explicit sheet+address
    if (TRIGGER.fallback && TRIGGER.fallback.sheet && TRIGGER.fallback.address) {
      return TRIGGER.fallback.sheet + "!" + TRIGGER.fallback.address;
    }
    return null;
  }

  async function _handleSelectionChanged(args) {
    try {
      const now = Date.now();
      if (now - _autoOpenDebounceAt < 750) return; // avoid repeated opens while user is nudging selection
      await Excel.run(async (ctx) => {
        // Selected range address (fully-qualified with sheet)
        const sel = ctx.workbook.getSelectedRange();
        sel.load("address");
        await ctx.sync();
        const selectedAddress = sel.address; // e.g., 'Sheet1!B2' or 'Sheet1!B2:B2'

        // Target to compare
        const targetAddress = await resolveTriggerAddress(ctx);
        if (!targetAddress) return;

        if (addressesMatch(selectedAddress, targetAddress)) {
          await Office.addin.showAsTaskpane();
          _autoOpenDebounceAt = now;
        }
      });
    } catch (err) {
      safeHandle(err, { action: "autoOpen.onSelectionChanged", userMessage: "Couldn't handle selection-change auto-open." });
    }
  }

  let _selectionHandlers = []; // keep EventHandlerResult objects per-worksheet so we could remove later if needed

  async function registerAutoOpenOnSelection() {
    try {
      await Excel.run(async (ctx) => {
        const sheets = ctx.workbook.worksheets;
        sheets.load("items/name");
        await ctx.sync();

        // Register on every current sheet (simple & reliable). We keep results in _selectionHandlers.
        for (const ws of sheets.items) {
          const result = ws.onSelectionChanged.add(_handleSelectionChanged);
          _selectionHandlers.push(result);
        }
        await ctx.sync();
      });
    } catch (err) {
      safeHandle(err, { action: "autoOpen.register", userMessage: "Couldn't register selection-change listener." });
    }
  }
  /* ====================================================================================== */

  async function showTaskpane() {
    try {
      await Office.addin.showAsTaskpane();
      return true;
    } catch (err) {
      return safeHandle(err, { action: "commands.showTaskpane", userMessage: "Couldn't show the task pane." });
    }
  }

  async function quickFillYellow() {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = "#FFF7AB";
        await context.sync();
      });
      return true;
    } catch (err) {
      return safeHandle(err, { action: "commands.quickFillYellow", userMessage: "Couldn't apply the yellow fill." });
    }
  }

  // Associate functions and register selection listener (shared runtime page)
  Office.onReady(() => {
    try {
      Office.actions.associate("showTaskpane", showTaskpane);
      Office.actions.associate("quickFillYellow", quickFillYellow);
      registerAutoOpenOnSelection();
    } catch (err) {
      safeHandle(err, { action: "commands.associate", userMessage: "Couldn't wire ribbon actions." });
    }
  });
})();
