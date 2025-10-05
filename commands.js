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

  /* ================= Auto-open/close task pane based on selection ========================
     Configure either a named range (recommended) or a fallback sheet+address.
     Create a Named Range:  OpenColorPickerCell  -> points to the hot cell (e.g., Sheet1!$B$2)
  ========================================================================================= */
  const TRIGGER = {
    name: "PrimaryColor",                // workbook-scoped name (preferred)
    fallback: { sheet: "Overview", address: "B19" } // used only if the name doesn't exist
  };

  let _paneOpenedByTrigger = false;
  let _autoOpenDebounceAt = 0;

  function upperNoDollarQuotes(s) {
    return (s || "").toString().toUpperCase().replace(/'/g, "").replace(/\$+/g, "");
  }

  function stripSheetPrefix(addr) {
    // Converts "Sheet!A1" or "'Sheet Name'!$A$1" -> "A1"
    const m = /^([^!]+)!(.+)$/.exec(addr);
    return m ? m[2] : addr;
  }

  async function resolveTriggerAddress(ctx) {
    // Try named item first (workbook-scoped)
    if (TRIGGER.name) {
      const ni = ctx.workbook.names.getItemOrNullObject(TRIGGER.name);
      ni.load(["name", "referenceAddress", "isNullObject"]);
      await ctx.sync();
      if (!ni.isNullObject && ni.referenceAddress) {
        return ni.referenceAddress; // e.g., "Sheet1!$B$2"
      }
    }
    // Fallback to explicit sheet+address
    if (TRIGGER.fallback && TRIGGER.fallback.sheet && TRIGGER.fallback.address) {
      return TRIGGER.fallback.sheet + "!" + TRIGGER.fallback.address;
    }
    return null;
  }

  async function selectionIntersectsTarget(event) {
    // Uses event.address and event.worksheetId to test intersection with the target cell.
    return Excel.run(async (ctx) => {
      const targetAddrFull = await resolveTriggerAddress(ctx);
      if (!targetAddrFull) return false;

      const eventAddrFull = event.address; // already includes sheet
      if (!eventAddrFull) return false;

      // If sheets differ, there's no intersection.
      const eventSheetPart = upperNoDollarQuotes(eventAddrFull.split("!")[0]);
      const targetSheetPart = upperNoDollarQuotes(targetAddrFull.split("!")[0]);
      if (eventSheetPart !== targetSheetPart) return false;

      const ws = ctx.workbook.worksheets.getItemOrNullObject(event.worksheetId || eventSheetPart);
      ws.load(["name", "id", "isNullObject"]);
      await ctx.sync();
      if (ws.isNullObject) return false;

      const selNoSheet = stripSheetPrefix(eventAddrFull);
      const tgtNoSheet = stripSheetPrefix(targetAddrFull);

      const selection = ws.getRange(selNoSheet);
      const target = ws.getRange(tgtNoSheet);
      const inter = selection.getIntersectionOrNullObject(target);
      inter.load("address");
      await ctx.sync();
      return !inter.isNullObject;
    }).catch((_e) => false);
  }

  async function handleSelectionChanged(event) {
    try {
      const now = Date.now();
      const intersects = await selectionIntersectsTarget(event);

      if (intersects) {
        // Debounce opens to avoid repeated UI churn while user nudges selection
        if (now - _autoOpenDebounceAt >= 400) {
          await Office.addin.showAsTaskpane();
          _autoOpenDebounceAt = now;
        }
        _paneOpenedByTrigger = true;
      } else if (_paneOpenedByTrigger) {
        // We previously opened it due to the trigger; close now that we've exited
        try { await Office.addin.hide(); } catch (e) { /* ignore */ }
        _paneOpenedByTrigger = false;
      }
    } catch (err) {
      safeHandle(err, { action: "autoOpenClose.onSelectionChanged", userMessage: "Couldn't auto-open/close the task pane." });
    }
  }

  // Register the best-available selection event (WorksheetCollection > fallback per-sheet)
  async function registerSelectionListener() {
    try {
      await Excel.run(async (ctx) => {
        const coll = ctx.workbook.worksheets;
        // Prefer the collection-level event (ExcelApi 1.9). It fires for any sheet.
        if (coll.onSelectionChanged && typeof coll.onSelectionChanged.add === "function") {
          coll.onSelectionChanged.add(handleSelectionChanged);
          await ctx.sync();
          return;
        }

        // Fallback: attach to each worksheet individually.
        coll.load("items/name");
        await ctx.sync();
        for (const ws of coll.items) {
          ws.onSelectionChanged.add(handleSelectionChanged);
        }
        await ctx.sync();
      });
    } catch (err) {
      safeHandle(err, { action: "registerSelectionListener", userMessage: "Couldn't register selection-change listener." });
    }
  }

  // LaunchEvent handler so the shared runtime initializes at workbook open.
  async function onDocumentOpened(event) {
    try {
      // Nothing critical here; Office.onReady will wire the listeners.
      // If you'd like, we could check the current selection and open immediately.
    } catch (err) {
      safeHandle(err, { action: "onDocumentOpened", userMessage: "Startup failed." });
    } finally {
      try { event.completed(); } catch (_) {}
    }
  }

  // Optional: legacy buttons still work in shared runtime; keep these if needed elsewhere.
  async function showTaskpane() {
    try { await Office.addin.showAsTaskpane(); }
    catch (err) { return safeHandle(err, { action: "commands.showTaskpane", userMessage: "Couldn't show the task pane." }); }
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

  // Wire everything once Excel is ready
  Office.onReady(() => {
    try {
      registerSelectionListener();
      Office.actions.associate("onDocumentOpened", onDocumentOpened);
      Office.actions.associate("showTaskpane", showTaskpane);
      Office.actions.associate("quickFillYellow", quickFillYellow);
    } catch (err) {
      safeHandle(err, { action: "commands.associate", userMessage: "Couldn't initialize the add-in." });
    }
  });
})();
