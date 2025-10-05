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
  let _lastOpenAt = 0;
  const OPEN_DEBOUNCE_MS = 350;

  function upperNoDollarQuotes(s) {
    return (s || "").toString().toUpperCase().replace(/'/g, "").replace(/\$+/g, "");
  }

  function stripSheetPrefix(addr) {
    // Converts "Sheet!A1" or "'Sheet Name'!$A$1" -> "A1"
    const m = /^([^!]+)!(.+)$/.exec(addr);
    return m ? m[2] : addr;
  }

  async function resolveTriggerAddress(ctx) {
    if (TRIGGER.name) {
      const ni = ctx.workbook.names.getItemOrNullObject(TRIGGER.name);
      ni.load(["name", "referenceAddress", "isNullObject"]);
      await ctx.sync();
      if (!ni.isNullObject && ni.referenceAddress) {
        return ni.referenceAddress; // e.g., "Sheet1!$B$2"
      }
    }
    if (TRIGGER.fallback && TRIGGER.fallback.sheet && TRIGGER.fallback.address) {
      return TRIGGER.fallback.sheet + "!" + TRIGGER.fallback.address;
    }
    return null;
  }

  async function selectionIntersectsTriggerUsingContext(ctx, addressFull) {
    const targetAddrFull = await resolveTriggerAddress(ctx);
    if (!targetAddrFull || !addressFull) return false;

    // If sheets differ, there's no intersection.
    const eventSheetPart = upperNoDollarQuotes(addressFull.split("!")[0]);
    const targetSheetPart = upperNoDollarQuotes(targetAddrFull.split("!")[0]);
    if (eventSheetPart !== targetSheetPart) return false;

    const ws = ctx.workbook.worksheets.getItemOrNullObject(eventSheetPart);
    ws.load(["name", "isNullObject"]);
    await ctx.sync();
    if (ws.isNullObject) return false;

    const selNoSheet = stripSheetPrefix(addressFull);
    const tgtNoSheet = stripSheetPrefix(targetAddrFull);

    const selection = ws.getRange(selNoSheet);
    const target = ws.getRange(tgtNoSheet);
    const inter = selection.getIntersectionOrNullObject(target);
    inter.load("address");
    await ctx.sync();
    return !inter.isNullObject;
  }

  async function getCurrentSelectionAddressFull(ctx) {
    const sel = ctx.workbook.getSelectedRange();
    sel.load(["address", "worksheet/name"]);
    await ctx.sync();
    return sel.worksheet.name + "!" + stripSheetPrefix(sel.address);
  }

  async function evaluateAndTogglePane(addressFull) {
    try {
      await Excel.run(async (ctx) => {
        const intersects = await selectionIntersectsTriggerUsingContext(ctx, addressFull || await getCurrentSelectionAddressFull(ctx));
        const now = Date.now();
        if (intersects) {
          if (now - _lastOpenAt >= OPEN_DEBOUNCE_MS) {
            await Office.addin.showAsTaskpane();
            _lastOpenAt = now;
          }
          _paneOpenedByTrigger = true;
        } else if (_paneOpenedByTrigger) {
          try { await Office.addin.hide(); } catch (e) {}
          _paneOpenedByTrigger = false;
        }
      });
    } catch (err) {
      safeHandle(err, { action: "evaluateAndTogglePane", userMessage: "Couldn't auto-open/close the task pane." });
    }
  }

  // ====== Event handlers (compatible with Excel for the web) ======
  async function onSelectionChanged(event) {
    // event.address should be like "Sheet1!B2" or "Sheet1!B2:C5"
    const address = event && event.address ? event.address : null;
    await evaluateAndTogglePane(address);
  }

  async function onSingleClicked(event) {
    // Backup for some web cases: click fires even if selection event lags
    const address = event && event.address ? event.address : null;
    await evaluateAndTogglePane(address);
  }

  async function onWorksheetActivated(event) {
    // Re-check when user switches sheets
    await evaluateAndTogglePane(null);
  }

  // Register per-worksheet listeners and keep them up-to-date
  const _subscribed = new Map(); // worksheetId -> { selToken, clickToken, actvToken }

  async function registerListenersForWorksheet(ws) {
    const id = ws.id;
    if (_subscribed.has(id)) return;

    const tokens = {};
    tokens.selToken = await ws.onSelectionChanged.add(onSelectionChanged);
    // onSingleClicked exists on many hosts including web; ignore errors if not present
    if (ws.onSingleClicked && typeof ws.onSingleClicked.add === "function") {
      try { tokens.clickToken = await ws.onSingleClicked.add(onSingleClicked); } catch (_) {}
    }
    tokens.actvToken = await ws.onActivated.add(onWorksheetActivated);

    _subscribed.set(id, tokens);
  }

  async function registerAllWorksheetsAndWatch() {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/id");
      await ctx.sync();

      for (const ws of sheets.items) {
        await registerListenersForWorksheet(ws);
      }

      // Try to watch for newly added worksheets (if supported on this host).
      if (sheets.onAdded && typeof sheets.onAdded.add === "function") {
        try {
          await sheets.onAdded.add(async (ev) => {
            try {
              await Excel.run(async (ctx2) => {
                const ws2 = ctx2.workbook.worksheets.getItem(ev.worksheetId);
                await registerListenersForWorksheet(ws2);
                await ctx2.sync();
              });
            } catch (err) { /* ignore */ }
          });
        } catch (_) {}
      }
    });
  }

  // LaunchEvent handler so the shared runtime initializes at workbook open.
  async function onDocumentOpened(event) {
    try {
      // Nothing critical; Office.onReady will wire listeners.
    } catch (err) {
      safeHandle(err, { action: "onDocumentOpened", userMessage: "Startup failed." });
    } finally {
      try { event.completed(); } catch (_) {}
    }
  }

  // Optional legacy actions (kept for compatibility)
  async function showTaskpane() {
    try { await Office.addin.showAsTaskpane(); } catch (err) { return safeHandle(err, { action: "showTaskpane" }); }
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
      return safeHandle(err, { action: "quickFillYellow" });
    }
  }

  // Wire everything once Excel is ready
  Office.onReady(async (info) => {
    try {
      if (info.host === Office.HostType.Excel) {
        // Ensure the runtime autoloads on future document opens (works on Excel for the web)
        try { await Office.addin.setStartupBehavior(Office.StartupBehavior.load); } catch (_) {}
        await registerAllWorksheetsAndWatch();
        // Check current selection immediately in case we're already on the trigger
        await evaluateAndTogglePane(null);
        // Associate any actions still referenced by the manifest
        if (Office.actions && Office.actions.associate) {
          Office.actions.associate("onDocumentOpened", onDocumentOpened);
          Office.actions.associate("showTaskpane", showTaskpane);
          Office.actions.associate("quickFillYellow", quickFillYellow);
        }
      }
    } catch (err) {
      safeHandle(err, { action: "onReady", userMessage: "Couldn't initialize the add-in." });
    }
  });
})();
