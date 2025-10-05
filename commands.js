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

  // Associate functions with the ribbon commands defined in manifest
  Office.actions.associate("showTaskpane", showTaskpane);
  Office.actions.associate("quickFillYellow", quickFillYellow);
})();
