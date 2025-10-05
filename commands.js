
(function () {
  async function showTaskpane() {
    const env = await window.ErrorHandler.getEnvFlag();
    try {
      await Office.addin.showAsTaskpane();
      return true;
    } catch (err) {
      return window.ErrorHandler.handle(err, { action: "showTaskpane", userMessage: err.message });
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
    } catch (e) {
      return window.ErrorHandler.handle(err, { action: "showTaskpane", userMessage: err.message });
    }
  }

  // Associate functions with the ribbon commands defined in manifest
  Office.actions.associate("showTaskpane", showTaskpane);
  Office.actions.associate("quickFillYellow", quickFillYellow);
})();
