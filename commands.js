
(function () {
  async function showTaskpane() {
    try {
      await Office.addin.showAsTaskpane();
      return true;
    } catch (err) {
      console.error(err);
      return false;
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
      console.error(err);
      return false;
    }
  }

  // Associate functions with the ribbon commands defined in manifest
  Office.actions.associate("showTaskpane", showTaskpane);
  Office.actions.associate("quickFillYellow", quickFillYellow);
})();
