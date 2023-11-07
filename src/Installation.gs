function createSpreadsheetOpenTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("onInsertRow")
    .forSpreadsheet(ss)
    .onChange()
    .create();
};
