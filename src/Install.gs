/*
  Copyright (c) 2023, drtingtp <dr.tingtp@gmail.com>
  All rights reserved.

  This source code is licensed under the BSD-style license found in the
  LICENSE file in the root directory of this source tree.

  https://raw.githubusercontent.com/drtingtp/sheets-kpi-dentur/main/LICENSE
*/

function createTrigger() {
  const ss = SpreadsheetApp.getActive();

  var allTriggers = ScriptApp.getProjectTriggers();
  var onInsertRow_installed = false;
  for (var i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction().toString() == "onInsertRow") {
      onInsertRow_installed = true;
      break;
    };
  };

  if (onInsertRow_installed === false) {
    ScriptApp.newTrigger("onInsertRow")
    .forSpreadsheet(ss)
    .onChange()
    .create();

    SpreadsheetApp.getUi().alert("onInsertRow trigger created.")
  };
};
