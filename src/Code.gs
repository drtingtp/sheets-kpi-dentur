const sheetNameDentur = "Dentur";

/*
  not used for now
  need to install onEdit trigger
  unable to automate without getting extra permission
*/
function onInsertRow(e) {
  
  if (e.changeType == "INSERT_ROW") {
    var spreadsheet = e.source
    //console.log("sheet name " + spreadsheet.getSheetName())

    if (spreadsheet.getSheetName() == sheetNameDentur){
      //console.log("sheet name matched")
      FillDenturFormula()
      // WIP: extend filter range
    };
  };
};

function FillDenturFormula() {
  const firstColumn = "R";
  const lastColumn = "AO";
  const sourceRow = "3";

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameDentur);

  var sourceRangeNotation = `${firstColumn}${sourceRow}:${lastColumn}${sourceRow}`;
  //console.log("sourceRangeNotation " + sourceRangeNotation)
  var sourceRange = sheet.getRange(sourceRangeNotation);
  var maxRow = sheet.getMaxRows();
  //console.log("maxRow " + maxRow)

  // find last row in 1st column
  var lastFilledRow = sheet
    .getRange(`${firstColumn}${maxRow}`)
    .getNextDataCell(SpreadsheetApp.Direction.UP)
    .getRow();

  console.log("lastFilledRow " + lastFilledRow, "maxRow " + maxRow)

  // early return if formula already filled - i.e. when getNextDataCell() selects row 1
  if (lastFilledRow === 1) {
    return;
  };

  var fillDownRange = sheet.getRange(
    lastFilledRow + 1,
    sourceRange.getColumn(),
    maxRow - lastFilledRow,
    sourceRange.getNumColumns()
  );

  sourceRange.copyTo(fillDownRange);
};

function ResetPeringkat() {
  // define range names for loop
  var ranges = [
    "Peringkat_BilDenturDiisu",
    "Peringkat_TarikhJT",
    "Peringkat_TarikhPeringkat",
    "Peringkat_Ulang",
  ];

  // the cell to return to after clearing
  var initialCellName = "D3"

  // shortcut to current sheet
  var sheet = SpreadsheetApp.getActive()

  // loop names and clear cells
  for (var i = 0; i < ranges.length; i++) {
    var range = sheet.getRangeByName(ranges[i]);

    if (ranges[i] == "Peringkat_Ulang") {
      range.setValue("FALSE");
    } else {
      range.clearContent();
    }
  }

  // return to initialCellName
  var range = sheet.getRange(initialCellName)
  sheet.setActiveRange(range)
}
