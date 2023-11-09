/*
  Copyright (c) 2023, drtingtp <dr.tingtp@gmail.com>
  All rights reserved.

  This source code is licensed under the BSD-style license found in the
  LICENSE file in the root directory of this source tree.

  https://raw.githubusercontent.com/drtingtp/sheets-kpi-dentur/main/LICENSE
*/

const sheetDenturName = "Dentur";
const sheetDentur_firstDataCol = 1;
const sheetDentur_firstDataRow = 3;
const sheetDentur_firstPeringkatCol = 10;
const sheetDentur_lastPeringkatCol = 16;

const formulaFirstColumn = "R";
const formulaLastColumn = "AO";

const sheetPeringkatName = "Peringkat";

const sheetPeringkat_dataRanges = [
  "Peringkat_BilDenturDiisu",
  "Peringkat_TarikhJT",
  "Peringkat_TarikhOfficer",
  "Peringkat_TarikhSesuai",
  "Peringkat_Ulang",
];

function onOpen(e) {
  // Add an Addon menu to the spreadsheet.
  SpreadsheetApp.getUi()
    .createMenu("KPI Dentur")
    .addItem("Install", "createTrigger")
    .addItem("Reset Peringkat", "ResetPeringkat")
    .addItem("Fix Dentur formula", "FillDenturFormula")
    .addToUi();
};

function onInsertRow(e) {
  if (e.changeType == "INSERT_ROW") {
    var spreadsheet = e.source
    //console.log("sheet name " + spreadsheet.getSheetName())

    if (spreadsheet.getSheetName() == sheetDenturName) {
      //console.log("sheet name matched")
      FillDenturFormula()
      ResizeFilter()
    };
  };
};

function FillDenturFormula() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetDenturName);
  const sourceRangeNotation = `${formulaFirstColumn}${sheetDentur_firstDataRow}:${formulaLastColumn}${sheetDentur_firstDataRow}`;
  //console.log("sourceRangeNotation " + sourceRangeNotation)
  var sourceRange = sheet.getRange(sourceRangeNotation);
  var maxRow = sheet.getMaxRows();
  //console.log("maxRow " + maxRow)

  // find last row in 1st column
  var lastFilledRow = sheet
    .getRange(`${formulaFirstColumn}${maxRow}`)
    .getNextDataCell(SpreadsheetApp.Direction.UP)
    .getRow();

  //console.log("lastFilledRow " + lastFilledRow, "maxRow " + maxRow)

  // early return if formula already filled - i.e. when getNextDataCell() selects row 1
  if (lastFilledRow === 1) {
    return;
  };

  // define the destination range
  var fillDownRange = sheet.getRange(
    lastFilledRow + 1,
    sourceRange.getColumn(),
    maxRow - lastFilledRow,
    sourceRange.getNumColumns()
  );

  sourceRange.copyTo(fillDownRange);
};

function CopyPeringkatSummary() {
  var ss = SpreadsheetApp.getActive()
  var sourceSheet = ss.getSheetByName(sheetPeringkatName)
  var sourceRange = sourceSheet.getRange("Peringkat_Summary")

  var destSheet = SpreadsheetApp.getActive().getSheetByName(sheetDenturName)
  ss.setActiveSheet(destSheet, true)
  // var destRange = destSheet.getCurrentCell()

  // // check whether destination current cell is in first peringkat col
  // if (destRange.getColumn() != sheetDentur_firstPeringkatCol) {
  //   SpreadsheetApp.getUi().alert("Active cell in Dentur not in the correct column.");
  //   return;
  // }

  // sourceRange.copyTo(destRange, {contentsOnly:true})
};

function ResetPeringkat() {
  // the cell to return to after reset
  var initialCellName = "D2"

  // shortcut to current sheet
  var sheet = SpreadsheetApp.getActive()

  // loop names and reset cells
  for (var i = 0; i < sheetPeringkat_dataRanges.length; i++) {
    var range = sheet.getRangeByName(sheetPeringkat_dataRanges[i]);

    if (sheetPeringkat_dataRanges[i] == "Peringkat_Ulang") {
      range.setValue("FALSE");
    } else {
      range.clearContent();
    }
  }

  // return to initialCellName
  var range = sheet.getRange(initialCellName)
  sheet.setActiveRange(range)
};

function ResizeFilter() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetDenturName)
  var filter = sheet.getFilter()

  var range = sheet.getRange(
    sheetDentur_firstDataRow - 1,
    sheetDentur_firstDataCol,
    sheet.getMaxRows() - sheetDentur_firstDataRow + 2,
    sheet.getLastColumn()
  );

  if (filter != null) {
    filter.remove()
  };

  range.createFilter();
};
