//
// Code that builds a front-end menu, and functions that it calls.
//

// The onOpen function is executed automatically every time a Spreadsheet is loaded
// We use it to create our menu
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  // menuEntries.push({name: "Prepare Theme Camps", functionName: "setupThemeCamps"});
  // menuEntries.push(null); // line separator
  menuEntries.push({name: "Toggle Wrap", functionName: "toggleWrap"});
  menuEntries.push(null); // line separator
  menuEntries.push({name: "Export this Sheet", functionName: "appConfirmExport"});

  ss.addMenu("-= LoF =-", menuEntries);
}


// protect the header row of current sheet
function protectThisHeaderRow() {
  doc = SpreadsheetApp.getActiveSpreadsheet()
  sheet = doc.getActiveSheet()
  protectHeaderRow(sheet)
}


// toggle wordwrap for entire sheet
function toggleWrap() {
  doc = SpreadsheetApp.getActiveSpreadsheet()
  sheet = doc.getActiveSheet()
  // get wrap status from A1
  cell = sheet.getRange("A1")
  // get entire range of cells
  r = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn())
  // toggle the wrap
  r.setWrap(!cell.getWrap())  
}


// destructive function to wipe all rows 2 and below
// be careful
function wipeRows() {
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("[dev] theme camps SORTABLE")
  // this works, but menu buttons are global across all sheets so it's dangerous.
  // sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  wipeRowRange(sheet, 2)
}




