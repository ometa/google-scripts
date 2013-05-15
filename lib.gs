
n createCONTINUEs( sheet, referencing, row, col, amount_rows, amount_cols ) {
  // debug alerts
  //sheet.getParent().toast( "@"+row+", "+col+": " + "=CONTINUE( "+referencing+" ; "+amount_rows+" ; "+amount_cols+" )" )
  var range = sheet.getRange( row, col, amount_rows, amount_cols )

  // save our upper leftmost original value
  var orig = range.getFormula()  
  
  var matrix = new Array( amount_rows ) ;
  for ( i = 0 ; i < matrix.length ; i++ ) { 
    matrix[i] = new Array( amount_cols )
    for ( j = 0 ; j < matrix[i].length ; j++ ) 
      matrix[i][j] = "=CONTINUE( "+referencing+" ; "+(i+1)+" ; "+(j+1)+" )"
  }
  matrix[0][0] = orig 
  //sheet.getParent().toast( matrix.toString() )
  range.setFormulas( matrix )        
}


/*
Create absolute =CONTINUEs down a single column.

sheet - the sheet object to edit 
referencing: Absolute cell reference containing the ArrayFormula            
  ex: "$A$1"
name: figure out how many rows we need to make by counting 
      the rows inside this named range.
  ex: "GlobalTimestamp"
*/
function createCONTINUEcolumn(sheet, referencing, name) {
  var row = sheet.getRange(referencing).getRow()
  var col = sheet.getRange(referencing).getColumn()
  
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(name)
  var num_cols = 1
  var num_rows = range.getNumRows()
  
//  sheet.getParent().toast(referencing+","+row+","+col+","+num_rows+","+num_cols)
  sheet.getParent().toast("Creating CONTINUE column of "+num_rows+" rows at "+referencing)
  createCONTINUEs (sheet, referencing, row, col, num_rows, num_cols)
}


/*
Create a range of =CONTINUEs that encompases the size of the named
range range we're using in our ArrayFormula.

sheet - the sheet object to edit 
referencing: Absolute reference to cell containing the ArrayFormula            
  ex: "$B$1"
name: the named range to import.
  ex: "GlobalTimestamp"
*/
function createCONTINUEsFromNamedRange(sheet, referencing, name) {
  var row = sheet.getRange(referencing).getRow()
  var col = sheet.getRange(referencing).getColumn()
  
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(name)
  var num_cols = range.getNumColumns()
  var num_rows = range.getNumRows()
  
  //sheet.getParent().toast(referencing+","+row+","+col+","+num_rows+","+num_cols)
  sheet.getParent().toast("Creating CONTINUES for "+name+" at "+referencing)
  createCONTINUEs (sheet, referencing, row, col, num_rows, num_cols)
}

function wipeRowRange(sheet, starting_row) {
  var rows = sheet.getLastRow()
  sheet.getParent().toast("Wiping rows "+starting_row+" through "+rows)
  sheet.deleteRows(starting_row, rows-starting_row)
}



// set the first row (row A) to be protected to prevent mess-ups
// I don't think one can protect just a row via code.
//function protectHeaderRow(sheet) {
//  var doc = SpreadsheetApp.getActiveSpreadsheet()
//  var sheet = doc.getActiveSheet()
//  r = sheet.getRange(1,1,sheet.getLastRow())
//}
  

/*
copySheet: duplicate a sheet by copy/pasting VALUES only.

filterCol - skip rows without columns in filterCol
column count starts from 0 in this case.
*/
function copySheet(filterCol) {
  var suffix = '[export]'
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  oldsheet = doc.getActiveSheet()
  filterCol == filterCol - 1
//  oldsheet.getParent().toast("filtering on column: "+filterCol)
  
  var sheetName = oldsheet.getName() + " " + suffix
  var sheetIndex = oldsheet.getIndex()
  doc.insertSheet(sheetName, sheetIndex)
  var newsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
   
  var matrix = oldsheet.getSheetValues(1,1, oldsheet.getLastRow(), oldsheet.getLastColumn())
  for ( i=0; i < matrix.length; i++) {
    if (matrix[i][filterCol] != "")
      newsheet.appendRow(matrix[i])
  }
}





// old way to create a single column
// createCONTINUEs( sheet, "$A$1" , 1, 1 , 800 , 1 ) ;

//var dataRange = sheet.getRange(1,1,data.length,headers.length);
//dataRange.setValues(data);
//dataRange.setWrap(false);


// SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
