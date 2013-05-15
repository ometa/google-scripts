//
// Code that builds UI elements (like pop-ups)
//

// create a dialog box 'app', prompting the user
// to enter a column that they know should never be empty on this sheet
function appConfirmExport() {
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = doc.getActiveSheet()
  var app = UiApp.createApplication().setTitle('Export Sheet').setWidth(700).setHeight(150)

  var panel = app.createVerticalPanel();

  panel.add(app.createLabel("Select the column name in this sheet that should never be blank."))

  var lb = app.createListBox().setName('column')  
  var col_names = sheet.getSheetValues(1,1,1,sheet.getLastColumn())[0]
  for (i = 0; i < col_names.length; i++) {
    lb.addItem(col_names[i], i)    
  }
  panel.add(lb);
  
  var button = app.createButton('Export');
  var handler = app.createServerHandler('serverHandlerCopyTeams').addCallbackElement(panel).validateInteger(lb)
  button.addClickHandler(handler)
  
  var closeButton = app.createButton('Close');
  var closeHandler = app.createServerClickHandler('close');
  closeButton.addClickHandler(closeHandler);
 
  var grid = app.createGrid(1,2)
  grid.setWidget(0,0,button)
  grid.setWidget(0,1,closeButton)
  panel.add(grid)
  
  app.add(panel)
  doc.show(app)
}

// pass our column parameter to copySheet function
function serverHandlerCopyTeams(e) {
  copySheet(e.parameter.column)
  close()
}


// Close everything & return when the close button is clicked
function close() {
  var app = UiApp.getActiveApplication()
  app.close()
  // The following line is REQUIRED for the widget to actually close.
  return app
}
