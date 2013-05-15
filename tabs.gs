//
// Code that operates on tabs.
//

// Helper methods

function setupStandardTab(sheet_name, named_range) {
  doc = SpreadsheetApp.getActiveSpreadsheet()
  sheet = doc.getSheetByName(sheet_name)  
  if(!sheet) {
    doc.toast("sheet '"+sheet_name+"' not found")
    return
  }
  createCONTINUEsFromNamedRange( sheet, "$A$1", "GlobalTimestamp" )
  createCONTINUEsFromNamedRange( sheet, "$I$1", named_range )
}

// Runnable methods that set up specific tabs.

function setupThemeCamps() {
  doc = SpreadsheetApp.getActiveSpreadsheet()
  sheet = doc.getSheetByName(":: theme camps ::")
  if(!sheet) {
    doc.toast("sheet not found")
    return
  }
     
  createCONTINUEcolumn( sheet, "$A$1", "GlobalTimestamp" )  
  createCONTINUEsFromNamedRange( sheet, "$B$1" , "GlobalTimestamp" )
  createCONTINUEsFromNamedRange( sheet, "$J$1" , "ThemeCamps" )
}

function setupArtProjects() {
  setupStandardTab(":: art projects ::", "ArtProjects")
}

function setupConclave() {
  setupStandardTab(":: conclave ::", "Conclave")
}  

function setupFlameEffects() {
  setupStandardTab(":: flame effects ::", "FEFlameEffect")
}  

function setupOpenFlame() {
  setupStandardTab(":: open flame ::", "FEOpenFire")
}  

function setupDMV() {
  setupStandardTab(":: dmv ::", "DMV")
}  

function setupSound() {
  setupStandardTab(":: sound ::", "Sound")
}  

function setupVolunteer() {
  setupStandardTab(":: volunteer ::", "Volunteer")
}  

