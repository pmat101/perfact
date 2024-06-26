//   Set trigger function to run 'onChange' at Event type 'onChange' 

function onChange(e) {
   if(e.changeType == 'INSERT_ROW') {
     const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
     let activeSheet = e.source.getActiveSheet();
     if(activeSheet.getName() == "Timelines") {
       let editedRow = activeSheet.getSelection().getActiveRange().getRow();
       let targetSheet = spreadSheet.getSheetByName("FAE Dashboard");
       if(editedRow < targetSheet.getLastRow()) {
         targetSheet.insertRowAfter(editedRow - 1);
       }
     }
   }
}

function setUpTrigger() {
  ScriptApp.newTrigger('onChange')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange(e)
    .create();
}
