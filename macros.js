function lightGreen() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('8:8').activate();
  spreadsheet.getActiveRangeList().setBackground('#b6d7a8');
  spreadsheet.getRange('G11').activate();
};