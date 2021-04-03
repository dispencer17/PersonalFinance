/** @OnlyCurrentDoc */
const COLORS = {
  YELLOW: '#ffe599',
  BLUE: '#a4c2f4',
  GREEN: '#b6d7a8',
  PURPLE: '#b4a7d6',
  RED: '#ea9999',
  LIMEGREEN: '#bada55'
}
const LAYOUT = {
  NUM_HEADING_ROWS: 5,
  NUM_COLS_BETWEEN_ACCOUNTS: 1,
  FIRST_ENVELOPE_COLUMN: 2,
}
const MAIN_SAVINGS_ACCOUNT = {
  NAME: 'Membership Savings',
  STARTING_CELL: 'A18',
  NUM_ENVELOPES: 24
}
const LONG_TERM_SAVINGS_ACCOUNT = {
  NAME: 'Mission Savings Fund',
  STARTING_CELL: 'A4',
  NUM_ENVELOPES: 19
}
const ADDITONAL_SAVINGS_ACCOUNT_ONE = {
  NAME: 'SAVINGS - XXXXX5756',
  NUM_ENVELOPES: 0
}
const CHECKING_ACCOUNT_ONE = {
  NAME: 'CHECKING - XXXXX4438',
  NUM_ENVELOPES: 3
}
const CHECKING_ACCOUNT_TWO = {
  NAME: 'First Choice Platinum',
  NUM_ENVELOPES: 0
}
const CREDIT_CARD_ONE = {}
const CREDIT_CARD_TWO = {}
const HEALTH_SAVINGS_ACCOUNT = {
  NAME: 'Health Savings Account',
  NUM_ENVELOPES: 0
}
/* Global variables */
var spreadsheet = SpreadsheetApp.getActive();
spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Transactions'), true)
var tsheet = spreadsheet.getActiveSheet();

// Transaction data. Set by gatherTransactionData().
var oldmonth; var date; var month; var lineItem; var envelope; var category; var account; var label; var note;

//Spreadsheet data
var msheet; var rowOffset; var envelopeCol;

// Sumation rows data
var summationRow; var firstRow; var rsFirstCol; var cbcFirstCol; var rowSummationRange; var rsTotalColumn; var cbcTotalColumn;
var rsSummationRange; var cbcSummationRange;

// Line item entry data
var row;

//k is the RS index tracker
var k = 0;

//j is the Long Term Savings index tracker
var j = 0;

//i is the transaction index tracker
var i = 2;


function containsTerm(str, term, caseSensitive = false) {
  if (caseSensitive == false) {
    str = str.toLowerCase();
    term = term.toLowerCase();
  }
  return (str.indexOf(term) > -1)
}
function skipTransaction() {
  setOldMonth();
  markOffTransactionInTsheet('Skipped');
}
function setRowOffSetByAccount(actnt = account) {
  if(actnt == LONG_TERM_SAVINGS_ACCOUNT.NAME){
    rowOffset = msheet.createTextFinder(envelope).startFrom(msheet.getRange(LONG_TERM_SAVINGS_ACCOUNT.STARTING_CELL)).matchCase(true).findNext().getRow() + LAYOUT.NUM_HEADING_ROWS;      
  }
  else if(envelope == 'N/A'){}
  //Find envelope if in CBC or Regular Savings
  else {
    rowOffset = msheet.createTextFinder(envelope).startFrom(msheet.getRange(MAIN_SAVINGS_ACCOUNT.STARTING_CELL)).matchCase(true).findNext().getRow() + LAYOUT.NUM_HEADING_ROWS;
  }
}
function setEnvelopeColByAccount(actnt = account) {
  if(actnt == LONG_TERM_SAVINGS_ACCOUNT.NAME){
    envelopeCol = msheet.createTextFinder(envelope).startFrom(msheet.getRange(LONG_TERM_SAVINGS_ACCOUNT.STARTING_CELL)).matchCase(true).findNext().getColumn();    
  }
  else if(envelope == 'N/A'){}
  //Find envelope if in CBC or Regular Savings
  else {
    envelopeCol = msheet.createTextFinder(envelope).startFrom(msheet.getRange(MAIN_SAVINGS_ACCOUNT.STARTING_CELL)).matchCase(true).findNext().getColumn();
  }
}
function setOldMonth() {
  oldmonth = month
}
function updateRow(incrementer) {
  row = rowOffset + incrementer;
}
function addNewRowInMsheet() {
  // msheet.getRange(row+':'+row).activate();
  // var lr = spreadsheet.getActiveRange().getLastRow();
  msheet.insertRowsAfter(row, 1);
  spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
}
function deleteExtraRowAtEndOfMonth(monthToSumSheet) {
  // Delete extra row above the summation row that was added after the last transaction.
  var extraRow = summationRow - 1;
  monthToSumSheet.getRange(extraRow+':'+extraRow).activate();
  spreadsheet.getActiveSheet().deleteRows(monthToSumSheet.getActiveRange().getRow(), monthToSumSheet.getActiveRange().getNumRows());
       
  // reset Long term savings and Regular savings row incrementation
  k = 0;
  j = 0;
}
function setLineItem(tithing = false, tithe, tithingCol, tithingRow) {
  if(tithing == true){
      setTithing(tithe, tithingCol, tithingRow);
  }
  else {
    msheet.getRange(row,envelopeCol).activate();
    msheet.getCurrentCell().setValue(lineItem);
  }
}
function setTithing(tithe, tithingCol, tithingRow) {
  msheet.getRange(tithingRow,tithingCol).activate();
  msheet.getCurrentCell().setValue(tithe);
  msheet.getRange(row,envelopeCol).activate();
  msheet.getCurrentCell().setValue(lineItem-tithe);
}
function setDate() {
  msheet.getRange(row,1).activate();
  msheet.getCurrentCell().setValue(date);
}
function balanceCredit(catchAllEnvelope) {
  if(catchAllEnvelope == 'CBC Transfer' || catchAllEnvelope == 'Mission Savings Fund Transfer') {
    var caCol = msheet.createTextFinder(catchAllEnvelope).matchCase(true).findNext().getColumn();
    setLineItem();
    msheet.getRange(row,caCol).activate();
    msheet.getCurrentCell().setValue(-1*lineItem);
  }
  else if(catchAllEnvelope == 'None') {}
  else {
    var caCol = msheet.createTextFinder(catchAllEnvelope).matchCase(true).findNext().getColumn();
    msheet.getRange(row,caCol).activate();
    msheet.getCurrentCell().setValue(-1*lineItem);
  }
}
function makeRecordedColumnInTsheet() {
  tsheet.getRange('M1').activate();
  tsheet.getCurrentCell().setValue('Recorded');
  tsheet.getRange('N1').activate();
  tsheet.getCurrentCell().setValue('Row in the month');
}
function getRecordedState() {
  tsheet.getRange(i,13).activate();
  var recorded = tsheet.getActiveCell().getValue();
  return recorded
}
function findCurrentTransaction() {
  var recordedState = getRecordedState();  
  if(recordedState != ""){
    skipToCurrentTransaction();
  }
}
function filloutTithing() {
  var gross = note
  var tithe = gross*0.1
  //get tithing column
  if(account == MAIN_SAVINGS_ACCOUNT.NAME) {
    var tithingCol = msheet.createTextFinder('Tithing').startFrom(msheet.getRange(MAIN_SAVINGS_ACCOUNT.STARTING_CELL)).matchCase(true).findNext().getColumn();
    //Update row, Set date and line item
    updateRow(k);
    setLineItem(true, tithe, tithingCol, row)
    setDate();
  } 
  else if(account == LONG_TERM_SAVINGS_ACCOUNT.NAME) {
    var tithingCol = msheet.createTextFinder('To Regular Savings').startFrom(msheet.getRange(LONG_TERM_SAVINGS_ACCOUNT.STARTING_CELL)).matchCase(true).findNext().getColumn();
    //Update row, Set date and line item in MSF
    updateRow(j);
    setLineItem(true, tithe, tithingCol, row);
    setDate();
    //Balance in Regular savings
    envelopeCol = msheet.createTextFinder('Tithing').startFrom(msheet.getRange(MAIN_SAVINGS_ACCOUNT.STARTING_CELL)).matchCase(true).findNext().getColumn();
    lineItem = tithe;
    setRowOffSetByAccount(MAIN_SAVINGS_ACCOUNT.NAME);
    updateRow(k);
    balanceCredit('Mission Savings Fund Transfer');
    setDate();
  }
}
function markOffTransactionInTsheet(recordedState) {
  tsheet.getRange(i,13).activate();
  tsheet.getActiveCell().setValue(recordedState);
  tsheet.getRange(i,14).activate();
  tsheet.getActiveCell().setValue(k-1);
}
function skipToCurrentTransaction() {
  tsheet.getRange(1,13).activate();
  tsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  i = tsheet.getActiveCell().getRow();
  tsheet.getRange(i,2).activate();
  oldmonth = tsheet.getActiveCell().getValue();
  tsheet.getRange(i,14).activate();
  k = tsheet.getActiveCell().getValue();
  gatherTransactionData();
  gatherSpreadsheetData();
  updateRow(k);
  addNewRowInMsheet();
  k++;
  i++;
}
function updateMsheet(catchAllEnvelope, skip = false, tithing = false) {
  if(skip == true) {
    skipTransaction();
  }
  else if(tithing == true) {
    //Subtract gross tithing from net and set both line items
    filloutTithing();
    
    //set old month
    setOldMonth()

    //Increment row and add new row if not MSF
    if(account == LONG_TERM_SAVINGS_ACCOUNT.NAME) {
      j++;
      addNewRowInMsheet();
      k++;
      markOffTransactionInTsheet('In MSF');
    }
    else {
      addNewRowInMsheet();
      k++;
      markOffTransactionInTsheet('Yes');
    }
    
  }
  else {
    // update then increment row
    if(account == LONG_TERM_SAVINGS_ACCOUNT.NAME) {
      updateRow(j);
      j++;
    }
    else {
      updateRow(k);
      addNewRowInMsheet();
      k++;
    }
    //Set date and line item
    setLineItem();
    setDate()
    
    //Set balance credit in respective catch all column
    balanceCredit(catchAllEnvelope);
    
    //Set oldmonth and add new row
    setOldMonth();
    
    //mark off in tSheet
    markOffTransactionInTsheet('Yes');    
  }
}
function envelopeRSSums(monthToSumSheet) {
  monthToSumSheet.getRange(summationRow,rsFirstCol).activate();
  monthToSumSheet.getCurrentCell().setFormulaR1C1('=SUM(R['+rowSummationRange+']C[0]:R[-1]C[0])');
  var destinationRange = monthToSumSheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), rsTotalColumn - 2);
  monthToSumSheet.getActiveRange().autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}
function envelopeCBCSums(monthToSumSheet) {
  monthToSumSheet.getRange(summationRow,rsTotalColumn + 1).activate();
  monthToSumSheet.getCurrentCell().setFormulaR1C1('=SUM(R['+rowSummationRange+']C[0]:R[-1]C[0])');
  var destinationRange = monthToSumSheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 3);
  monthToSumSheet.getActiveRange().autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}
function totalRSSums(monthToSumSheet) {
  monthToSumSheet.getRange(firstRow,rsTotalColumn).activate();
  monthToSumSheet.getCurrentCell().setFormulaR1C1('=Sum(R[0]C['+rsSummationRange+']:R[0]C[-1])');
  var colSummationRange = spreadsheet.getActiveRange().getNumColumns();
  var destinationRange = monthToSumSheet.getActiveRange().offset(0, 0, -1*rowSummationRange, colSummationRange);
  monthToSumSheet.getActiveRange().autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}
function totalCBCSums(monthToSumSheet) {
  monthToSumSheet.getRange(firstRow,cbcTotalColumn).activate();
  monthToSumSheet.getCurrentCell().setFormulaR1C1('=Sum(R[0]C['+cbcSummationRange+']:R[0]C[-1])');
  var colSummationRange = spreadsheet.getActiveRange().getNumColumns();
  var destinationRange = monthToSumSheet.getActiveRange().offset(0, 0, -1*rowSummationRange, colSummationRange);
  monthToSumSheet.getActiveRange().autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);   
}
function rsSumCheck(monthToSumSheet) {
  monthToSumSheet.getRange(summationRow,rsTotalColumn).activate();
  monthToSumSheet.getCurrentCell().setFormulaR1C1('=IF(SUM(R[0]C['+rsSummationRange+']:R[0]C[-1])=Sum(R['+rowSummationRange+']C[0]:R[-1]C[0]),SUM(R[0]C['+rsSummationRange+']:R[0]C[-1]),"Sums don\'t match")')
}
function cbcSumCheck(monthToSumSheet) {
  monthToSumSheet.getRange(summationRow,cbcTotalColumn).activate();
  monthToSumSheet.getCurrentCell().setFormulaR1C1('=IF(SUM(R[0]C['+cbcSummationRange+']:R[0]C[-1])=Sum(R['+rowSummationRange+']C[0]:R[-1]C[0]),SUM(R[0]C['+cbcSummationRange+']:R[0]C[-1]),"Sums don\'t match")')
}
function finalizeMonthWithSumsAndDeleteExtraRows(month) {
  var monthToSumSheet = spreadsheet.getSheetByName(month);
  gatherSummationRowData(monthToSumSheet);
  envelopeRSSums(monthToSumSheet);
  envelopeCBCSums(monthToSumSheet);
  totalRSSums(monthToSumSheet);
  totalCBCSums(monthToSumSheet);
  rsSumCheck(monthToSumSheet);
  cbcSumCheck(monthToSumSheet);
  deleteExtraRowAtEndOfMonth(monthToSumSheet)
}
function gatherTransactionData() {
  tsheet.getRange(i,1).activate();
  date = tsheet.getActiveCell().getValue();
  tsheet.getRange(i,2).activate();
  month = tsheet.getActiveCell().getValue();
  tsheet.getRange(i,6).activate();
  lineItem = tsheet.getActiveCell().getValue();
  tsheet.getRange(i,9).activate();
  envelope = tsheet.getActiveCell().getValue();
  tsheet.getRange(i,8).activate();
  category = tsheet.getActiveCell().getValue();
  tsheet.getRange(i,10).activate();
  account = tsheet.getActiveCell().getValue();
  tsheet.getRange(i,11).activate();
  label = tsheet.getActiveCell().getValue();
  tsheet.getRange(i,12).activate();
  note = tsheet.getActiveCell().getValue();
}
function gatherSpreadsheetData() {
   //Set new msheet to current month
  msheet = spreadsheet.getSheetByName(month);
  //find envelope if in Mission Savings Fund
  setRowOffSetByAccount();
  setEnvelopeColByAccount();
}
function gatherSummationRowData(monthToSumSheet) {
  summationRow = monthToSumSheet.createTextFinder('Monthly Total').startFrom(monthToSumSheet.getRange(MAIN_SAVINGS_ACCOUNT.STARTING_CELL)).matchCase(true).findNext().getRow();
  firstRow = rowOffset;
  rowSummationRange = firstRow - summationRow;
  rsFirstCol = LAYOUT.FIRST_ENVELOPE_COLUMN;
  rsTotalColumn = monthToSumSheet.createTextFinder('Total').startFrom(monthToSumSheet.getRange(MAIN_SAVINGS_ACCOUNT.STARTING_CELL)).matchCase(true).findNext().getColumn();
  rsSummationRange = rsFirstCol - rsTotalColumn;
  cbcFirstCol = rsTotalColumn + LAYOUT.NUM_COLS_BETWEEN_ACCOUNTS;
  cbcTotalColumn = cbcFirstCol + CHECKING_ACCOUNT_ONE.NUM_ENVELOPES; 
  cbcSummationRange = cbcFirstCol - cbcTotalColumn;
}
function processTransactions() {
  tsheet.getRange('I1').activate();
  envelope = tsheet.getActiveCell().getValue();
  var currentRow = 2;
  var check = '';
  if(envelope == 'Envelope'){
    tsheet.getRange('I2').activate();
    check = tsheet.getActiveCell().getValue();
    if(check != ''){
      tsheet.getRange('I1').activate();
      tsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
      currentRow = tsheet.getActiveCell().getRow();
      currentRow++;
    }   
  }
  else {
    spreadsheet.getRange('B:B').activate();
    spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 1);
    spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 1).activate();
    spreadsheet.getRange('B1').activate();
    spreadsheet.getCurrentCell().setValue('Month');
    spreadsheet.getRange('B2').activate();
    spreadsheet.getCurrentCell().setFormula('=Text(A2,"MMMM")');
    spreadsheet.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    
    spreadsheet.getRange('E:E').activate();
    spreadsheet.getActiveSheet().insertColumnsAfter(spreadsheet.getActiveRange().getLastColumn(), 1);
    spreadsheet.getActiveRange().offset(0, spreadsheet.getActiveRange().getNumColumns(), spreadsheet.getActiveRange().getNumRows(), 1).activate();
    spreadsheet.getRange('F1').activate();
    spreadsheet.getCurrentCell().setValue('Line Item');
    spreadsheet.getRange('F2').activate();
    spreadsheet.getCurrentCell().setFormula('=E2*IF(G2="debit",-1,1)');
    spreadsheet.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    
    tsheet = spreadsheet.getActiveSheet();
    tsheet.getRange(1, 1, tsheet.getMaxRows(), tsheet.getMaxColumns()).activate();
    tsheet = spreadsheet.getActiveSheet();
    tsheet.getRange(1, 1, tsheet.getMaxRows(), tsheet.getMaxColumns()).createFilter();
    spreadsheet.getCurrentCell().activate();
    spreadsheet.getActiveSheet().getFilter().sort(spreadsheet.getActiveRange().getColumn(), true);
    spreadsheet.getActiveSheet().getFilter().remove();
    
    spreadsheet.getRange('H:H').activate();
    spreadsheet.getActiveSheet().insertColumnsAfter(spreadsheet.getActiveRange().getLastColumn(), 1);
    spreadsheet.getRange('I1').activate();
    spreadsheet.getCurrentCell().setValue('Envelope');
  }
  
  //Declaring variables
  lastRow = tsheet.getLastRow();
  
  for(var i = currentRow;i<lastRow+1;i++){
    tsheet.getRange(i,8).activate();
    category = spreadsheet.getActiveCell().getValue();
    tsheet.getRange(i,10).activate();
    account = tsheet.getActiveCell().getValue();
  
    //Color row by account
    if(account == CHECKING_ACCOUNT_ONE.NAME && category != 'Mortgage & Rent'){
      tsheet.getRange(i+':'+i).activate();
      tsheet.getActiveRangeList().setBackground(COLORS.YELLOW)
    }
    if(account == HEALTH_SAVINGS_ACCOUNT.NAME){
      tsheet.getRange(i+':'+i).activate();
      tsheet.getActiveRangeList().setBackground(COLORS.LIMEGREEN)
    }
    if(account == LONG_TERM_SAVINGS_ACCOUNT.NAME && category != 'Interest Income'){
      tsheet.getRange(i+':'+i).activate()
      tsheet.getActiveRangeList().setBackground(COLORS.BLUE);
      tsheet.getRange(i,7).activate();
      var sign = tsheet.getActiveCell().getValue();
      if(sign == 'debit'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('To Regular Savings');
      }
      if(sign == 'credit'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('From Regular Savings');
      }
    }
    //Set envelope
    else if(account == ADDITONAL_SAVINGS_ACCOUNT_ONE.NAME){
      tsheet.getRange(i,9).activate();
      tsheet.getCurrentCell().setValue('Savings');
    }
    else if(account == 'Roth Contributory IRA'||account == 'Joint Tenant                     '||
      account == 'Individual Brokerage'||
      account == 'Robinhood Investments'||containsTerm(account, 'DCRA')){                   
      tsheet.getRange(i+':'+i).activate();
      spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
      --i;
    }
    else{
      if(category == 'Activities'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Activities');
      }
      else if(category == 'Addison\'s Fun Money'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Fun Money - Addison');
      }
      else if(category == 'Auto Insurance'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Car Maintenance');
      }
      else if(category == 'Baby'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Baby');
      }
      else if(category == 'Baby Supplies'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Monthly Rollover');
      }
      else if(category == 'Babysitter & Daycare'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Monthly Rollover');
      }
      else if(category == 'Books & Supplies'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('School Supplies/electronics');
      }
      else if(category == 'Camera'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Camera');
      }
      else if(category == 'Chiropractor'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Medical (medication, appts, dental)');
      }
      else if(containsTerm(category, 'Christmas')){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Gift');
      }
      else if(category == 'Clothing'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Clothing');
      }
      else if(category == 'Credit Card Payment'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Credit Card Debt');
      }
      else if(category == 'Date Night'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Date Night');
      }
      else if(category == 'Dentist'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Medical (medication, appts, dental)');
      }
      else if(category == 'Doctor'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Medical (medication, appts, dental)');
      }
      else if(category == 'Electronics & Software'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('School Supplies/electronics');
      } 
      else if(category == 'Fast Food'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Monthly Rollover');
      }
      else if(category == 'Gas & Fuel'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Monthly Rollover');
      }
      else if(category == 'Gift'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Gift');
      }
      else if(category == 'Groceries'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Monthly Rollover');
      }
      else if(category == 'Health & Fitness'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Medical (medication, appts, dental)');
      }
      else if(category == 'Health Insurance'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Medical (medication, appts, dental)');
      }
      else if(category == 'Home Supplies'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Monthly Rollover');
      }
      else if(category == 'Income'){
        if(account == CHECKING_ACCOUNT_ONE.NAME){
          tsheet.getRange(i,9).activate();
          tsheet.getCurrentCell().setValue('Rent');
        }
        else{
          tsheet.getRange(i,9).activate();
          tsheet.getCurrentCell().setValue('Monthly Rollover');
        }
      }
      else if(category == 'Interest Income'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Interest');
      }
      else if(category == 'Investments'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Investments');
      }
      else if(category == 'Late Fee'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Monthly Rollover');
      }
      else if(category == 'LDS Tithing'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Tithing');
      }
      else if(category == 'Mortgage & Rent'){
        if(account == MAIN_SAVINGS_ACCOUNT.NAME || containsTerm(account,'credit card') || containsTerm(account,'Platinum Rewards')){
          tsheet.getRange(i,9).activate();
          tsheet.getCurrentCell().setValue('CBC Transfer');
        }
        else{
          tsheet.getRange(i,9).activate();
          tsheet.getCurrentCell().setValue('Rent');
        }
      }
      else if(category == 'Paycheck'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Monthly Rollover');
      }
      else if(category == 'Personal Care'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Monthly Rollover');
      }
      else if(category == 'Pharmacy'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Medical (medication, appts, dental)');
      }
      else if(category == 'Possessions'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Possessions');
      }
      else if(category == 'Service & Parts'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Car Maintenance');
      }
      else if(category == 'Special Occasions'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Special Occassions');
      }
      else if(category == 'Spencer\'s Fun Money'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Fun Money - Spencer');
      }
      else if(category == 'Stock investment'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Investments');
      }
      else if(category == 'Subscriptions'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Subscriptions');
      }
      else if(category == 'Temple Clothes'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Clothing');
      }
      else if(category == 'Transfer'){
        if(account == CHECKING_ACCOUNT_TWO.NAME){
          tsheet.getRange(i,9).activate();
          tsheet.getCurrentCell().setValue('N/A');
          tsheet.getRange(i+':'+i).activate();
          tsheet.getActiveRangeList().setBackground(COLORS.GREEN);
        }
        else if(account == CHECKING_ACCOUNT_ONE.NAME){
          tsheet.getRange(i,9).activate();
          tsheet.getCurrentCell().setValue('Rent');
          tsheet.getRange(i+':'+i).activate();
          tsheet.getActiveRangeList().setBackground(COLORS.PURPLE);
        }
        else{
          tsheet.getRange(i,9).activate();
          tsheet.getCurrentCell().setValue('Monthly Rollover');
          tsheet.getRange(i+':'+i).activate();
          tsheet.getActiveRangeList().setBackground(COLORS.PURPLE);
        }
      }
      else if(category == 'Travel'){
        tsheet.getRange(i,9).activate();
        tsheet.getCurrentCell().setValue('Travel');
      }
      else if(category == 'Tuition'){
        if(account == CHECKING_ACCOUNT_ONE.NAME){
          tsheet.getRange(i,9).activate();
          tsheet.getCurrentCell().setValue('Rent');
        }
        else{
          tsheet.getRange(i,9).activate();
          tsheet.getCurrentCell().setValue('School Supplies/electronics');
        }
      }
      else{
        tsheet.getRange(i+':'+i).activate();
        tsheet.getActiveRangeList().setBackground(COLORS.RED);
      }
    }
  }
  spreadsheet.getActiveSheet().setFrozenRows(1);
  spreadsheet.getActiveSheet().setFrozenColumns(1);
}
function fillLineItems(){
  var lastRow = tsheet.getLastRow();
  //var lastLine;
  var catchAllEnvelope;

  makeRecordedColumnInTsheet();
  findCurrentTransaction();
  
  //loop through all transactions and populate variables
  //i is the transaction index tracker
  for(i = i;i<lastRow+1;i++){  
    //Populate all global transaction data variables  
    gatherTransactionData();
    gatherSpreadsheetData();
    
    //Account if statments
    if(account == CHECKING_ACCOUNT_TWO.NAME && envelope == 'Monthly Rollover'){
      updateMsheet(catchAllEnvelope, true);
    }
    else if(envelope == 'N/A'){
      updateMsheet(catchAllEnvelope, true);
    }
    else if(account == CHECKING_ACCOUNT_ONE.NAME && envelope != 'Rent'){
      var rentCol = msheet.createTextFinder('Rent').matchCase(true).findNext().getColumn();
      catchAllEnvelope = 'CBC Transfer';
      updateMsheet(catchAllEnvelope);
    }
    else if(account == HEALTH_SAVINGS_ACCOUNT.NAME){
      updateMsheet(catchAllEnvelope, true);
    }
    else if(account == CHECKING_ACCOUNT_TWO.NAME && envelope != 'Monthly Rollover'){
      catchAllEnvelope = 'Monthly Rollover';
      updateMsheet(catchAllEnvelope);
    }
    else if(label == 'Tithing Unpaid' || label == 'Tithing Paid') {
      catchAllEnvelope = 'None';
      updateMsheet(catchAllEnvelope, false, true);
    }
    else if(account == MAIN_SAVINGS_ACCOUNT.NAME){
      catchAllEnvelope = 'None';
      updateMsheet(catchAllEnvelope);
    }
    else if(account == LONG_TERM_SAVINGS_ACCOUNT.NAME){
      catchAllEnvelope = 'None';
      updateMsheet(catchAllEnvelope);
    }
    else if((containsTerm(account,'credit card') || containsTerm(account,'Platinum Rewards')) && envelope == 'Credit Card Debt'){
      updateMsheet(catchAllEnvelope, true);
    }
    else if(containsTerm(account,'credit card') || containsTerm(account,'Platinum Rewards')){
      catchAllEnvelope = 'Credit Card Debt';
      updateMsheet(catchAllEnvelope);
    }
    else if(account == 'Cash'){
      if(containsTerm(note,'Missing credit card transaction')){
        catchAllEnvelope = 'Credit Card Debt';
      }
      else{
        catchAllEnvelope = 'Monthly Rollover';
      }
      updateMsheet(catchAllEnvelope);
    }
    else {
      catchAllEnvelope = 'None';
      updateMsheet(catchAllEnvelope);
    }

    //Start new month and add old month summations
    if(month != oldmonth && i != 2){
      finalizeMonthWithSumsAndDeleteExtraRows(oldmonth);
    }
    else if(i == lastRow){
      finalizeMonthWithSumsAndDeleteExtraRows(month);
    }
  }//for loop
};


