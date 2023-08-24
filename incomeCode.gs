var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet1 = ss.getSheetByName("Input");
var sheet2 = ss.getSheetByName("Calculations");
var sheet3 = ss.getSheetByName("Output");
var companyName = sheet1.getRange('E2').getValue();
var timePeriod = sheet1.getRange('F2').getValue();
var date = sheet1.getRange('G2').getValue();
var revenueNum = sheet2.getRange("A2").getValue();
var expenseNum = sheet2.getRange("B2").getValue();
var revenueSum = Number(ss.getRange("A4").getValue());
var expenseSum = Number(ss.getRange("B4").getValue());
var difference = sheet2.getRange("C2").getValue();
var rowNum = revenueNum + expenseNum + 6;

function income() {
  //sort revenue
  sheet1.getRange('A2:B').sort({column : 2,ascending : false});
  //creating/getting doc
  var doc = DocumentApp.create("Income Statement");
  sheet3.getRange(2,1).setValue(doc.getUrl());
  var body = doc.getBody();
  body.clear();
  //setting header
  var header = body.insertParagraph(0,'');
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  header.appendText(companyName + '\n' + 'Income Statement\n' + 'For the ' + timePeriod + ' ending ' + date).setFontFamily("Times");
  //building the table
  var table = body.appendTable([['']]);
  for (i=1; i<rowNum; i++) {
    table.appendTableRow().appendTableCell('');
  }
  for (i=0; i<rowNum; i++) {
    for (j=0; j<10; j++) {
      table.getRow(i).appendTableCell('');
    }
  }
  table.setColumnWidth(0,200);
  table.setColumnWidth(1,40);
  table.setColumnWidth(6,40);
  table.setColumnWidth(5,30);
  table.setColumnWidth(10,30);
  var profit = 'Income';
  if (revenueSum-expenseSum < 0) {
    profit = 'Loss';
  }
  table.getCell(0,0).setText('Revenue').setBold(true).setFontFamily("Times");
  table.getCell(revenueNum+1,0).setText('Total Revenue').setBold(true).setFontFamily("Times");
  table.getCell(revenueNum+3,0).setText('Expenses').setBold(true).setFontFamily("Times");
  table.getCell(revenueNum+expenseNum+4,0).setText('Total Expenses').setBold(true).setFontFamily("Times");
  table.getCell(revenueNum+expenseNum+5,0).setText('Net ' + profit).setBold(true).setFontFamily("Times");
  for (i=0; i<revenueNum; i++) {
    table.getCell(i+1,0).setText(sheet1.getRange(i+2,1).getValue()).setFontFamily("Times");
  }
  for (i=0; i<expenseNum; i++) {
    table.getCell(revenueNum+4+i, 0).setText(sheet1.getRange(i+2,3).getValue()).setFontFamily("Times");
  }
  var thousandsDigit = 0;
  if (sheet1.getRange(2,2).getValue() < 1000) {
    thousandsDigit = '';
  }
  else {
    thousandsDigit = Math.floor(parseInt(sheet1.getRange(2,2).getValue())/1000)
  }
  table.getCell(1,1).setText('$ ' + thousandsDigit).setFontFamily("Times");
  for (i=1; i<revenueNum; i++) {
    var revenueAmount = sheet1.getRange(i+2,2).getValue();
    var thousandsDigit = 0;
    if (parseInt(revenueAmount) < 1000) {
      thousandsDigit = '';
    }
    else {
      thousandsDigit = Math.floor(parseInt(revenueAmount)/1000);
    }
    table.getCell(i+1,1).setText(thousandsDigit).setFontFamily("Times");
  }
  for (i=0; i<revenueNum; i++) {
    var revenueAmount = sheet1.getRange(i+2,2).getValue();
    var thousandsDigit = 0;
    if (parseInt(revenueAmount) < 1000) {
      thousandsDigit = '';
    }
    else {
      thousandsDigit = Math.floor(parseInt(revenueAmount)/1000);
    }
    var hundredsDigit = '';
    if (parseInt(revenueAmount) >= 100) {
      hundredsDigit = Math.floor((parseInt(revenueAmount) - 1000*thousandsDigit)/100);
    }
    table.getCell(i+1,2).setText(hundredsDigit).setFontFamily("Times");
    var tensDigit = '';
    if (parseInt(revenueAmount) >= 10) {
      tensDigit = Math.floor((parseInt(revenueAmount)-1000*thousandsDigit-100*hundredsDigit)/10);
    }
    table.getCell(i+1,3).setText(tensDigit).setFontFamily("Times");
    var onesDigit = '';
    if (parseInt(revenueAmount) >= 10) {
      onesDigit = Math.floor(parseInt(revenueAmount)-1000*thousandsDigit-100*hundredsDigit-10*tensDigit);
    }
    table.getCell(i+1,4).setText(onesDigit).setFontFamily("Times");
    var cents = '--';
    if (Number(revenueAmount) - 1000*thousandsDigit-100*hundredsDigit-10*tensDigit - onesDigit > 0) {
      cents = Math.round(100*(Number(revenueAmount) - 1000*thousandsDigit-100*hundredsDigit-10*tensDigit - onesDigit));
    }
    if (cents < 10) {
      table.getCell(i+1,5).setText('0' + cents).setFontFamily("Times");
    }
    else {
      table.getCell(i+1,5).setText(cents).setFontFamily("Times");
    }
  }
  var thousandsDigit = 0;
  if (sheet1.getRange(2,4).getValue() < 1000) {
    thousandsDigit = '';
  }
  else {
    thousandsDigit = Math.floor(parseInt(sheet1.getRange(2,4).getValue())/1000)
  }
  table.getCell(revenueNum+4,1).setText('$ ' + thousandsDigit).setFontFamily("Times");
  for (i=1; i<expenseNum; i++) {
    var expenseAmount = sheet1.getRange(i+2,4).getValue();
    var thousandsDigit = 0;
    if (parseInt(expenseAmount) < 1000) {
      thousandsDigit = '';
    }
    else {
      thousandsDigit = Math.floor(parseInt(expenseAmount)/1000);
    }
    table.getCell(revenueNum+4+i,1).setText(thousandsDigit).setFontFamily("Times");
  }
  for (i=0; i<expenseNum; i++) {
    var expenseAmount = sheet1.getRange(i+2,4).getValue();
    var thousandsDigit = 0;
    if (parseInt(expenseAmount) < 1000) {
      thousandsDigit = '';
    }
    else {
      thousandsDigit = Math.floor(parseInt(expenseAmount)/1000);
    }
    var hundredsDigit = '';
    if (parseInt(expenseAmount) >= 100) {
      hundredsDigit = Math.floor((parseInt(expenseAmount) - 1000*thousandsDigit)/100);
    }
    table.getCell(revenueNum+4+i,2).setText(hundredsDigit).setFontFamily("Times");
    var tensDigit = '';
    if (parseInt(expenseAmount) >= 10) {
      tensDigit = Math.floor((parseInt(expenseAmount)-1000*thousandsDigit-100*hundredsDigit)/10);
    }
    table.getCell(revenueNum+4+i,3).setText(tensDigit).setFontFamily("Times");
    var onesDigit = '';
    if (parseInt(expenseAmount) >= 10) {
      onesDigit = Math.floor(parseInt(expenseAmount)-1000*thousandsDigit-100*hundredsDigit-10*tensDigit);
    }
    table.getCell(revenueNum+4+i,4).setText(onesDigit).setFontFamily("Times");
    var cents = '--';
    if (Number(expenseAmount) - 1000*thousandsDigit-100*hundredsDigit-10*tensDigit - onesDigit > 0) {
      cents = Math.round(100*(Number(expenseAmount) - 1000*thousandsDigit-100*hundredsDigit-10*tensDigit - onesDigit));
    }
    if (cents < 10) {
      table.getCell(revenueNum+4+i,5).setText('0' + cents).setFontFamily("Times");
    }
    else {
      table.getCell(revenueNum+4+i,5).setText(cents).setFontFamily("Times");
    }
  }
  var thousandsDigit = 0;
  if (revenueSum < 1000) {
    thousandsDigit = '';
  }
  else {
    thousandsDigit = Math.floor(revenueSum/1000);
  }
  table.getCell(revenueNum+1,6).setText('$ ' + thousandsDigit).setFontFamily("Times");
  var hundredsDigit = '';
  if (revenueSum >= 100) {
    hundredsDigit = Math.floor((revenueSum - 1000*thousandsDigit)/100);
  }
  table.getCell(revenueNum+1,7).setText(hundredsDigit).setFontFamily("Times");
  var tensDigit = '';
  if (revenueSum >= 10) {
    tensDigit = Math.floor((revenueSum-1000*thousandsDigit-100*hundredsDigit)/10);
  }
  table.getCell(revenueNum+1,8).setText(tensDigit).setFontFamily("Times");
  var onesDigit = '';
  if (revenueSum >= 10) {
    onesDigit = Math.floor(revenueSum-1000*thousandsDigit-100*hundredsDigit-10*tensDigit);
  }
  table.getCell(revenueNum+1,9).setText(onesDigit).setFontFamily("Times");
  var cents = '--';
  if (revenueSum - 1000*thousandsDigit-100*hundredsDigit-10*tensDigit - onesDigit > 0) {
    cents = Math.round(100*(revenueSum - 1000*thousandsDigit-100*hundredsDigit-10*tensDigit - onesDigit));
  }
  if (cents < 10) {
    table.getCell(revenueNum+1,10).setText('0' + cents).setFontFamily("Times");
  }
  else {
    table.getCell(revenueNum+1,10).setText(cents).setFontFamily("Times");
  }
  var thousandsDigit = '';
  if (expenseSum >= 1000) {
    thousandsDigit = Math.floor(expenseSum/1000);
  }
  table.getCell(revenueNum+expenseNum+4,6).setText(thousandsDigit).setFontFamily("Times");
  var hundredsDigit = '';
  if (expenseSum >= 100) {
    hundredsDigit = Math.floor((expenseSum - 1000*thousandsDigit)/100);
  }
  table.getCell(revenueNum+expenseNum+4,7).setText(hundredsDigit).setFontFamily("Times");
  var tensDigit = '';
  if (expenseSum >= 10) {
    tensDigit = Math.floor((expenseSum-1000*thousandsDigit-100*hundredsDigit)/10);
  }
  table.getCell(revenueNum+expenseNum+4,8).setText(tensDigit).setFontFamily("Times");
  var onesDigit = '';
  if (expenseSum >= 10) {
    onesDigit = Math.floor(expenseSum-1000*thousandsDigit-100*hundredsDigit-10*tensDigit);
  }
  table.getCell(revenueNum+expenseNum+4,9).setText(onesDigit).setFontFamily("Times");
  var cents = '--';
  if (expenseSum - 1000*thousandsDigit-100*hundredsDigit-10*tensDigit - onesDigit > 0) {
    cents = Math.round(100*(expenseSum - 1000*thousandsDigit-100*hundredsDigit-10*tensDigit - onesDigit));
  }
  if (cents < 10) {
    table.getCell(revenueNum+expenseNum+4,10).setText('0' + cents).setFontFamily("Times");
  }
  else {
    table.getCell(revenueNum+expenseNum+4,10).setText(cents).setFontFamily("Times");
  }
  var thousandsDigit = 0;
  if (difference < 1000) {
    thousandsDigit = '';
  }
  else {
    thousandsDigit = Math.floor(difference/1000);
  }
  table.getCell(revenueNum+expenseNum+5,6).setText('$ ' + thousandsDigit).setFontFamily("Times");
  var hundredsDigit = '';
  if (difference >= 100) {
    hundredsDigit = Math.floor((difference - 1000*thousandsDigit)/100);
  }
  table.getCell(revenueNum+expenseNum+5,7).setText(hundredsDigit).setFontFamily("Times");
  var tensDigit = '';
  if (difference >= 10) {
    tensDigit = Math.floor((difference-1000*thousandsDigit-100*hundredsDigit)/10);
  }
  table.getCell(revenueNum+expenseNum+5,8).setText(tensDigit).setFontFamily("Times");
  var onesDigit = '';
  if (difference >= 10) {
    onesDigit = Math.floor(difference-1000*thousandsDigit-100*hundredsDigit-10*tensDigit);
  }
  table.getCell(revenueNum+expenseNum+5,9).setText(onesDigit).setFontFamily("Times");
  var cents = '--';
  if (difference - 1000*thousandsDigit-100*hundredsDigit-10*tensDigit - onesDigit > 0) {
    cents = Math.round(100*(difference - 1000*thousandsDigit-100*hundredsDigit-10*tensDigit - onesDigit));
  }
  if (cents < 10) {
    table.getCell(revenueNum+expenseNum+5,10).setText('0' + cents).setFontFamily("Times");
  }
  else {
    table.getCell(revenueNum+expenseNum+5,10).setText(cents).setFontFamily("Times");
  }
  for (i=1; i<=5; i++) {
    table.getCell(revenueNum,i).editAsText().setUnderline(true);
  }
  for (i=6; i<=10; i++) {
    table.getCell(revenueNum+1,i).editAsText().setUnderline(true);
  }
  for (i=1; i<=5; i++) {
    table.getCell(revenueNum+expenseNum+3,i).editAsText().setUnderline(true);
  }
  for (i=6; i<=10; i++) {
    table.getCell(revenueNum+expenseNum+4,i).editAsText().setUnderline(true);
  }
  for (i=6; i<=10; i++) {
    table.getCell(revenueNum+expenseNum+5,i).editAsText().setUnderline(true);
    table.getCell(revenueNum+expenseNum+5,i).editAsText().setBold(true);
  }
}
function onOpen() {
  var menuItems = [
    {name: 'Create statement', functionName: 'income'},
  ];
  ss.addMenu('Create statement', menuItems);
}