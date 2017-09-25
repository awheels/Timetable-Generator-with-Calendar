function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open Sidebar', 'openSidebar')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function openSidebar() {
  var html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate();
  
  SpreadsheetApp.getUi().showSidebar(html);
}

function createTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var date = Utilities.formatDate(new Date(), tz, 'MM-dd-yyyy HH:MM:SS');
  ss.insertSheet('Template'+date, 1);
  var activeSheet = ss.getActiveSheet();
  for (x=0; x<5; x++){
    activeSheet.getRange(5*x+5,1,5,1).mergeVertically();   
    activeSheet.getRange(5*x+5,1,5,1).setBorder(true, true, true, true, false, false);
  }
}

function setDates(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var activeSh = ss.getActiveSheet();
  
  previousCell(ss, activeSh);
  previousCellDate(ss, activeSh);
}

function previousCell(spreadSheet, activeSheet) {
  var activeSheetIndex = activeSheet.getIndex();
  var preSheetIndex = activeSheetIndex - 2;
  var preSheet = spreadSheet.getSheets()[preSheetIndex];
  var range = preSheet.getRange(2,6);

  var data = range.getValues();
  data = +data + 1;

  if (data == 9) {
    activeSheet.getRange('B2').setValue(1);
  } else {
    activeSheet.getRange('B2').setValue(data);
  }
}

function previousCellDate(spreadSheet, activeSheet) {
  var activeSheetIndex = activeSheet.getIndex();
  var preSheetIndex = activeSheetIndex - 2;
  var preSheet = spreadSheet.getSheets()[preSheetIndex];
  var range = preSheet.getRange(3,6);

  var date = range.getValues();
  var dateFormatted = new Date(date[0][0]);
  var newDate = new Date(dateFormatted.getTime()+3*3600000*24);

  activeSheet.getRange('B3').setValue(newDate);
}


function copyFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  
  var sh1 = ss.getSheetByName('Template'); //Get sheet 1
  var activeSh = ss.getActiveSheet(); //Get the active sheet, you should be on the sheet just added
  
  for(var i=0; i<5; i++){
    var range = activeSh.getRange(2,2+i);
    var result = range.getValues();
    copyColumn(sh1, activeSh, result, 2+i);
  }
}


function copyColumn(template, activeSheet, dayNumber, day) {
  if (dayNumber == 1){
    template.getRange("A5:A36").copyFormatToRange(activeSheet, day, day, 5, 36); 
    template.getRange("A5:A36").copyValuesToRange(activeSheet, day, day, 5, 36); 
  } else if (dayNumber == 2){
    template.getRange("B5:B36").copyFormatToRange(activeSheet, day, day, 5, 36); 
    template.getRange("B5:B36").copyValuesToRange(activeSheet, day, day, 5, 36); 
  } else if (dayNumber == 3){
    template.getRange("C5:C36").copyFormatToRange(activeSheet, day, day, 5, 36); 
    template.getRange("C5:C36").copyValuesToRange(activeSheet, day, day, 5, 36); 
  } else if (dayNumber == 4){
    template.getRange("D5:D36").copyFormatToRange(activeSheet, day, day, 5, 36); 
    template.getRange("D5:D36").copyValuesToRange(activeSheet, day, day, 5, 36); 
  } else if (dayNumber == 5){
    template.getRange("E5:E36").copyFormatToRange(activeSheet, day, day, 5, 36); 
    template.getRange("E5:E36").copyValuesToRange(activeSheet, day, day, 5, 36); 
  } else if (dayNumber == 6){
    template.getRange("F5:F36").copyFormatToRange(activeSheet, day, day, 5, 36); 
    template.getRange("F5:F36").copyValuesToRange(activeSheet, day, day, 5, 36); 
  } else if (dayNumber == 7){
    template.getRange("G5:G36").copyFormatToRange(activeSheet, day, day, 5, 36);
    template.getRange("G5:G36").copyValuesToRange(activeSheet, day, day, 5, 36); 
  } else if (dayNumber == 8){
    template.getRange("H5:H36").copyFormatToRange(activeSheet, day, day, 5, 36);
    template.getRange("H5:H36").copyValuesToRange(activeSheet, day, day, 5, 36); 
  }
}
