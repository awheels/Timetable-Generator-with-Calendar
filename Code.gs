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

function createTemplate(cycles, periods) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  
  //***TIMETABLE TEMPLATE CREATION***//
  
  //creating timetable/schedule template
  var old_schedule = ss.getSheetByName('Timetable');
  var date = Utilities.formatDate(new Date(), tz, 'MM-dd-yyyy HH:MM:SS');
  old_schedule.setName('Old Timetable ' + date);
  ss.insertSheet('Timetable', 1);
  var activeSheet = ss.getActiveSheet();
  for (x=1; x<=cycles; x++){
      activeSheet.getRange(2,x,3,1).setBorder(true, true, true, true, false, true);
      activeSheet.getRange(3,x,1,1).setBackground("#cfe2f3");
      activeSheet.getRange(4,x,1,1).setBackground("#38761d");
      activeSheet.getRange(2,x,1,1)
        .setValue(x)
        .setHorizontalAlignment("center")
        .setFontWeight("bold");
    for (y=0; y<periods; y++){
      activeSheet.getRange(5*y+6,x,4,1).mergeVertically();   
      activeSheet.getRange(5*y+5,x,5,1).setBorder(true, true, true, true, false, false);
    }
  }

  //creating instruction message for timetable/schedule
  activeSheet.getRange(2,parseInt(cycles)+2,7,2).merge()
    .setBorder(true, true, true, true, false, false)
    .setValue("Please enter your timetable/schedule to the template on the left. I suggest coloring and titling similarly to the example below. This template will be used to generate your weekly schedule.")
    .setVerticalAlignment('top')
    .setWrap(true); 
  
  //creating example formating for schedule/timetable
  activeSheet.getRange(11,parseInt(cycles)+2,4,1).merge();
  activeSheet.getRange(10,parseInt(cycles)+2,5,1)
    .setBorder(true, true, true, true, false, false)
    .setBackground('#fff2cc');
  activeSheet.getRange(10,parseInt(cycles)+2,1,1)
    .setValue("Math 9")
    .setHorizontalAlignment("center")
    .setFontWeight("bold");  
  
  
  //***BLANK WEEKLY SCHEDULE TEMPLATE CREATION***//

  var weekday=new Array(7);
  weekday[0]="Monday";
  weekday[1]="Tuesday";
  weekday[2]="Wednesday";
  weekday[3]="Thursday";
  weekday[4]="Friday";

  var old_blank = ss.getSheetByName("Blank");
  var date2 = Utilities.formatDate(new Date(), tz, 'MM-dd-yyyy HH:MM:SS');
  old_blank.setName('Old Blank ' + date2);
  ss.insertSheet('Blank', ss.getNumSheets());
  var blank = ss.getSheetByName("Blank");

  for (y=0; y<periods; y++){
    //creating for column of blank template
    blank.getRange(5+5*y,1,5,1).mergeVertically();   
    blank.getRange(5+5*y,1,5,1).setBorder(true, true, true, true, false, false);
    blank.getRange(2,1,3,1).setBorder(true, true, true, true, false, true);
    blank.getRange(3,1,1,1).setBackground("#cfe2f3");
    blank.getRange(4,1,1,1).setBackground("#38761d");  
    for (x=1; x<=5; x++){
      //creating columns 1-5 (M-F) of blank template
      blank.getRange(5*y+6,x+1,4,1).mergeVertically();   
      blank.getRange(5*y+5,x+1,5,1).setBorder(true, true, true, true, false, false);
      blank.getRange(2,x+1,3,1).setBorder(true, true, true, true, false, true);
      blank.getRange(3,x+1,1,1).setBackground("#cfe2f3");
      blank.getRange(4,x+1,1,1).setBackground("#38761d");
      blank.getRange(4,x+1,1,1).setFontColor("white").setHorizontalAlignment("center").setValue(weekday[x-1]);
    }
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
