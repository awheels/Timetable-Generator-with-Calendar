//add pop up message at beginning describing flow
//error messages (duplicate sheet, etc)
//wrap text and align at top

function onOpen() {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu("Template Generator")
    .addItem('Open Sidebar', 'openSidebar')
    .addItem('Update Dates for Current Sheet', 'updateDates')
    .addItem('Update Timetable for Current Sheet', 'copyFormat')
    .addItem('Add Next Week', 'addNextWeek')
    .addItem('Create New Calendar', 'createNewCalendar')
    .addItem('Update Calendar for Current Sheet', 'updateCalendar')
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

function createNewCalendar(){
  var calendar = CalendarApp.createCalendar('SWIS Timetable');

  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  
  var sh1 = ss.getSheetByName('Timetable');
  var activeSh = ss.getActiveSheet();
  
  var c = activeSh.getRange(1,2,1,1).getValue();
  
  activeSh.getRange(16,parseInt(c)+2,1,1)
    .setValue("Current Calendar")
    .setBorder(true, true, true, true, false, false)
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setWrap(true); 
  activeSh.getRange(16,parseInt(c)+3,1,1)
    .setBorder(true, true, true, true, false, false)
    .setValue(calendar.getName())
    .setWrap(true)
    .setVerticalAlignment('center');
  activeSh.getRange(16,parseInt(c)+4,1,3).merge()
    .setBorder(true, true, true, true, false, false)
    .setValue(calendar.getId())
    .setWrap(true)
    .setVerticalAlignment('center');    
  activeSh.getRange(17,parseInt(c)+2,3,5).merge()
    .setBorder(true, true, true, true, false, false)
    .setValue("This identifies the currently linked calendar. It is important that the calendar name here matches the name of the calendar in the Google Calendar application.")
    .setVerticalAlignment('top')
    .setWrap(true);     
}

function updateCalendar(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sh1 = ss.getSheetByName('Timetable');
  var activeSh = ss.getActiveSheet();
  var c = activeSh.getRange(1,2,1,1).getValue();

  var calendarName = activeSh.getRange(16,parseInt(c)+3,1,1).getValue();
  var calendarId = activeSh.getRange(16,parseInt(c)+4,1,1).getValue();
  
  Logger.log(calendarName);
  Logger.log(calendarId);
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  
  Logger.log(calendar.getName());  
}

function createTemplate(cycles, periods, selectedDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  
  //creating timetable template
  if (ss.getSheetByName('Timetable') != null){
    var old_schedule = ss.getSheetByName('Timetable');
    var date = Utilities.formatDate(new Date(), tz, 'MM-dd-yyyy HH:MM:SS');
    old_schedule.setName('Old Timetable ' + date);
  }
  
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
      activeSheet.getRange(5*y+6,x,4,1).mergeVertically().setVerticalAlignment("top").setWrap(true);
      activeSheet.getRange(5*y+5,x,5,1).setBorder(true, true, true, true, false, false);
      activeSheet.getRange(5*y+5,x,1,1).setHorizontalAlignment("center").setFontWeight("bold");
    }
  }

  //creating instruction message for timetable
  activeSheet.getRange(2,parseInt(cycles)+2,7,2).merge()
    .setBorder(true, true, true, true, false, false)
    .setValue("Please enter your timetable/schedule to the template on the left. I suggest coloring and titling similarly to the example below. This template will be used to generate your weekly schedule.")
    .setVerticalAlignment('top')
    .setWrap(true); 
  
  //creating example formating for timetable
  activeSheet.getRange(11,parseInt(cycles)+2,4,1).merge().setWrap(true).setVerticalAlignment("top");
  activeSheet.getRange(10,parseInt(cycles)+2,5,1)
    .setBorder(true, true, true, true, false, false)
    .setBackground('#fff2cc');
  activeSheet.getRange(10,parseInt(cycles)+2,1,1)
    .setValue("Math 9")
    .setHorizontalAlignment("center")
    .setFontWeight("bold");  
  
  //set cycle and periods
  activeSheet.getRange(1,1,1,1).setValue("Cycles");
  activeSheet.getRange(1,2,1,1).setValue(cycles);
  activeSheet.getRange(1,3,1,1).setValue("Periods");
  activeSheet.getRange(1,4,1,1).setValue(periods);

  //creating first weekly plan template
  var startEndDate = getStartAndEndWeekDates(false, new Date(selectedDate));
  ss.insertSheet(startEndDate[0]+"_"+startEndDate[1]);
  var blank = ss.getSheetByName(startEndDate[0]+"_"+startEndDate[1]);

  writeDates(blank, new Date(selectedDate));
  writeDaysOfWeek(blank);
  createFirstBlankColumn(blank, periods);
  writeCycleDays(blank, cycles, 1);
  
  for (y=0; y<periods; y++){
    for (x=1; x<=5; x++){
      blank.getRange(5*y+6,x+1,4,1).mergeVertically();
      blank.getRange(5*y+5,x+1,5,1).setBorder(true, true, true, true, false, false);
      blank.setColumnWidth(x+1, 175);
    }
  }                              
}

function writeDates(sheet, initialDate){
  var rule = SpreadsheetApp.newDataValidation().requireDate().build();

  for (x=1; x<=5; x++){
    sheet.getRange(3,x+1,1,1).setBackground("#cfe2f3");
    sheet.getRange(3,x+1,1,1).setDataValidation(rule)
      .setNumberFormat("mmm d")
      .setHorizontalAlignment("center")
      .setFontWeight('bold')
      .setValue(initialDate);
    if (x != 1) {
      sheet.getRange(3,x+1,1,1).setFormulaR1C1("=IF(ISBLANK(R[0]C[-1]),,R[0]C[-1]+1)")
        .setNumberFormat("mmm d");
    }
  }
}

function writeCycleDays(sheet, numCycles, cycleDayNum){
  for (x=1; x<=5; x++){
    sheet.getRange(2,x+1,3,1).setBorder(true, true, true, true, false, true);
    if (x != 1) {
      sheet.getRange(2,x+1,1,1).setFormulaR1C1("=IF(ISBLANK(R[0]C[-1]),,IF(R[0]C[-1]="+parseInt(numCycles)+",1,R[0]C[-1]+1))")
        .setHorizontalAlignment("center")
        .setFontWeight('bold');
    }
  }
  sheet.getRange(2,2,1,1).setValue(cycleDayNum)
    .setHorizontalAlignment("center")
    .setFontWeight('bold');
}

function createFirstBlankColumn(sheet, periods){
  for (y=0; y<periods; y++){
    sheet.getRange(5+5*y,1,5,1).mergeVertically();   
    sheet.getRange(5+5*y,1,5,1).setBorder(true, true, true, true, false, false);
    sheet.getRange(2,1,3,1).setBorder(true, true, true, true, false, true);
    sheet.getRange(3,1,1,1).setBackground("#cfe2f3");
    sheet.getRange(4,1,1,1).setBackground("#38761d");
  }
}

function addNextWeek(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  
  var cyclesAndPeriodsArray = getMasterCyclesAndPeriods(ss);
  var prevSheetInfo = getSheetsEndDateAndCycles(ss, false);
  
  ss.insertSheet(prevSheetInfo[0]+"_"+prevSheetInfo[1], ss.getNumSheets());
  var blank = ss.getSheetByName(prevSheetInfo[0]+"_"+prevSheetInfo[1]);

  writeDates(blank, prevSheetInfo[0]);
  writeDaysOfWeek(blank);
  createFirstBlankColumn(blank, cyclesAndPeriodsArray[1]);
  writeCycleDays(blank, cyclesAndPeriodsArray[0], prevSheetInfo[2]);

  for (y=0; y<cyclesAndPeriodsArray[1]; y++){
    for (x=1; x<=5; x++){
      //creating columns 1-5 (M-F) of blank template
      blank.getRange(5*y+6,x+1,4,1).mergeVertically();   
      blank.getRange(5*y+5,x+1,5,1).setBorder(true, true, true, true, false, false);
      blank.setColumnWidth(x+1, 175);
    }
  }  
}

function getMasterCyclesAndPeriods(ss){
  var template = ss.getSheetByName("Timetable");
  var cycles = template.getRange(1,2,1,1).getValue();
  var periods = template.getRange(1,4,1,1).getValue();
  
  var cyclesAndPeriodsArray = [cycles, periods];
  return cyclesAndPeriodsArray;
}

function writeDaysOfWeek(sheet){
  var weekday=new Array(7);
  weekday[0]="Monday";
  weekday[1]="Tuesday";
  weekday[2]="Wednesday";
  weekday[3]="Thursday";
  weekday[4]="Friday";

  for (x=1; x<=5; x++){
    sheet.getRange(4,x+1,1,1).setBackground("#38761d")
      .setFontColor("white")
      .setHorizontalAlignment("center")
      .setValue(weekday[x-1]);
  }
}

function getSheetsEndDateAndCycles(ss, byIndex){
  if (byIndex){
    var activeSh = ss.getActiveSheet();
    var activeSheetIndex = activeSh.getIndex();
    var preSheetIndex = activeSheetIndex - 2;
    var prevSheet = ss.getSheets()[preSheetIndex];
  }else{
    var totalSheets = ss.getNumSheets();
    var prevSheet = ss.getSheets()[totalSheets-1];
  }
  var lastCycleNum = prevSheet.getRange(2,6,1,1).getValue();
  var lastMonday = prevSheet.getRange(3,2,1,1).getValue();

  var dateArray = getStartAndEndWeekDates(true, lastMonday);

  var infoArray = [dateArray[0], dateArray[1], lastCycleNum, ss];
  
  return infoArray;
}

function getStartAndEndWeekDates(prevWeek, monday){
  if (prevWeek){
    monday.setDate(monday.getDate()+7);
    var mondayDate = Utilities.formatDate(monday, Session.getScriptTimeZone(), "MMMd");
    monday.setDate(monday.getDate()+4);
    var fridayDate = Utilities.formatDate(monday, Session.getScriptTimeZone(), "MMMd");
  }else{
    var mondayDate = Utilities.formatDate(monday, Session.getScriptTimeZone(), "MMMd");
    monday.setDate(monday.getDate()+4);
    var fridayDate = Utilities.formatDate(monday, Session.getScriptTimeZone(), "MMMd");  
  }
  
  return [mondayDate, fridayDate];
}

function updateDates(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 

  var prevSheetInfo = getSheetsEndDateAndCycles(ss, true);
  var cyclesAndPeriods = getMasterCyclesAndPeriods(ss);
  
  setDatesAndCycleForSheet(cyclesAndPeriods[0], prevSheetInfo);

  //var date = range.getValues();
  //var dateFormatted = new Date(date[0][0]);
  //var newDate = new Date(dateFormatted.getTime()+3*3600000*24);
}

function copyFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  
  var sh1 = ss.getSheetByName('Timetable');
  var activeSh = ss.getActiveSheet();
  
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