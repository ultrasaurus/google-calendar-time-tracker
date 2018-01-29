/** This Script goes with a spreadsheet where the top three rows include:

Calendar	whatever.org_k...e68@group.calendar.google.com
Start Date	12/13/2017
Date	Meeting	Duration

All meetings on the calendar will be tracked
**/

function AppLog() {
};
// note: had planned to implement multiple levels
// but limited logging doesn't seem to have a significant performance impact
AppLog.LEVELS = {
    //error: 1000,
    //warn:  900,
    //info:  800,
    debug: 0
  };

AppLog.level = "debug";

AppLog.debug = function() {
  if (typeof this.level === "undefined") this.level = "debug";
  if (this.LEVELS[this.level] <= this.LEVELS.debug) {
    //Logger.log.apply(null, arguments);  hmm, not sure why this doesn't work
    switch(arguments.length) {
      case 1:
        Logger.log(arguments[0]);
        break;
      case 2:
        Logger.log(arguments[0], arguments[1]);
        break;
    }
  }
}

AppLog.warn = function() {
  if (AppLog.LEVELS[AppLog.level] <= AppLog.LEVELS.warn) {
    Logger.prototype.log.apply(null, arguments);
  }
}


// returns reference to spreadsheet object
function activate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Hours"));
  return ss;
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name:"Calculate Hours", functionName: "calculateHours"}];
  ss.addMenu("Hours", menuEntries);
}


function meetingSummary(cal_id, start_date){
  AppLog.debug('meetingSummary -->' + cal_id + '<--');
  var hours = 0;
  var cal = CalendarApp.getCalendarById(cal_id);
  AppLog.debug(cal);
  AppLog.debug('The calendar is named "%s".', cal.getName());

  var now = new Date()
  var events = cal.getEvents(start_date, now);
  var results = []

  for ( i = events.length -1 ; i >= 0 ; i--){
    var event = events[i];
    title = event.getTitle();
    AppLog.debug(title);
    var start = event.getStartTime() ;
    var end =  event.getEndTime();
    start = new Date(start);
    end = new Date(end);
    hours = ( end - start ) / ( 1000 * 60 * 60 );
    guests = event.getGuestList();

    data = {meeting:title, duration:hours};
//    for (g in guests) {
//      var email = guests[g].getEmail();
//      var name = email.split('@')[0];
//     Logger.log(email, name);
//      data[name] = hours;
//    }
    data['date'] = start.toDateString();
    results.push(data)
  }

  AppLog.debug(results);
  return results;
}

function calculateHours(ss){
  Logger.clear();
  AppLog.debug("calculateHours");
  AppLog.debug("ss", ss);
  if (typeof(ss) === "undefined") {
    ss = activate();
  }
  var s = ss.getSheets()[0];
  // in Calendar settings, the following is the calendar address
  var cal_id = s.getRange("B2").getValue();
  var start_date = s.getRange("B3").getValue();
  //var cal_id = "bridgefoundry.org_kms7c430qq5649o2joja9iue68@group.calendar.google.com";
  var headerRange = s.getRange("A4:C4");
  var headerValues = headerRange.getValues();
  headerRange.setBackground("#000000");
  AppLog.debug(headerValues[0].length);
  s.getRange('A5:C').clearContent();
  var columns = {}
  var num_columns = headerValues[0].length;
  for (i=0; i < num_columns; i++) {
    columns[headerValues[0][i].toLowerCase()] = i+1;
  }
  AppLog.debug(columns);
  //{'meeting':2, 'duration':3, 'date':1}
  var title_column = 2;

  // from second row
  var results = meetingSummary(cal_id, start_date);
  AppLog.debug("results.length="+results.length);

  var start_row = 4;
  for ( var i = 0; i < results.length ; i++){
    row_number = start_row + i;
    var result = results[i];
    AppLog.debug("row_number="+row_number);
    AppLog.debug(result);
    for (name in columns) {
      AppLog.debug("..."+name+"   "+result[name]);
      if (typeof result[name] == 'undefined') result[name] = "";
      s.getRange(row_number+1, columns[name]).setValue(result[name]);
    }
  }
  headerRange.setBackground("#99ff99");
}