// ═══════════════════════════════════════════════════════════════════════════
//  COMPLETE ATTENDANCE SYSTEM — VERSION 5.0
//  Features: QR Attendance + Performance Dashboard + Half Day Detection
//           + Fast Geolocation + Individual Employee Performance
//
//  TIMING RULES:
//  Late Check-in     : After 10:20 AM
//  Early Check-out   : Before 6:30 PM
//  Half Day (CI)     : Check-in between 11:30 AM and 2:00 PM
//  Half Day (CO)     : Check-out before 5:00 PM
//
//  Updated: May 2026
// ═══════════════════════════════════════════════════════════════════════════

const SHEET_ID   = "1bQpiCMMZyMQo9WCJlOjnV3UsaoGb0XPKuLFWwiRbUg0";
const DATA_SHEET = "Employees Code List";
const ATT_SHEET  = "Attendance";
const HOL_SHEET  = "Holidays Unigen";
const OFFICE_LAT = 28.6812156;
const OFFICE_LNG = 77.017456;
const ALLOWED_RADIUS_METERS = 100;
const LATE_CHECKIN_HOUR = 10;
const LATE_CHECKIN_MIN = 20;
const EARLY_CHECKOUT_HOUR = 18;
const EARLY_CHECKOUT_MIN = 30;
const HALF_DAY_CI_START_HOUR = 11;
const HALF_DAY_CI_START_MIN = 30;
const HALF_DAY_CI_END_HOUR = 14;
const HALF_DAY_CI_END_MIN = 0;
const HALF_DAY_CO_HOUR = 17;
const HALF_DAY_CO_MIN = 0;
const LATE_BG = "#FFCCCC";
const LATE_FG = "#CC0000";
const EARLY_BG = "#FFCCCC";
const EARLY_FG = "#CC0000";
const HALF_BG = "#FFF3CD";
const HALF_FG = "#856404";


// ═══════════════════════════════════════════════════════════════════════════
//  MENU
// ═══════════════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi().createMenu("📋 Attendance Tools")
    .addItem("🗓️ Generate Master + Summary (Choose Month)", "buildMasterAndSummary")
    .addItem("📋 Generate Enhanced Summary (Half Day)", "buildEnhancedSummaryAuto")
    .addSeparator()
    .addItem("⚡ Auto-Refresh ON (Har 1 Ghanta)", "setupAutoRefresh")
    .addItem("🛑 Auto-Refresh Band Karo", "removeAutoRefresh")
    .addToUi();
}

function buildEnhancedSummaryAuto() {
  var now = new Date();
  generateEnhancedSummarySheet(now.getMonth()+1, now.getFullYear());
  SpreadsheetApp.getUi().alert("✅ Enhanced Summary Ready!");
}

function buildMasterAndSummary() {
  var ui = SpreadsheetApp.getUi();
  var mr = ui.prompt("Month?","1-12:",ui.ButtonSet.OK_CANCEL);
  if(mr.getSelectedButton()!==ui.Button.OK) return;
  var yr = ui.prompt("Year?","e.g. 2026:",ui.ButtonSet.OK_CANCEL);
  if(yr.getSelectedButton()!==ui.Button.OK) return;
  var month=parseInt(mr.getResponseText().trim()), year=parseInt(yr.getResponseText().trim());
  if(isNaN(month)||month<1||month>12||isNaN(year)){ui.alert("❌ Invalid");return;}
  generateMonthlySheet(month,year);
  generateSummarySheet(month,year);
  ui.alert("✅ Done! "+getMonthName(month)+" "+year);
}

// ═══════════════════════════════════════════════════════════════════════════
//  WEB APP ROUTER
// ═══════════════════════════════════════════════════════════════════════════
function doGet(e) {
  var page = e&&e.parameter&&e.parameter.page?e.parameter.page:"attendance";
  if(page==="api"){
    var action=e.parameter.action||"",month=parseInt(e.parameter.month)||(new Date().getMonth()+1),year=parseInt(e.parameter.year)||new Date().getFullYear();
    var result={};
    if(action==="getDashboard") result=getDashboardData(month,year);
    else if(action==="getTodayStatus") result=getTodayStatusData();
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }
  if(page==="dashboard") return HtmlService.createHtmlOutput(getDashboardHTML()).setTitle("Dashboard").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return HtmlService.createHtmlOutput(getAttendanceHTML()).setTitle("Attendance QR 2026").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ═══════════════════════════════════════════════════════════════════════════
//  SHARED HELPERS
// ═══════════════════════════════════════════════════════════════════════════
function parseTimeStr(s){s=s.toString().trim();var pm=s.toUpperCase().indexOf("PM")>=0,am=s.toUpperCase().indexOf("AM")>=0;var cl=s.replace(/(AM|PM)/gi,"").trim().split(":");var h=parseInt(cl[0])||0,m=parseInt(cl[1])||0;if(pm&&h!==12)h+=12;if(am&&h===12)h=0;return{h:h,m:m};}
function formatTimeShort(ts){if(!ts||ts.trim()==="")return"";try{var s=ts.toString().trim();if(/AM|PM/i.test(s)){var t=parseTimeStr(s);return(t.h%12||12)+":"+t.m.toString().padStart(2,"0")+" "+(t.h>=12?"PM":"AM");}var p=s.split(":");if(p.length>=2){var h=parseInt(p[0]),m=parseInt(p[1]);if(!isNaN(h)&&!isNaN(m)&&h>=0&&h<=23)return(h%12||12)+":"+m.toString().padStart(2,"0")+" "+(h>=12?"PM":"AM");}return ts;}catch(x){return ts;}}
function isLateCI(ts){if(!ts||ts.trim()==="")return false;try{var t=parseTimeStr(ts),tm=t.h*60+t.m;if(tm<360||tm>840)return false;return tm>(LATE_CHECKIN_HOUR*60+LATE_CHECKIN_MIN);}catch(x){return false;}}
function isEarlyCO(ts){if(!ts||ts.trim()===""||ts==="-")return false;try{var t=parseTimeStr(ts),tm=t.h*60+t.m;if(tm<720)return false;return tm<(EARLY_CHECKOUT_HOUR*60+EARLY_CHECKOUT_MIN);}catch(x){return false;}}
function isHalfDayCI(ts){if(!ts||ts.trim()==="")return false;try{var t=parseTimeStr(ts),tm=t.h*60+t.m;return tm>=(HALF_DAY_CI_START_HOUR*60+HALF_DAY_CI_START_MIN)&&tm<=(HALF_DAY_CI_END_HOUR*60+HALF_DAY_CI_END_MIN);}catch(x){return false;}}
function isHalfDayCO(ts){if(!ts||ts.trim()===""||ts==="-")return false;try{var t=parseTimeStr(ts),tm=t.h*60+t.m;if(tm<360)return false;return tm<(HALF_DAY_CO_HOUR*60+HALF_DAY_CO_MIN);}catch(x){return false;}}
function getMonthName(m){return["","January","February","March","April","May","June","July","August","September","October","November","December"][m]||m;}
function getHolidays(month,year){var ss=SpreadsheetApp.openById(SHEET_ID);var h=ss.getSheetByName(HOL_SHEET);if(!h)return{};var d=h.getDataRange().getValues(),hol={},tz=Session.getScriptTimeZone();for(var i=1;i<d.length;i++){var r=d[i];if(!r[0])continue;try{var ds=r[0] instanceof Date?Utilities.formatDate(r[0],tz,"dd-MM-yyyy"):r[0].toString().trim();var p=ds.split("-");if(p.length<3||parseInt(p[1])!==month||parseInt(p[2])!==year)continue;hol[ds]=r[1]?r[1].toString().trim():"Holiday";}catch(x){}}return hol;}


// ═══════════════════════════════════════════════════════════════════════════
//  EMPLOYEE PERFORMANCE API (NEW V5)
// ═══════════════════════════════════════════════════════════════════════════
function getEmployeePerformance(empCode, month, year) {
  try {
    var ss=SpreadsheetApp.openById(SHEET_ID),empSheet=ss.getSheetByName(DATA_SHEET),attSheet=ss.getSheetByName(ATT_SHEET);
    var tz=Session.getScriptTimeZone(),today=new Date();
    if(!month)month=today.getMonth()+1; if(!year)year=today.getFullYear();
    var empData=empSheet.getDataRange().getValues(),empInfo=null;
    for(var i=1;i<empData.length;i++){var r=empData[i];if(!r[2]||r[2].toString().trim()==="")continue;if(r[2].toString().trim()===empCode){empInfo={name:r[0]?r[0].toString().trim():"",designation:r[6]?r[6].toString().trim():"",dept:r[5]?r[5].toString().trim():"General",code:r[2].toString().trim()};break;}}
    if(!empInfo)return{success:false,error:"Employee not found"};
    var attData=attSheet?attSheet.getDataRange().getValues():[[]];
    var dayMap={};
    for(var j=1;j<attData.length;j++){var row=attData[j];if(!row[0])continue;var ds="";try{ds=row[0].toString().indexOf("-")>-1?row[0].toString().trim():Utilities.formatDate(new Date(row[0]),tz,"dd-MM-yyyy");}catch(x){continue;}var pts=ds.split("-");if(pts.length<3||parseInt(pts[1])!==month||parseInt(pts[2])!==year)continue;var code=row[2]?row[2].toString().trim():"";if(code!==empCode)continue;var rawIn="",rawOut="";try{if(row[5])rawIn=row[5] instanceof Date?Utilities.formatDate(row[5],tz,"hh:mm a"):formatTimeShort(row[5].toString());if(row[6])rawOut=row[6] instanceof Date?Utilities.formatDate(row[6],tz,"hh:mm a"):formatTimeShort(row[6].toString());}catch(f){}dayMap[ds]={timeIn:rawIn,timeOut:rawOut,duration:row[7]?(parseFloat(row[7].toString().replace(/[^0-9.]/g,""))||0):0};}
    var holidays=getHolidays(month,year),daysInMonth=new Date(year,month,0).getDate();
    var workingDays=[];
    for(var d=1;d<=daysInMonth;d++){var dt=new Date(year,month-1,d);if(dt.getDay()===0||dt>today)continue;var dKey=d.toString().padStart(2,"0")+"-"+month.toString().padStart(2,"0")+"-"+year;if(holidays[dKey])continue;workingDays.push(dKey);}
    var totalWD=workingDays.length,present=0,absent=0,late=0,earlyOut=0,halfDay=0,totalHrs=0;
    var lateDates=[],earlyDates=[],halfDates=[],absentDates=[],dailyLog=[];
    var currentStreak=0,maxStreak=0;
    for(var w=0;w<workingDays.length;w++){
      var dateKey=workingDays[w],rec=dayMap[dateKey]||null;
      var entry={date:dateKey,status:"absent",timeIn:"",timeOut:"",hours:0,flags:[]};
      if(!rec||rec.timeIn===""){absent++;absentDates.push(dateKey);currentStreak=0;}
      else{present++;currentStreak++;if(currentStreak>maxStreak)maxStreak=currentStreak;entry.status="present";entry.timeIn=rec.timeIn;entry.timeOut=rec.timeOut||"";entry.hours=rec.duration||0;totalHrs+=rec.duration||0;
        if(isLateCI(rec.timeIn)){late++;lateDates.push(dateKey);entry.flags.push("late");}
        if(rec.timeOut&&rec.timeOut!=="-"&&rec.timeOut!==""&&isEarlyCO(rec.timeOut)){earlyOut++;earlyDates.push(dateKey);entry.flags.push("early");}
        if(isHalfDayCI(rec.timeIn)||(rec.timeOut&&rec.timeOut!=="-"&&rec.timeOut!==""&&isHalfDayCO(rec.timeOut))){halfDay++;halfDates.push(dateKey);entry.flags.push("halfday");}
      }
      dailyLog.push(entry);
    }
    var pct=totalWD>0?Math.round((present/totalWD)*100):0;
    var avgHrs=present>0?Math.round((totalHrs/present)*10)/10:0;
    var onTimePct=present>0?Math.round(((present-late)/present)*100):0;
    return{success:true,employee:empInfo,month:month,year:year,monthName:getMonthName(month),totalWorkingDays:totalWD,present:present,absent:absent,late:late,earlyOut:earlyOut,halfDay:halfDay,totalHrs:Math.round(totalHrs*10)/10,avgHrsPerDay:avgHrs,attendancePct:pct,onTimePct:onTimePct,fullDays:present-halfDay,currentStreak:currentStreak,maxStreak:maxStreak,lateDates:lateDates,earlyDates:earlyDates,halfDates:halfDates,absentDates:absentDates,last10Days:dailyLog.slice(-10),generatedAt:Utilities.formatDate(today,tz,"dd-MM-yyyy hh:mm a")};
  }catch(err){return{success:false,error:err.message};}
}


// ═══════════════════════════════════════════════════════════════════════════
//  QR ATTENDANCE BACKEND FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════
function getEmployeeByLastTwoDigits(code2){try{if(code2===null||code2===undefined||code2.toString().trim()==="")return null;var ss=SpreadsheetApp.openById(SHEET_ID),sheet=ss.getSheetByName(DATA_SHEET);if(!sheet)return null;var data=sheet.getDataRange().getValues(),padded=code2.toString().trim().padStart(2,"0");for(var i=1;i<data.length;i++){var row=data[i];if(!row||row.length<3||!row[2]||row[2].toString().trim()==="")continue;var ec=row[2].toString().trim();if(ec.slice(-2)===padded)return{name:row[0]?row[0].toString():"",code:ec,uc:row[3]?row[3].toString():"",email:row[4]&&row[4].toString().trim()!==""?row[4].toString():""};}return null;}catch(e){return null;}}

function getAttSheet(){var ss=SpreadsheetApp.openById(SHEET_ID),sh=ss.getSheetByName(ATT_SHEET);if(!sh){sh=ss.insertSheet(ATT_SHEET);sh.appendRow(["Date","Employee Name","Employee Code","Unique Code","Email","Time In","Time Out","Duration (hrs)","Location","Device ID"]);sh.getRange(1,1,1,10).setFontWeight("bold").setBackground("#4F46E5").setFontColor("white");}return sh;}

function getTodayRow(sheet,empCode){var tz=Session.getScriptTimeZone(),today=Utilities.formatDate(new Date(),tz,"dd-MM-yyyy"),data=sheet.getDataRange().getValues();for(var i=1;i<data.length;i++){if(!data[i][0])continue;try{var rd=data[i][0].toString().indexOf("-")>-1?data[i][0].toString().trim():Utilities.formatDate(new Date(data[i][0]),tz,"dd-MM-yyyy");if(rd===today&&data[i][2].toString().trim()===empCode.trim())return{rowIndex:i+1,rowData:data[i]};}catch(x){}}return null;}

function getDeviceActiveEmployee(deviceId){var ss=SpreadsheetApp.openById(SHEET_ID),sh=ss.getSheetByName(ATT_SHEET);if(!sh)return null;var tz=Session.getScriptTimeZone(),today=Utilities.formatDate(new Date(),tz,"dd-MM-yyyy"),data=sh.getDataRange().getValues();for(var i=1;i<data.length;i++){if(!data[i][0])continue;try{var rd=data[i][0].toString().indexOf("-")>-1?data[i][0].toString().trim():Utilities.formatDate(new Date(data[i][0]),tz,"dd-MM-yyyy");var devId=data[i][9]?data[i][9].toString().trim():"",to=data[i][6]?data[i][6].toString().trim():"";if(rd===today&&devId===deviceId&&to==="")return{name:data[i][1].toString(),code:data[i][2].toString()};}catch(x){}}return null;}

function lookupEmployee(code2,deviceId){try{var ae=getDeviceActiveEmployee(deviceId);if(ae&&ae.code.slice(-2)!==code2.toString().padStart(2,"0"))return{success:false,deviceLocked:true,lockedName:ae.name,lockedCode:ae.code,msg:"Is device par <b>"+ae.name+"</b> ("+ae.code+") pehle se checked-in hai."};var emp=getEmployeeByLastTwoDigits(code2);if(!emp)return{success:false,msg:"Code nahi mila. Sahi last 2 digits dalein."};var attSheet=getAttSheet(),todayRow=getTodayRow(attSheet,emp.code);var checkedIn=!!todayRow,checkedOut=false,timeIn="",timeOut="";if(todayRow){timeIn=todayRow.rowData[5]?todayRow.rowData[5].toString():"";timeOut=todayRow.rowData[6]?todayRow.rowData[6].toString():"";checkedOut=timeOut.trim()!=="";}return{success:true,name:emp.name,code:emp.code,uc:emp.uc,email:emp.email,checkedIn:checkedIn,checkedOut:checkedOut,timeIn:timeIn,timeOut:timeOut};}catch(e){return{success:false,msg:e.message};}}


function recordAttendance(empCode,lat,lng,deviceId){
  try{
    var ae=getDeviceActiveEmployee(deviceId);if(ae&&ae.code!==empCode)return{success:false,msg:"Is device par "+ae.name+" pehle se checked-in hai."};
    var emp=getEmployeeByLastTwoDigits(empCode.slice(-2));if(!emp)return{success:false,msg:"Employee not found."};
    var attSheet=getAttSheet(),todayRow=getTodayRow(attSheet,emp.code),now=new Date(),tz=Session.getScriptTimeZone();
    var nowStr=Utilities.formatDate(now,tz,"hh:mm:ss a"),dateStr=Utilities.formatDate(now,tz,"dd-MM-yyyy"),locStr="Lat:"+lat+" Lng:"+lng;
    var nowH=parseInt(Utilities.formatDate(now,tz,"H")),nowM=parseInt(Utilities.formatDate(now,tz,"m")),totalNow=nowH*60+nowM;
    if(!todayRow){
      var isLate=totalNow>(LATE_CHECKIN_HOUR*60+LATE_CHECKIN_MIN),lateBy="";
      if(isLate){var df=totalNow-(LATE_CHECKIN_HOUR*60+LATE_CHECKIN_MIN);lateBy=(Math.floor(df/60)>0?Math.floor(df/60)+" ghante ":"")+(df%60)+" minute";}
      var isHD=totalNow>=(HALF_DAY_CI_START_HOUR*60+HALF_DAY_CI_START_MIN)&&totalNow<=(HALF_DAY_CI_END_HOUR*60+HALF_DAY_CI_END_MIN);
      attSheet.appendRow([dateStr,emp.name,emp.code,emp.uc,emp.email,nowStr,"","",locStr,deviceId]);
      return{success:true,action:"in",time:nowStr,date:dateStr,isLate:isLate,lateBy:lateBy,isHalfDay:isHD};
    }
    var to=todayRow.rowData[6]?todayRow.rowData[6].toString().trim():"";
    if(to===""){
      var isEarly=totalNow<(EARLY_CHECKOUT_HOUR*60+EARLY_CHECKOUT_MIN),earlyBy="",isHDOut=totalNow<(HALF_DAY_CO_HOUR*60+HALF_DAY_CO_MIN);
      if(isEarly){var ed=(EARLY_CHECKOUT_HOUR*60+EARLY_CHECKOUT_MIN)-totalNow;earlyBy=(Math.floor(ed/60)>0?Math.floor(ed/60)+" ghante ":"")+(ed%60)+" minute";}
      var inStr=todayRow.rowData[5].toString(),diffHrs="N/A";
      try{var toM=function(s){var pm=s.toUpperCase().indexOf("PM")>=0;var p=s.replace(/(AM|PM)/gi,"").trim().split(":");var h=parseInt(p[0]),m=parseInt(p[1]);if(pm&&h!==12)h+=12;if(!pm&&h===12)h=0;return h*60+m;};var df2=toM(nowStr)-toM(inStr);if(df2<0)df2+=1440;diffHrs=(df2/60).toFixed(2);}catch(x){}
      attSheet.getRange(todayRow.rowIndex,7).setValue(nowStr);attSheet.getRange(todayRow.rowIndex,8).setValue(diffHrs);attSheet.getRange(todayRow.rowIndex,9).setValue(locStr);attSheet.getRange(todayRow.rowIndex,10).setValue(deviceId);
      return{success:true,action:"out",time:nowStr,duration:diffHrs,timeIn:inStr,date:dateStr,isEarly:isEarly,earlyBy:earlyBy,isHalfDay:isHDOut};
    }
    return{success:false,msg:"Aaj ki attendance already complete hai."};
  }catch(e){return{success:false,msg:e.message};}
}


// ═══════════════════════════════════════════════════════════════════════════
//  AUTO REFRESH + DASHBOARD DATA
// ═══════════════════════════════════════════════════════════════════════════
function setupAutoRefresh(){var t=ScriptApp.getProjectTriggers();for(var i=0;i<t.length;i++)if(t[i].getHandlerFunction()==="autoRefreshCurrentMonth")ScriptApp.deleteTrigger(t[i]);ScriptApp.newTrigger("autoRefreshCurrentMonth").timeBased().everyHours(1).create();SpreadsheetApp.getUi().alert("✅ Auto-Refresh ON!");}
function removeAutoRefresh(){var t=ScriptApp.getProjectTriggers(),c=0;for(var i=0;i<t.length;i++)if(t[i].getHandlerFunction()==="autoRefreshCurrentMonth"){ScriptApp.deleteTrigger(t[i]);c++;}SpreadsheetApp.getUi().alert("🛑 Removed "+c+" triggers.");}
function autoRefreshCurrentMonth(){var n=new Date();generateMonthlySheet(n.getMonth()+1,n.getFullYear());generateSummarySheet(n.getMonth()+1,n.getFullYear());generateEnhancedSummarySheet(n.getMonth()+1,n.getFullYear());}

function getDashboardData(month,year){try{var ss=SpreadsheetApp.openById(SHEET_ID),empSheet=ss.getSheetByName(DATA_SHEET),attSheet=ss.getSheetByName(ATT_SHEET),tz=Session.getScriptTimeZone(),today=new Date();var empData=empSheet.getDataRange().getValues(),employees=[];for(var i=1;i<empData.length;i++){var r=empData[i];if(!r[2]||r[2].toString().trim()==="")continue;employees.push({name:r[0]?r[0].toString().trim():"",designation:r[6]?r[6].toString().trim():"",dept:r[5]?r[5].toString().trim():"General",code:r[2].toString().trim()});}var attData=attSheet?attSheet.getDataRange().getValues():[[]],attMap={};for(var j=1;j<attData.length;j++){var row=attData[j];if(!row[0])continue;var ds="";try{ds=row[0].toString().indexOf("-")>-1?row[0].toString().trim():Utilities.formatDate(new Date(row[0]),tz,"dd-MM-yyyy");}catch(x){continue;}var code=row[2]?row[2].toString().trim():"";if(!code)continue;if(!attMap[code])attMap[code]={};var rI="",rO="";try{if(row[5])rI=row[5] instanceof Date?Utilities.formatDate(row[5],tz,"hh:mm a"):formatTimeShort(row[5].toString());if(row[6])rO=row[6] instanceof Date?Utilities.formatDate(row[6],tz,"hh:mm a"):formatTimeShort(row[6].toString());}catch(f){}attMap[code][ds]={timeIn:rI,timeOut:rO,duration:row[7]?(parseFloat(row[7].toString().replace(/[^0-9.]/g,""))||0):0};}var holidays=getHolidays(month,year),dIM=new Date(year,month,0).getDate(),wd=[];for(var d=1;d<=dIM;d++){var dt=new Date(year,month-1,d);if(dt.getDay()===0||dt>today)continue;var dk=d.toString().padStart(2,"0")+"-"+month.toString().padStart(2,"0")+"-"+year;if(holidays[dk])continue;wd.push(dk);}var tStr=Utilities.formatDate(today,tz,"dd-MM-yyyy"),tP=0,tA=0,tL=0,tE=0,tHD=0,empSummary=[];employees.forEach(function(emp){var p=0,a=0,l=0,eo=0,hd=0,hrs=0;wd.forEach(function(w){var rec=attMap[emp.code]&&attMap[emp.code][w]?attMap[emp.code][w]:null;if(!rec||!rec.timeIn){a++;return;}p++;if(isLateCI(rec.timeIn))l++;if(rec.timeOut&&rec.timeOut!=="-"&&isEarlyCO(rec.timeOut))eo++;if(isHalfDayCI(rec.timeIn)||(rec.timeOut&&rec.timeOut!=="-"&&rec.timeOut!==""&&isHalfDayCO(rec.timeOut)))hd++;var dv=parseFloat(rec.duration);if(!isNaN(dv))hrs+=dv;});var pct=wd.length>0?Math.round((p/wd.length)*100):0;empSummary.push({name:emp.name,code:emp.code,dept:emp.dept,designation:emp.designation,present:p,absent:a,late:l,earlyOut:eo,halfDay:hd,hrs:Math.round(hrs*10)/10,pct:pct});var tr=attMap[emp.code]&&attMap[emp.code][tStr]?attMap[emp.code][tStr]:null;if(tr&&tr.timeIn){tP++;if(isLateCI(tr.timeIn))tL++;if(tr.timeOut&&tr.timeOut!=="-"&&isEarlyCO(tr.timeOut))tE++;if(isHalfDayCI(tr.timeIn)||(tr.timeOut&&tr.timeOut!=="-"&&tr.timeOut!==""&&isHalfDayCO(tr.timeOut)))tHD++;}else tA++;});var totalP=empSummary.reduce(function(s,e){return s+e.present;},0),totalAll=wd.length*employees.length;return{success:true,month:month,year:year,monthName:getMonthName(month),workingDays:wd.length,totalEmployees:employees.length,todayPresent:tP,todayAbsent:tA,todayLate:tL,todayEarly:tE,todayHalfDay:tHD,avgPct:totalAll>0?Math.round(totalP/totalAll*100):0,empSummary:empSummary,generatedAt:Utilities.formatDate(today,tz,"dd-MM-yyyy hh:mm a")};}catch(e){return{success:false,error:e.message};}}
function getTodayStatusData(){try{var ss=SpreadsheetApp.openById(SHEET_ID),att=ss.getSheetByName(ATT_SHEET),tz=Session.getScriptTimeZone(),today=Utilities.formatDate(new Date(),tz,"dd-MM-yyyy"),data=att?att.getDataRange().getValues():[[]],rows=[];for(var i=1;i<data.length;i++){var r=data[i];if(!r[0])continue;var ds=r[0].toString().indexOf("-")>-1?r[0].toString().trim():Utilities.formatDate(new Date(r[0]),tz,"dd-MM-yyyy");if(ds!==today)continue;rows.push({name:r[1]?r[1].toString():"",code:r[2]?r[2].toString():"",timeIn:r[5]?r[5].toString():"",timeOut:r[6]?r[6].toString():"",duration:r[7]?r[7].toString():""});}return{success:true,today:today,rows:rows};}catch(e){return{success:false,error:e.message};}}


// ═══════════════════════════════════════════════════════════════════════════
//  ENHANCED SUMMARY SHEET
// ═══════════════════════════════════════════════════════════════════════════
function generateEnhancedSummarySheet(month,year){var ss=SpreadsheetApp.openById(SHEET_ID),empSheet=ss.getSheetByName(DATA_SHEET),attSheet=ss.getSheetByName(ATT_SHEET);var sheetName="ESummary - "+getMonthName(month)+" "+year,sh=ss.getSheetByName(sheetName);if(!sh)sh=ss.insertSheet(sheetName);sh.clearContents();sh.clearFormats();sh.setFrozenRows(0);sh.setFrozenColumns(0);var tz=Session.getScriptTimeZone(),today=new Date(),holidays=getHolidays(month,year);var empData=empSheet.getDataRange().getValues(),employees=[];for(var i=1;i<empData.length;i++){var r=empData[i];if(!r[2]||r[2].toString().trim()==="")continue;employees.push({name:r[0]?r[0].toString().trim():"",designation:r[6]?r[6].toString().trim():"",dept:r[5]?r[5].toString().trim():"General",code:r[2].toString().trim()});}var attData=attSheet?attSheet.getDataRange().getValues():[[]],attMap={};for(var j=1;j<attData.length;j++){var row=attData[j];if(!row[0])continue;var ds="";try{ds=row[0].toString().indexOf("-")>-1?row[0].toString().trim():Utilities.formatDate(new Date(row[0]),tz,"dd-MM-yyyy");}catch(x){continue;}var pts=ds.split("-");if(pts.length<3||parseInt(pts[1])!==month||parseInt(pts[2])!==year)continue;var code=row[2]?row[2].toString().trim():"";if(!code)continue;if(!attMap[code])attMap[code]={};var rI="",rO="";try{if(row[5])rI=row[5] instanceof Date?Utilities.formatDate(row[5],tz,"hh:mm a"):formatTimeShort(row[5].toString());if(row[6])rO=row[6] instanceof Date?Utilities.formatDate(row[6],tz,"hh:mm a"):formatTimeShort(row[6].toString());}catch(f){}attMap[code][ds]={timeIn:rI,timeOut:rO,duration:row[7]?(parseFloat(row[7].toString().replace(/[^0-9.]/g,""))||0):0};}var dIM=new Date(year,month,0).getDate(),workingDays=[];for(var d=1;d<=dIM;d++){var dt=new Date(year,month-1,d);if(dt.getDay()===0||dt>today)continue;var dKey=d.toString().padStart(2,"0")+"-"+month.toString().padStart(2,"0")+"-"+year;if(holidays[dKey])continue;workingDays.push(dKey);}var totalWD=workingDays.length,sRow=1;
sh.getRange(sRow,1,1,13).merge().setValue("📊 ENHANCED SUMMARY — "+getMonthName(month).toUpperCase()+" "+year).setFontSize(16).setFontWeight("bold").setBackground("#1E1B4B").setFontColor("white").setHorizontalAlignment("center");sh.setRowHeight(sRow,48);sRow++;
sh.getRange(sRow,1,1,13).merge().setValue("Working Days: "+totalWD+" | Employees: "+employees.length+" | Generated: "+Utilities.formatDate(today,tz,"dd-MM-yyyy hh:mm a")).setFontSize(9).setBackground("#E0E7FF").setFontColor("#3730A3").setHorizontalAlignment("center").setWrap(true);sh.setRowHeight(sRow,30);sRow+=2;
var headers=["#","Name","Designation","Code","Present","Absent","Half Day","Late","Early Out","Hrs","Attend %","Trend","Status"];
sh.getRange(sRow,1,1,13).setValues([headers]).setFontWeight("bold").setBackground("#1E40AF").setFontColor("white").setHorizontalAlignment("center").setWrap(true);sh.setRowHeight(sRow,40);var headerRow=sRow;sRow++;
var gP=0,gA=0,gHD=0,gL=0,gE=0,gH=0;
for(var e=0;e<employees.length;e++){var emp=employees[e],present=0,absent=0,halfDay=0,late=0,earlyOut=0,hrs=0;for(var w=0;w<workingDays.length;w++){var rec=attMap[emp.code]&&attMap[emp.code][workingDays[w]]?attMap[emp.code][workingDays[w]]:null;if(!rec||rec.timeIn===""){absent++;continue;}present++;if(isHalfDayCI(rec.timeIn)||(rec.timeOut&&rec.timeOut!=="-"&&rec.timeOut!==""&&isHalfDayCO(rec.timeOut)))halfDay++;if(isLateCI(rec.timeIn))late++;if(rec.timeOut&&rec.timeOut!=="-"&&rec.timeOut!==""&&isEarlyCO(rec.timeOut))earlyOut++;var dv=parseFloat(rec.duration);if(!isNaN(dv))hrs+=dv;}gP+=present;gA+=absent;gHD+=halfDay;gL+=late;gE+=earlyOut;gH+=hrs;var pct=totalWD>0?(present/totalWD)*100:0;var trend="",l7=workingDays.slice(-7);for(var t7=0;t7<l7.length;t7++){var r7=attMap[emp.code]&&attMap[emp.code][l7[t7]]?attMap[emp.code][l7[t7]]:null;if(!r7||!r7.timeIn)trend+="✗";else if(isHalfDayCI(r7.timeIn)||(r7.timeOut&&r7.timeOut!=="-"&&isHalfDayCO(r7.timeOut)))trend+="½";else trend+="✓";}
var status=pct>=95?"✅ Excellent":pct>=75?"👍 Good":"⚠ Critical";if(late>3)status+=" | 🔴Late x"+late;if(absent>3)status+=" | 🔴Abs x"+absent;
var bg=e%2===0?"#FFFFFF":"#F0F4FF";sh.getRange(sRow,1,1,13).setValues([[e+1,emp.name,emp.designation,emp.code,present,absent,halfDay,late,earlyOut,hrs.toFixed(1),pct.toFixed(1)+"%",trend,status]]).setBackground(bg).setHorizontalAlignment("center");sh.getRange(sRow,2).setHorizontalAlignment("left").setFontWeight("bold");
var pc=sh.getRange(sRow,11);if(pct>=95)pc.setBackground("#DCFCE7").setFontColor("#166534").setFontWeight("bold");else if(pct>=75)pc.setBackground("#FEF9C3").setFontColor("#854D0E").setFontWeight("bold");else pc.setBackground("#FEE2E2").setFontColor("#991B1B").setFontWeight("bold");
if(late>3)sh.getRange(sRow,8).setBackground(LATE_BG).setFontColor(LATE_FG).setFontWeight("bold");
if(halfDay>2)sh.getRange(sRow,7).setBackground(HALF_BG).setFontColor(HALF_FG).setFontWeight("bold");
sh.setRowHeight(sRow,30);sRow++;}
var avgPct=employees.length>0&&totalWD>0?((gP/(totalWD*employees.length))*100).toFixed(1)+"%":"0%";
sh.getRange(sRow,1,1,13).setValues([["","TOTAL","","",gP,gA,gHD,gL,gE,gH.toFixed(1),avgPct,"",""]]).setFontWeight("bold").setBackground("#1E40AF").setFontColor("white").setHorizontalAlignment("center");
sh.setColumnWidth(1,38);sh.setColumnWidth(2,140);sh.setColumnWidth(3,130);sh.setColumnWidth(4,85);sh.setColumnWidth(5,65);sh.setColumnWidth(6,65);sh.setColumnWidth(7,70);sh.setColumnWidth(8,65);sh.setColumnWidth(9,75);sh.setColumnWidth(10,60);sh.setColumnWidth(11,80);sh.setColumnWidth(12,90);sh.setColumnWidth(13,180);
sh.setFrozenRows(headerRow);sh.setFrozenColumns(2);Logger.log("✅ ESummary done");}


// ═══════════════════════════════════════════════════════════════════════════
//  MASTER + SUMMARY SHEET GENERATORS
// ═══════════════════════════════════════════════════════════════════════════
function generateMonthlySheet(month,year){var ss=SpreadsheetApp.openById(SHEET_ID),empSheet=ss.getSheetByName(DATA_SHEET),attSheet=ss.getSheetByName(ATT_SHEET);var sheetName="Master - "+getMonthName(month)+" "+year,ms=ss.getSheetByName(sheetName);if(!ms)ms=ss.insertSheet(sheetName);ms.setFrozenRows(0);ms.setFrozenColumns(0);ms.clearContents();ms.clearFormats();var holidays=getHolidays(month,year),empData=empSheet.getDataRange().getValues(),employees=[];for(var i=1;i<empData.length;i++){var r=empData[i];if(!r[2]||r[2].toString().trim()==="")continue;employees.push({name:r[0]?r[0].toString().trim():"",designation:r[6]?r[6].toString().trim():"",code:r[2].toString().trim()});}var attData=attSheet?attSheet.getDataRange().getValues():[[]],attMap={},tz=Session.getScriptTimeZone();for(var j=1;j<attData.length;j++){var row=attData[j];if(!row[0])continue;var ds="";try{ds=row[0].toString().indexOf("-")>-1?row[0].toString().trim():Utilities.formatDate(new Date(row[0]),tz,"dd-MM-yyyy");}catch(x){continue;}var pts=ds.split("-");if(pts.length<3||parseInt(pts[1])!==month||parseInt(pts[2])!==year)continue;var code=row[2]?row[2].toString().trim():"";if(!code)continue;if(!attMap[code])attMap[code]={};var rI="",rO="";try{if(row[5])rI=row[5] instanceof Date?Utilities.formatDate(row[5],tz,"hh:mm a"):formatTimeShort(row[5].toString());if(row[6])rO=row[6] instanceof Date?Utilities.formatDate(row[6],tz,"hh:mm a"):formatTimeShort(row[6].toString());}catch(f){}attMap[code][ds]={timeIn:rI,timeOut:rO,duration:row[7]?row[7].toString():""};}var dIM=new Date(year,month,0).getDate(),days=[],today=new Date();for(var d=1;d<=dIM;d++)days.push({day:d,dow:new Date(year,month-1,d).getDay()});var totalCols=2+(dIM*2)+1;
ms.getRange(1,1,1,totalCols).merge().setValue("UG ATTENDANCE — "+getMonthName(month).toUpperCase()+" "+year).setFontSize(14).setFontWeight("bold").setBackground("#4F46E5").setFontColor("white").setHorizontalAlignment("center");ms.setRowHeight(1,38);
ms.getRange(2,1).setValue("Name").setFontWeight("bold").setBackground("#1E40AF").setFontColor("white").setHorizontalAlignment("center");ms.getRange(2,2).setValue("Desig").setFontWeight("bold").setBackground("#1E40AF").setFontColor("white").setHorizontalAlignment("center");
var col=3;for(var d2=0;d2<days.length;d2++){var isSun=days[d2].dow===0,dk=days[d2].day.toString().padStart(2,"0")+"-"+month.toString().padStart(2,"0")+"-"+year,isHol=!!holidays[dk];ms.getRange(2,col,1,2).merge().setValue(days[d2].day).setHorizontalAlignment("center").setFontWeight("bold").setFontColor(isSun?"#B91C1C":isHol?"#3730A3":"#1E3A8A").setBackground(isSun?"#FEF9C3":isHol?"#E0E7FF":"#DBEAFE");col+=2;}ms.getRange(2,col).setValue("P").setFontWeight("bold").setBackground("#DCFCE7").setFontColor("#166534").setHorizontalAlignment("center");ms.setRowHeight(2,28);
ms.getRange(3,1).setBackground("#374151");ms.getRange(3,2).setBackground("#374151");col=3;for(var d3=0;d3<days.length;d3++){var bg3=days[d3].dow===0?"#FEF9C3":"#EFF6FF";ms.getRange(3,col).setValue("IN").setFontSize(9).setFontWeight("bold").setBackground(bg3).setHorizontalAlignment("center");ms.getRange(3,col+1).setValue("OUT").setFontSize(9).setFontWeight("bold").setBackground(bg3).setHorizontalAlignment("center");col+=2;}ms.setRowHeight(3,22);
for(var e=0;e<employees.length;e++){var emp=employees[e],dR=4+e,pC=0,rBg=e%2===0?"#FFFFFF":"#F0F4FF";ms.getRange(dR,1).setValue(emp.name).setFontWeight("bold").setBackground(rBg).setFontSize(11);ms.getRange(dR,2).setValue(emp.designation).setBackground(rBg).setFontSize(9).setFontColor("#6B7280");col=3;for(var d4=0;d4<days.length;d4++){var dO=days[d4],isSun3=dO.dow===0,dk4=dO.day.toString().padStart(2,"0")+"-"+month.toString().padStart(2,"0")+"-"+year,isFut=new Date(year,month-1,dO.day)>today;if(isSun3){ms.getRange(dR,col,1,2).merge().setValue("SUN").setBackground("#FEF9C3").setFontColor("#92400E").setFontWeight("bold").setFontSize(8).setHorizontalAlignment("center");}else if(holidays[dk4]){ms.getRange(dR,col,1,2).merge().setValue("HOL").setBackground("#E0E7FF").setFontColor("#3730A3").setFontWeight("bold").setFontSize(8).setHorizontalAlignment("center");}else if(isFut){ms.getRange(dR,col).setBackground("#F9FAFB");ms.getRange(dR,col+1).setBackground("#F9FAFB");}else{var rec=attMap[emp.code]&&attMap[emp.code][dk4]?attMap[emp.code][dk4]:null;if(rec&&rec.timeIn!==""){pC++;var lt=isLateCI(rec.timeIn),er=rec.timeOut&&rec.timeOut!=="-"&&isEarlyCO(rec.timeOut),hd=isHalfDayCI(rec.timeIn)||(rec.timeOut&&rec.timeOut!=="-"&&rec.timeOut!==""&&isHalfDayCO(rec.timeOut));ms.getRange(dR,col).setValue(rec.timeIn).setHorizontalAlignment("center").setFontSize(9).setBackground(hd?HALF_BG:lt?LATE_BG:rBg).setFontColor(hd?HALF_FG:lt?LATE_FG:"#111827");ms.getRange(dR,col+1).setValue(rec.timeOut||"-").setHorizontalAlignment("center").setFontSize(9).setBackground(hd?HALF_BG:er?EARLY_BG:rBg).setFontColor(hd?HALF_FG:er?EARLY_FG:"#111827");}else{ms.getRange(dR,col,1,2).merge().setValue("AB").setBackground("#FCA5A5").setFontColor("#7F1D1D").setFontWeight("bold").setHorizontalAlignment("center");}}col+=2;}ms.getRange(dR,col).setValue(pC).setHorizontalAlignment("center").setFontWeight("bold").setBackground("#DCFCE7").setFontColor("#166534");ms.setRowHeight(dR,28);}
ms.setColumnWidth(1,130);ms.setColumnWidth(2,120);for(var cw=3;cw<=2+dIM*2;cw++)ms.setColumnWidth(cw,62);ms.setColumnWidth(2+dIM*2+1,40);ms.setFrozenRows(3);ms.setFrozenColumns(2);Logger.log("✅ Master done");}

function generateSummarySheet(month,year){var ss=SpreadsheetApp.openById(SHEET_ID),empSheet=ss.getSheetByName(DATA_SHEET),attSheet=ss.getSheetByName(ATT_SHEET);var sheetName="Summary - "+getMonthName(month)+" "+year,sumSheet=ss.getSheetByName(sheetName);if(!sumSheet)sumSheet=ss.insertSheet(sheetName);sumSheet.setFrozenRows(0);sumSheet.setFrozenColumns(0);sumSheet.clearContents();sumSheet.clearFormats();var empData=empSheet.getDataRange().getValues(),employees=[];for(var i=1;i<empData.length;i++){var r=empData[i];if(!r[2]||r[2].toString().trim()==="")continue;employees.push({name:r[0]?r[0].toString().trim():"",designation:r[6]?r[6].toString().trim():"",code:r[2].toString().trim()});}var attData=attSheet?attSheet.getDataRange().getValues():[[]],attMap={},tz=Session.getScriptTimeZone();for(var j=1;j<attData.length;j++){var row=attData[j];if(!row[0])continue;var ds="";try{ds=row[0].toString().indexOf("-")>-1?row[0].toString().trim():Utilities.formatDate(new Date(row[0]),tz,"dd-MM-yyyy");}catch(x){continue;}var pts=ds.split("-");if(pts.length<3||parseInt(pts[1])!==month||parseInt(pts[2])!==year)continue;var code=row[2]?row[2].toString().trim():"";if(!code)continue;if(!attMap[code])attMap[code]={};var rI="",rO="";try{if(row[5])rI=row[5] instanceof Date?Utilities.formatDate(row[5],tz,"hh:mm a"):row[5].toString();if(row[6])rO=row[6] instanceof Date?Utilities.formatDate(row[6],tz,"hh:mm a"):row[6].toString();}catch(f){}attMap[code][ds]={timeIn:rI,timeOut:rO,duration:row[7]?(parseFloat(row[7].toString().replace(/[^0-9.]/g,""))||0):0};}var dIM=new Date(year,month,0).getDate(),today=new Date(),holidays=getHolidays(month,year),wd=[];for(var d=1;d<=dIM;d++){var dt=new Date(year,month-1,d);if(dt.getDay()===0||dt>today)continue;var dk=d.toString().padStart(2,"0")+"-"+month.toString().padStart(2,"0")+"-"+year;if(holidays[dk])continue;wd.push(dk);}var totalWD=wd.length,sRow=1;
sumSheet.getRange(sRow,1,1,11).merge().setValue("📊 SUMMARY — "+getMonthName(month).toUpperCase()+" "+year).setFontSize(15).setFontWeight("bold").setBackground("#4F46E5").setFontColor("white").setHorizontalAlignment("center");sumSheet.setRowHeight(sRow,42);sRow+=2;
sumSheet.getRange(sRow,1,1,11).setValues([["#","Name","Designation","Code","Present","Absent","Half Day","Late","Early Out","Hrs","Attend %"]]).setFontWeight("bold").setBackground("#1E40AF").setFontColor("white").setHorizontalAlignment("center");var hR=sRow;sumSheet.setRowHeight(sRow,36);sRow++;
var gP=0,gA=0,gHD=0,gL=0,gE=0,gH=0;for(var e=0;e<employees.length;e++){var emp=employees[e],p=0,a=0,hd=0,l=0,eo=0,hrs=0;for(var w=0;w<wd.length;w++){var rec=attMap[emp.code]&&attMap[emp.code][wd[w]]?attMap[emp.code][wd[w]]:null;if(!rec||rec.timeIn===""){a++;continue;}p++;if(isHalfDayCI(rec.timeIn)||(rec.timeOut&&rec.timeOut!=="-"&&rec.timeOut!==""&&isHalfDayCO(rec.timeOut)))hd++;if(isLateCI(rec.timeIn))l++;if(rec.timeOut&&rec.timeOut!=="-"&&isEarlyCO(rec.timeOut))eo++;var dv=parseFloat(rec.duration);if(!isNaN(dv))hrs+=dv;}gP+=p;gA+=a;gHD+=hd;gL+=l;gE+=eo;gH+=hrs;var pct=totalWD>0?((p/totalWD)*100).toFixed(1)+"%":"0%";var bg=e%2===0?"#FFFFFF":"#F8F9FF";sumSheet.getRange(sRow,1,1,11).setValues([[e+1,emp.name,emp.designation,emp.code,p,a,hd,l,eo,hrs.toFixed(1),pct]]).setBackground(bg).setHorizontalAlignment("center");sumSheet.getRange(sRow,2).setHorizontalAlignment("left");sumSheet.setRowHeight(sRow,28);sRow++;}
sumSheet.getRange(sRow,1,1,11).setValues([["","TOTAL","","",gP,gA,gHD,gL,gE,gH.toFixed(1),employees.length>0?((gP/(totalWD*employees.length))*100).toFixed(1)+"%":"0%"]]).setFontWeight("bold").setBackground("#4F46E5").setFontColor("white").setHorizontalAlignment("center");
sumSheet.setColumnWidth(1,40);sumSheet.setColumnWidth(2,140);sumSheet.setColumnWidth(3,140);sumSheet.setColumnWidth(4,85);sumSheet.setColumnWidth(5,65);sumSheet.setColumnWidth(6,65);sumSheet.setColumnWidth(7,70);sumSheet.setColumnWidth(8,65);sumSheet.setColumnWidth(9,75);sumSheet.setColumnWidth(10,60);sumSheet.setColumnWidth(11,85);sumSheet.setFrozenRows(hR);Logger.log("✅ Summary done");}


// ═══════════════════════════════════════════════════════════════════════════
//  DASHBOARD HTML (for ?page=dashboard)
// ═══════════════════════════════════════════════════════════════════════════
function getDashboardHTML(){return '<!DOCTYPE html><html><head><title>Dashboard</title><meta name="viewport" content="width=device-width,initial-scale=1.0"></head><body style="font-family:sans-serif;padding:20px;text-align:center;"><h2>📊 Dashboard available via API</h2><p>Use main app for attendance</p></body></html>';}

// ═══════════════════════════════════════════════════════════════════════════
//  MAIN QR ATTENDANCE HTML — V5 WITH PERFORMANCE + FAST GEOLOCATION
// ═══════════════════════════════════════════════════════════════════════════
function getAttendanceHTML() {
  return '<!DOCTYPE html><html lang="hi"><head>'
+'<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">'
+'<meta name="mobile-web-app-capable" content="yes"><meta name="apple-mobile-web-app-capable" content="yes">'
+'<title>Attendance QR 2026</title>'
+'<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap" rel="stylesheet">'
+'<style>'
+'*{box-sizing:border-box;margin:0;padding:0}html{height:100%}body{margin:0;min-height:100%;font-family:"Nunito",sans-serif;background:#4F46E5}'
+'.shell{min-height:100vh;display:flex;flex-direction:column;align-items:center}'
+'.topbar{width:100%;background:#4F46E5;padding:16px 18px 12px;display:flex;align-items:center;justify-content:space-between}'
+'.topbar h1{color:#fff;font-size:18px;font-weight:900;margin:0 0 2px}.topbar p{color:rgba(255,255,255,.7);font-size:11px;margin:0;font-weight:600}'
+'.lbtn{background:rgba(255,255,255,.18);border:1.5px solid rgba(255,255,255,.3);color:#fff;padding:7px 14px;border-radius:9px;font-size:12px;font-weight:700;cursor:pointer;font-family:inherit}'
+'.body{width:100%;background:#F5F6FA;flex:1;padding:12px}@media(min-width:500px){.body{max-width:460px;margin:0 auto;border-radius:18px 18px 0 0}}'
+'.locbar{display:flex;align-items:center;gap:8px;border-radius:11px;padding:11px 14px;margin-bottom:12px;font-size:12px;font-weight:700;flex-wrap:wrap}'
+'.locbar.ok{background:#DCFCE7;border:1.5px solid #86EFAC;color:#166534}.locbar.err{background:#FEE2E2;border:1.5px solid #FECACA;color:#991B1B}.locbar.chk{background:#DBEAFE;border:1.5px solid #93C5FD;color:#1E40AF}'
+'.locdot{width:9px;height:9px;border-radius:50%}.locdot.ok{background:#16A34A}.locdot.err{background:#DC2626}.locdot.chk{background:#3B82F6;animation:pulse 1s infinite}'
+'@keyframes pulse{0%,100%{opacity:1}50%{opacity:.2}}'
+'.retrybtn{background:#1E40AF;color:#fff;border:none;border-radius:7px;padding:4px 10px;font-size:11px;font-weight:700;cursor:pointer;margin-left:6px;font-family:inherit}'
+'.fg{margin-bottom:10px}.lbl{font-size:10px;font-weight:800;color:#6B7280;text-transform:uppercase;letter-spacing:.6px;margin-bottom:4px}'
+'input[type=number]{width:100%;padding:13px;border:2px solid #E5E7EB;border-radius:11px;font-size:28px;font-weight:900;letter-spacing:10px;outline:none;color:#111;background:#fff;-webkit-appearance:none;text-align:center;font-family:inherit}'
+'input:focus{border-color:#4F46E5;box-shadow:0 0 0 3px rgba(79,70,229,.12)}input::-webkit-inner-spin-button,input::-webkit-outer-spin-button{-webkit-appearance:none}'
+'.hint{font-size:11px;color:#9CA3AF;text-align:center;margin-top:4px;font-weight:600}'
+'.fv{background:#fff;border:2px solid #E5E7EB;border-radius:11px;padding:11px 14px;font-size:14px;font-weight:700;color:#111}'
+'.btn{width:100%;border:none;border-radius:13px;padding:15px;font-size:15px;font-weight:800;cursor:pointer;margin-top:7px;font-family:inherit;transition:transform .1s}.btn:active{transform:scale(.98)}'
+'.btn.primary{background:#4F46E5;color:#fff}.btn.out{background:#D97706;color:#fff}.btn.done{background:#16A34A;color:#fff}.btn:disabled{background:#D1D5DB;color:#9CA3AF;cursor:not-allowed}'
+'.msg{border-radius:11px;padding:12px;font-size:12px;margin-top:9px;text-align:center;font-weight:700}.msg.ok{background:#DCFCE7;color:#166534;border:1.5px solid #86EFAC}.msg.err{background:#FEE2E2;color:#991B1B;border:1.5px solid #FECACA}'
+'.spin{display:inline-block;width:14px;height:14px;border:2.5px solid rgba(255,255,255,.4);border-top-color:#fff;border-radius:50%;animation:sp .7s linear infinite;vertical-align:middle;margin-right:6px}@keyframes sp{to{transform:rotate(360deg)}}'
+'.tgrid{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:11px}'
+'.tc{background:#fff;border:2px solid #E5E7EB;border-radius:11px;padding:11px 13px}.tl{font-size:9px;font-weight:800;color:#9CA3AF;text-transform:uppercase;margin-bottom:4px}.tv{font-size:12px;font-weight:800;color:#111}'
+'.badge{display:inline-block;padding:2px 9px;border-radius:18px;font-size:10px;font-weight:800;margin-top:5px}.b-in{background:#DCFCE7;color:#166534}.b-out{background:#FEF9C3;color:#854D0E}.b-no{background:#F3F4F6;color:#9CA3AF}'
+'hr{border:none;border-top:2px solid #E5E7EB;margin:11px 0}'
+'.lockbanner{background:#FEF3C7;border:2px solid #FCD34D;border-radius:11px;padding:11px 13px;margin-bottom:11px;font-size:12px;color:#92400E;font-weight:700;line-height:1.5}'
+'</style>'
+'<style>'
+'/* Performance Screen */'
+'.perf-screen{width:100%;display:none}'
+'.ph{background:linear-gradient(135deg,#0F172A,#1E293B);padding:16px 18px 12px}'
+'.ph h1{color:#fff;font-size:17px;font-weight:900}.ph p{color:rgba(255,255,255,.6);font-size:11px;font-weight:600}'
+'.pb{background:#F5F6FA;padding:12px}@media(min-width:500px){.pb{max-width:460px;margin:0 auto}}'
+'.hero{background:linear-gradient(135deg,#1E1B4B,#4F46E5);border-radius:16px;padding:16px;margin-bottom:10px;display:flex;align-items:center;gap:12px}'
+'.hero-av{width:48px;height:48px;border-radius:50%;background:rgba(255,255,255,.15);display:flex;align-items:center;justify-content:center;font-size:20px;font-weight:900;color:#fff;border:2px solid rgba(255,255,255,.25)}'
+'.hero-info h2{color:#fff;font-size:15px;font-weight:900;margin:0 0 2px}.hero-info p{color:rgba(255,255,255,.6);font-size:10px;font-weight:600;margin:0}'
+'.hero-badge{background:rgba(255,255,255,.15);border:1.5px solid rgba(255,255,255,.2);border-radius:7px;padding:2px 9px;font-size:10px;font-weight:800;color:#fff;margin-top:4px;display:inline-block}'
+'.pstats{display:grid;grid-template-columns:1fr 1fr 1fr;gap:7px;margin-bottom:10px}'
+'.ps{background:#fff;border-radius:11px;padding:10px 8px;text-align:center;box-shadow:0 1px 2px rgba(0,0,0,.05)}'
+'.ps-icon{font-size:18px;margin-bottom:3px}.ps-val{font-size:24px;font-weight:900;line-height:1}.ps-lbl{font-size:8px;font-weight:800;color:#6B7280;text-transform:uppercase;margin-top:2px}'
+'.ps-green .ps-val{color:#15803D}.ps-red .ps-val{color:#B91C1C}.ps-orange .ps-val{color:#B45309}.ps-purple .ps-val{color:#6D28D9}.ps-yellow .ps-val{color:#92400E}.ps-blue .ps-val{color:#0369A1}'
+'.pct-card{background:#fff;border-radius:13px;padding:13px 15px;margin-bottom:10px;box-shadow:0 1px 2px rgba(0,0,0,.05)}'
+'.pct-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:7px}.pct-lbl{font-size:11px;font-weight:800;color:#374151}.pct-num{font-size:26px;font-weight:900}'
+'.pct-high{color:#15803D}.pct-med{color:#B45309}.pct-low{color:#B91C1B}'
+'.bar-track{width:100%;height:9px;background:#F3F4F6;border-radius:5px;overflow:hidden}.bar-fill{height:100%;border-radius:5px;transition:width .8s ease}'
+'.bar-high{background:linear-gradient(90deg,#16A34A,#4ADE80)}.bar-med{background:linear-gradient(90deg,#D97706,#FCD34D)}.bar-low{background:linear-gradient(90deg,#DC2626,#FCA5A5)}'
+'.pct-sub{font-size:10px;color:#9CA3AF;margin-top:5px;font-weight:600}'
+'.info-row{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-bottom:10px}'
+'.pinfo{background:#fff;border-radius:11px;padding:10px 13px;box-shadow:0 1px 2px rgba(0,0,0,.05)}'
+'.pinfo-lbl{font-size:9px;font-weight:800;color:#9CA3AF;text-transform:uppercase;margin-bottom:3px}.pinfo-val{font-size:17px;font-weight:900;color:#111}.pinfo-sub{font-size:9px;color:#9CA3AF;font-weight:600;margin-top:2px}'
+'.trend-card{background:#fff;border-radius:13px;padding:12px 14px;margin-bottom:10px;box-shadow:0 1px 2px rgba(0,0,0,.05)}'
+'.trend-lbl{font-size:9px;font-weight:800;color:#6B7280;text-transform:uppercase;margin-bottom:7px}'
+'.tdots{display:flex;gap:4px;flex-wrap:wrap}'
+'.td{width:28px;height:28px;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:900}'
+'.td.tp{background:#DCFCE7;color:#16A34A}.td.ta{background:#FEE2E2;color:#DC2626}.td.th{background:#FEF9C3;color:#D97706}.td.tl{background:#FFCCCC;color:#CC0000}'
+'.td-legend{display:flex;gap:8px;margin-top:6px;flex-wrap:wrap}.td-legend span{font-size:9px;font-weight:700;color:#6B7280;display:flex;align-items:center;gap:3px}.dsm{width:8px;height:8px;border-radius:2px;display:inline-block}'
+'.det-sec{background:#fff;border-radius:12px;margin-bottom:8px;box-shadow:0 1px 2px rgba(0,0,0,.05);overflow:hidden}'
+'.det-hdr{padding:10px 14px;display:flex;align-items:center;justify-content:space-between;cursor:pointer}.det-hdr:active{background:#F9FAFB}'
+'.dhl{display:flex;align-items:center;gap:7px}.dh-icon{font-size:16px}.dh-title{font-size:11px;font-weight:800;color:#374151;text-transform:uppercase}'
+'.dh-cnt{background:#EEF2FF;color:#4338CA;border-radius:10px;padding:2px 8px;font-size:10px;font-weight:800}.dh-arrow{font-size:12px;color:#9CA3AF;transition:transform .2s}'
+'.det-body{padding:0 14px 10px;display:none}.det-body.open{display:block}'
+'.dtag{display:inline-block;background:#F3F4F6;border-radius:5px;padding:2px 8px;font-size:10px;font-weight:700;color:#374151;margin:2px 3px 2px 0}'
+'.dtag.red{background:#FEE2E2;color:#991B1B}.dtag.orange{background:#FEF3C7;color:#92400E}.dtag.purple{background:#EDE9FE;color:#6D28D9}'
+'.st-card{border-radius:13px;padding:12px 14px;margin-bottom:10px}'
+'.st-card.sg{background:#DCFCE7;border:2px solid #86EFAC}.st-card.sy{background:#FEF9C3;border:2px solid #FCD34D}.st-card.sr{background:#FEE2E2;border:2px solid #FECACA}'
+'.st-title{font-size:14px;font-weight:900;margin-bottom:3px}.sg .st-title{color:#15803D}.sy .st-title{color:#854D0E}.sr .st-title{color:#991B1B}'
+'.st-msg{font-size:11px;font-weight:700;line-height:1.5}.sg .st-msg{color:#166534}.sy .st-msg{color:#92400E}.sr .st-msg{color:#B91C1B}'
+'.skel{background:linear-gradient(90deg,#F3F4F6 25%,#E5E7EB 50%,#F3F4F6 75%);background-size:200% 100%;animation:shm 1.4s infinite;border-radius:10px;height:16px;margin-bottom:7px}@keyframes shm{0%{background-position:200% 0}100%{background-position:-200% 0}}'
+'</style>'
+'<style>'
+'/* Success + Modal */'
+'.swrap{text-align:center;padding:24px 14px 18px}.sicon{font-size:64px;margin-bottom:8px}.swrap h2{font-size:20px;font-weight:900;color:#111;margin-bottom:6px}.swrap p{font-size:13px;color:#6B7280;font-weight:600}'
+'.sdet{background:#fff;border:2px solid #E5E7EB;border-radius:13px;padding:14px;margin-top:12px;font-size:13px;color:#374151;line-height:2;text-align:left;font-weight:600}.sdet strong{color:#111;font-weight:800}'
+'.alert-late,.alert-early{background:#FFCCCC;border:2px solid #CC0000;border-radius:12px;padding:12px 14px;margin-top:10px;display:flex;gap:9px;align-items:flex-start}'
+'.alert-half{background:#FFF3CD;border:2px solid #D97706;border-radius:12px;padding:12px 14px;margin-top:10px;display:flex;gap:9px;align-items:flex-start}'
+'.al-icon{font-size:22px;flex-shrink:0}.al-title{font-size:13px;font-weight:900;margin-bottom:2px}.alert-late .al-title,.alert-early .al-title{color:#CC0000}.alert-half .al-title{color:#92400E}'
+'.al-sub{font-size:11px;font-weight:700;line-height:1.5}.alert-late .al-sub,.alert-early .al-sub{color:#990000}.alert-half .al-sub{color:#78350F}'
+'.modal-bg{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:999;display:flex;align-items:center;justify-content:center;padding:16px}'
+'.modal{background:#fff;border-radius:18px;padding:24px 18px 18px;max-width:360px;width:100%;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.3)}'
+'.m-icon{font-size:56px;margin-bottom:8px}.modal h3{font-size:18px;font-weight:900;margin:0 0 6px}.modal-late h3,.modal-early h3{color:#CC0000}.modal-half h3{color:#92400E}'
+'.modal p{font-size:12px;color:#6B7280;line-height:1.6;margin:0 0 14px;font-weight:600}'
+'.m-badge{display:inline-block;padding:6px 16px;border-radius:28px;font-size:14px;font-weight:900;margin-bottom:12px}'
+'.modal-late .m-badge,.modal-early .m-badge{background:#FFCCCC;color:#CC0000;border:2px solid #CC0000}.modal-half .m-badge{background:#FFF3CD;color:#92400E;border:2px solid #D97706}'
+'.m-ok{width:100%;border:none;border-radius:11px;padding:12px;font-size:14px;font-weight:800;cursor:pointer;font-family:inherit}.modal-late .m-ok,.modal-early .m-ok{background:#CC0000;color:#fff}.modal-half .m-ok{background:#D97706;color:#fff}'
+'</style>'
+'</head><body><div class="shell">'


/* VIEW 1: LOGIN */
+'<div id="vLogin" style="width:100%">'
+'<div class="topbar"><div><h1>Attendance Register</h1><p>QR 2026 — Apna Code dalein</p></div></div>'
+'<div class="body">'
+'<div id="locBar" class="locbar chk"><div class="locdot chk" id="locDot"></div><span id="locTxt">📍 Location check ho rahi hai...</span></div>'
+'<div id="lockBanner" style="display:none"></div>'
+'<div class="fg"><div class="lbl">🔢 Employee Code (Last 2 Digits)</div>'
+'<input type="number" id="codeIn" min="1" max="99" inputmode="numeric" oninput="if(this.value.length>2)this.value=this.value.slice(0,2)"/>'
+'<div class="hint">Example: UG0008 hai to sirf <strong>08</strong> dalein</div></div>'
+'<button class="btn primary" id="loginBtn" onclick="doLookup()" disabled>📊 Performance & Attendance →</button>'
+'<div id="loginMsg"></div>'
+'</div></div>'

/* VIEW 2: PERFORMANCE */
+'<div id="vPerf" class="perf-screen">'
+'<div class="ph"><h1 id="pName">Loading...</h1><p id="pSub">📊 Monthly Performance</p></div>'
+'<div class="pb">'
+'<div id="pLoad"><div class="skel" style="height:60px;border-radius:13px;margin-bottom:8px"></div><div class="skel" style="height:45px;border-radius:11px;margin-bottom:8px"></div><div class="skel" style="height:35px;border-radius:11px"></div></div>'
+'<div id="pContent" style="display:none">'
+'<div class="hero"><div class="hero-av" id="hAv">👤</div><div class="hero-info"><h2 id="hName"></h2><p id="hDesig"></p><div class="hero-badge" id="hBadge"></div></div></div>'
+'<div class="pstats">'
+'<div class="ps ps-green"><div class="ps-icon">✅</div><div class="ps-val" id="s1">—</div><div class="ps-lbl">Present</div></div>'
+'<div class="ps ps-red"><div class="ps-icon">❌</div><div class="ps-val" id="s2">—</div><div class="ps-lbl">Absent</div></div>'
+'<div class="ps ps-orange"><div class="ps-icon">⏰</div><div class="ps-val" id="s3">—</div><div class="ps-lbl">Late</div></div>'
+'<div class="ps ps-purple"><div class="ps-icon">🏃</div><div class="ps-val" id="s4">—</div><div class="ps-lbl">Early Out</div></div>'
+'<div class="ps ps-yellow"><div class="ps-icon">🟡</div><div class="ps-val" id="s5">—</div><div class="ps-lbl">Half Day</div></div>'
+'<div class="ps ps-blue"><div class="ps-icon">⏱</div><div class="ps-val" id="s6">—</div><div class="ps-lbl">Hours</div></div>'
+'</div>'
+'<div class="pct-card"><div class="pct-top"><div class="pct-lbl">📊 Attendance %</div><div class="pct-num" id="pctN">—</div></div><div class="bar-track"><div class="bar-fill" id="pctB" style="width:0%"></div></div><div class="pct-sub" id="pctS"></div></div>'
+'<div class="info-row"><div class="pinfo"><div class="pinfo-lbl">🔥 Streak</div><div class="pinfo-val" id="strV">—</div><div class="pinfo-sub" id="strS">days</div></div><div class="pinfo"><div class="pinfo-lbl">⏱ Avg Hrs/Day</div><div class="pinfo-val" id="avgV">—</div><div class="pinfo-sub" id="avgS">per day</div></div></div>'
+'<div class="trend-card"><div class="trend-lbl">📅 Last 10 Working Days</div><div class="tdots" id="tDots"></div><div class="td-legend"><span><div class="dsm" style="background:#DCFCE7"></div>Present</span><span><div class="dsm" style="background:#FEE2E2"></div>Absent</span><span><div class="dsm" style="background:#FEF9C3"></div>Half</span><span><div class="dsm" style="background:#FFCCCC"></div>Late</span></div></div>'
+'<div id="dLate" class="det-sec" style="display:none"><div class="det-hdr" onclick="togDet(this)"><div class="dhl"><span class="dh-icon">⏰</span><span class="dh-title">Late Dates</span></div><span class="dh-cnt" id="lcnt">0</span><span class="dh-arrow">▼</span></div><div class="det-body" id="lBody"></div></div>'
+'<div id="dEarly" class="det-sec" style="display:none"><div class="det-hdr" onclick="togDet(this)"><div class="dhl"><span class="dh-icon">🏃</span><span class="dh-title">Early Out Dates</span></div><span class="dh-cnt" id="ecnt">0</span><span class="dh-arrow">▼</span></div><div class="det-body" id="eBody"></div></div>'
+'<div id="dHalf" class="det-sec" style="display:none"><div class="det-hdr" onclick="togDet(this)"><div class="dhl"><span class="dh-icon">🟡</span><span class="dh-title">Half Day Dates</span></div><span class="dh-cnt" id="hcnt">0</span><span class="dh-arrow">▼</span></div><div class="det-body" id="hBody"></div></div>'
+'<div id="dAbsent" class="det-sec" style="display:none"><div class="det-hdr" onclick="togDet(this)"><div class="dhl"><span class="dh-icon">❌</span><span class="dh-title">Absent Dates</span></div><span class="dh-cnt" id="acnt">0</span><span class="dh-arrow">▼</span></div><div class="det-body" id="aBody"></div></div>'
+'<div class="st-card" id="stCard"><div class="st-title" id="stT"></div><div class="st-msg" id="stM"></div></div>'
+'</div>'
+'<button class="btn primary" onclick="goAtt()" style="margin-top:4px">✅ Attendance Lagao →</button>'
+'<button class="btn" style="background:#fff;color:#6B7280;border:2px solid #E5E7EB;margin-top:8px" onclick="doLogout()">← Back</button>'
+'</div></div>'


/* VIEW 3: ATTENDANCE FORM */
+'<div id="vForm" style="width:100%;display:none">'
+'<div class="topbar"><div><h1>Attendance</h1><p id="hdrC"></p></div><button class="lbtn" onclick="doLogout()">Logout</button></div>'
+'<div class="body">'
+'<div class="fg"><div class="lbl">👤 Name</div><div class="fv" id="fN"></div></div>'
+'<div class="fg"><div class="lbl">💾 Code</div><div class="fv" id="fC"></div></div>'
+'<div class="fg"><div class="lbl">🔑 UC</div><div class="fv" id="fU"></div></div>'
+'<div class="fg"><div class="lbl">📧 Email</div><div class="fv" id="fE"></div></div>'
+'<hr>'
+'<div class="tgrid">'
+'<div class="tc"><div class="tl">🕐 Time In</div><div class="tv" id="tIn">—</div><div id="bIn"><span class="badge b-no">Pending</span></div></div>'
+'<div class="tc"><div class="tl">🕑 Time Out</div><div class="tv" id="tOut">—</div><div id="bOut"><span class="badge b-no">Pending</span></div></div>'
+'</div>'
+'<button class="btn primary" id="attBtn" onclick="doAtt()">Check In →</button>'
+'<div id="attMsg"></div>'
+'</div></div>'

/* VIEW 4: SUCCESS */
+'<div id="vSuccess" style="width:100%;display:none">'
+'<div class="topbar"><div><h1>Attendance</h1><p>QR 2026</p></div></div>'
+'<div class="body"><div class="swrap">'
+'<div class="sicon" id="sIco"></div><h2 id="sT"></h2><p id="sM"></p>'
+'<div class="sdet" id="sD"></div><div id="sAl"></div>'
+'<button class="btn primary" style="margin-top:16px" onclick="doLogout()">Done / Logout</button>'
+'</div></div></div>'

/* MODAL */
+'<div id="modBg" class="modal-bg" style="display:none" onclick="closeMod(event)">'
+'<div class="modal" id="modBox" onclick="event.stopPropagation()">'
+'<div class="m-icon" id="mI"></div><h3 id="mT"></h3><div class="m-badge" id="mB"></div><p id="mP"></p>'
+'<button class="m-ok" id="mO" onclick="closeMod()">OK ✔</button>'
+'</div></div>'

+'</div>'


/* ═══════ JAVASCRIPT ═══════ */
+'<script>'
+'var emp={},uLat=null,uLng=null,locOk=false,OLat=28.6812156,OLng=77.017456,MaxD=100;'

/* Device/Lock */
+'function gDI(){var d=localStorage.getItem("attDI");if(!d){d="D-"+Math.random().toString(36).substr(2,8).toUpperCase()+Date.now().toString(36);localStorage.setItem("attDI",d);}return d;}'
+'function gLk(){return{code:localStorage.getItem("lkC"),name:localStorage.getItem("lkN")};}'
+'function sLk(c,n){localStorage.setItem("lkC",c);localStorage.setItem("lkN",n);}'
+'function cLk(){localStorage.removeItem("lkC");localStorage.removeItem("lkN");}'
+'function showLB(){var lk=gLk(),el=document.getElementById("lockBanner");if(lk.code){el.style.display="block";el.className="lockbanner";el.innerHTML="🔒 <strong>"+lk.name+"</strong> ("+lk.code+") checked-in hai. Pehle Check-Out karein.";}else el.style.display="none";}'

/* Geo */
+'function dist(a,b,c,d){var R=6371000,dL=(c-a)*Math.PI/180,dl=(d-b)*Math.PI/180,x=Math.sin(dL/2)*Math.sin(dL/2)+Math.cos(a*Math.PI/180)*Math.cos(c*Math.PI/180)*Math.sin(dl/2)*Math.sin(dl/2);return R*2*Math.atan2(Math.sqrt(x),Math.sqrt(1-x));}'
+'function sLoc(st,msg){var b=document.getElementById("locBar"),d=document.getElementById("locDot"),t=document.getElementById("locTxt");b.className="locbar "+st;d.className="locdot "+st;t.innerHTML=msg;}'

/* FAST GEOLOCATION — dual strategy with auto-retry */
+'var geoOK=false,geoTry=0;'
+'function chkLoc(){geoOK=false;geoTry=0;sLoc("chk","📍 Detecting...");document.getElementById("loginBtn").disabled=true;if(!navigator.geolocation){sLoc("err","❌ No GPS");return;}tryFast();}'
+'function tryFast(){navigator.geolocation.getCurrentPosition(onGeo,function(){tryHigh();},{enableHighAccuracy:false,timeout:3000,maximumAge:30000});}'
+'function tryHigh(){navigator.geolocation.getCurrentPosition(onGeo,onFail,{enableHighAccuracy:true,timeout:5000,maximumAge:10000});}'
+'function onGeo(p){if(geoOK)return;geoOK=true;uLat=p.coords.latitude;uLng=p.coords.longitude;var dm=Math.round(dist(uLat,uLng,OLat,OLng));if(dm<=MaxD){locOk=true;sLoc("ok","✅ Office ("+dm+"m)");document.getElementById("loginBtn").disabled=false;navigator.geolocation.getCurrentPosition(function(p2){uLat=p2.coords.latitude;uLng=p2.coords.longitude;var d2=Math.round(dist(uLat,uLng,OLat,OLng));if(d2<=MaxD)sLoc("ok","✅ Office ("+d2+"m precise)");else{locOk=false;document.getElementById("loginBtn").disabled=true;sLoc("err","⚠ Precise: "+d2+"m door");}},function(){},{enableHighAccuracy:true,timeout:7000,maximumAge:0});}else{locOk=false;sLoc("err","⚠ Office se "+dm+"m door!");addRB();}}'
+'function onFail(e){if(geoOK)return;geoTry++;if(geoTry<3){sLoc("chk","📍 Retry #"+geoTry+"...");setTimeout(tryHigh,500);return;}geoOK=true;sLoc("err",e.code===1?"⚠ Location blocked! Allow karein":"⚠ Location error");addRB();}'
+'function addRB(){var lb=document.getElementById("locBar");if(!lb.querySelector(".retrybtn"))lb.innerHTML+=\' <button class="retrybtn" onclick="chkLoc()">🔄 Retry</button>\';}'


/* Helpers */
+'function showMsg(id,t,c){document.getElementById(id).innerHTML=\'<div class="msg \'+c+\'">\'+t+"</div>";}'
+'function show(id){["vLogin","vPerf","vForm","vSuccess"].forEach(function(v){document.getElementById(v).style.display=v===id?"block":"none";});}'
+'function openMod(type,badge,title,msg){var box=document.getElementById("modBox");box.className="modal modal-"+type;document.getElementById("mI").textContent={late:"⏰",early:"🏃",half:"🟡"}[type]||"⚠";document.getElementById("mT").textContent=title;document.getElementById("mB").textContent=badge;document.getElementById("mP").textContent=msg;document.getElementById("modBg").style.display="flex";}'
+'function closeMod(ev){if(ev&&ev.target!==document.getElementById("modBg"))return;document.getElementById("modBg").style.display="none";}'
+'function togDet(el){var b=el.parentElement.querySelector(".det-body"),a=el.querySelector(".dh-arrow");if(b.classList.contains("open")){b.classList.remove("open");a.style.transform="";}else{b.classList.add("open");a.style.transform="rotate(180deg)";}}'

/* LOOKUP */
+'function doLookup(){'
+'if(!locOk){showMsg("loginMsg","⚠ Office pe aao","err");return;}'
+'var code=document.getElementById("codeIn").value.trim();if(!code){showMsg("loginMsg","Code dalo","err");return;}'
+'var lk=gLk();if(lk.code&&lk.code.slice(-2)!==code.padStart(2,"0")){showMsg("loginMsg","⚠ "+lk.name+" pehle se checked-in","err");return;}'
+'var btn=document.getElementById("loginBtn");btn.innerHTML=\'<span class="spin"></span>Loading...\';btn.disabled=true;'
+'google.script.run.withSuccessHandler(function(r){btn.innerHTML="📊 Performance & Attendance →";btn.disabled=false;if(!r.success){if(r.deviceLocked){sLk(r.lockedCode,r.lockedName);showLB();}showMsg("loginMsg",r.msg,"err");return;}emp=r;showPerf(r);}).withFailureHandler(function(e){btn.innerHTML="📊 Performance & Attendance →";btn.disabled=false;showMsg("loginMsg",e.message,"err");}).lookupEmployee(code,gDI());}'

/* PERFORMANCE SCREEN */
+'function showPerf(ed){'
+'show("vPerf");document.getElementById("pName").textContent=ed.name;document.getElementById("pSub").textContent=ed.code+" — Performance";'
+'document.getElementById("pLoad").style.display="block";document.getElementById("pContent").style.display="none";'
+'google.script.run.withSuccessHandler(function(d){document.getElementById("pLoad").style.display="none";document.getElementById("pContent").style.display="block";if(!d||!d.success){perfFB(ed);return;}renderPerf(d,ed);}).withFailureHandler(function(){document.getElementById("pLoad").style.display="none";document.getElementById("pContent").style.display="block";perfFB(ed);}).getEmployeePerformance(ed.code,null,null);}'

+'function perfFB(ed){document.getElementById("hAv").textContent=ed.name.charAt(0);document.getElementById("hName").textContent=ed.name;document.getElementById("hDesig").textContent=ed.code;document.getElementById("hBadge").textContent="Data unavailable";document.getElementById("stCard").className="st-card sy";document.getElementById("stT").textContent="⚠ Data nahi mila";document.getElementById("stM").textContent="Aap attendance laga sakte hain.";}'


/* RENDER PERFORMANCE DATA */
+'function renderPerf(d,ed){'
+'var ini=ed.name.split(" ").map(function(w){return w[0]||"";}).join("").substr(0,2).toUpperCase();'
+'document.getElementById("hAv").textContent=ini||"👤";'
+'document.getElementById("hName").textContent=d.employee.name;'
+'document.getElementById("hDesig").textContent=d.employee.code+(d.employee.designation?" | "+d.employee.designation:"");'
+'document.getElementById("hBadge").textContent=d.attendancePct+"% | "+d.present+"/"+d.totalWorkingDays+" Days | "+d.monthName;'
+'document.getElementById("s1").textContent=d.present;document.getElementById("s2").textContent=d.absent;document.getElementById("s3").textContent=d.late;document.getElementById("s4").textContent=d.earlyOut;document.getElementById("s5").textContent=d.halfDay;document.getElementById("s6").textContent=d.totalHrs+"h";'
/* % bar */
+'var pct=d.attendancePct,pe=document.getElementById("pctN");pe.textContent=pct+"%";pe.className="pct-num "+(pct>=90?"pct-high":pct>=75?"pct-med":"pct-low");'
+'var bar=document.getElementById("pctB");bar.className="bar-fill "+(pct>=90?"bar-high":pct>=75?"bar-med":"bar-low");setTimeout(function(){bar.style.width=Math.min(pct,100)+"%";},100);'
+'document.getElementById("pctS").textContent=d.present+" present / "+d.totalWorkingDays+" working days | On-time: "+d.onTimePct+"%";'
/* Streak + Avg */
+'document.getElementById("strV").textContent=d.currentStreak;document.getElementById("strS").textContent="Max: "+d.maxStreak+" days";'
+'document.getElementById("avgV").textContent=d.avgHrsPerDay+"h";document.getElementById("avgS").textContent="Total: "+d.totalHrs+"h";'
/* Trend dots */
+'var tE=document.getElementById("tDots");tE.innerHTML="";'
+'if(d.last10Days&&d.last10Days.length>0){d.last10Days.forEach(function(day){var dot=document.createElement("div");dot.className="td ";var lb=day.date.substr(0,2);if(day.status==="absent"){dot.className+="ta";}else if(day.flags.indexOf("halfday")>=0){dot.className+="th";}else if(day.flags.indexOf("late")>=0){dot.className+="tl";}else{dot.className+="tp";}dot.textContent=lb;dot.title=day.date+(day.timeIn?" In:"+day.timeIn:"")+(day.timeOut?" Out:"+day.timeOut:"");tE.appendChild(dot);});}else{tE.innerHTML="<span style=\'font-size:10px;color:#9CA3AF\'>No data</span>";}'
/* Detail sections */
+'if(d.lateDates&&d.lateDates.length>0){document.getElementById("dLate").style.display="block";document.getElementById("lcnt").textContent=d.lateDates.length;document.getElementById("lBody").innerHTML=d.lateDates.map(function(x){return\'<span class="dtag red">\'+x+"</span>";}).join("");}else document.getElementById("dLate").style.display="none";'
+'if(d.earlyDates&&d.earlyDates.length>0){document.getElementById("dEarly").style.display="block";document.getElementById("ecnt").textContent=d.earlyDates.length;document.getElementById("eBody").innerHTML=d.earlyDates.map(function(x){return\'<span class="dtag purple">\'+x+"</span>";}).join("");}else document.getElementById("dEarly").style.display="none";'
+'if(d.halfDates&&d.halfDates.length>0){document.getElementById("dHalf").style.display="block";document.getElementById("hcnt").textContent=d.halfDates.length;document.getElementById("hBody").innerHTML=d.halfDates.map(function(x){return\'<span class="dtag orange">\'+x+"</span>";}).join("");}else document.getElementById("dHalf").style.display="none";'
+'if(d.absentDates&&d.absentDates.length>0){document.getElementById("dAbsent").style.display="block";document.getElementById("acnt").textContent=d.absentDates.length;document.getElementById("aBody").innerHTML=d.absentDates.map(function(x){return\'<span class="dtag red">\'+x+"</span>";}).join("");}else document.getElementById("dAbsent").style.display="none";'
/* Status */
+'var sc=document.getElementById("stCard"),st=document.getElementById("stT"),sm=document.getElementById("stM");'
+'if(pct>=90){sc.className="st-card sg";st.textContent="✅ Excellent!";sm.textContent="Bahut badhiya performance! Keep it up 🎉";}'
+'else if(pct>=75){sc.className="st-card sy";st.textContent="👍 Good";var tips=[];if(d.late>2)tips.push("Late: "+d.late+" baar");if(d.earlyOut>2)tips.push("Early: "+d.earlyOut+" baar");if(d.absent>2)tips.push("Absent: "+d.absent+" din");sm.textContent=tips.length>0?tips.join(" | "):"Achha hai, aur improve karo!";}'
+'else if(d.totalWorkingDays<3){sc.className="st-card sy";st.textContent="📅 Month shuru hua";sm.textContent="Kam working days abhi. Regular raho!";}'
+'else{sc.className="st-card sr";st.textContent="⚠ Attendance Kam!";sm.textContent="Present: "+d.present+"/"+d.totalWorkingDays+" | Late: "+d.late+" | Absent: "+d.absent+". Improve karo!";}'
+'}'


/* GO TO ATTENDANCE */
+'function goAtt(){fillForm(emp);}'
+'function fillForm(r){'
+'document.getElementById("fN").textContent=r.name;document.getElementById("fC").textContent=r.code;document.getElementById("fU").textContent=r.uc;document.getElementById("fE").textContent=r.email||"N/A";document.getElementById("hdrC").textContent=r.code;'
+'var ab=document.getElementById("attBtn");'
+'if(r.checkedIn&&r.checkedOut){document.getElementById("tIn").textContent=r.timeIn||"—";document.getElementById("bIn").innerHTML=\'<span class="badge b-in">✓</span>\';document.getElementById("tOut").textContent=r.timeOut||"—";document.getElementById("bOut").innerHTML=\'<span class="badge b-out">✓</span>\';ab.textContent="Complete ✓";ab.className="btn done";ab.disabled=true;cLk();}'
+'else if(r.checkedIn&&!r.checkedOut){document.getElementById("tIn").textContent=r.timeIn||"—";document.getElementById("bIn").innerHTML=\'<span class="badge b-in">✓ In</span>\';document.getElementById("tOut").textContent="—";document.getElementById("bOut").innerHTML=\'<span class="badge b-no">Pending</span>\';ab.textContent="Check Out →";ab.className="btn out";ab.disabled=false;sLk(r.code,r.name);}'
+'else{document.getElementById("tIn").textContent="—";document.getElementById("tOut").textContent="—";document.getElementById("bIn").innerHTML=\'<span class="badge b-no">Pending</span>\';document.getElementById("bOut").innerHTML=\'<span class="badge b-no">Pending</span>\';ab.textContent="Check In →";ab.className="btn primary";ab.disabled=false;}'
+'show("vForm");document.getElementById("attMsg").innerHTML="";}'

/* RECORD ATTENDANCE */
+'function doAtt(){'
+'if(!locOk){showMsg("attMsg","⚠ Office se bahar!","err");return;}'
+'var btn=document.getElementById("attBtn");btn.innerHTML=\'<span class="spin"></span>Saving...\';btn.disabled=true;'
+'google.script.run.withSuccessHandler(function(r){'
+'if(!r.success){showMsg("attMsg",r.msg,"err");btn.disabled=false;btn.innerHTML=emp.checkedIn?"Check Out →":"Check In →";return;}'
+'var sal=document.getElementById("sAl");sal.innerHTML="";'
+'if(r.action==="in"){'
+'sLk(emp.code,emp.name);document.getElementById("sIco").textContent="✅";document.getElementById("sT").textContent="Check In Done!";document.getElementById("sM").textContent=emp.name+" ki attendance lag gayi.";'
+'document.getElementById("sD").innerHTML="<strong>Name:</strong> "+emp.name+"<br><strong>Code:</strong> "+emp.code+"<br><strong>Date:</strong> "+r.date+"<br><strong>Time In:</strong> "+r.time+"<br><strong>Location:</strong> ✅ Office";'
+'if(r.isHalfDay){sal.innerHTML+=\'<div class="alert-half"><div class="al-icon">🟡</div><div><div class="al-title">Half Day</div><div class="al-sub">11:30–2:00 PM check-in = Half Day</div></div></div>\';openMod("half","Half Day","Half Day!","11:30 AM–2:00 PM ke beech aaye. Half Day count hogi.");}'
+'if(r.isLate){sal.innerHTML+=\'<div class="alert-late"><div class="al-icon">⏰</div><div><div class="al-title">Late!</div><div class="al-sub">\'+r.lateBy+\' late hain.</div></div></div>\';if(!r.isHalfDay)openMod("late","Late","Late!","10:20 AM ke baad. "+r.lateBy+" late.");}'
+'}else{'
+'cLk();document.getElementById("sIco").textContent="🏁";document.getElementById("sT").textContent="Check Out Done!";document.getElementById("sM").textContent=emp.name+" — day complete.";'
+'document.getElementById("sD").innerHTML="<strong>Name:</strong> "+emp.name+"<br><strong>Date:</strong> "+r.date+"<br><strong>In:</strong> "+r.timeIn+"<br><strong>Out:</strong> "+r.time+"<br><strong>Total:</strong> "+r.duration+" hrs";'
+'if(r.isHalfDay){sal.innerHTML+=\'<div class="alert-half"><div class="al-icon">🟡</div><div><div class="al-title">Half Day</div><div class="al-sub">5 PM se pehle = Half Day</div></div></div>\';openMod("half","Half Day","Half Day!","5 PM se pehle checkout = Half Day.");}'
+'if(r.isEarly){sal.innerHTML+=\'<div class="alert-early"><div class="al-icon">🏃</div><div><div class="al-title">Early!</div><div class="al-sub">\'+r.earlyBy+\' pehle ja rahe.</div></div></div>\';if(!r.isHalfDay)openMod("early","Early","Jaldi!","6:30 PM se "+r.earlyBy+" pehle.");}'
+'}'
+'show("vSuccess");'
+'}).withFailureHandler(function(e){showMsg("attMsg",e.message,"err");btn.disabled=false;btn.innerHTML=emp.checkedIn?"Check Out →":"Check In →";}).recordAttendance(emp.code,uLat,uLng,gDI());}'

/* LOGOUT */
+'function doLogout(){emp={};document.getElementById("codeIn").value="";document.getElementById("loginMsg").innerHTML="";show("vLogin");showLB();chkLoc();}'
+'window.onload=function(){chkLoc();showLB();};'
+'<\/script></body></html>';
}
