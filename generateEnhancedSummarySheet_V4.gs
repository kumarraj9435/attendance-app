// ═══════════════════════════════════════════════════════════════════════════
//  ORIGINAL V4 — generateEnhancedSummarySheet
//  This is your ORIGINAL function that produces the detailed ESummary
//  with Half Day Detail, Legend, Color bars, Info bar etc.
//  
//  INSTRUCTIONS: Copy this ENTIRE function and REPLACE the 
//  generateEnhancedSummarySheet in your Code.gs
// ═══════════════════════════════════════════════════════════════════════════

function generateEnhancedSummarySheet(month, year) {
  var ss        = SpreadsheetApp.openById(SHEET_ID);
  var empSheet  = ss.getSheetByName(DATA_SHEET);
  var attSheet  = ss.getSheetByName(ATT_SHEET);
  var sheetName = "ESummary - " + getMonthName(month) + " " + year;
  var sh        = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  sh.clearContents(); sh.clearFormats();
  sh.setFrozenRows(0); sh.setFrozenColumns(0);

  var tz       = Session.getScriptTimeZone();
  var today    = new Date();
  var holidays = getHolidays(month, year);

  // ── Load Employees ──
  var empData   = empSheet.getDataRange().getValues();
  var employees = [];
  for (var i = 1; i < empData.length; i++) {
    var r = empData[i];
    if (!r[2] || r[2].toString().trim() === "") continue;
    employees.push({
      name:        r[0] ? r[0].toString().trim() : "",
      designation: r[6] ? r[6].toString().trim() : "",
      dept:        r[5] ? r[5].toString().trim() : "General",
      code:        r[2].toString().trim()
    });
  }

  // ── Load Attendance Map ──
  var attData = attSheet ? attSheet.getDataRange().getValues() : [[]];
  var attMap  = {};
  for (var j = 1; j < attData.length; j++) {
    var row = attData[j];
    if (!row[0]) continue;
    var ds = "";
    try {
      ds = row[0].toString().indexOf("-") > -1
           ? row[0].toString().trim()
           : Utilities.formatDate(new Date(row[0]), tz, "dd-MM-yyyy");
    } catch(ex) { continue; }
    var pts = ds.split("-");
    if (pts.length < 3 || parseInt(pts[1]) !== month || parseInt(pts[2]) !== year) continue;
    var code = row[2] ? row[2].toString().trim() : "";
    if (!code) continue;
    if (!attMap[code]) attMap[code] = {};
    var rawIn = "", rawOut = "";
    try {
      if (row[5]) rawIn  = row[5] instanceof Date ? Utilities.formatDate(row[5], tz, "hh:mm a") : formatTimeShort(row[5].toString());
      if (row[6]) rawOut = row[6] instanceof Date ? Utilities.formatDate(row[6], tz, "hh:mm a") : formatTimeShort(row[6].toString());
    } catch(fe) {}
    attMap[code][ds] = {
      timeIn:   rawIn,
      timeOut:  rawOut,
      duration: row[7] ? (parseFloat(row[7].toString().replace(/[^0-9.]/g,"")) || 0) : 0
    };
  }

  // ── Working Days ──
  var daysInMonth = new Date(year, month, 0).getDate();
  var workingDays = [];
  for (var d = 1; d <= daysInMonth; d++) {
    var dt = new Date(year, month - 1, d);
    if (dt.getDay() === 0 || dt > today) continue;
    var dKey = d.toString().padStart(2,"0") + "-" + month.toString().padStart(2,"0") + "-" + year;
    if (holidays[dKey]) continue;
    workingDays.push(dKey);
  }
  var totalWD = workingDays.length;

  // ════ BUILD SHEET ════
  var sRow = 1;

  // ── Title Banner ──
  sh.getRange(sRow, 1, 1, 13).merge()
    .setValue("📊  ENHANCED ATTENDANCE SUMMARY  —  " + getMonthName(month).toUpperCase() + " " + year)
    .setFontSize(16).setFontWeight("bold")
    .setBackground("#1E1B4B").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(sRow, 48); sRow++;

  // ── Info Bar ──
  sh.getRange(sRow, 1, 1, 13).merge()
    .setValue(
      "📅 Working Days (so far): " + totalWD +
      "    👥 Total Employees: " + employees.length +
      "    🎉 Holidays: " + Object.keys(holidays).length +
      "    ⏰ Late After: 10:20 AM" +
      "    🚪 Early Before: 6:30 PM" +
      "    🟡 Half Day CI: 11:30 AM – 2:00 PM  |  Half Day CO: Before 5:00 PM" +
      "    🔄 Generated: " + Utilities.formatDate(today, tz, "dd-MM-yyyy hh:mm a")
    )
    .setFontSize(9).setBackground("#E0E7FF").setFontColor("#3730A3")
    .setHorizontalAlignment("center").setWrap(true);
  sh.setRowHeight(sRow, 36); sRow++;

  // ── Color Legend ──
  var legendData = [
    ["🟢 ≥95% Excellent", "#DCFCE7", "#166534"],
    ["🟡 75–94% Good",    "#FEF9C3", "#854D0E"],
    ["🔴 <75% Critical",  "#FEE2E2", "#991B1B"],
    ["🟠 Late (>10:20AM)","#FFCCCC", "#CC0000"],
    ["🟠 Early (<6:30PM)","#FFCCCC", "#CC0000"],
    ["🟡 Half Day CI (11:30–2PM)", "#FFF3CD", "#856404"],
    ["🟡 Half Day CO (<5PM)", "#FFF3CD", "#856404"]
  ];
  for (var lg = 0; lg < legendData.length; lg++) {
    sh.getRange(sRow, lg + 1)
      .setValue(legendData[lg][0])
      .setBackground(legendData[lg][1]).setFontColor(legendData[lg][2])
      .setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  }
  sh.getRange(sRow, 8, 1, 6).merge().setValue("").setBackground("#F8FAFF");
  sh.setRowHeight(sRow, 26); sRow++;

  // ── Spacer ──
  sh.getRange(sRow, 1, 1, 13).merge().setBackground("#F8FAFF");
  sh.setRowHeight(sRow, 6); sRow++;

  // ── Table Header ──
  var headers = [
    "#", "Employee Name", "Designation", "Code",
    "Present", "Absent", "🟡 Half\nDay",
    "⏰ Late\nCheck-In", "🚪 Early\nCheck-Out",
    "Total\nHrs", "Attend.\n%", "Trend", "Status"
  ];
  sh.getRange(sRow, 1, 1, 13).setValues([headers])
    .setFontWeight("bold").setFontSize(11)
    .setBackground("#1E40AF").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setWrap(true);
  sh.setRowHeight(sRow, 44);
  var headerRow = sRow; sRow++;

  // ── Employee Rows ──
  var gP=0, gA=0, gHD=0, gL=0, gE=0, gH=0;
  var empResults = [];

  for (var e = 0; e < employees.length; e++) {
    var emp = employees[e];
    var present=0, absent=0, halfDay=0, late=0, earlyOut=0, hrs=0;

    for (var w = 0; w < workingDays.length; w++) {
      var rec = attMap[emp.code] && attMap[emp.code][workingDays[w]]
                ? attMap[emp.code][workingDays[w]] : null;

      if (!rec || rec.timeIn === "") { absent++; continue; }
      present++;

      var hdCI = isHalfDayCI(rec.timeIn);
      var hdCO = rec.timeOut && rec.timeOut !== "-" && rec.timeOut !== ""
                 && isHalfDayCO(rec.timeOut);
      if (hdCI || hdCO) halfDay++;
      if (isLateCI(rec.timeIn)) late++;
      if (rec.timeOut && rec.timeOut !== "-" && rec.timeOut !== "" && isEarlyCO(rec.timeOut)) earlyOut++;
      var dVal = parseFloat(rec.duration);
      if (!isNaN(dVal)) hrs += dVal;
    }

    gP+=present; gA+=absent; gHD+=halfDay; gL+=late; gE+=earlyOut; gH+=hrs;
    var pct    = totalWD > 0 ? (present / totalWD) * 100 : 0;
    var pctTxt = pct.toFixed(1) + "%";

    // Trend sparkline (last 7 days)
    var trendStr = "";
    var last7 = workingDays.slice(-7);
    for (var t7 = 0; t7 < last7.length; t7++) {
      var r7 = attMap[emp.code] && attMap[emp.code][last7[t7]] ? attMap[emp.code][last7[t7]] : null;
      if (!r7 || !r7.timeIn) trendStr += "✗";
      else if (isHalfDayCI(r7.timeIn) || (r7.timeOut && r7.timeOut !== "-" && isHalfDayCO(r7.timeOut))) trendStr += "½";
      else trendStr += "✓";
    }

    // Status text
    var statusParts = [];
    if (pct >= 95) statusParts.push("✅ Excellent");
    else if (pct >= 75) statusParts.push("👍 Good");
    else statusParts.push("⚠ Critical");
    if (halfDay > 0) statusParts.push("½ " + halfDay + " Half Day");
    if (late > 3)    statusParts.push("🔴 Late x" + late);
    else if (late > 0) statusParts.push("🟡 Late x" + late);
    if (earlyOut > 3)   statusParts.push("🔴 Early x" + earlyOut);
    else if (earlyOut > 0) statusParts.push("🟡 Early x" + earlyOut);
    if (absent > 3) statusParts.push("🔴 Abs x" + absent);

    empResults.push({
      emp, present, absent, halfDay, late, earlyOut,
      hrs, pct, pctTxt, trendStr, statusTxt: statusParts.join("  |  ")
    });
  }

  // Write rows
  for (var er = 0; er < empResults.length; er++) {
    var res = empResults[er];
    var bg  = er % 2 === 0 ? "#FFFFFF" : "#F0F4FF";

    sh.getRange(sRow, 1, 1, 13).setValues([[
      er + 1, res.emp.name, res.emp.designation, res.emp.code,
      res.present, res.absent, res.halfDay, res.late, res.earlyOut,
      res.hrs.toFixed(2), res.pctTxt, res.trendStr, res.statusTxt
    ]]).setBackground(bg).setHorizontalAlignment("center").setFontSize(11);

    sh.getRange(sRow, 2).setHorizontalAlignment("left").setFontWeight("bold");
    sh.getRange(sRow, 3).setHorizontalAlignment("left").setFontSize(10).setFontColor("#6B7280");
    sh.getRange(sRow, 12).setFontSize(9).setFontColor("#374151").setHorizontalAlignment("center");
    sh.getRange(sRow, 13).setHorizontalAlignment("left").setFontSize(9);

    // Conditional: Attendance %
    var pctCell = sh.getRange(sRow, 11);
    if (res.pct >= 95)      pctCell.setBackground("#DCFCE7").setFontColor("#166534").setFontWeight("bold");
    else if (res.pct >= 75) pctCell.setBackground("#FEF9C3").setFontColor("#854D0E").setFontWeight("bold");
    else                    pctCell.setBackground("#FEE2E2").setFontColor("#991B1B").setFontWeight("bold");

    if (res.absent > 3) sh.getRange(sRow, 6).setBackground("#FEE2E2").setFontColor("#991B1B").setFontWeight("bold");

    var hdCell = sh.getRange(sRow, 7);
    if (res.halfDay > 2)      hdCell.setBackground(HALF_BG).setFontColor(HALF_FG).setFontWeight("bold");
    else if (res.halfDay > 0) hdCell.setBackground("#FFFBEB").setFontColor("#92400E");

    var lCell = sh.getRange(sRow, 8);
    if (res.late > 3)      lCell.setBackground(LATE_BG).setFontColor(LATE_FG).setFontWeight("bold");
    else if (res.late > 0) lCell.setBackground("#FFF3CD").setFontColor("#856404");

    var eCell = sh.getRange(sRow, 9);
    if (res.earlyOut > 3)      eCell.setBackground(EARLY_BG).setFontColor(EARLY_FG).setFontWeight("bold");
    else if (res.earlyOut > 0) eCell.setBackground("#FFF3CD").setFontColor("#856404");

    sh.setRowHeight(sRow, 32); sRow++;
  }

  // ── Totals Row ──
  var avgPct = employees.length > 0 && totalWD > 0
               ? ((gP / (totalWD * employees.length)) * 100).toFixed(1) + "%" : "0%";
  sh.getRange(sRow, 1, 1, 13).setValues([[
    "", "▶ TOTAL / AVERAGE", "", "",
    gP, gA, gHD, gL, gE,
    gH.toFixed(2), avgPct, "", ""
  ]]).setFontWeight("bold").setFontSize(12)
    .setBackground("#1E40AF").setFontColor("white")
    .setHorizontalAlignment("center");
  sh.getRange(sRow, 2).setHorizontalAlignment("left");
  sh.setRowHeight(sRow, 38); sRow += 2;

  // ════════════════════════════════════════════════════════
  //  HALF DAY DETAIL SECTION
  // ════════════════════════════════════════════════════════
  sh.getRange(sRow, 1, 1, 13).merge()
    .setValue("🟡  HALF DAY DETAIL  —  Employee + Date wise breakdown")
    .setFontSize(13).setFontWeight("bold")
    .setBackground("#78350F").setFontColor("white")
    .setHorizontalAlignment("center");
  sh.setRowHeight(sRow, 34); sRow++;

  sh.getRange(sRow, 1, 1, 8).setValues([[
    "Date", "Employee Name", "Code", "Check-In", "Check-Out", "Hours", "Half Day Reason", ""
  ]]).setFontWeight("bold").setFontSize(11)
    .setBackground("#FEF3C7").setFontColor("#92400E")
    .setHorizontalAlignment("center");
  sh.setRowHeight(sRow, 30); sRow++;

  var hdCount = 0;
  for (var w2 = 0; w2 < workingDays.length; w2++) {
    for (var e2 = 0; e2 < employees.length; e2++) {
      var emp2 = employees[e2];
      var rec2 = attMap[emp2.code] && attMap[emp2.code][workingDays[w2]]
                 ? attMap[emp2.code][workingDays[w2]] : null;
      if (!rec2 || rec2.timeIn === "") continue;

      var hdCI2 = isHalfDayCI(rec2.timeIn);
      var hdCO2 = rec2.timeOut && rec2.timeOut !== "-" && rec2.timeOut !== ""
                  && isHalfDayCO(rec2.timeOut);
      if (!hdCI2 && !hdCO2) continue;

      var reason = "";
      if (hdCI2 && hdCO2) reason = "⚠ Both: Check-in 11:30AM–2PM + Check-out before 5:00 PM";
      else if (hdCI2)     reason = "⏰ Check-in between 11:30 AM – 2:00 PM  →  Half Day";
      else                reason = "🚪 Check-out before 5:00 PM  →  Half Day";

      var rowBg2 = hdCount % 2 === 0 ? "#FFFBEB" : "#FEF9C3";
      var durTxt = rec2.duration ? rec2.duration + " hrs" : "—";

      sh.getRange(sRow, 1, 1, 7).setValues([[
        workingDays[w2], emp2.name, emp2.code,
        rec2.timeIn, rec2.timeOut || "—",
        durTxt, reason
      ]]).setBackground(rowBg2).setHorizontalAlignment("center");
      sh.getRange(sRow, 2).setHorizontalAlignment("left").setFontWeight("bold");
      sh.getRange(sRow, 7).setHorizontalAlignment("left").setFontColor(HALF_FG).setFontWeight("bold");
      sh.getRange(sRow, 8, 1, 6).merge().setBackground(rowBg2);
      sh.setRowHeight(sRow, 26); sRow++; hdCount++;
    }
  }

  if (hdCount === 0) {
    sh.getRange(sRow, 1, 1, 13).merge()
      .setValue("✅ Is month mein koi Half Day record nahi mila.")
      .setBackground("#DCFCE7").setFontColor("#166534")
      .setFontWeight("bold").setHorizontalAlignment("center");
    sh.setRowHeight(sRow, 30); sRow++;
  } else {
    sRow++;
    sh.getRange(sRow, 1, 1, 13).merge()
      .setValue("📊  Half Day Count per Employee")
      .setFontSize(12).setFontWeight("bold")
      .setBackground("#92400E").setFontColor("white")
      .setHorizontalAlignment("center");
    sh.setRowHeight(sRow, 28); sRow++;

    sh.getRange(sRow, 1, 1, 4).setValues([["Employee Name", "Code", "Half Day Count", "Status"]])
      .setFontWeight("bold").setBackground("#FEF3C7").setFontColor("#92400E").setHorizontalAlignment("center");
    sh.setRowHeight(sRow, 26); sRow++;

    for (var hc = 0; hc < empResults.length; hc++) {
      if (empResults[hc].halfDay === 0) continue;
      var hcBg = empResults[hc].halfDay > 2 ? "#FFF3CD" : "#FFFBEB";
      sh.getRange(sRow, 1, 1, 4).setValues([[
        empResults[hc].emp.name,
        empResults[hc].emp.code,
        empResults[hc].halfDay,
        empResults[hc].halfDay > 4 ? "⚠ High — Review Needed" :
        empResults[hc].halfDay > 2 ? "🟡 Moderate" : "✓ Low"
      ]]).setBackground(hcBg).setHorizontalAlignment("center");
      sh.getRange(sRow, 1).setHorizontalAlignment("left").setFontWeight("bold");
      sh.getRange(sRow, 4).setHorizontalAlignment("left")
        .setFontColor(empResults[hc].halfDay > 4 ? "#991B1B" : empResults[hc].halfDay > 2 ? "#856404" : "#166534")
        .setFontWeight("bold");
      sh.setRowHeight(sRow, 26); sRow++;
    }
  }
  sRow++;

  // ── Holidays ──
  var holKeys = Object.keys(holidays);
  if (holKeys.length > 0) {
    sh.getRange(sRow, 1, 1, 13).merge()
      .setValue("🎉  HOLIDAYS THIS MONTH")
      .setFontSize(12).setFontWeight("bold")
      .setBackground("#3730A3").setFontColor("white")
      .setHorizontalAlignment("center");
    sh.setRowHeight(sRow, 28); sRow++;
    for (var h = 0; h < holKeys.length; h++) {
      sh.getRange(sRow, 1).setValue(holKeys[h]).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#E0E7FF").setFontColor("#3730A3");
      sh.getRange(sRow, 2, 1, 12).merge().setValue(holidays[holKeys[h]]).setBackground("#E0E7FF").setFontColor("#3730A3").setFontWeight("bold");
      sh.setRowHeight(sRow, 24); sRow++;
    }
    sRow++;
  }

  // ── Legend ──
  sh.getRange(sRow, 1, 1, 13).merge()
    .setValue("🗝  LEGEND")
    .setFontWeight("bold").setBackground("#374151").setFontColor("white")
    .setHorizontalAlignment("center");
  sh.setRowHeight(sRow, 26); sRow++;

  var legendFull = [
    ["🟢 Green (Attendance %)",   "≥ 95% — Excellent Attendance"],
    ["🟡 Yellow (Attendance %)",  "75–94% — Good, keep improving"],
    ["🔴 Red (Attendance %)",     "< 75% — Needs Improvement / HR Action"],
    ["🟡 Yellow (Half Day CI)",   "Check-in between 11:30 AM – 2:00 PM = Half Day"],
    ["🟡 Yellow (Half Day CO)",   "Check-out before 5:00 PM = Half Day"],
    ["🔴 Red (Half Day col)",     "> 2 Half Days in month — Warning"],
    ["🔴 Red (Late Check-In)",    "> 3 Late entries (after 10:20 AM)"],
    ["🔴 Red (Early Checkout)",   "> 3 Early departures (before 6:30 PM)"],
    ["🔴 Red (Absent)",           "> 3 Absents — Warning"],
    ["Trend (✓ ½ ✗)",            "Last 7 working days: ✓ Full day  |  ½ Half day  |  ✗ Absent"],
    ["🎉 HOL",                    "Holiday — Not counted in Absent or Working Days"],
    ["☀️ SUN",                   "Sunday — Weekly Off (not counted)"]
  ];
  for (var l = 0; l < legendFull.length; l++) {
    sh.getRange(sRow, 1).setValue(legendFull[l][0]).setFontWeight("bold");
    sh.getRange(sRow, 2, 1, 12).merge().setValue(legendFull[l][1]);
    sh.setRowHeight(sRow, 22); sRow++;
  }

  // ── Column Widths ──
  sh.setColumnWidth(1,  38);
  sh.setColumnWidth(2,  145);
  sh.setColumnWidth(3,  145);
  sh.setColumnWidth(4,  85);
  sh.setColumnWidth(5,  68);
  sh.setColumnWidth(6,  68);
  sh.setColumnWidth(7,  72);
  sh.setColumnWidth(8,  85);
  sh.setColumnWidth(9,  90);
  sh.setColumnWidth(10, 75);
  sh.setColumnWidth(11, 80);
  sh.setColumnWidth(12, 95);
  sh.setColumnWidth(13, 200);

  sh.setFrozenRows(headerRow);
  sh.setFrozenColumns(2);

  Logger.log("✅ Enhanced Summary done: " + sheetName + " | HalfDays: " + hdCount);
}
