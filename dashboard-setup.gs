/**
 * BLUCOR Daily Activity Report — KPI Dashboard Setup
 *
 * PREREQUISITES:
 *   1. You already ran form-generator.gs to create the Form + linked Sheet
 *   2. Open that linked Google Sheet
 *   3. Extensions → Apps Script → paste this file → Run setupDashboard()
 *
 * CREATES:
 *   - "KPI Engine" tab   — per-employee metric calculations
 *   - "Dashboard" tab    — management scorecard with charts
 *   - "Config" tab       — settings (expected work days, thresholds)
 */

function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Find the Form Responses sheet (auto-named by Google)
  var responseSheet = findResponseSheet(ss);
  if (!responseSheet) {
    throw new Error(
      'Could not find a Form Responses sheet. Make sure this spreadsheet ' +
      'is linked to the BLUCOR Daily Activity Report form.'
    );
  }

  var respName = responseSheet.getName();
  Logger.log('Found response sheet: ' + respName);

  // Map column positions from the response sheet headers
  var headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
  var col = mapColumns(headers);
  Logger.log('Column mapping complete: ' + JSON.stringify(col));

  // ═══════════════════════════════════════════════════════
  //  CONFIG TAB
  // ═══════════════════════════════════════════════════════
  var configSheet = getOrCreateSheet(ss, 'Config');
  configSheet.clear();
  configSheet.appendRow(['Setting', 'Value', 'Notes']);
  configSheet.appendRow(['Expected Work Days per Week', 5, 'Mon–Fri default']);
  configSheet.appendRow(['KPI Green Threshold', 80, 'Score >= this = green']);
  configSheet.appendRow(['KPI Yellow Threshold', 50, 'Score >= this = yellow; below = red']);
  configSheet.appendRow(['Dashboard Start Date', '', 'Leave blank for all-time; or enter YYYY-MM-DD']);
  configSheet.appendRow(['Dashboard End Date', '', 'Leave blank for today']);
  configSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  configSheet.setColumnWidth(1, 260);
  configSheet.setColumnWidth(2, 100);
  configSheet.setColumnWidth(3, 300);

  // ═══════════════════════════════════════════════════════
  //  KPI ENGINE TAB
  // ═══════════════════════════════════════════════════════
  var kpiSheet = getOrCreateSheet(ss, 'KPI Engine');
  kpiSheet.clear();

  // Header row
  var kpiHeaders = [
    'Employee Name',
    'Total Submissions',
    'Avg Effectiveness (1–10)',
    'Avg Time Mgmt (1–10)',
    'Productivity Score',
    'PPE Done Rate %',
    'Avg Safety Incidents',
    'Safety Compliance Score',
    'Extra Tasks Yes %',
    'Avg Proactive Tasks',
    'Improvements Yes %',
    'Initiative Score',
    'Error Check Yes %',
    'Avg Error Check Count',
    'Quality Score',
    'Avg Internal Comms',
    'Avg External Comms',
    'Avg Daily Comm Volume',
    'Mentoring Yes %',
    'Submission Consistency %'
  ];
  kpiSheet.appendRow(kpiHeaders);
  kpiSheet.getRange(1, 1, 1, kpiHeaders.length).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');

  // Build unique employee list formula in a helper column (col V = 22)
  // We'll use a script-driven approach instead of complex array formulas
  // to keep it readable and maintainable. The refresh function recalculates.

  // Write a note explaining the tab
  kpiSheet.getRange('A2').setNote(
    'This sheet is populated by running updateKPIs() from the script editor, ' +
    'or automatically via the daily trigger set up below.'
  );

  // ═══════════════════════════════════════════════════════
  //  DASHBOARD TAB
  // ═══════════════════════════════════════════════════════
  var dashSheet = getOrCreateSheet(ss, 'Dashboard');
  dashSheet.clear();

  // Title
  dashSheet.getRange('A1').setValue('BLUCOR — Employee KPI Dashboard').setFontSize(16).setFontWeight('bold').setFontColor('#1a3a5c');
  dashSheet.getRange('A2').setValue('Last updated:').setFontWeight('bold');
  dashSheet.getRange('B2').setFormula('=NOW()').setNumberFormat('yyyy-mm-dd hh:mm AM/PM');

  // Employee selector
  dashSheet.getRange('A4').setValue('Employee:').setFontWeight('bold');
  dashSheet.getRange('B4').setValue('(All)');
  dashSheet.getRange('B4').setNote('Type an employee name to filter, or leave as (All) for team view');

  // Team summary table header (row 6)
  var dashHeaders = [
    'Employee',
    'Submissions',
    'Productivity',
    'Safety',
    'Initiative',
    'Quality',
    'Comms/Day',
    'Mentoring %',
    'Consistency %',
    'Overall Grade'
  ];
  dashSheet.getRange(6, 1, 1, dashHeaders.length).setValues([dashHeaders]);
  dashSheet.getRange(6, 1, 1, dashHeaders.length)
    .setFontWeight('bold')
    .setBackground('#1a3a5c')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // Column widths
  dashSheet.setColumnWidth(1, 180);
  for (var c = 2; c <= dashHeaders.length; c++) {
    dashSheet.setColumnWidth(c, 110);
  }

  // KPI definitions reference (right side)
  dashSheet.getRange('L6').setValue('KPI Definitions').setFontWeight('bold').setFontSize(11);
  var defs = [
    ['Productivity', 'AVG(Effectiveness + Time Mgmt ratings) / 10 × 100'],
    ['Safety', '(PPE Done % + Zero-Incident %) / 2'],
    ['Initiative', '(Extra Tasks % + Has Proactive % + Improvements %) / 3'],
    ['Quality', '(Error Check % + min(Avg Checks, 3)/3 × 100) / 2'],
    ['Comms/Day', 'AVG(Internal + External) per submission'],
    ['Mentoring %', 'Days mentored / Total submissions × 100'],
    ['Consistency', 'Actual submissions / Expected work days × 100'],
    ['Overall', 'Weighted avg: Prod 25%, Safety 20%, Init 15%, Quality 15%, Consist 15%, Mentor 10%']
  ];
  dashSheet.getRange(7, 12, defs.length, 2).setValues(defs);
  dashSheet.setColumnWidth(12, 120);
  dashSheet.setColumnWidth(13, 400);

  // ═══════════════════════════════════════════════════════
  //  CREATE TRIGGER — auto-refresh KPIs daily at 5 AM
  // ═══════════════════════════════════════════════════════
  removeTriggers('updateKPIs'); // clear any existing
  ScriptApp.newTrigger('updateKPIs')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();
  Logger.log('Daily trigger set: updateKPIs runs at 5 AM');

  // ═══════════════════════════════════════════════════════
  //  INITIAL KPI CALCULATION
  // ═══════════════════════════════════════════════════════
  updateKPIs();

  Logger.log('═══════════════════════════════════════════');
  Logger.log('  Dashboard setup complete!');
  Logger.log('═══════════════════════════════════════════');
  Logger.log('Tabs created: Config, KPI Engine, Dashboard');
  Logger.log('Trigger: updateKPIs() runs daily at 5 AM');
  Logger.log('To refresh manually: Run updateKPIs()');
}


// ═══════════════════════════════════════════════════════════
//  updateKPIs() — Reads Form Responses, calculates all KPIs,
//  writes to KPI Engine and Dashboard tabs
// ═══════════════════════════════════════════════════════════

function updateKPIs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = findResponseSheet(ss);
  if (!responseSheet) { Logger.log('No response sheet found.'); return; }

  var data = responseSheet.getDataRange().getValues();
  if (data.length < 2) { Logger.log('No responses yet.'); return; }

  var headers = data[0];
  var col = mapColumns(headers);

  // ── Aggregate per employee ──
  var employees = {};

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var name = String(row[col.employeeName] || '').trim();
    if (!name) continue;

    if (!employees[name]) {
      employees[name] = {
        submissions: 0,
        effectivenessSum: 0, effectivenessCount: 0,
        timeMgmtSum: 0, timeMgmtCount: 0,
        ppeDone: 0,
        incidentSum: 0,
        zeroIncident: 0,
        extraTasksYes: 0,
        hasProactive: 0,
        improvementsYes: 0,
        errorCheckYes: 0,
        errorCheckCountSum: 0, errorCheckCountN: 0,
        internalCommsSum: 0,
        externalCommsSum: 0,
        mentoringYes: 0,
        dates: {}
      };
    }

    var emp = employees[name];
    emp.submissions++;

    // Track unique dates for consistency calc
    var dateVal = row[col.reportDate];
    if (dateVal instanceof Date) {
      emp.dates[dateVal.toISOString().slice(0, 10)] = true;
    }

    // Effectiveness (1–10 scale)
    var eff = parseFloat(row[col.effectiveness]);
    if (!isNaN(eff)) { emp.effectivenessSum += eff; emp.effectivenessCount++; }

    // Time management (1–10 scale)
    var tm = parseFloat(row[col.timeMgmt]);
    if (!isNaN(tm)) { emp.timeMgmtSum += tm; emp.timeMgmtCount++; }

    // PPE
    var ppe = String(row[col.ppe] || '').toLowerCase();
    if (ppe.indexOf('done') >= 0 || ppe.indexOf('good') >= 0) emp.ppeDone++;

    // Safety incidents
    var inc = parseFloat(row[col.incidents]);
    if (!isNaN(inc)) {
      emp.incidentSum += inc;
      if (inc === 0) emp.zeroIncident++;
    }

    // Extra tasks
    var extra = String(row[col.extraTasks] || '').toLowerCase();
    if (extra === 'yes') emp.extraTasksYes++;

    // Proactive tasks (count non-empty across 3 fields)
    var proCount = 0;
    [col.proactive1, col.proactive2, col.proactive3].forEach(function(ci) {
      if (ci >= 0 && String(row[ci] || '').trim()) proCount++;
    });
    if (proCount > 0) emp.hasProactive++;

    // Process improvements
    var improv = String(row[col.improvements] || '').toLowerCase();
    if (improv === 'yes') emp.improvementsYes++;

    // Error checking
    var errChk = String(row[col.errorCheck] || '').toLowerCase();
    if (errChk === 'yes') emp.errorCheckYes++;

    var errCount = parseFloat(row[col.errorCheckCount]);
    if (!isNaN(errCount)) { emp.errorCheckCountSum += errCount; emp.errorCheckCountN++; }

    // Communications
    var intComms = parseFloat(row[col.internalComms]);
    if (!isNaN(intComms)) emp.internalCommsSum += intComms;

    var extComms = parseFloat(row[col.externalComms]);
    if (!isNaN(extComms)) emp.externalCommsSum += extComms;

    // Mentoring
    var mentor = String(row[col.mentoring] || '').toLowerCase();
    if (mentor === 'yes') emp.mentoringYes++;
  }

  // ── Get config ──
  var configSheet = ss.getSheetByName('Config');
  var expectedDaysPerWeek = 5;
  if (configSheet) {
    var cfgVal = configSheet.getRange('B2').getValue();
    if (cfgVal) expectedDaysPerWeek = parseFloat(cfgVal) || 5;
  }

  // ── Calculate KPIs and build output rows ──
  var kpiRows = [];
  var dashRows = [];
  var names = Object.keys(employees).sort();

  names.forEach(function(name) {
    var e = employees[name];
    var s = e.submissions;
    if (s === 0) return;

    // Averages
    var avgEff = e.effectivenessCount > 0 ? e.effectivenessSum / e.effectivenessCount : 0;
    var avgTM = e.timeMgmtCount > 0 ? e.timeMgmtSum / e.timeMgmtCount : 0;

    // Productivity Score: avg of both ratings, scaled to 100
    var productivity = ((avgEff + avgTM) / 2) / 10 * 100;

    // Safety Compliance: (PPE done rate + zero-incident rate) / 2
    var ppeDoneRate = (e.ppeDone / s) * 100;
    var zeroIncRate = (e.zeroIncident / s) * 100;
    var avgIncidents = e.incidentSum / s;
    var safety = (ppeDoneRate + zeroIncRate) / 2;

    // Initiative: (extra tasks % + has proactive % + improvements %) / 3
    var extraRate = (e.extraTasksYes / s) * 100;
    var proactiveRate = (e.hasProactive / s) * 100;
    var improvRate = (e.improvementsYes / s) * 100;
    var initiative = (extraRate + proactiveRate + improvRate) / 3;

    // Quality: (error check rate + normalized avg check count) / 2
    var errorCheckRate = (e.errorCheckYes / s) * 100;
    var avgCheckCount = e.errorCheckCountN > 0 ? e.errorCheckCountSum / e.errorCheckCountN : 0;
    var normalizedChecks = Math.min(avgCheckCount / 3, 1) * 100; // 3+ checks = 100%
    var quality = (errorCheckRate + normalizedChecks) / 2;

    // Communication volume
    var avgIntComms = e.internalCommsSum / s;
    var avgExtComms = e.externalCommsSum / s;
    var avgDailyComms = avgIntComms + avgExtComms;

    // Mentoring rate
    var mentoringRate = (e.mentoringYes / s) * 100;

    // Submission consistency
    var uniqueDates = Object.keys(e.dates).length;
    // Estimate expected days: span from first to last date / 7 * expectedDaysPerWeek
    var dateKeys = Object.keys(e.dates).sort();
    var expectedDays = s; // fallback
    if (dateKeys.length >= 2) {
      var firstDate = new Date(dateKeys[0]);
      var lastDate = new Date(dateKeys[dateKeys.length - 1]);
      var spanDays = Math.max(1, Math.ceil((lastDate - firstDate) / 86400000) + 1);
      var spanWeeks = Math.max(1, spanDays / 7);
      expectedDays = Math.ceil(spanWeeks * expectedDaysPerWeek);
    }
    var consistency = Math.min((uniqueDates / expectedDays) * 100, 100);

    // Overall grade (weighted average)
    var overall = (
      productivity * 0.25 +
      safety * 0.20 +
      initiative * 0.15 +
      quality * 0.15 +
      consistency * 0.15 +
      mentoringRate * 0.10
    );

    // KPI Engine row
    kpiRows.push([
      name,
      s,
      round2(avgEff),
      round2(avgTM),
      round2(productivity),
      round2(ppeDoneRate),
      round2(avgIncidents),
      round2(safety),
      round2(extraRate),
      round2(avgCheckCount > 0 ? e.hasProactive / s : proactiveRate / 100), // avg proactive tasks
      round2(improvRate),
      round2(initiative),
      round2(errorCheckRate),
      round2(avgCheckCount),
      round2(quality),
      round2(avgIntComms),
      round2(avgExtComms),
      round2(avgDailyComms),
      round2(mentoringRate),
      round2(consistency)
    ]);

    // Dashboard row
    dashRows.push([
      name,
      s,
      round2(productivity),
      round2(safety),
      round2(initiative),
      round2(quality),
      round2(avgDailyComms),
      round2(mentoringRate),
      round2(consistency),
      round2(overall)
    ]);
  });

  // ── Write KPI Engine ──
  var kpiSheet = ss.getSheetByName('KPI Engine');
  if (kpiSheet && kpiRows.length > 0) {
    // Clear data rows (keep header)
    if (kpiSheet.getLastRow() > 1) {
      kpiSheet.getRange(2, 1, kpiSheet.getLastRow() - 1, 20).clearContent();
    }
    kpiSheet.getRange(2, 1, kpiRows.length, kpiRows[0].length).setValues(kpiRows);
  }

  // ── Write Dashboard ──
  var dashSheet = ss.getSheetByName('Dashboard');
  if (dashSheet && dashRows.length > 0) {
    // Clear data rows (keep header at row 6)
    if (dashSheet.getLastRow() > 6) {
      dashSheet.getRange(7, 1, dashSheet.getLastRow() - 6, 10).clearContent();
    }
    dashSheet.getRange(7, 1, dashRows.length, dashRows[0].length).setValues(dashRows);

    // Apply conditional formatting (green/yellow/red) to score columns (3–6, 8–10)
    applyThresholdFormatting(dashSheet, dashRows.length);

    // Update timestamp
    dashSheet.getRange('B2').setFormula('=NOW()');
  }

  Logger.log('KPIs updated for ' + names.length + ' employee(s), ' + (data.length - 1) + ' total submissions.');
}


// ═══════════════════════════════════════════════════════════
//  CONDITIONAL FORMATTING — green/yellow/red thresholds
// ═══════════════════════════════════════════════════════════

function applyThresholdFormatting(sheet, numRows) {
  if (numRows < 1) return;

  // Get thresholds from Config
  var ss = sheet.getParent();
  var configSheet = ss.getSheetByName('Config');
  var green = 80, yellow = 50;
  if (configSheet) {
    green = parseFloat(configSheet.getRange('B3').getValue()) || 80;
    yellow = parseFloat(configSheet.getRange('B4').getValue()) || 50;
  }

  // Score columns: C(3), D(4), E(5), F(6), H(8), I(9), J(10)
  var scoreCols = [3, 4, 5, 6, 8, 9, 10];

  // Clear existing conditional format rules
  sheet.clearConditionalFormatRules();
  var rules = [];

  scoreCols.forEach(function(c) {
    var range = sheet.getRange(7, c, numRows, 1);

    // Green: >= green threshold
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(green)
      .setBackground('#d5f5e3')
      .setFontColor('#1e7e34')
      .setRanges([range])
      .build());

    // Yellow: >= yellow and < green
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(yellow, green - 0.01)
      .setBackground('#fef9e7')
      .setFontColor('#9a7d0a')
      .setRanges([range])
      .build());

    // Red: < yellow
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(yellow)
      .setBackground('#fadbd8')
      .setFontColor('#922b21')
      .setRanges([range])
      .build());
  });

  sheet.setConditionalFormatRules(rules);
}


// ═══════════════════════════════════════════════════════════
//  COLUMN MAPPER — Finds column indices by header text
// ═══════════════════════════════════════════════════════════

function mapColumns(headers) {
  var map = {
    employeeName: -1,
    reportDate: -1,
    effectiveness: -1,
    timeMgmt: -1,
    ppe: -1,
    incidents: -1,
    extraTasks: -1,
    proactive1: -1,
    proactive2: -1,
    proactive3: -1,
    improvements: -1,
    errorCheck: -1,
    errorCheckCount: -1,
    internalComms: -1,
    externalComms: -1,
    mentoring: -1
  };

  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().trim();

    if (h === 'employee name') map.employeeName = i;
    else if (h === 'report date') map.reportDate = i;
    else if (h.indexOf('effectiveness') >= 0 && h.indexOf('rating') >= 0) map.effectiveness = i;
    else if (h.indexOf('time management') >= 0 && h.indexOf('rating') >= 0) map.timeMgmt = i;
    else if (h.indexOf('ppe') >= 0 && h.indexOf('inspection') >= 0) map.ppe = i;
    else if (h.indexOf('safety incidents') >= 0) map.incidents = i;
    else if (h.indexOf('additional tasks') >= 0 && h.indexOf('ask') >= 0) map.extraTasks = i;
    else if (h === 'proactive task 1') map.proactive1 = i;
    else if (h === 'proactive task 2') map.proactive2 = i;
    else if (h === 'proactive task 3') map.proactive3 = i;
    else if (h.indexOf('process') >= 0 && h.indexOf('improvement') >= 0) map.improvements = i;
    else if (h.indexOf('check') >= 0 && h.indexOf('error') >= 0 && h.indexOf('how many') < 0 && h.indexOf('count') < 0) map.errorCheck = i;
    else if (h.indexOf('how many times') >= 0) map.errorCheckCount = i;
    else if (h.indexOf('internal') >= 0 && h.indexOf('count') >= 0) map.internalComms = i;
    else if (h.indexOf('external') >= 0 && h.indexOf('count') >= 0) map.externalComms = i;
    else if (h.indexOf('mentor') >= 0 && h.indexOf('train') >= 0) map.mentoring = i;
  }

  return map;
}


// ═══════════════════════════════════════════════════════════
//  UTILITY HELPERS
// ═══════════════════════════════════════════════════════════

function findResponseSheet(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf('Form Responses') >= 0) return sheets[i];
  }
  // Fallback: look for 'Timestamp' in row 1 of the first sheet
  if (sheets.length > 0) {
    var firstHeader = String(sheets[0].getRange('A1').getValue()).toLowerCase();
    if (firstHeader === 'timestamp') return sheets[0];
  }
  return null;
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function removeTriggers(funcName) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === funcName) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function round2(n) {
  return Math.round(n * 100) / 100;
}
