/**
 * BLUCOR Daily Activity Report — Google Apps Script Backend
 * 
 * SETUP:
 * 1. Go to https://script.google.com and create a new project
 * 2. Paste this entire file into Code.gs (replacing any existing code)
 * 3. Click Deploy > New deployment
 * 4. Type = "Web app"
 * 5. Execute as = "Me"
 * 6. Who has access = "Anyone"
 * 7. Click Deploy and copy the web app URL
 * 8. Paste that URL into the SCRIPT_URL constant in index.html
 *
 * The script auto-creates a "Daily Reports" sheet on first submission.
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Daily Reports');

    // Create sheet with headers if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Daily Reports');
      sheet.appendRow([
        'Timestamp',
        'Report Date',
        'Employee Name',
        'Role/Team',
        'Tasks (summary)',
        'Total Hours',
        'Work Orders (summary)',
        'Hours Worked',
        'Safety Incidents',
        'Safety Details',
        'PPE Check',
        'Calls/Meetings',
        'Change Orders',
        'Info Requests',
        'Client Details',
        'Challenges',
        'Tomorrow Focus',
        'Asked for Extra Tasks',
        'Clarified Instructions',
        'Proactive Tasks',
        'Effectiveness Rating',
        'Effectiveness Reason',
        'Process Improvements',
        'Mentored Colleague',
        'Time Mgmt Rating',
        'Time Mgmt Improve',
        'Checked for Errors',
        'Error Check Count',
        'Issues Found/Fixed',
        'Remarks',
        'Submitted By',
        'Submit Date/Time',
        'Full JSON'
      ]);
      // Bold + freeze header row
      sheet.getRange(1, 1, 1, 33).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // ── Build task summary + total hours ──
    var taskLines = [];
    var totalHrs = 0;
    (data.tasks || []).forEach(function(t) {
      if (t.description || t.notes || t.hours) {
        taskLines.push(t.description + ' (' + (t.hours || '0') + 'h)');
        totalHrs += parseFloat(t.hours) || 0;
      }
    });

    // ── Build work orders summary ──
    var woLines = [];
    (data.workOrders || []).forEach(function(wo) {
      if (wo.number || wo.location || wo.percent || wo.notes) {
        woLines.push('#' + wo.number + ' — ' + wo.location + ' (' + wo.percent + '%)');
      }
    });

    // ── Build bullet lists ──
    var challenges = (data.challenges || []).filter(Boolean).join('\n');
    var tomorrowFocus = (data.tomorrowFocus || []).filter(Boolean).join('\n');
    var proactiveTasks = (data.proactiveTasks || []).filter(Boolean).join('\n');
    var remarks = (data.remarks || []).filter(Boolean).join('\n');

    // ── Append the row ──
    sheet.appendRow([
      new Date(),                                     // Timestamp
      data.reportDate || '',                          // Report Date
      data.employeeName || '',                        // Employee Name
      data.roleTeam || '',                            // Role/Team
      taskLines.join('\n'),                           // Tasks summary
      totalHrs,                                       // Total Hours
      woLines.join('\n'),                             // Work Orders summary
      data.crew.hoursWorked || '',                    // Hours Worked
      data.crew.safetyIncidents || '',                // Safety Incidents
      data.crew.safetyDetails || '',                  // Safety Details
      data.crew.ppeCheck || '',                       // PPE Check
      data.client.callsMeetings || '',                // Calls/Meetings
      data.client.changeOrders || '',                 // Change Orders
      data.client.infoRequests || '',                 // Info Requests
      data.client.details || '',                      // Client Details
      challenges,                                     // Challenges
      tomorrowFocus,                                  // Tomorrow Focus
      data.assessment.askedExtraTasks || '',          // Asked for Extra Tasks
      data.assessment.clarifiedInstructions || '',    // Clarified Instructions
      proactiveTasks,                                 // Proactive Tasks
      data.assessment.effectivenessRating || '',      // Effectiveness Rating
      data.assessment.effectivenessReason || '',      // Effectiveness Reason
      data.assessment.processImprovements || '',      // Process Improvements
      data.assessment.mentoredColleague || '',        // Mentored Colleague
      data.assessment.timeMgmtRating || '',           // Time Mgmt Rating
      data.assessment.timeMgmtImprove || '',          // Time Mgmt Improve
      data.assessment.checkedErrors || '',            // Checked for Errors
      data.assessment.errorCheckCount || '',          // Error Check Count
      data.assessment.issuesFoundFixed || '',         // Issues Found/Fixed
      remarks,                                        // Remarks
      data.submittedBy || '',                         // Submitted By
      data.submitDateTime || '',                      // Submit Date/Time
      JSON.stringify(data)                            // Full JSON backup
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'Report saved' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Required for CORS preflight (though Apps Script handles this automatically for web apps)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Blucor Daily Report API is live.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
