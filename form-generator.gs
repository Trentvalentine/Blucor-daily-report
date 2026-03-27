/**
 * BLUCOR Daily Activity Report — Google Form Generator
 *
 * USAGE:
 *   1. Go to https://script.google.com → New project
 *   2. Paste this entire file into Code.gs
 *   3. Run  createBluCorDailyForm()
 *   4. Check the Execution Log for the Form URL and linked Sheet ID
 *
 * The script creates:
 *   - A fully structured Google Form with all 7 sections
 *   - Task blocks with conditional routing (up to 6 tasks)
 *   - Work order blocks with conditional routing (up to 4 WOs)
 *   - Yes/No questions with follow-up detail fields
 *   - Linear scale ratings (1–10)
 *   - Automatic link to a new Google Sheet for responses
 */

function createBluCorDailyForm() {
  var form = FormApp.create('BLUCOR — Employee Daily Activity Report');
  form.setDescription(
    'Complete this form at the end of each work day. Be specific and honest — ' +
    'your responses drive team KPIs and help us all improve.\n\n' +
    'Estimated time: 5–8 minutes.'
  );
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(false);
  form.setProgressBar(true);
  form.setConfirmationMessage('Report submitted. Thank you!');

  // ── Time slot options (shared across task blocks) ──
  var timeOptions = buildTimeOptions();

  // ═══════════════════════════════════════════════════════
  //  SECTION 1 — Employee Info  (default first page)
  // ═══════════════════════════════════════════════════════
  form.addDateItem()
    .setTitle('Report Date')
    .setHelpText('Date this report covers')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Employee Name')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Role / Team');

  // ═══════════════════════════════════════════════════════
  //  SECTION 2 — Today's Work  (up to 6 task blocks)
  // ═══════════════════════════════════════════════════════

  // We need page breaks for routing. Each task block is its own page.
  // Task 1 is always shown. Tasks 2–6 are gated by "Add another task?"

  // ── Task 1 (always shown) ──
  var taskPage1 = form.addPageBreakItem().setTitle("Today's Work — Task 1");
  addTaskBlock(form, 1, timeOptions);

  var moreTask1 = form.addMultipleChoiceItem()
    .setTitle('Add another task?')
    .setHelpText('Select Yes to log Task 2, or No to skip to Work Orders');

  // ── Task 2 ──
  var taskPage2 = form.addPageBreakItem().setTitle("Today's Work — Task 2");
  addTaskBlock(form, 2, timeOptions);

  var moreTask2 = form.addMultipleChoiceItem()
    .setTitle('Add another task?')
    .setHelpText('Select Yes to log Task 3, or No to skip to Work Orders');

  // ── Task 3 ──
  var taskPage3 = form.addPageBreakItem().setTitle("Today's Work — Task 3");
  addTaskBlock(form, 3, timeOptions);

  var moreTask3 = form.addMultipleChoiceItem()
    .setTitle('Add another task?')
    .setHelpText('Select Yes to log Task 4, or No to skip to Work Orders');

  // ── Task 4 ──
  var taskPage4 = form.addPageBreakItem().setTitle("Today's Work — Task 4");
  addTaskBlock(form, 4, timeOptions);

  var moreTask4 = form.addMultipleChoiceItem()
    .setTitle('Add another task?')
    .setHelpText('Select Yes to log Task 5, or No to skip to Work Orders');

  // ── Task 5 ──
  var taskPage5 = form.addPageBreakItem().setTitle("Today's Work — Task 5");
  addTaskBlock(form, 5, timeOptions);

  var moreTask5 = form.addMultipleChoiceItem()
    .setTitle('Add another task?')
    .setHelpText('Select Yes to log Task 6, or No to skip to Work Orders');

  // ── Task 6 (last) ──
  var taskPage6 = form.addPageBreakItem().setTitle("Today's Work — Task 6");
  addTaskBlock(form, 6, timeOptions);

  // ═══════════════════════════════════════════════════════
  //  SECTION 3 — Work Orders
  // ═══════════════════════════════════════════════════════
  var woPage = form.addPageBreakItem().setTitle('Work Orders & Repair Requests');

  form.addTextItem()
    .setTitle('Total Work Orders completed today')
    .setHelpText('Enter a number (0 if none)')
    .setValidation(FormApp.createTextValidation()
      .requireNumber().build());

  // WO 1 (always shown on this page)
  addWoBlock(form, 1);

  var moreWo1 = form.addMultipleChoiceItem()
    .setTitle('Add another work order?')
    .setHelpText('Yes to log WO 2, or No to continue');

  // WO 2
  var woPage2 = form.addPageBreakItem().setTitle('Work Order 2');
  addWoBlock(form, 2);

  var moreWo2 = form.addMultipleChoiceItem()
    .setTitle('Add another work order?');

  // WO 3
  var woPage3 = form.addPageBreakItem().setTitle('Work Order 3');
  addWoBlock(form, 3);

  var moreWo3 = form.addMultipleChoiceItem()
    .setTitle('Add another work order?');

  // WO 4
  var woPage4 = form.addPageBreakItem().setTitle('Work Order 4');
  addWoBlock(form, 4);

  // ═══════════════════════════════════════════════════════
  //  SECTION 4 — Crew & Safety
  // ═══════════════════════════════════════════════════════
  var crewPage = form.addPageBreakItem().setTitle('Crew & Safety Log');

  form.addTextItem()
    .setTitle('Hours worked today')
    .setRequired(true)
    .setValidation(FormApp.createTextValidation()
      .requireNumber().build());

  form.addTextItem()
    .setTitle('Safety incidents (count)')
    .setHelpText('Enter 0 if none')
    .setValidation(FormApp.createTextValidation()
      .requireNumber().build());

  form.addParagraphTextItem()
    .setTitle('Incident / near-miss details')
    .setHelpText('Leave blank if none');

  form.addMultipleChoiceItem()
    .setTitle('PPE & equipment inspection')
    .setChoiceValues(['Done — all good', 'Issues found'])
    .setRequired(true);

  form.addTextItem()
    .setTitle('PPE issue details')
    .setHelpText('Describe any equipment issues (leave blank if Done)');

  // ═══════════════════════════════════════════════════════
  //  SECTION 5 — Internal & External Communiqué
  // ═══════════════════════════════════════════════════════
  var commsPage = form.addPageBreakItem().setTitle('Internal & External Communiqué');

  form.addTextItem()
    .setTitle('Internal communications (count)')
    .setHelpText('Calls, meetings, messages with team/company')
    .setValidation(FormApp.createTextValidation()
      .requireNumber().build());

  form.addTextItem()
    .setTitle('Internal communication details')
    .setHelpText('Who and what was discussed');

  form.addTextItem()
    .setTitle('External communications (count)')
    .setHelpText('Clients, subcontractors, vendors')
    .setValidation(FormApp.createTextValidation()
      .requireNumber().build());

  form.addTextItem()
    .setTitle('External communication details');

  form.addParagraphTextItem()
    .setTitle('Follow-up notes')
    .setHelpText('Pending responses, escalations, action items');

  // ═══════════════════════════════════════════════════════
  //  SECTION 6 — Tomorrow's Priorities
  // ═══════════════════════════════════════════════════════
  var tomorrowPage = form.addPageBreakItem().setTitle("Tomorrow's Plan & Priorities");

  for (var p = 1; p <= 5; p++) {
    var item = form.addTextItem().setTitle('Priority ' + p);
    if (p === 1) item.setRequired(true);
  }

  // ═══════════════════════════════════════════════════════
  //  SECTION 7 — Self-Assessment
  // ═══════════════════════════════════════════════════════
  var assessPage = form.addPageBreakItem().setTitle('Productivity & Self-Assessment');

  form.addMultipleChoiceItem()
    .setTitle('Did you ask for additional tasks when you finished early?')
    .setChoiceValues(['Yes', 'No', 'N/A — stayed busy all day'])
    .setRequired(true);

  form.addTextItem()
    .setTitle('If No, why not?');

  form.addMultipleChoiceItem()
    .setTitle('Did you need to clarify instructions after starting a task?')
    .setChoiceValues(['Yes', 'No'])
    .setRequired(true);

  form.addTextItem()
    .setTitle('Clarification details');

  // Proactive tasks
  form.addSectionHeaderItem()
    .setTitle('Proactive Work')
    .setHelpText('Tasks you initiated without being asked');

  form.addTextItem().setTitle('Proactive task 1');
  form.addTextItem().setTitle('Proactive task 2');
  form.addTextItem().setTitle('Proactive task 3');

  // Effectiveness
  form.addScaleItem()
    .setTitle('Overall effectiveness rating')
    .setBounds(1, 10)
    .setLabels('Poor', 'Excellent')
    .setRequired(true);

  form.addTextItem()
    .setTitle('What drove your effectiveness rating today?');

  // Process improvements
  form.addMultipleChoiceItem()
    .setTitle('Did you identify any process or workflow improvements?')
    .setChoiceValues(['Yes', 'No'])
    .setRequired(true);

  form.addTextItem()
    .setTitle('Improvement details');

  // Mentoring
  form.addMultipleChoiceItem()
    .setTitle('Did you mentor, train, or help a colleague today?')
    .setChoiceValues(['Yes', 'No'])
    .setRequired(true);

  form.addTextItem()
    .setTitle('Mentoring details');

  // Time management
  form.addScaleItem()
    .setTitle('Time management rating')
    .setBounds(1, 10)
    .setLabels('Poor', 'Excellent')
    .setRequired(true);

  form.addTextItem()
    .setTitle('How could you manage time better?');

  // Error checking
  form.addMultipleChoiceItem()
    .setTitle('Did you check your work for errors before finishing?')
    .setChoiceValues(['Yes', 'No'])
    .setRequired(true);

  form.addTextItem()
    .setTitle('How many times did you check?')
    .setValidation(FormApp.createTextValidation()
      .requireNumber().build());

  form.addTextItem()
    .setTitle('Issues found and fixed');

  // ═══════════════════════════════════════════════════════
  //  SECTION 8 — Submission Signature
  // ═══════════════════════════════════════════════════════
  var sigPage = form.addPageBreakItem().setTitle('Submission');

  form.addTextItem()
    .setTitle('Submitted by (your name)')
    .setRequired(true);

  // ═══════════════════════════════════════════════════════
  //  ROUTING — Wire up "Add another?" questions
  // ═══════════════════════════════════════════════════════

  // Task routing: Yes → next task page, No → work orders page
  wireRoute(moreTask1, taskPage2, woPage);
  wireRoute(moreTask2, taskPage3, woPage);
  wireRoute(moreTask3, taskPage4, woPage);
  wireRoute(moreTask4, taskPage5, woPage);
  wireRoute(moreTask5, taskPage6, woPage);

  // Work order routing: Yes → next WO page, No → crew/safety page
  wireRoute(moreWo1, woPage2, crewPage);
  wireRoute(moreWo2, woPage3, crewPage);
  wireRoute(moreWo3, woPage4, crewPage);

  // ═══════════════════════════════════════════════════════
  //  LINK TO GOOGLE SHEET
  // ═══════════════════════════════════════════════════════
  form.setDestination(FormApp.DestinationType.SPREADSHEET,
    SpreadsheetApp.create('BLUCOR Daily Reports — Data').getId());

  // ── Output ──
  Logger.log('═══════════════════════════════════════════');
  Logger.log('  BLUCOR Daily Form created successfully!');
  Logger.log('═══════════════════════════════════════════');
  Logger.log('Form URL (edit):    ' + form.getEditUrl());
  Logger.log('Form URL (submit):  ' + form.getPublishedUrl());
  Logger.log('Linked Sheet ID:    ' + form.getDestinationId());
  Logger.log('Sheet URL:          https://docs.google.com/spreadsheets/d/' + form.getDestinationId());
  Logger.log('═══════════════════════════════════════════');
  Logger.log('NEXT STEP: Open the Sheet and run dashboard-setup.gs');

  return {
    formUrl: form.getPublishedUrl(),
    editUrl: form.getEditUrl(),
    sheetId: form.getDestinationId(),
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + form.getDestinationId()
  };
}


// ─────────────────────────────────────────────────────────
//  WEB APP TRIGGER — visit the deployed URL to run setup
// ─────────────────────────────────────────────────────────
function doGet() {
  var result = createBluCorDailyForm();
  var html = '<h1>BLUCOR Form Created!</h1>' +
    '<p><b>Form URL:</b> <a href="' + result.formUrl + '">' + result.formUrl + '</a></p>' +
    '<p><b>Edit URL:</b> <a href="' + result.editUrl + '">' + result.editUrl + '</a></p>' +
    '<p><b>Sheet URL:</b> <a href="' + result.sheetUrl + '">' + result.sheetUrl + '</a></p>' +
    '<p>You can close this tab now.</p>';
  return HtmlService.createHtmlOutput(html);
}
// ─────────────────────────────────────────────────────────
function addTaskBlock(form, num, timeOptions) {
  form.addListItem()
    .setTitle('Task ' + num + ' — Start Time')
    .setChoiceValues(timeOptions)
    .setRequired(num === 1);

  form.addListItem()
    .setTitle('Task ' + num + ' — End Time')
    .setChoiceValues(timeOptions)
    .setRequired(num === 1);

  form.addParagraphTextItem()
    .setTitle('Task ' + num + ' — Description')
    .setHelpText('What did you do? Be specific — include location, equipment, and outcome.')
    .setRequired(num === 1);
}


// ─────────────────────────────────────────────────────────
//  HELPER: Add a work order block
// ─────────────────────────────────────────────────────────
function addWoBlock(form, num) {
  form.addTextItem()
    .setTitle('WO ' + num + ' — Number');

  form.addTextItem()
    .setTitle('WO ' + num + ' — Location / Asset');

  form.addScaleItem()
    .setTitle('WO ' + num + ' — Completion %')
    .setBounds(0, 10)
    .setLabels('0%', '100%');

  form.addTextItem()
    .setTitle('WO ' + num + ' — Notes');
}


// ─────────────────────────────────────────────────────────
//  HELPER: Wire a Yes/No routing question
// ─────────────────────────────────────────────────────────
function wireRoute(mcItem, yesPage, noPage) {
  mcItem.setChoices([
    mcItem.createChoice('Yes', yesPage),
    mcItem.createChoice('No',  noPage)
  ]);
}


// ─────────────────────────────────────────────────────────
//  HELPER: Build 30-min time slot options (4:00 AM – 4:00 PM)
// ─────────────────────────────────────────────────────────
function buildTimeOptions() {
  var opts = [];
  for (var h = 4; h <= 16; h++) {
    for (var m = 0; m < 60; m += 30) {
      if (h === 16 && m > 0) break; // stop at 4:00 PM
      var hr = h > 12 ? h - 12 : (h === 0 ? 12 : h);
      var ampm = h >= 12 ? 'PM' : 'AM';
      var mm = m < 10 ? '0' + m : '' + m;
      opts.push(hr + ':' + mm + ' ' + ampm);
    }
  }
  return opts;
}
