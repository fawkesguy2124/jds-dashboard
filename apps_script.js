// ============================================================
// JDS LOGISTICS — GOOGLE APPS SCRIPT
// Paste this ENTIRE file into Google Apps Script editor
// Instructions below in the setup guide
// ============================================================

const SHEET_NAME_JOBS = 'Jobs';
const SHEET_NAME_LOG = 'ActivityLog';
const SHEET_NAME_DASHBOARD = 'Dashboard';

// Column headers for Jobs sheet
const JOB_HEADERS = [
  'ID', 'Job No', 'Client', 'Type', 'Mode', 'BL Number', 'BE Number',
  'Containers', 'Vessel', 'ETA', 'CFS Name', 'CFS Amount', 'Destination',
  'Assigned To', 'Priority', 'HS Code', 'Goods Description', 'RMS Status',
  'BG Required', 'Free Days', 'Detention Risk', 'Stage No', 'Stage Name',
  'Delay Reason', 'Notes', 'CFS Status', 'Pay Status', 'Created At', 'Last Updated', 'Updated By'
];

// Column headers for Activity Log sheet
const LOG_HEADERS = [
  'Log ID', 'Job ID', 'Job No', 'Action', 'Done By', 'Note', 'Timestamp'
];

// ============================================================
// MAIN ENTRY POINT — handles all website requests
// ============================================================
function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // Allow all origins for CORS
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const params = e.parameter || {};
    const postData = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
    const action = params.action || postData.action;

    let result;
    switch (action) {
      case 'getJobs':       result = getJobs(); break;
      case 'getLog':        result = getLog(); break;
      case 'upsertJob':     result = upsertJob(postData.job); break;
      case 'deleteJob':     result = deleteJob(postData.id); break;
      case 'addLog':        result = addLog(postData.log); break;
      case 'updatePayment': result = updatePayment(postData.id, postData.payStatus, postData.cfsStatus, postData.by); break;
      case 'setup':         result = setupSheets(); break;
      default:              result = { error: 'Unknown action: ' + action };
    }

    output.setContent(JSON.stringify({ success: true, data: result }));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.toString() }));
  }

  return output;
}

// ============================================================
// SETUP — creates sheets with correct headers
// Run this ONCE manually from the Apps Script editor
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Setup Jobs sheet
  let jobSheet = ss.getSheetByName(SHEET_NAME_JOBS);
  if (!jobSheet) jobSheet = ss.insertSheet(SHEET_NAME_JOBS);
  if (jobSheet.getLastRow() === 0) {
    jobSheet.getRange(1, 1, 1, JOB_HEADERS.length).setValues([JOB_HEADERS]);
    jobSheet.getRange(1, 1, 1, JOB_HEADERS.length)
      .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    jobSheet.setFrozenRows(1);
    jobSheet.setColumnWidth(1, 120);  // ID
    jobSheet.setColumnWidth(2, 90);   // Job No
    jobSheet.setColumnWidth(3, 130);  // Client
    jobSheet.setColumnWidth(23, 160); // Stage Name
    jobSheet.setColumnWidth(24, 200); // Delay Reason
    jobSheet.setColumnWidth(25, 250); // Notes
  }

  // Setup Activity Log sheet
  let logSheet = ss.getSheetByName(SHEET_NAME_LOG);
  if (!logSheet) logSheet = ss.insertSheet(SHEET_NAME_LOG);
  if (logSheet.getLastRow() === 0) {
    logSheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
    logSheet.getRange(1, 1, 1, LOG_HEADERS.length)
      .setBackground('#0f9d58').setFontColor('#ffffff').setFontWeight('bold');
    logSheet.setFrozenRows(1);
  }

  // Setup Dashboard sheet
  let dashSheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  if (!dashSheet) dashSheet = ss.insertSheet(SHEET_NAME_DASHBOARD);
  updateDashboard();

  return { message: 'Sheets created successfully. You can now use the system.' };
}

// ============================================================
// GET ALL JOBS
// ============================================================
function getJobs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_JOBS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, JOB_HEADERS.length).getValues();
  return data
    .filter(row => row[0] !== '')  // filter empty rows
    .map(row => ({
      id:            row[0],
      jobNo:         row[1],
      client:        row[2],
      type:          row[3],
      mode:          row[4],
      bl:            row[5],
      be:            row[6],
      containers:    row[7],
      vessel:        row[8],
      eta:           row[9],
      cfs:           row[10],
      cfsAmt:        Number(row[11]) || 0,
      destination:   row[12],
      assignedTo:    row[13],
      priority:      row[14],
      hsCode:        row[15],
      goods:         row[16],
      rms:           row[17],
      bgRequired:    row[18],
      freeDays:      Number(row[19]) || 14,
      detentionRisk: row[20],
      stage:         Number(row[21]) || 0,
      stageName:     row[22],
      delayReason:   row[23],
      notes:         row[24],
      cfsStatus:     row[25],
      payStatus:     row[26],
      ts:            new Date(row[27]).getTime() || Date.now(),
      lastUpdated:   row[28],
      updatedBy:     row[29],
    }));
}

// ============================================================
// UPSERT JOB (insert or update)
// ============================================================
function upsertJob(job) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_JOBS);
  const stages = ['Job Created','Documents Received','BL Confirmation','BE Filing','RMS Processing','Assessment','Duty Payment','Examination','OOC Pending','CFS Coordination','Transporter Assigned','Vehicle Out','Delivery Pending','POD Pending','Billing Pending','Payment Collection','Job Closed'];
  const now = new Date().toLocaleString('en-IN', {timeZone: 'Asia/Kolkata'});

  const row = [
    job.id, job.jobNo, job.client, job.type, job.mode,
    job.bl || '', job.be || '', job.containers || '', job.vessel || '', job.eta || '',
    job.cfs || '', job.cfsAmt || 0, job.destination || '',
    job.assignedTo || '', job.priority || 'normal',
    job.hsCode || '', job.goods || '',
    job.rms || 'pend', job.bgRequired || 'no',
    job.freeDays || 14, job.detentionRisk || 'low',
    job.stage || 0, stages[job.stage] || stages[0],
    job.delayReason || '', job.notes || '',
    job.cfsStatus || 'pending', job.payStatus || 'pending',
    job.createdAt || now, now, job.updatedBy || ''
  ];

  // Check if job exists (search by ID)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const existingRow = ids.indexOf(job.id);
    if (existingRow !== -1) {
      // Update existing row
      sheet.getRange(existingRow + 2, 1, 1, row.length).setValues([row]);
      applyRowColor(sheet, existingRow + 2, job.priority, job.stage);
      updateDashboard();
      return { action: 'updated', jobNo: job.jobNo };
    }
  }

  // Insert new row
  sheet.appendRow(row);
  applyRowColor(sheet, sheet.getLastRow(), job.priority, job.stage);
  updateDashboard();
  return { action: 'inserted', jobNo: job.jobNo };
}

// ============================================================
// DELETE JOB
// ============================================================
function deleteJob(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_JOBS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { error: 'No jobs found' };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIndex = ids.indexOf(id);
  if (rowIndex === -1) return { error: 'Job not found' };

  sheet.deleteRow(rowIndex + 2);
  updateDashboard();
  return { action: 'deleted', id };
}

// ============================================================
// GET ACTIVITY LOG
// ============================================================
function getLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, LOG_HEADERS.length).getValues();
  return data
    .filter(row => row[0] !== '')
    .map(row => ({
      id:     row[0],
      jobId:  row[1],
      jobNo:  row[2],
      action: row[3],
      by:     row[4],
      note:   row[5],
      ts:     new Date(row[6]).getTime() || Date.now(),
    }))
    .reverse(); // newest first
}

// ============================================================
// ADD ACTIVITY LOG ENTRY
// ============================================================
function addLog(entry) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_LOG);
  const now = new Date().toLocaleString('en-IN', {timeZone: 'Asia/Kolkata'});
  sheet.appendRow([
    entry.id || ('l' + Date.now()),
    entry.jobId || '',
    entry.jobNo || '',
    entry.action || '',
    entry.by || '',
    entry.note || '',
    now
  ]);
  return { action: 'logged' };
}

// ============================================================
// UPDATE PAYMENT STATUS
// ============================================================
function updatePayment(id, payStatus, cfsStatus, by) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_JOBS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { error: 'No jobs' };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIndex = ids.indexOf(id);
  if (rowIndex === -1) return { error: 'Job not found' };

  const sheetRow = rowIndex + 2;
  sheet.getRange(sheetRow, 26).setValue(cfsStatus); // CFS Status col
  sheet.getRange(sheetRow, 27).setValue(payStatus);  // Pay Status col
  sheet.getRange(sheetRow, 29).setValue(new Date().toLocaleString('en-IN', {timeZone:'Asia/Kolkata'}));
  sheet.getRange(sheetRow, 30).setValue(by);

  addLog({ id: 'l'+Date.now(), jobId: id, jobNo: sheet.getRange(sheetRow, 2).getValue(), action: 'Payment updated: '+payStatus, by, note: 'CFS: '+cfsStatus });
  updateDashboard();
  return { action: 'payment_updated' };
}

// ============================================================
// COLOR ROWS BASED ON PRIORITY / STAGE
// ============================================================
function applyRowColor(sheet, rowNum, priority, stage) {
  let color = '#ffffff';
  if (stage >= 16) color = '#e8f5e9';        // closed — light green
  else if (priority === 'urgent') color = '#fce8e8';  // urgent — light red
  else if (priority === 'high') color = '#fff8e1';    // high — light amber
  sheet.getRange(rowNum, 1, 1, JOB_HEADERS.length).setBackground(color);
}

// ============================================================
// UPDATE DASHBOARD SUMMARY SHEET
// ============================================================
function updateDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dash = ss.getSheetByName(SHEET_NAME_DASHBOARD);
    if (!dash) return;

    const jobs = getJobs();
    const active = jobs.filter(j => j.stage < 16);
    const urgent = jobs.filter(j => j.priority === 'urgent' && j.stage < 16);
    const payPend = jobs.filter(j => j.cfsStatus === 'approved' && j.payStatus === 'pending');
    const totalPay = payPend.reduce((s, j) => s + (j.cfsAmt || 0), 0);

    dash.clearContents();
    dash.getRange('A1').setValue('JDS LOGISTICS — LIVE DASHBOARD SUMMARY');
    dash.getRange('A1').setFontSize(14).setFontWeight('bold').setFontColor('#1a73e8');
    dash.getRange('A2').setValue('Last Updated: ' + new Date().toLocaleString('en-IN', {timeZone:'Asia/Kolkata'}));
    dash.getRange('A2').setFontColor('#888888');

    const summary = [
      ['', ''],
      ['METRIC', 'VALUE'],
      ['Total Jobs', jobs.length],
      ['Active Jobs', active.length],
      ['Urgent Jobs', urgent.length],
      ['In Customs (BE→OOC)', jobs.filter(j=>j.stage>=3&&j.stage<=8).length],
      ['In Transit', jobs.filter(j=>j.stage===11||j.stage===12).length],
      ['Completed/Closed', jobs.filter(j=>j.stage>=16).length],
      ['', ''],
      ['PAYMENTS', ''],
      ['CFS Invoices Approved (Pay Now)', payPend.length],
      ['Total Amount Pending (Rs.)', totalPay],
      ['Total CFS Paid', jobs.filter(j=>j.payStatus==='paid').reduce((s,j)=>s+(j.cfsAmt||0),0)],
      ['', ''],
      ['TEAM WORKLOAD', ''],
    ];

    ['Akshyani','Kiran','Vrn','Mayur','Anil','Updhy','Trupti'].forEach(emp => {
      summary.push([emp + ' — Active Jobs', jobs.filter(j=>j.assignedTo===emp&&j.stage<16).length]);
    });

    dash.getRange(3, 1, summary.length, 2).setValues(summary);
    dash.getRange('A5').setFontWeight('bold').setBackground('#e8f0fe');
    dash.getRange('A14').setFontWeight('bold').setBackground('#e8f0fe');
    dash.getRange('A18').setFontWeight('bold').setBackground('#e8f0fe');
    dash.setColumnWidth(1, 280);
    dash.setColumnWidth(2, 150);
  } catch(e) {
    // Dashboard update is non-critical, ignore errors
  }
}
