// ============================================================
// JDS LOGISTICS — API LAYER
// This file connects the website to Google Sheets
// Place in same folder as jose.html and employee.html
// ============================================================

// *** PASTE YOUR GOOGLE APPS SCRIPT WEB APP URL HERE ***
// After deploying Apps Script, you get a URL like:
// https://script.google.com/macros/s/AKfycb.../exec
// Paste it between the quotes below:

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxp4k4AymyTLKUtFL3GQWyEcppzhoFAxE-KUnNpcFk3YsRNzavBMcA_Q9Dk28wagVKbzw/exec';

// ============================================================
// STAGES AND CONSTANTS
// ============================================================
const STAGES = [
  'Job Created','Documents Received','BL Confirmation','BE Filing',
  'RMS Processing','Assessment','Duty Payment','Examination',
  'OOC Pending','CFS Coordination','Transporter Assigned','Vehicle Out',
  'Delivery Pending','POD Pending','Billing Pending','Payment Collection','Job Closed'
];
const EMPLOYEES = ['Akshyani','Kiran','Vrn','Anil','Mayur','Updhy','Trupti','Jose','DPM'];

// ============================================================
// CORE API CALLS
// ============================================================
async function apiCall(action, data = {}) {
  try {
    showLoader(true);
    const payload = { action, ...data };
    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify(payload),
      headers: { 'Content-Type': 'text/plain' } // Apps Script needs text/plain for CORS
    });
    const json = await res.json();
    if (!json.success) throw new Error(json.error || 'API error');
    return json.data;
  } catch (err) {
    console.error('API Error:', err);
    showToast('Connection error: ' + err.message, true);
    throw err;
  } finally {
    showLoader(false);
  }
}

// ============================================================
// DATA FUNCTIONS — called by jose.html and employee.html
// ============================================================
const API = {
  async getJobs() {
    return await apiCall('getJobs');
  },
  async getLog() {
    return await apiCall('getLog');
  },
  async upsertJob(job) {
    const result = await apiCall('upsertJob', { job });
    await API.addLog({
      id: 'l' + Date.now(),
      jobId: job.id, jobNo: job.jobNo,
      action: job._isNew ? 'New job created' : 'Job updated',
      by: job.updatedBy || 'System',
      note: job.goods || ''
    });
    return result;
  },
  async deleteJob(id, jobNo, by) {
    await API.addLog({ id:'l'+Date.now(), jobId:id, jobNo, action:'Job deleted', by, note:'' });
    return await apiCall('deleteJob', { id });
  },
  async addLog(entry) {
    return await apiCall('addLog', { log: entry });
  },
  async updateStage(job, newStage, by, note) {
    const oldStage = job.stage;
    job.stage = newStage;
    job.updatedBy = by;
    job._isNew = false;
    const result = await API.upsertJob(job);
    await API.addLog({
      id: 'l' + Date.now(),
      jobId: job.id, jobNo: job.jobNo,
      action: `Stage: "${STAGES[oldStage]}" → "${STAGES[newStage]}"`,
      by, note
    });
    return result;
  },
  async markPaid(job, by) {
    return await apiCall('updatePayment', {
      id: job.id, payStatus: 'paid', cfsStatus: 'paid', by
    });
  },
  async markApproved(job, by) {
    return await apiCall('updatePayment', {
      id: job.id, payStatus: 'pending', cfsStatus: 'approved', by
    });
  }
};

// ============================================================
// UTILITY
// ============================================================
function timeAgo(ts) {
  const d = Math.floor((Date.now() - ts) / 1000);
  if (d < 60) return 'just now';
  if (d < 3600) return Math.floor(d/60) + 'm ago';
  if (d < 86400) return Math.floor(d/3600) + 'h ago';
  return Math.floor(d/86400) + 'd ago';
}
function daysOld(job) { return Math.floor((Date.now() - job.ts) / 86400000); }
function fmtMoney(n) { return 'Rs.' + Number(n||0).toLocaleString('en-IN'); }
function fmtDate(ts) {
  return new Date(ts).toLocaleString('en-IN', { day:'numeric', month:'short', hour:'2-digit', minute:'2-digit' });
}

function showLoader(show) {
  const el = document.getElementById('loader');
  if (el) el.style.display = show ? 'flex' : 'none';
}
function showToast(msg, isError) {
  const t = document.getElementById('toast');
  if (!t) return;
  t.textContent = msg;
  t.className = 'toast show' + (isError ? ' error' : '');
  setTimeout(() => { t.className = 'toast'; }, 3500);
}
