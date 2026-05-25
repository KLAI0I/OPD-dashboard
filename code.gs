/*******************************************************
 MINA NEW STREET - OPD FLOW | VERSION 6.3 - PERIOD REPORTS API ALLOWLIST FIX
 Google Apps Script backend
 Paste this file as Code.gs
*******************************************************/

const SPREADSHEET_ID = '1o6tJ2uJqzh0d0iqLhj2ztlWwDvSE08Dq6ySW713j34c';

const SHEETS = {
  CLINICS: 'Clinics',
  VISITS: 'Visits',
  DOCTORS: 'Doctors',
  ACCESS: 'Access',
  SETTINGS: 'Settings',
  SUMMARY: 'Summary',
  ARCHIVE: 'Daily_Backup_Archive'
};

const DEFAULT_SETTINGS = {
  hospitalHeader: 'Mina new street - OPD Flow',
  footerText: 'Designed By Mr. Mohamed Ghonim',
  longDurationThreshold: 30,
  refreshSeconds: 3
};

function doGet(e) {
  // Normal Apps Script web app hosting
  if (!e || !e.parameter || !e.parameter.api) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Mina new street - OPD Flow')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // GitHub Pages / external static hosting bridge via JSONP
  const callback = String(e.parameter.callback || 'callback').replace(/[^a-zA-Z0-9_.$]/g, '');
  const fn = String(e.parameter.fn || '');
  const allowed = {
    login: login,
    getInitialData: getInitialData,
    getDashboardData: getDashboardData,
    startPatient: startPatient,
    dischargePatient: dischargePatient,
    backupToday: backupToday,
    manualDailyReset: manualDailyReset,
    exportTodayCsv: exportTodayCsv,
    getReportData: getReportData,
    exportReportCsv: exportReportCsv
  };

  let payload;
  try {
    if (!allowed[fn]) throw new Error('API function not allowed: ' + fn);
    const args = e.parameter.args ? JSON.parse(e.parameter.args) : [];
    payload = { ok: true, data: allowed[fn].apply(null, args) };
  } catch (err) {
    payload = { ok: false, error: err && err.message ? err.message : String(err) };
  }

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sh_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Missing sheet: ' + name);
  return s;
}
function tz_() { return Session.getScriptTimeZone(); }
function now_() { return new Date(); }
function date_(d) { return Utilities.formatDate(new Date(d), tz_(), 'yyyy-MM-dd'); }
function dt_(d) { return Utilities.formatDate(new Date(d), tz_(), 'yyyy-MM-dd HH:mm:ss'); }
function hr_(d) { return Number(Utilities.formatDate(new Date(d), tz_(), 'H')); }
function headers_(sheet) { return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String); }
function indexMap_(headers) { const m = {}; headers.forEach((h,i)=>m[h]=i); return m; }

function rows_(sheetName) {
  const s = sh_(sheetName);
  const values = s.getDataRange().getValues();
  if (values.length < 2) return [];
  const h = values[0].map(String);
  return values.slice(1).filter(r => r.some(c => c !== '')).map((r, i) => {
    const o = {_row: i + 2};
    h.forEach((k, j) => o[k] = r[j]);
    return o;
  });
}

function setRow_(sheetName, rowNumber, object) {
  const s = sh_(sheetName);
  const h = headers_(s);
  const current = s.getRange(rowNumber, 1, 1, h.length).getValues()[0];
  h.forEach((key, i) => {
    if (Object.prototype.hasOwnProperty.call(object, key)) current[i] = object[key];
  });
  s.getRange(rowNumber, 1, 1, h.length).setValues([current]);
}

function appendRow_(sheetName, object) {
  const s = sh_(sheetName);
  const h = headers_(s);
  s.appendRow(h.map(k => Object.prototype.hasOwnProperty.call(object, k) ? object[k] : ''));
}

function getSettings_() {
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  try {
    rows_(SHEETS.SETTINGS).forEach(r => {
      const key = String(r.Setting || '').trim().toLowerCase();
      const val = r.Value;
      if (key === 'hospital header') settings.hospitalHeader = String(val || settings.hospitalHeader);
      if (key === 'footer text') settings.footerText = String(val || settings.footerText);
      if (key === 'long duration threshold minutes') settings.longDurationThreshold = Number(val) || settings.longDurationThreshold;
      if (key === 'auto refresh seconds') settings.refreshSeconds = Number(val) || settings.refreshSeconds;
    });
  } catch (err) {}
  return settings;
}

function login(pin) {
  pin = String(pin || '').trim();
  if (!pin) throw new Error('Enter PIN / password.');

  const users = rows_(SHEETS.ACCESS);
  const user = users.find(u =>
    String(u['PIN/Password Placeholder'] || '').trim() === pin &&
    String(u['Active Status'] || '').toLowerCase().trim() === 'active'
  );
  if (!user) throw new Error('Invalid or inactive PIN. Check Access sheet.');

  return {
    userId: String(user['User ID'] || ''),
    userName: String(user['User Name'] || user['Login User/Email'] || user['Role'] || ''),
    role: String(user['Role'] || ''),
    assignedClinicId: String(user['Assigned Clinic ID'] || ''),
    assignedClinicName: String(user['Assigned Clinic Name'] || ''),
    doctorName: String(user['Doctor Name'] || user['User Name'] || user['Login User/Email'] || 'Doctor')
  };
}

function getInitialData(user) {
  dailyResetIfNeeded_();
  return getDashboardData(user);
}

function getDashboardData(user) {
  return {
    settings: getSettings_(),
    clinics: getClinics_(user),
    stats: getStats_(user),
    serverTime: dt_(now_())
  };
}

function getClinics_(user) {
  const threshold = getSettings_().longDurationThreshold;
  const n = now_();
  let clinics = rows_(SHEETS.CLINICS).map(c => {
    const occupied = String(c['Current Status'] || '').toLowerCase() === 'occupied';
    const st = c['Start Time'] ? new Date(c['Start Time']) : null;
    const mins = occupied && st ? Math.floor((n - st) / 60000) : 0;
    return {
      clinicId: String(c['Clinic ID'] || ''),
      clinicName: String(c['Clinic Name'] || ''),
      status: occupied ? 'Occupied' : 'Vacant',
      assignedDoctor: String(c['Assigned Doctor'] || ''),
      currentDoctor: String(c['Current Doctor'] || ''),
      startTime: st ? dt_(st) : '',
      rawStartTime: st ? st.toISOString() : '',
      currentDurationMin: mins,
      lastVisitId: String(c['Last Visit ID'] || ''),
      occupiedLock: String(c['Occupied Lock'] || 'No'),
      longWarning: occupied && mins >= threshold ? 'Yes' : 'No'
    };
  });

  if (user && String(user.role).toLowerCase() === 'doctor') {
    clinics = clinics.filter(c => String(c.clinicId) === String(user.assignedClinicId));
  }
  return clinics;
}

function startPatient(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const user = payload.user;
    if (String(user.role || '').toLowerCase() === 'admin') throw new Error('Admin is read-only. Start/discharge is doctor-only.');
    const clinicId = String(payload.clinicId || '').trim();
    enforceAccess_(user, clinicId);

    const s = sh_(SHEETS.CLINICS);
    const h = headers_(s);
    const idx = indexMap_(h);
    const values = s.getDataRange().getValues();
    let rowNumber = -1, row = null;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idx['Clinic ID']]) === clinicId) { rowNumber = i + 1; row = values[i]; break; }
    }
    if (!row) throw new Error('Clinic not found.');

    const currentStatus = String(row[idx['Current Status']] || '').toLowerCase();
    const currentLock = String(row[idx['Occupied Lock']] || '').toLowerCase();
    if (currentStatus === 'occupied' || currentLock === 'yes') {
      throw new Error('Already occupied. Discharge current patient first.');
    }

    const start = now_();
    const clinicName = String(row[idx['Clinic Name']] || '');
    const doctorName = String(user.doctorName || user.userName || row[idx['Assigned Doctor']] || 'Doctor').trim();
    const visitId = 'V' + Utilities.formatDate(start, tz_(), 'yyyyMMddHHmmssSSS') + '-' + clinicId;

    const updates = {
      'Current Status': 'Occupied',
      'Current Doctor': doctorName,
      'Patient Number': '',
      'Start Time': start,
      'Current Duration (Min)': 0,
      'Last Visit ID': visitId,
      'Occupied Lock': 'Yes',
      'Long Duration Warning': 'No',
      'Last Reset Date': date_(start)
    };
    h.forEach((key, i) => { if (Object.prototype.hasOwnProperty.call(updates, key)) row[i] = updates[key]; });
    s.getRange(rowNumber, 1, 1, h.length).setValues([row]);

    appendRow_(SHEETS.VISITS, {
      'Visit ID': visitId,
      'Date': date_(start),
      'Patient Number': '',
      'Clinic ID': clinicId,
      'Clinic Name': clinicName,
      'Doctor Name': doctorName,
      'Start Time': start,
      'End Time': '',
      'Duration in Minutes': '',
      'Hour of Visit': hr_(start),
      'Status': 'Open',
      'Started By User': user.userName,
      'Ended By User': '',
      'Role': user.role,
      'Exported/Archived': 'No'
    });

    return {
      ok: true,
      action: 'start',
      clinicId: clinicId,
      clinicName: clinicName,
      doctorName: doctorName,
      visitId: visitId,
      status: 'Occupied',
      startTime: dt_(start),
      rawStartTime: start.toISOString(),
      stats: null,
      serverTime: dt_(now_())
    };
  } finally {
    lock.releaseLock();
  }
}

function dischargePatient(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const user = payload.user;
    if (String(user.role || '').toLowerCase() === 'admin') throw new Error('Admin is read-only. Start/discharge is doctor-only.');
    const clinicId = String(payload.clinicId || '').trim();
    enforceAccess_(user, clinicId);

    const clinics = rows_(SHEETS.CLINICS);
    const clinic = clinics.find(c => String(c['Clinic ID']) === clinicId);
    if (!clinic) throw new Error('Clinic not found.');
    if (String(clinic['Current Status'] || '').toLowerCase() !== 'occupied') throw new Error('Clinic is already vacant.');

    const visitId = String(clinic['Last Visit ID'] || '');
    const start = clinic['Start Time'] ? new Date(clinic['Start Time']) : now_();
    const end = now_();
    const duration = Math.max(0, Math.ceil((end - start) / 60000));

    const visits = rows_(SHEETS.VISITS);
    const visit = visits.find(v => String(v['Visit ID']) === visitId && String(v.Status || '').toLowerCase() === 'open');
    if (!visit) throw new Error('Open visit record not found.');

    setRow_(SHEETS.VISITS, visit._row, {
      'End Time': end,
      'Duration in Minutes': duration,
      'Status': 'Completed',
      'Ended By User': user.userName
    });

    setRow_(SHEETS.CLINICS, clinic._row, {
      'Current Status': 'Vacant',
      'Current Doctor': '',
      'Patient Number': '',
      'Start Time': '',
      'Current Duration (Min)': '',
      'Last Visit ID': '',
      'Occupied Lock': 'No',
      'Long Duration Warning': 'No',
      'Last Reset Date': date_(end)
    });

    return {
      ok: true,
      action: 'discharge',
      clinicId: clinicId,
      visitId: visitId,
      duration: duration,
      status: 'Vacant',
      stats: null,
      serverTime: dt_(now_())
    };
  } finally {
    lock.releaseLock();
  }
}

function enforceAccess_(user, clinicId) {
  if (!user) throw new Error('Session expired. Login again.');
  if (!clinicId) throw new Error('Missing clinic.');
  if (String(user.role || '').toLowerCase() === 'doctor' && String(user.assignedClinicId) !== String(clinicId)) {
    throw new Error('Access denied. Doctor can control only assigned clinic.');
  }
}

function getStats_(user) {
  const today = date_(now_());
  let visits = rows_(SHEETS.VISITS).filter(v => String(v.Date) === today || (v.Date instanceof Date && date_(v.Date) === today));
  if (user && String(user.role).toLowerCase() === 'doctor') {
    visits = visits.filter(v => String(v['Clinic ID']) === String(user.assignedClinicId));
  }

  const clinics = getClinics_(user);
  const occupied = clinics.filter(c => c.status === 'Occupied').length;
  const vacant = clinics.length - occupied;
  const completed = visits.filter(v => String(v.Status || '').toLowerCase() === 'completed');
  const durations = completed.map(v => Number(v['Duration in Minutes'])).filter(n => !isNaN(n));
  const avg = durations.length ? Math.round((durations.reduce((a,b)=>a+b,0) / durations.length) * 10) / 10 : 0;

  const byHour = {}, byClinic = {}, byDoctor = {};
  visits.forEach(v => {
    const h = String(v['Hour of Visit'] || '');
    byHour[h] = (byHour[h] || 0) + 1;
    const c = String(v['Clinic Name'] || 'Unknown');
    byClinic[c] = (byClinic[c] || 0) + 1;
    const d = String(v['Doctor Name'] || 'Doctor');
    if (!byDoctor[d]) byDoctor[d] = {doctor:d, visits:0, completed:0, totalDuration:0, avg:0};
    byDoctor[d].visits++;
    const dur = Number(v['Duration in Minutes']);
    if (!isNaN(dur) && String(v.Status || '').toLowerCase() === 'completed') {
      byDoctor[d].completed++;
      byDoctor[d].totalDuration += dur;
    }
  });
  Object.keys(byDoctor).forEach(k => byDoctor[k].avg = byDoctor[k].completed ? Math.round((byDoctor[k].totalDuration / byDoctor[k].completed) * 10) / 10 : 0);

  return {
    today: today,
    totalVisits: visits.length,
    completedVisits: completed.length,
    openVisits: visits.filter(v => String(v.Status || '').toLowerCase() === 'open').length,
    occupiedClinics: occupied,
    vacantClinics: vacant,
    avgDuration: avg,
    byHour: byHour,
    byClinic: byClinic,
    byDoctor: Object.keys(byDoctor).map(k => byDoctor[k]),
    visits: visits.map(v => ({
      visitId: v['Visit ID'],
      clinicId: v['Clinic ID'],
      clinicName: v['Clinic Name'],
      doctorName: v['Doctor Name'],
      startTime: v['Start Time'] ? dt_(v['Start Time']) : '',
      endTime: v['End Time'] ? dt_(v['End Time']) : '',
      duration: v['Duration in Minutes'],
      status: v.Status
    }))
  };
}

function dailyResetIfNeeded_() {
  const today = date_(now_());
  const clinics = rows_(SHEETS.CLINICS);
  clinics.forEach(c => {
    const last = String(c['Last Reset Date'] || '');
    const occupied = String(c['Current Status'] || '').toLowerCase() === 'occupied';
    // Do not silently close an active visit; reset only vacant/empty old rows.
    if (last !== today && !occupied) {
      setRow_(SHEETS.CLINICS, c._row, {
        'Current Status': 'Vacant',
        'Current Doctor': '',
        'Patient Number': '',
        'Start Time': '',
        'Current Duration (Min)': '',
        'Last Visit ID': '',
        'Occupied Lock': 'No',
        'Long Duration Warning': 'No',
        'Last Reset Date': today
      });
    }
  });
}

function manualDailyReset(user) {
  if (String(user.role || '').toLowerCase() !== 'admin') throw new Error('Admin only.');
  const today = date_(now_());
  rows_(SHEETS.CLINICS).forEach(c => setRow_(SHEETS.CLINICS, c._row, {
    'Current Status': 'Vacant',
    'Current Doctor': '',
    'Patient Number': '',
    'Start Time': '',
    'Current Duration (Min)': '',
    'Last Visit ID': '',
    'Occupied Lock': 'No',
    'Long Duration Warning': 'No',
    'Last Reset Date': today
  }));
  return getDashboardData(user);
}

function backupToday(user) {
  if (String(user.role || '').toLowerCase() !== 'admin') throw new Error('Admin only.');
  const today = date_(now_());
  const visits = rows_(SHEETS.VISITS).filter(v => String(v.Date) === today && String(v.Status).toLowerCase() === 'completed');
  const existing = rows_(SHEETS.ARCHIVE).map(a => String(a['Visit ID']));
  let count = 0;
  visits.forEach(v => {
    if (!existing.includes(String(v['Visit ID']))) {
      appendRow_(SHEETS.ARCHIVE, {
        'Archive ID': 'A' + Utilities.formatDate(now_(), tz_(), 'yyyyMMddHHmmss') + '-' + count,
        'Backup Date': dt_(now_()),
        'Visit ID': v['Visit ID'],
        'Date': v.Date,
        'Patient Number': '',
        'Clinic ID': v['Clinic ID'],
        'Clinic Name': v['Clinic Name'],
        'Doctor Name': v['Doctor Name'],
        'Start Time': v['Start Time'],
        'End Time': v['End Time'],
        'Duration in Minutes': v['Duration in Minutes'],
        'Hour of Visit': v['Hour of Visit'],
        'Status': v.Status,
        'Started By User': v['Started By User'],
        'Ended By User': v['Ended By User'],
        'Role': v.Role,
        'Exported/Archived': 'Yes',
        'Backup Notes': 'Daily manual backup'
      });
      count++;
    }
  });
  return {message: count + ' completed visits copied to Daily_Backup_Archive.'};
}

function exportTodayCsv(user) {
  const data = getStats_(user).visits;
  const headers = ['Visit ID','Clinic ID','Clinic Name','Doctor Name','Start Time','End Time','Duration in Minutes','Status'];
  const lines = [headers.join(',')];
  data.forEach(v => lines.push([v.visitId,v.clinicId,v.clinicName,v.doctorName,v.startTime,v.endTime,v.duration,v.status]
    .map(x => '"' + String(x || '').replace(/"/g, '""') + '"').join(',')));
  return lines.join('\n');
}

/*******************************************************
 RETROSPECTIVE REPORTS: specific date / date period
*******************************************************/
function parseDateStart_(value) {
  if (!value) return null;
  if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0, 0);
  const s = String(value).trim();
  if (!s) return null;
  const parts = s.split('-').map(Number);
  if (parts.length >= 3 && parts.every(n => !isNaN(n))) return new Date(parts[0], parts[1] - 1, parts[2], 0, 0, 0, 0);
  const d = new Date(s);
  if (isNaN(d.getTime())) throw new Error('Invalid start date: ' + value);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
}

function parseDateEnd_(value) {
  const d = parseDateStart_(value);
  if (!d) return null;
  d.setHours(23, 59, 59, 999);
  return d;
}

function visitDate_(v) {
  if (v.Date instanceof Date) return parseDateStart_(v.Date);
  if (v.Date) return parseDateStart_(v.Date);
  if (v['Start Time']) return parseDateStart_(v['Start Time']);
  return null;
}

function filterVisitsByPeriod_(user, startDate, endDate) {
  const start = parseDateStart_(startDate) || parseDateStart_(date_(now_()));
  const end = parseDateEnd_(endDate || startDate) || parseDateEnd_(date_(now_()));
  if (start.getTime() > end.getTime()) throw new Error('Start date cannot be after end date.');

  let visits = rows_(SHEETS.VISITS).filter(v => {
    const d = visitDate_(v);
    if (!d) return false;
    return d.getTime() >= start.getTime() && d.getTime() <= end.getTime();
  });

  if (user && String(user.role).toLowerCase() === 'doctor') {
    visits = visits.filter(v => String(v['Clinic ID']) === String(user.assignedClinicId));
  }
  return {start:start, end:end, visits:visits};
}

function buildReportStats_(user, startDate, endDate) {
  const filtered = filterVisitsByPeriod_(user, startDate, endDate);
  const visits = filtered.visits;
  const completed = visits.filter(v => String(v.Status || '').toLowerCase() === 'completed');
  const open = visits.filter(v => String(v.Status || '').toLowerCase() === 'open');
  const durations = completed.map(v => Number(v['Duration in Minutes'])).filter(n => !isNaN(n));
  const totalDuration = durations.reduce((a,b)=>a+b,0);
  const avgDuration = durations.length ? Math.round((totalDuration / durations.length) * 10) / 10 : 0;

  const byHour = {}, byClinic = {}, byDoctor = {}, byStatus = {}, byDate = {};
  visits.forEach(v => {
    const dObj = visitDate_(v);
    const day = dObj ? date_(dObj) : String(v.Date || 'Unknown');
    const hour = String(v['Hour of Visit'] || (v['Start Time'] ? hr_(v['Start Time']) : 'Unknown'));
    const clinic = String(v['Clinic Name'] || 'Unknown');
    const doctor = String(v['Doctor Name'] || 'Doctor');
    const status = String(v.Status || 'Unknown');
    byDate[day] = (byDate[day] || 0) + 1;
    byHour[hour] = (byHour[hour] || 0) + 1;
    byClinic[clinic] = (byClinic[clinic] || 0) + 1;
    byStatus[status] = (byStatus[status] || 0) + 1;
    if (!byDoctor[doctor]) byDoctor[doctor] = {doctor:doctor, visits:0, completed:0, totalDuration:0, avg:0};
    byDoctor[doctor].visits++;
    const dur = Number(v['Duration in Minutes']);
    if (!isNaN(dur) && String(status).toLowerCase() === 'completed') {
      byDoctor[doctor].completed++;
      byDoctor[doctor].totalDuration += dur;
    }
  });
  Object.keys(byDoctor).forEach(k => byDoctor[k].avg = byDoctor[k].completed ? Math.round((byDoctor[k].totalDuration / byDoctor[k].completed) * 10) / 10 : 0);

  return {
    startDate: date_(filtered.start),
    endDate: date_(filtered.end),
    totalVisits: visits.length,
    completedVisits: completed.length,
    openVisits: open.length,
    avgDuration: avgDuration,
    totalDuration: totalDuration,
    byDate: byDate,
    byHour: byHour,
    byClinic: byClinic,
    byDoctor: Object.keys(byDoctor).map(k => byDoctor[k]),
    byStatus: byStatus,
    visits: visits.map(v => ({
      visitId: v['Visit ID'],
      date: v.Date instanceof Date ? date_(v.Date) : String(v.Date || ''),
      clinicId: v['Clinic ID'],
      clinicName: v['Clinic Name'],
      doctorName: v['Doctor Name'],
      startTime: v['Start Time'] ? dt_(v['Start Time']) : '',
      endTime: v['End Time'] ? dt_(v['End Time']) : '',
      duration: v['Duration in Minutes'],
      hour: v['Hour of Visit'],
      status: v.Status
    }))
  };
}

function getReportData(user, startDate, endDate) {
  if (String(user.role || '').toLowerCase() !== 'admin') throw new Error('Admin only.');
  return buildReportStats_(user, startDate, endDate);
}

function exportReportCsv(user, startDate, endDate) {
  if (String(user.role || '').toLowerCase() !== 'admin') throw new Error('Admin only.');
  const report = buildReportStats_(user, startDate, endDate);
  const headers = ['Visit ID','Date','Clinic ID','Clinic Name','Doctor Name','Start Time','End Time','Duration in Minutes','Hour of Visit','Status'];
  const lines = [headers.join(',')];
  report.visits.forEach(v => lines.push([v.visitId,v.date,v.clinicId,v.clinicName,v.doctorName,v.startTime,v.endTime,v.duration,v.hour,v.status]
    .map(x => '"' + String(x || '').replace(/"/g, '""') + '"').join(',')));
  return lines.join('\n');
}

