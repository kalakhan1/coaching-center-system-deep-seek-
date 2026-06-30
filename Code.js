// ==================== GLOBAL CONFIGURATION ====================
const SPREADSHEET_ID = 'take from sheet url';
const CACHE_PREFIX = 'ERP_';
const SESSION_EXPIRY = 30;

// ==================== WEB APP ====================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Coaching Center ERP')
    .setFaviconUrl('https://img.icons8.com/color/48/000000/classroom.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== CACHE SERVICE ====================
function getC(key) {
  try {
    const c = CacheService.getScriptCache();
    const v = c.get(CACHE_PREFIX + key);
    return v;
  } catch(e) { return null; }
}

function setC(key, val, sec = 21600) {
  try {
    CacheService.getScriptCache().put(CACHE_PREFIX + key, val, sec);
  } catch(e) {}
}

function delC(key) {
  try {
    if (key) {
      CacheService.getScriptCache().remove(CACHE_PREFIX + key);
    } else {
      const keys = ['SETTINGS','STUDENTS','TUTORS','FEE_HISTORY','SALARY_HISTORY','ARCHIVE_STUDENTS','ARCHIVE_TUTORS'];
      const c = CacheService.getScriptCache();
      keys.forEach(k => c.remove(CACHE_PREFIX + k));
    }
  } catch(e) {}
}

// ==================== ERROR LOGGING ====================
function logErr(src, desc, err) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('ErrorLogs');
    if (!sheet) {
      sheet = doc.insertSheet('ErrorLogs');
      sheet.getRange(1,1,1,4).setValues([['Timestamp','Source','Description','Error']]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), src, desc, err ? err.toString() : '']);
    lock.releaseLock();
  } catch(e) { console.error(src, desc, err); }
}

// ==================== AUTHENTICATION ====================
function adminLogin(pwd) {
  try {
    const s = getSettings();
    if (!s || !s.adminPassword) {
      createDefaultSettings();
      return { success: false, message: 'System initialized. Default: admin123' };
    }
    if (pwd === s.adminPassword) {
      const token = Utilities.getUuid();
      setC('SESSION_' + token, JSON.stringify({ role: 'admin', time: new Date().toISOString() }), 3600);
      return { success: true, token: token, settings: s };
    }
    return { success: false, message: 'Invalid password' };
  } catch(e) {
    return { success: false, message: 'Login error' };
  }
}

function validateSession(token) {
  if (!token) return false;
  const d = getC('SESSION_' + token);
  if (!d) return false;
  try {
    const session = JSON.parse(d);
    const exp = new Date(session.time);
    exp.setMinutes(exp.getMinutes() + SESSION_EXPIRY);
    if (new Date() > exp) { delC('SESSION_' + token); return false; }
    session.time = new Date().toISOString();
    setC('SESSION_' + token, JSON.stringify(session), 3600);
    return true;
  } catch(e) { return false; }
}

function adminLogout(token) {
  if (token) delC('SESSION_' + token);
  return { success: true };
}

// ==================== SETTINGS ====================
function getSettings() {
  const cached = getC('SETTINGS');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = doc.getSheetByName('Settings');
    if (!sheet) return createDefaultSettings();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return createDefaultSettings();
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) settings[data[i][0]] = data[i][1] || '';
    }
    setC('SETTINGS', JSON.stringify(settings));
    return settings;
  } catch(e) { return {}; }
}

function createDefaultSettings() {
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('Settings');
    if (sheet) sheet.clear();
    else sheet = doc.insertSheet('Settings');
    
    const defaults = [
      ['instituteName', 'Iqbal Coaching Academy'],
      ['adminPassword', 'admin123'],
      ['logoUrl', 'https://img.icons8.com/color/96/000000/classroom.png'],
      ['address', '123 Main Street, Model Town, Lahore'],
      ['phone', '+92-300-1234567'],
      ['email', 'info@coachingcenter.com'],
      ['website', 'www.coachingcenter.com'],
      ['footerText', '© 2024 Coaching Center'],
      ['primaryColor', '#0d6efd'],
      ['secondaryColor', '#6c757d'],
      ['theme', 'light'],
      ['language', 'en'],
      ['qrApiUrl', 'https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=']
    ];
    sheet.getRange(1,1,1,2).setValues([['Key','Value']]);
    sheet.getRange(2,1,defaults.length,2).setValues(defaults);
    sheet.setFrozenRows(1);
    delC('SETTINGS');
    const obj = Object.fromEntries(defaults);
    setC('SETTINGS', JSON.stringify(obj));
    return obj;
  } catch(e) {
    logErr('createDefaultSettings', 'Init error', e);
    return {};
  }
}

function updateSettings(obj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('Settings');
    if (!sheet) { sheet = doc.insertSheet('Settings'); sheet.getRange(1,1,1,2).setValues([['Key','Value']]); }
    const existing = {};
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0]) existing[data[i][0]] = data[i][1]; }
    const merged = { ...existing, ...obj };
    const entries = Object.entries(merged);
    sheet.clear();
    sheet.getRange(1,1,1,2).setValues([['Key','Value']]);
    if (entries.length > 0) sheet.getRange(2,1,entries.length,2).setValues(entries);
    sheet.setFrozenRows(1);
    delC('SETTINGS');
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

// ==================== STUDENT CRUD ====================
function addStudent(data, imgBase64) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('Students');
    if (!sheet) { sheet = doc.insertSheet('Students'); createStudentHeaders(sheet); }
    
    const sid = 'STU-' + Utilities.getUuid().substring(0,8).toUpperCase();
    const adm = 'ADM-' + new Date().getFullYear() + '-' + String(sheet.getLastRow()).padStart(4,'0');
    const img = imgBase64 ? saveImage(sid, imgBase64) : (data.imageUrl || '');
    
    sheet.appendRow([
      sid, adm, data.fullName||'', data.fatherName||'', data.gender||'', data.dob||'',
      data.cnic||'', data.phone||'', data.whatsapp||'', data.email||'', data.address||'',
      data.city||'', data.country||'Pakistan', data.institute||'', data.qualification||'',
      data.emergencyContact||'', data.guardianContact||'', data.course||'', data.class||'',
      data.batch||'', data.slot||'', new Date(), data.fee||0, data.discount||0,
      data.status||'Active', data.notes||'', img
    ]);
    delC('STUDENTS');
    return { success: true, studentId: sid, admissionNo: adm };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateStudent(sid, updates) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Students');
    if (!sheet) return { success: false, message: 'Not found' };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sid) {
        Object.keys(updates).forEach(k => {
          const idx = headers.indexOf(k);
          if (idx !== -1) sheet.getRange(i+1, idx+1).setValue(updates[k]);
        });
        delC('STUDENTS');
        return { success: true };
      }
    }
    return { success: false, message: 'Student not found' };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function deleteStudent(sid) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = doc.getSheetByName('Students');
    if (!sheet) return { success: false, message: 'Not found' };
    const data = sheet.getDataRange().getValues();
    for (let i = data.length-1; i >= 1; i--) {
      if (data[i][0] === sid) {
        // Archive before delete
        archiveRecord('ArchiveStudents', data[i]);
        sheet.deleteRow(i+1);
        delC('STUDENTS');
        return { success: true };
      }
    }
    return { success: false, message: 'Not found' };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getStudents() {
  const cached = getC('STUDENTS');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Students');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx] || '');
      result.push(obj);
    }
    setC('STUDENTS', JSON.stringify(result));
    return result;
  } catch(e) { return []; }
}

function getStudentById(sid) {
  const students = getStudents();
  return students.find(s => s.StudentID === sid) || null;
}

// ==================== TUTOR CRUD ====================
function addTutor(data, imgBase64) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('Tutors');
    if (!sheet) { sheet = doc.insertSheet('Tutors'); createTutorHeaders(sheet); }
    
    const tid = 'TUT-' + Utilities.getUuid().substring(0,8).toUpperCase();
    const img = imgBase64 ? saveImage(tid, imgBase64) : (data.imageUrl || '');
    
    sheet.appendRow([
      tid, data.name||'', data.cnic||'', data.qualification||'', data.experience||'',
      data.subjects||'', data.phone||'', data.whatsapp||'', data.email||'',
      data.address||'', data.city||'', new Date(), data.salaryType||'Monthly',
      data.perLecturePay||0, data.monthlyPackage||0, data.status||'Active', img
    ]);
    delC('TUTORS');
    return { success: true, tutorId: tid };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateTutor(tid, updates) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Tutors');
    if (!sheet) return { success: false };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tid) {
        Object.keys(updates).forEach(k => {
          const idx = headers.indexOf(k);
          if (idx !== -1) sheet.getRange(i+1, idx+1).setValue(updates[k]);
        });
        delC('TUTORS');
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function deleteTutor(tid) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = doc.getSheetByName('Tutors');
    if (!sheet) return { success: false };
    const data = sheet.getDataRange().getValues();
    for (let i = data.length-1; i >= 1; i--) {
      if (data[i][0] === tid) {
        archiveRecord('ArchiveTutors', data[i]);
        sheet.deleteRow(i+1);
        delC('TUTORS');
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getTutors() {
  const cached = getC('TUTORS');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Tutors');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx] || '');
      result.push(obj);
    }
    setC('TUTORS', JSON.stringify(result));
    return result;
  } catch(e) { return []; }
}

// ==================== ARCHIVE ====================
function archiveRecord(sheetName, rowData) {
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName(sheetName);
    if (!sheet) {
      sheet = doc.insertSheet(sheetName);
      if (sheetName === 'ArchiveStudents') createStudentHeaders(sheet);
      else createTutorHeaders(sheet);
    }
    sheet.appendRow(rowData);
  } catch(e) { logErr('archiveRecord', 'Archive error', e); }
}

function getArchivedStudents() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ArchiveStudents');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx] || '');
      result.push(obj);
    }
    return result;
  } catch(e) { return []; }
}

function getArchivedTutors() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ArchiveTutors');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx] || '');
      result.push(obj);
    }
    return result;
  } catch(e) { return []; }
}

// ==================== FEE MANAGEMENT ====================
function addFeePayment(studentId, amount, month, receiptNo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('FeeHistory');
    if (!sheet) {
      sheet = doc.insertSheet('FeeHistory');
      sheet.getRange(1,1,1,5).setValues([['StudentID','Date','Amount','Month','ReceiptNo']]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([studentId, new Date(), amount, month, receiptNo || Utilities.getUuid().substring(0,8)]);
    delC('FEE_HISTORY');
    return { success: true };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getFeeHistory(studentId) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('FeeHistory');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    return data.slice(1).filter(r => r[0] === studentId);
  } catch(e) { return []; }
}

function getAllFeeHistory() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('FeeHistory');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx] || '');
      result.push(obj);
    }
    return result;
  } catch(e) { return []; }
}

// ==================== SALARY MANAGEMENT ====================
function addSalaryPayment(tutorId, month, amount, lectures, deductions) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('SalaryHistory');
    if (!sheet) {
      sheet = doc.insertSheet('SalaryHistory');
      sheet.getRange(1,1,1,7).setValues([['TutorID','Month','Amount','Lectures','Deductions','NetPay','Date']]);
      sheet.setFrozenRows(1);
    }
    const netPay = amount - (deductions || 0);
    sheet.appendRow([tutorId, month, amount, lectures||0, deductions||0, netPay, new Date()]);
    delC('SALARY_HISTORY');
    return { success: true };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getSalaryHistory(tutorId) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SalaryHistory');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    return data.slice(1).filter(r => r[0] === tutorId);
  } catch(e) { return []; }
}

function getAllSalaryHistory() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SalaryHistory');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx] || '');
      result.push(obj);
    }
    return result;
  } catch(e) { return []; }
}

// ==================== IMAGE HANDLING ====================
function saveImage(id, base64Data) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', id + '.png');
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    const parents = DriveApp.getFileById(SPREADSHEET_ID).getParents();
    if (!parents.hasNext()) return '';
    const parentFolder = parents.next();
    let folder;
    const folders = parentFolder.getFoldersByName('ERP_Images');
    if (folders.hasNext()) folder = folders.next();
    else folder = parentFolder.createFolder('ERP_Images');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch(e) { return ''; }
}

// ==================== MAINTENANCE ====================
function resetProject() {
  try {
    delC();
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = doc.getSheetByName('Settings');
    if (sheet) sheet.clear();
    createDefaultSettings();
    return { success: true, message: 'Project reset. Password: admin123' };
  } catch(e) { return { success: false }; }
}

function clearSystemCache() {
  delC();
  return { success: true, message: 'Cache cleared' };
}

function getErrorLogs() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ErrorLogs');
    if (!sheet) return [];
    return sheet.getDataRange().getValues().slice(1);
  } catch(e) { return []; }
}

// ==================== DUMMY DATA ====================
function generateDummyData() {
  try {
    addStudent({ fullName:'Ahmed Khan', fatherName:'Muhammad Khan', gender:'Male', dob:'2005-03-15', cnic:'35202-1234567-1', phone:'0300-1111111', whatsapp:'0300-1111111', email:'ahmed@test.com', address:'123 Main St', city:'Lahore', country:'Pakistan', institute:'Govt High School', qualification:'Matric', emergencyContact:'0300-9999999', guardianContact:'0300-8888888', course:'FSc Pre-Engineering', class:'1st Year', batch:'Morning A', slot:'8:00-10:00 AM', fee:5000, discount:500, status:'Active', notes:'Hardworking' });
    addStudent({ fullName:'Fatima Ali', fatherName:'Ali Raza', gender:'Female', dob:'2006-07-22', cnic:'35202-9876543-2', phone:'0300-2222222', whatsapp:'0300-2222222', email:'fatima@test.com', address:'456 Model Town', city:'Karachi', country:'Pakistan', institute:'Beaconhouse', qualification:'Matric', emergencyContact:'0300-7777777', guardianContact:'0300-6666666', course:'ICS', class:'2nd Year', batch:'Evening B', slot:'5:00-7:00 PM', fee:6000, discount:0, status:'Active' });
    addStudent({ fullName:'Usman Raza', fatherName:'Raza Ahmed', gender:'Male', dob:'2004-11-10', cnic:'35202-5555555-3', phone:'0300-3333333', whatsapp:'0300-3333333', email:'usman@test.com', address:'789 College Rd', city:'Islamabad', country:'Pakistan', institute:'Fazaia College', qualification:'FSc', emergencyContact:'0300-4444444', guardianContact:'0300-5555555', course:'Pre-Medical', class:'1st Year', batch:'Morning B', slot:'10:00-12:00 PM', fee:7000, discount:1000, status:'Active' });
    addTutor({ name:'Abdul Qadir', cnic:'35202-1111111-1', qualification:'M.Sc Math', experience:'5 years', subjects:'Math, Physics', phone:'0300-6666666', whatsapp:'0300-6666666', email:'qadir@test.com', address:'789 College Rd', city:'Lahore', salaryType:'Monthly', perLecturePay:0, monthlyPackage:45000, status:'Active' });
    addTutor({ name:'Sara Ahmed', cnic:'35202-2222222-2', qualification:'M.A English', experience:'3 years', subjects:'English, Urdu', phone:'0300-7777777', whatsapp:'0300-7777777', email:'sara@test.com', address:'123 Garden Town', city:'Lahore', salaryType:'PerLecture', perLecturePay:500, monthlyPackage:0, status:'Active' });
    return { success: true, message: 'Dummy data generated!' };
  } catch(e) { return { success: false }; }
}

// ==================== SHEET HEADERS ====================
function createStudentHeaders(sheet) {
  sheet.getRange(1,1,1,27).setValues([['StudentID','AdmissionNo','FullName','FatherName','Gender','DOB','CNIC','Phone','WhatsApp','Email','Address','City','Country','Institute','Qualification','EmergencyContact','GuardianContact','Course','Class','Batch','Slot','JoiningDate','Fee','Discount','Status','Notes','ImageURL']]);
  sheet.setFrozenRows(1);
}

function createTutorHeaders(sheet) {
  sheet.getRange(1,1,1,17).setValues([['TutorID','Name','CNIC','Qualification','Experience','Subjects','Phone','WhatsApp','Email','Address','City','JoiningDate','SalaryType','PerLecturePay','MonthlyPackage','Status','ImageURL']]);
  sheet.setFrozenRows(1);
}

function onOpen() {
  try { createDefaultSettings(); } catch(e) {}
}
