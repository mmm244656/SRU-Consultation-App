// ══════════════════════════════════════════════════════
//  CONFIGURATION
// ══════════════════════════════════════════════════════
const SHEET_ID             = '14ev6msUxlbJeP3XU3mL9YG9iw-eMPaCJDNGrwRUrkFE';
const ADMIN_EMAIL          = 'admin@rex.com';
const ADMIN_PASSWORD       = 'rex123123';
const STUDENT_EMAIL_DOMAIN = '@student.sru.ac.th';
const TEST_EMAIL           = 'test@student.sru.ac.th';
const TEST_PASSWORD        = 'test123';

// ══════════════════════════════════════════════════════
//  ENTRY POINTS
// ══════════════════════════════════════════════════════
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter);
  }
  return ContentService.createTextOutput('SRU Portal Backend running.').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    return handleApiRequest(params);
  } catch(err) {
    return jsonOut({ success: false, message: 'Bad request: ' + err.message });
  }
}

// ══════════════════════════════════════════════════════
//  API ROUTER
// ══════════════════════════════════════════════════════
function handleApiRequest(p) {
  var result;
  try {
    switch (p.action) {
      case 'loginStudent':        result = loginStudent(p.email, p.password);                      break;
      case 'loginAdmin':          result = loginAdmin(p.email, p.password);                        break;
      case 'registerStudent':     result = registerStudent(p.email, p.name, p.password);           break;
      case 'changePassword':      result = changePassword(p.email, p.oldPassword, p.newPassword);  break;
      case 'submitRequest':       result = submitRequest(p.studentEmail, p.category, p.details);   break;
      case 'getMyRequests':       result = getMyRequests(p.studentEmail);                          break;
      case 'getAllRequests':       result = getAllRequests();                                        break;
      case 'getStats':            result = getStats();                                              break;
      case 'updateRequestStatus': result = updateRequestStatus(Number(p.rowIndex), p.newStatus);   break;
      case 'replyToRequest':      result = replyToRequest(Number(p.rowIndex), p.replyText);        break;
      case 'deleteRequest':       result = deleteRequest(Number(p.rowIndex));                      break;
      default: result = { success: false, message: 'Unknown action: ' + p.action };
    }
  } catch(err) {
    result = { success: false, message: 'Server error: ' + err.message };
  }
  return jsonOut(result);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════
//  SHEET HELPERS
//  Users: A=email  B=name  C=role  D=password  E=created_at  F=verified
// ══════════════════════════════════════════════════════
function getUsersSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['email','name','role','password','created_at','verified']);
  }
  return sheet;
}

function getRequestsSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Requests');
  if (!sheet) {
    sheet = ss.insertSheet('Requests');
    sheet.appendRow(['id','student_email','category','details','status','reply','timestamp']);
  }
  return sheet;
}

function getStudentName(email) {
  try {
    var data = getUsersSheet().getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === email.toLowerCase()) return data[i][1];
    }
  } catch(e) {}
  return email.split('@')[0];
}

// ══════════════════════════════════════════════════════
//  AUTH
// ══════════════════════════════════════════════════════
function isValidStudentEmail(email) {
  return email.toLowerCase().endsWith(STUDENT_EMAIL_DOMAIN);
}

// ── REGISTER ──
function registerStudent(email, name, password) {
  email = email.toLowerCase().trim();

  if (!isValidStudentEmail(email))
    return { success: false, message: 'Email must end with ' + STUDENT_EMAIL_DOMAIN };
  if (!name || !password)
    return { success: false, message: 'All fields are required' };
  if (password.length < 6)
    return { success: false, message: 'Password must be at least 6 characters' };

  // Block test email from registering (it already exists as hardcoded)
  if (email === TEST_EMAIL)
    return { success: false, message: 'This email is reserved. Please use a different email.' };

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email)
      return { success: false, message: 'Email already registered. Please sign in.' };
  }

  sheet.appendRow([email, name, 'student', password, new Date(), 'verified']);
  return { success: true, user: { email: email, name: name, role: 'student' } };
}

// ── LOGIN STUDENT ──
function loginStudent(email, password) {
  email = email.toLowerCase().trim();

  if (!isValidStudentEmail(email))
    return { success: false, message: 'Email must end with ' + STUDENT_EMAIL_DOMAIN };

  // Test account shortcut
  if (email === TEST_EMAIL) {
    if (password === TEST_PASSWORD)
      return { success: true, user: { email: email, name: 'Test Student', role: 'student' } };
    return { success: false, message: 'Incorrect password' };
  }

  var data = getUsersSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email && data[i][2] === 'student') {
      if (data[i][3].toString() !== password.toString())
        return { success: false, message: 'Incorrect password' };
      return { success: true, user: { email: data[i][0], name: data[i][1], role: 'student' } };
    }
  }
  return { success: false, message: 'No account found. Please register first.' };
}

// ── LOGIN ADMIN ──
function loginAdmin(email, password) {
  email = email.toLowerCase().trim();
  if (email === ADMIN_EMAIL && password === ADMIN_PASSWORD)
    return { success: true, user: { email: ADMIN_EMAIL, name: 'Administrator', role: 'admin' } };
  return { success: false, message: 'Invalid admin credentials' };
}

// ── CHANGE PASSWORD ──
function changePassword(email, oldPassword, newPassword) {
  email = email.toLowerCase().trim();
  if (!oldPassword || !newPassword)   return { success: false, message: 'Please fill all fields' };
  if (newPassword.length < 6)         return { success: false, message: 'New password must be at least 6 characters' };
  if (oldPassword === newPassword)    return { success: false, message: 'New password must be different' };

  if (email === TEST_EMAIL) {
    if (oldPassword !== TEST_PASSWORD) return { success: false, message: 'Current password is incorrect' };
    return { success: true, message: 'Password changed successfully!' };
  }

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email && data[i][2] === 'student') {
      if (data[i][3].toString() !== oldPassword.toString())
        return { success: false, message: 'Current password is incorrect' };
      sheet.getRange(i + 1, 4).setValue(newPassword);
      return { success: true, message: 'Password changed successfully!' };
    }
  }
  return { success: false, message: 'Account not found' };
}

// ══════════════════════════════════════════════════════
//  REQUESTS
// ══════════════════════════════════════════════════════
function submitRequest(studentEmail, category, details) {
  if (!details || details.trim() === '') return { success: false, message: 'Please enter details' };
  getRequestsSheet().appendRow([new Date().getTime(), studentEmail, category, details, 'pending', '', new Date()]);
  return { success: true };
}

function getMyRequests(studentEmail) {
  var data = getRequestsSheet().getDataRange().getValues();
  var out  = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().toLowerCase() === studentEmail.toLowerCase()) {
      out.push({ id: data[i][0], category: data[i][2], details: data[i][3], status: data[i][4]||'pending', reply: data[i][5]||'', timestamp: data[i][6] ? new Date(data[i][6]).toLocaleString() : '' });
    }
  }
  return out.reverse();
}

function getAllRequests() {
  var data = getRequestsSheet().getDataRange().getValues();
  var out  = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var email = data[i][1] ? data[i][1].toString() : '';
      out.push({ rowIndex: i+1, id: data[i][0], studentEmail: email, studentName: getStudentName(email), category: data[i][2]||'', details: data[i][3]||'', status: data[i][4]||'pending', reply: data[i][5]||'', timestamp: data[i][6] ? new Date(data[i][6]).toLocaleString() : '' });
    }
  }
  return out.reverse();
}

function getStats() {
  var data = getRequestsSheet().getDataRange().getValues();
  var total=0, pending=0, completed=0, rejected=0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    total++;
    var s = (data[i][4]||'pending').toString().toLowerCase();
    if (s==='pending') pending++;
    else if (s==='completed') completed++;
    else if (s==='rejected') rejected++;
  }
  return { total: total, pending: pending, completed: completed, rejected: rejected };
}

function updateRequestStatus(rowIndex, newStatus) {
  try { getRequestsSheet().getRange(rowIndex, 5).setValue(newStatus); return { success: true }; }
  catch(e) { return { success: false, message: e.message }; }
}

function replyToRequest(rowIndex, replyText) {
  if (!replyText || replyText.trim() === '') return { success: false, message: 'Reply cannot be empty' };
  try {
    var sheet = getRequestsSheet();
    sheet.getRange(rowIndex, 6).setValue(replyText);
    sheet.getRange(rowIndex, 5).setValue('completed');
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteRequest(rowIndex) {
  try { getRequestsSheet().deleteRow(rowIndex); return { success: true }; }
  catch(e) { return { success: false, message: e.message }; }
}
