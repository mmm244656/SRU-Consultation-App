// ══════════════════════════════════════════════════════
//  CONFIGURATION — edit these values
// ══════════════════════════════════════════════════════
const SHEET_ID             = '14ev6msUxlbJeP3XU3mL9YG9iw-eMPaCJDNGrwRUrkFE';
const ADMIN_EMAIL          = 'admin@rex.com';
const ADMIN_PASSWORD       = 'rex123123';
const STUDENT_EMAIL_DOMAIN = '@student.sru.ac.th';
const SCRIPT_URL           = 'https://script.google.com/macros/s/AKfycbzaaxPomO468u6N2uCo7AUUiyzfKzBRsTkCx39nHTdpLRS_1cZxV93ZVOC5B6ONNOQY/exec';
const TEST_EMAIL           = 'test@student.sru.ac.th';

// ══════════════════════════════════════════════════════
//  ENTRY POINTS
// ══════════════════════════════════════════════════════

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter);
  }
  // Handle verification link clicked directly in browser
  if (e && e.parameter && e.parameter.token) {
    var result = verifyEmail(e.parameter.token);
    var html = result.success
      ? '<h2 style="font-family:sans-serif;color:green;">✅ Email Verified!</h2><p style="font-family:sans-serif;">Your account is now active. You can close this tab and log in.</p>'
      : '<h2 style="font-family:sans-serif;color:red;">❌ Verification Failed</h2><p style="font-family:sans-serif;">' + result.message + '</p>';
    return HtmlService.createHtmlOutput(html);
  }
  return ContentService
    .createTextOutput('Apps Script backend is running.')
    .setMimeType(ContentService.MimeType.TEXT);
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
  var action = p.action;
  var result;
  try {
    switch (action) {
      case 'loginStudent':
        result = loginStudent(p.email, p.password); break;
      case 'loginAdmin':
        result = loginAdmin(p.email, p.password); break;
      case 'registerStudent':
        result = registerStudent(p.email, p.name, p.password); break;
      case 'verifyEmail':
        result = verifyEmail(p.token); break;
      case 'forgotPassword':
        result = forgotPassword(p.email); break;
      case 'changePassword':
        result = changePassword(p.email, p.oldPassword, p.newPassword); break;
      case 'submitRequest':
        result = submitRequest(p.studentEmail, p.category, p.details); break;
      case 'getMyRequests':
        result = getMyRequests(p.studentEmail); break;
      case 'getAllRequests':
        result = getAllRequests(); break;
      case 'getStats':
        result = getStats(); break;
      case 'updateRequestStatus':
        result = updateRequestStatus(Number(p.rowIndex), p.newStatus); break;
      case 'replyToRequest':
        result = replyToRequest(Number(p.rowIndex), p.replyText); break;
      case 'deleteRequest':
        result = deleteRequest(Number(p.rowIndex)); break;
      default:
        result = { success: false, message: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { success: false, message: 'Server error: ' + err.message };
  }
  return jsonOut(result);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════
//  SHEET HELPERS
// ══════════════════════════════════════════════════════
function getUsersSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    // Column F (index 5) = verified status: 'verified' or a UUID token
    sheet.appendRow(['email', 'name', 'role', 'password', 'created_at', 'verified']);
  }
  return sheet;
}

function getRequestsSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Requests');
  if (!sheet) {
    sheet = ss.insertSheet('Requests');
    sheet.appendRow(['id', 'student_email', 'category', 'details', 'status', 'reply', 'timestamp']);
  }
  return sheet;
}

function getStudentName(email) {
  try {
    var data = getUsersSheet().getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        return data[i][1];
      }
    }
    return email.split('@')[0];
  } catch(e) {
    return email.split('@')[0];
  }
}

// ══════════════════════════════════════════════════════
//  AUTH
// ══════════════════════════════════════════════════════
function isValidStudentEmail(email) {
  return email.toLowerCase().endsWith(STUDENT_EMAIL_DOMAIN);
}

function registerStudent(email, name, password) {
  email = email.toLowerCase().trim();

  if (!isValidStudentEmail(email))
    return { success: false, message: 'Email must end with ' + STUDENT_EMAIL_DOMAIN };
  if (!name || !password)
    return { success: false, message: 'All fields are required' };
  if (password.length < 6)
    return { success: false, message: 'Password must be at least 6 characters' };

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email)
      return { success: false, message: 'Email already registered' };
  }

  // Test email and any existing pre-verified accounts skip verification
  var isTestEmail = (email === TEST_EMAIL);

  if (isTestEmail) {
    // Auto-verified — no email sent
    sheet.appendRow([email, name, 'student', password, new Date(), 'verified']);
    return { success: true, needsVerification: false };
  } else {
    // Generate a unique token and send verification email
    var token = Utilities.getUuid();
    sheet.appendRow([email, name, 'student', password, new Date(), token]);

    var verifyUrl = SCRIPT_URL + '?action=verifyEmail&token=' + token;
    GmailApp.sendEmail(
      email,
      'ยืนยันอีเมลของคุณ - ศูนย์ให้คำปรึกษา SRU',
      'สวัสดี ' + name + ',\n\n' +
      'กรุณากดลิงก์ด้านล่างเพื่อยืนยันอีเมลของคุณ:\n\n' +
      verifyUrl + '\n\n' +
      'หลังจากยืนยันแล้ว คุณสามารถเข้าสู่ระบบได้ทันที\n\n' +
      'ศูนย์ให้คำปรึกษา วิทยาลัยนานาชาติการท่องเที่ยว (SRU)'
    );
    return { success: true, needsVerification: true };
  }
}

function verifyEmail(token) {
  if (!token) return { success: false, message: 'Invalid verification link.' };

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var stored = data[i][5] ? data[i][5].toString() : '';
    if (stored === token) {
      sheet.getRange(i + 1, 6).setValue('verified');
      return { success: true, message: 'Email verified! You can now log in.' };
    }
  }
  return { success: false, message: 'Token not found or already used.' };
}

function loginStudent(email, password) {
  email = email.toLowerCase().trim();
  if (!isValidStudentEmail(email))
    return { success: false, code: 'INVALID_EMAIL', message: 'Email must end with ' + STUDENT_EMAIL_DOMAIN };

  var data = getUsersSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email && data[i][2] === 'student') {

      if (data[i][3].toString() !== password.toString())
        return { success: false, code: 'WRONG_PASSWORD', message: 'Incorrect password' };

      // Check verification status in column F (index 5)
      // Empty column F = old account created before this update = allow login
      var verifyStatus = data[i][5] ? data[i][5].toString().trim() : '';
      if (verifyStatus !== '' && verifyStatus !== 'verified') {
        return { success: false, code: 'NOT_VERIFIED', message: 'Please verify your email first. Check your @student.sru.ac.th inbox.' };
      }

      return { success: true, user: { email: data[i][0], name: data[i][1], role: 'student' } };
    }
  }
  return { success: false, code: 'NO_ACCOUNT', message: 'No account found. Please register first.' };
}

function loginAdmin(email, password) {
  email = email.toLowerCase().trim();
  if (email === ADMIN_EMAIL && password === ADMIN_PASSWORD)
    return { success: true, user: { email: ADMIN_EMAIL, name: 'Administrator', role: 'admin' } };
  return { success: false, message: 'Invalid admin credentials' };
}

function forgotPassword(email) {
  email = email.toLowerCase().trim();
  if (!isValidStudentEmail(email))
    return { success: false, message: 'Email must end with ' + STUDENT_EMAIL_DOMAIN };

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email && data[i][2] === 'student') {
      var chars    = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
      var tempPass = '';
      for (var j = 0; j < 8; j++) tempPass += chars.charAt(Math.floor(Math.random() * chars.length));
      sheet.getRange(i + 1, 4).setValue(tempPass);
      GmailApp.sendEmail(
        email,
        'รหัสผ่านชั่วคราว - ศูนย์ให้คำปรึกษา SRU',
        'สวัสดี ' + data[i][1] + ',\n\nรหัสผ่านชั่วคราวของคุณคือ: ' + tempPass +
        '\n\nกรุณาเข้าสู่ระบบและเปลี่ยนรหัสผ่านในหน้า Settings\n\nศูนย์ให้คำปรึกษา วิทยาลัยนานาชาติการท่องเที่ยว (SRU)'
      );
      return { success: true, message: 'Temporary password sent to your email!' };
    }
  }
  return { success: false, message: 'No account found with this email.' };
}

function changePassword(email, oldPassword, newPassword) {
  email = email.toLowerCase().trim();
  if (!isValidStudentEmail(email))      return { success: false, message: 'Invalid student email' };
  if (!oldPassword || !newPassword)     return { success: false, message: 'Please fill all fields' };
  if (newPassword.length < 6)           return { success: false, message: 'New password must be at least 6 characters' };
  if (oldPassword === newPassword)      return { success: false, message: 'New password must be different from old password' };

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
  if (!details || details.trim() === '')
    return { success: false, message: 'Please enter details' };
  var sheet = getRequestsSheet();
  sheet.appendRow([new Date().getTime(), studentEmail, category, details, 'pending', '', new Date()]);
  return { success: true };
}

function getMyRequests(studentEmail) {
  var data     = getRequestsSheet().getDataRange().getValues();
  var requests = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().toLowerCase() === studentEmail.toLowerCase()) {
      requests.push({
        id:        data[i][0],
        category:  data[i][2],
        details:   data[i][3],
        status:    data[i][4] || 'pending',
        reply:     data[i][5] || '',
        timestamp: data[i][6] ? new Date(data[i][6]).toLocaleString() : ''
      });
    }
  }
  return requests.reverse();
}

function getAllRequests() {
  var data     = getRequestsSheet().getDataRange().getValues();
  var requests = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var email = data[i][1] ? data[i][1].toString() : '';
      requests.push({
        rowIndex:     i + 1,
        id:           data[i][0],
        studentEmail: email,
        studentName:  getStudentName(email),
        category:     data[i][2] || '',
        details:      data[i][3] || '',
        status:       data[i][4] || 'pending',
        reply:        data[i][5] || '',
        timestamp:    data[i][6] ? new Date(data[i][6]).toLocaleString() : ''
      });
    }
  }
  return requests.reverse();
}

function getStats() {
  var data = getRequestsSheet().getDataRange().getValues();
  var total = 0, pending = 0, completed = 0, rejected = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      total++;
      var s = (data[i][4] || 'pending').toString().toLowerCase();
      if (s === 'pending')        pending++;
      else if (s === 'completed') completed++;
      else if (s === 'rejected')  rejected++;
    }
  }
  return { total: total, pending: pending, completed: completed, rejected: rejected };
}

function updateRequestStatus(rowIndex, newStatus) {
  try {
    getRequestsSheet().getRange(rowIndex, 5).setValue(newStatus);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function replyToRequest(rowIndex, replyText) {
  if (!replyText || replyText.trim() === '')
    return { success: false, message: 'Reply cannot be empty' };
  try {
    var sheet = getRequestsSheet();
    sheet.getRange(rowIndex, 6).setValue(replyText);
    sheet.getRange(rowIndex, 5).setValue('completed');
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteRequest(rowIndex) {
  try {
    getRequestsSheet().deleteRow(rowIndex);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}
