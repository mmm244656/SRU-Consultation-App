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
    var result = verifyEmailFromLink(e.parameter.token);
    var html = result.success
      ? '<html><head><style>body{font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f0f9ff;}</style></head><body><div style="background:white;padding:40px;border-radius:10px;box-shadow:0 4px 6px rgba(0,0,0,0.1);text-align:center;max-width:400px;"><h2 style="color:green;margin:0 0 10px 0;">✅ Email Verified!</h2><p style="color:#666;margin:0;">Your account is now active. You can close this tab and log in.</p></div></body></html>'
      : '<html><head><style>body{font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#fef2f2;}</style></head><body><div style="background:white;padding:40px;border-radius:10px;box-shadow:0 4px 6px rgba(0,0,0,0.1);text-align:center;max-width:400px;"><h2 style="color:red;margin:0 0 10px 0;">❌ Verification Failed</h2><p style="color:#666;margin:0;">' + result.message + '</p></div></body></html>';
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
      case 'checkEmailVerified':
        result = checkEmailVerified(p.email); break;
      case 'resendVerificationEmail':
        result = resendVerificationEmail(p.email); break;
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
//  PENDING REGISTRATIONS SHEET
// ══════════════════════════════════════════════════════
function getPendingSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('PendingRegistrations');
  if (!sheet) {
    sheet = ss.insertSheet('PendingRegistrations');
    sheet.appendRow(['email', 'name', 'password', 'token', 'created_at', 'expires_at']);
  }
  return sheet;
}

// ══════════════════════════════════════════════════════
//  EMAIL SENDER
// ══════════════════════════════════════════════════════
function sendVerificationEmail(email, name, token) {
  try {
    var verifyUrl = SCRIPT_URL + '?token=' + encodeURIComponent(token);
    
    var subject = 'ยืนยันอีเมลของคุณ - SRU Student Hub';
    
    var htmlBody = '<div style="font-family:Arial,sans-serif;padding:20px;max-width:600px;">' +
      '<h2 style="color:#333;">สวัสดี ' + name + '</h2>' +
      '<p style="font-size:16px;">ขอบคุณที่สมัครสมาชิก SRU Student Hub</p>' +
      '<p style="font-size:16px;">กรุณากดปุ่มด้านล่างเพื่อยืนยันอีเมลของคุณ:</p>' +
      '<p style="margin:30px 0;"><a href="' + verifyUrl + '" style="background:#6366f1;color:white;padding:14px 32px;text-decoration:none;border-radius:6px;display:inline-block;font-weight:bold;font-size:16px;">ยืนยันอีเมล</a></p>' +
      '<p style="color:#666;">หรือคัดลอกลิงก์นี้ลงในเบราว์เซอร์:</p>' +
      '<p style="word-break:break-all;background:#f5f5f5;padding:10px;border-radius:4px;"><code style="font-size:12px;">' + verifyUrl + '</code></p>' +
      '<p style="color:#999;font-size:12px;margin-top:30px;">⏱️ ลิงก์นี้จะหมดอายุใน 24 ชั่วโมง</p>' +
      '<p style="color:#999;font-size:12px;">หากคุณไม่ได้สมัครสมาชิก โปรดเพิกเฉยต่อข้อความนี้</p>' +
      '<hr style="border:none;border-top:1px solid #ddd;margin:30px 0;">' +
      '<p style="color:#999;font-size:11px;">SRU Student Hub | วิทยาลัยนานาชาติการท่องเที่ยว</p>' +
      '</div>';
    
    var textBody = 'สวัสดี ' + name + ',\n\n' +
      'ขอบคุณที่สมัครสมาชิก SRU Student Hub\n\n' +
      'กรุณากดลิงก์ด้านล่างเพื่อยืนยันอีเมลของคุณ:\n\n' +
      verifyUrl + '\n\n' +
      'ลิงก์นี้จะหมดอายุใน 24 ชั่วโมง\n\n' +
      'หากคุณไม่ได้สมัครสมาชิก โปรดเพิกเฉยต่อข้อความนี้\n\n' +
      '---\n' +
      'SRU Student Hub\n' +
      'วิทยาลัยนานาชาติการท่องเที่ยว';
    
    GmailApp.sendEmail(email, subject, textBody, {
      htmlBody: htmlBody,
      name: 'SRU Student Hub',
      noReply: true
    });
    
    Logger.log('Email sent to: ' + email);
    return { success: true };
    
  } catch(err) {
    Logger.log('Email error: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// ══════════════════════════════════════════════════════
//  SHEET HELPERS
// ══════════════════════════════════════════════════════
function getUsersSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
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

  var usersData = getUsersSheet().getDataRange().getValues();
  for (var i = 1; i < usersData.length; i++) {
    if (usersData[i][0] && usersData[i][0].toString().toLowerCase() === email) {
      return { success: false, message: 'Email already registered' };
    }
  }

  // Check pending registrations
  var pendingData = getPendingSheet().getDataRange().getValues();
  for (var i = 1; i < pendingData.length; i++) {
    if (pendingData[i][0] && pendingData[i][0].toString().toLowerCase() === email) {
      return { success: false, message: 'This email is already awaiting verification. Please check your inbox.' };
    }
  }

  // Test email - auto verify
  if (email === TEST_EMAIL) {
    getUsersSheet().appendRow([email, name, 'student', password, new Date(), 'verified']);
    return { success: true, autoVerified: true };
  }

  // Generate token and save to pending
  var token = Utilities.getUuid();
  var expiresAt = new Date(Date.now() + 24 * 60 * 60 * 1000); // 24 hours
  getPendingSheet().appendRow([email, name, password, token, new Date(), expiresAt]);

  // Send email
  var emailResult = sendVerificationEmail(email, name, token);

  if (!emailResult.success) {
    // Delete the pending entry if email failed
    var pendingSheet = getPendingSheet();
    var allData = pendingSheet.getDataRange().getValues();
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][0] && allData[i][0].toString().toLowerCase() === email) {
        pendingSheet.deleteRow(i + 1);
        break;
      }
    }
    return { success: false, message: 'Failed to send verification email. Please try again.' };
  }

  return { success: true, emailSent: true, message: 'Registration successful! Check your email for verification link.' };
}

function verifyEmailFromLink(token) {
  if (!token) return { success: false, message: 'Invalid verification link.' };

  var pendingSheet = getPendingSheet();
  var pendingData = pendingSheet.getDataRange().getValues();

  for (var i = 1; i < pendingData.length; i++) {
    var storedToken = pendingData[i][3] ? pendingData[i][3].toString() : '';
    var expiresAt = pendingData[i][5];

    if (storedToken === token) {
      // Check if token expired
      if (new Date() > new Date(expiresAt)) {
        return { success: false, message: 'Verification link has expired. Please register again.' };
      }

      // Move from pending to users
      var email = pendingData[i][0];
      var name = pendingData[i][1];
      var password = pendingData[i][2];

      // Add to Users sheet
      getUsersSheet().appendRow([email, name, 'student', password, new Date(), 'verified']);

      // Delete from pending
      pendingSheet.deleteRow(i + 1);

      return { success: true, message: 'Email verified! Your account is now active.' };
    }
  }

  return { success: false, message: 'Token not found or already used.' };
}

function checkEmailVerified(email) {
  email = email.toLowerCase().trim();
  
  var data = getUsersSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email && data[i][2] === 'student') {
      return { verified: true };
    }
  }

  // Check if still pending
  var pendingData = getPendingSheet().getDataRange().getValues();
  for (var i = 1; i < pendingData.length; i++) {
    if (pendingData[i][0] && pendingData[i][0].toString().toLowerCase() === email) {
      return { verified: false, pending: true };
    }
  }
  
  return { verified: false, pending: false };
}

function resendVerificationEmail(email) {
  email = email.toLowerCase().trim();
  if (!isValidStudentEmail(email))
    return { success: false, message: 'Invalid student email' };

  var pendingSheet = getPendingSheet();
  var pendingData = pendingSheet.getDataRange().getValues();

  for (var i = 1; i < pendingData.length; i++) {
    if (pendingData[i][0] && pendingData[i][0].toString().toLowerCase() === email) {
      var token = pendingData[i][3];
      var name = pendingData[i][1];

      // Send verification email
      var emailResult = sendVerificationEmail(email, name, token);

      if (!emailResult.success) {
        return { success: false, message: 'Failed to send email. Please try again later.' };
      }

      return { success: true, message: 'Verification email sent successfully!' };
    }
  }

  // Check if already verified
  var usersData = getUsersSheet().getDataRange().getValues();
  for (var i = 1; i < usersData.length; i++) {
    if (usersData[i][0] && usersData[i][0].toString().toLowerCase() === email && usersData[i][5] === 'verified') {
      return { success: false, message: 'Email is already verified. Please log in.' };
    }
  }

  return { success: false, message: 'Email not found in pending registrations.' };
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

      var verifyStatus = data[i][5] ? data[i][5].toString().trim() : '';
      if (verifyStatus !== 'verified') {
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
      
      try {
        GmailApp.sendEmail(
          email,
          'รหัสผ่านชั่วคราว - SRU Student Hub',
          'สวัสดี ' + data[i][1] + ',\n\nรหัสผ่านชั่วคราวของคุณคือ: ' + tempPass +
          '\n\nกรุณาเข้าสู่ระบบและเปลี่ยนรหัสผ่านในหน้า Settings\n\nSRU Student Hub'
        );
      } catch(e) {
        Logger.log('Password reset email failed: ' + e.toString());
      }
      
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
