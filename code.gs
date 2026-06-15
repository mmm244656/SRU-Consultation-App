// ══════════════════════════════════════════════════════
//  CONFIGURATION
// ══════════════════════════════════════════════════════
const SHEET_ID             = '14ev6msUxlbJeP3XU3mL9YG9iw-eMPaCJDNGrwRUrkFE';
const ADMIN_EMAIL          = 'admin@rex.com';
const ADMIN_PASSWORD       = 'rex123123';
const STUDENT_DOMAIN       = '@student.sru.ac.th';
const PROFESSOR_DOMAIN     = '@sru.ac.th';
const TEST_EMAIL           = 'test@student.sru.ac.th';
const TEST_PASSWORD        = 'teststudent';
const RESET_TOKEN_EXPIRY_MS = 30 * 60 * 1000; // 30 minutes
const SCRIPT_URL           = 'https://script.google.com/macros/s/AKfycbwYBwryZsJ-HL1cxsgQj9o3_2C4ETrznHZdrqCJKLtUx2DoBcYGRIrQnUyoSGK1y1leCg/exec';

// ══════════════════════════════════════════════════════
//  ENTRY POINTS
// ══════════════════════════════════════════════════════
function doGet(e) {
  if (e && e.parameter) {
    // Email verification link clicked in browser
    if (e.parameter.action === 'verifyEmail' && e.parameter.token) {
      var result = verifyEmail(e.parameter.token);
      var html = result.success
        ? '<div style="font-family:sans-serif;text-align:center;padding:40px;">' +
          '<div style="font-size:52px;margin-bottom:16px;">✅</div>' +
          '<h2 style="color:#10b981;">Email Verified!</h2>' +
          '<p style="color:#64748b;">Your account is now active.</p>' +
          '<p style="color:#94a3b8;margin-top:12px;font-size:14px;">You can now close this tab and log in to SRU Hub.</p>' +
          '</div>'
        : '<div style="font-family:sans-serif;text-align:center;padding:40px;">' +
          '<div style="font-size:52px;margin-bottom:16px;">❌</div>' +
          '<h2 style="color:#ef4444;">Verification Failed</h2>' +
          '<p style="color:#64748b;">' + result.message + '</p>' +
          '<a href="https://mmm244656.github.io/SRU-Consultation-App/" ' +
          'style="display:inline-block;margin-top:20px;padding:12px 24px;' +
          'background:#6366f1;color:white;border-radius:8px;text-decoration:none;font-weight:600;">' +
          '← Back to SRU Hub</a></div>';
      return HtmlService.createHtmlOutput(html);
    }
    if (e.parameter.action) {
      return handleApiRequest(e.parameter);
    }
  }
  return ContentService.createTextOutput('SRU Hub Backend running.').setMimeType(ContentService.MimeType.TEXT);
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
      case 'loginStudent':        result = loginStudent(p.email, p.password);                     break;
      case 'loginAdmin':          result = loginAdmin(p.email, p.password);                       break;
      case 'registerStudent':     result = registerStudent(p.email, p.name, p.password);          break;
      case 'verifyEmail':         result = verifyEmail(p.token);                                  break;
      case 'resendVerification':  result = resendVerification(p.email);                           break;
      case 'sendPasswordReset':   result = sendPasswordReset(p.email);                             break;
      case 'resetPasswordFromToken': result = resetPasswordFromToken(p.token, p.password);         break;
      case 'changePassword':      result = changePassword(p.email, p.oldPassword, p.newPassword); break;
      case 'submitRequest':       result = submitRequest(p.studentEmail, p.category, p.details);  break;
      case 'getMyRequests':       result = getMyRequests(p.studentEmail);                         break;
      case 'getAllRequests':       result = getAllRequests();                                       break;
      case 'getStats':            result = getStats();                                             break;
      case 'updateRequestStatus': result = updateRequestStatus(Number(p.rowIndex), p.newStatus);  break;
      case 'replyToRequest':      result = replyToRequest(Number(p.rowIndex), p.replyText);       break;
      case 'deleteRequest':       result = deleteRequest(Number(p.rowIndex));                     break;
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
//  Users: A=email B=name C=role D=password E=created_at F=verified
// ══════════════════════════════════════════════════════
function getUsersSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['email','name','role','password','created_at','verified','reset_token','reset_expires']);
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

function getUserInfo(email) {
  try {
    var data = getUsersSheet().getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        return { name: data[i][1], role: data[i][2] };
      }
    }
  } catch(e) {}
  return { name: email.split('@')[0], role: 'student' };
}

// ══════════════════════════════════════════════════════
//  AUTH HELPERS
// ══════════════════════════════════════════════════════
function isValidEmail(email) {
  return email.toLowerCase().endsWith(STUDENT_DOMAIN) || email.toLowerCase().endsWith(PROFESSOR_DOMAIN);
}

function getRoleByEmail(email) {
  if (email.toLowerCase().endsWith(PROFESSOR_DOMAIN)) return 'professor';
  return 'student';
}

// ══════════════════════════════════════════════════════
//  REGISTER
// ══════════════════════════════════════════════════════
function registerStudent(email, name, password) {
  email = email.toLowerCase().trim();

  if (!isValidEmail(email))
    return { success: false, message: 'Email must end with ' + STUDENT_DOMAIN + ' or ' + PROFESSOR_DOMAIN };
  if (!name || !password)
    return { success: false, message: 'All fields are required' };
  if (password.length < 6)
    return { success: false, message: 'Password must be at least 6 characters' };
  if (email === TEST_EMAIL)
    return { success: false, message: 'This email is reserved. Please use a different email.' };

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email)
      return { success: false, message: 'Email already registered. Please sign in.' };
  }

  // Auto assign role based on domain
  var role  = getRoleByEmail(email);
  var token = Utilities.getUuid();
  sheet.appendRow([email, name, role, password, new Date(), token]);

  // Send verification email
  var verifyUrl = SCRIPT_URL + '?action=verifyEmail&token=' + token;
      GmailApp.sendEmail(
        email,
        'ยืนยันอีเมลของคุณ - SRU Hub / Verify your email - SRU Hub',
        'สวัสดี ' + name + ' / Hello ' + name + ',\n\n' +
        'กรุณากดลิงก์ด้านล่างเพื่อยืนยันอีเมลและเริ่มใช้งาน SRU Hub:\n' +
        'Click the link below to verify your email and start using SRU Hub:\n\n' +
        verifyUrl + '\n\n' +
        'หากคุณไม่ได้สมัครสมาชิก กรุณาเพิกเฉยต่ออีเมลนี้\n' +
        'If you did not register, please ignore this email.\n\n' +
        'SRU Hub - วิทยาลัยนานาชาติการท่องเที่ยว'
      );

  return { success: true };
}

// ══════════════════════════════════════════════════════
//  VERIFY EMAIL
// ══════════════════════════════════════════════════════
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

// ══════════════════════════════════════════════════════
//  RESEND VERIFICATION
// ══════════════════════════════════════════════════════
function resendVerification(email) {
  email = email.toLowerCase().trim();
  if (!isValidEmail(email)) return { success: false, message: 'Invalid email' };

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email) {
      var verified = data[i][5] ? data[i][5].toString().trim() : '';
      if (verified === 'verified') return { success: false, message: 'Email already verified. Please login.' };
      var token = Utilities.getUuid();
      sheet.getRange(i + 1, 6).setValue(token);
      var verifyUrl = SCRIPT_URL + '?action=verifyEmail&token=' + token;
      GmailApp.sendEmail(
        email,
        'ยืนยันอีเมลของคุณ (ส่งใหม่) - SRU Hub',
        'สวัสดี ' + data[i][1] + ' / Hello ' + data[i][1] + ',\n\n' +
        'นี่คืออีเมลยืนยันใหม่ของคุณ / Here is your new verification link:\n\n' +
        verifyUrl + '\n\n' +
        'SRU Hub - วิทยาลัยนานาชาติการท่องเที่ยว'
      );
      return { success: true };
    }
  }
  return { success: false, message: 'No account found with this email.' };
}

// ══════════════════════════════════════════════════════
//  LOGIN STUDENT / PROFESSOR
// ══════════════════════════════════════════════════════
function loginStudent(email, password) {
  email = email.toLowerCase().trim();

  if (!isValidEmail(email))
    return { success: false, message: 'Email must end with ' + STUDENT_DOMAIN + ' or ' + PROFESSOR_DOMAIN };

  // Test account shortcut
  if (email === TEST_EMAIL) {
    if (password === TEST_PASSWORD)
      return { success: true, user: { email: email, name: 'Test Student', role: 'student' } };
    return { success: false, message: 'Incorrect password' };
  }

  var data = getUsersSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email) {
      if (data[i][3].toString() !== password.toString())
        return { success: false, message: 'Incorrect password' };
      var verified = data[i][5] ? data[i][5].toString().trim() : '';
      if (verified !== 'verified')
        return { success: false, message: 'Please verify your email first. Check your inbox.' };
      var role = data[i][2] ? data[i][2].toString() : getRoleByEmail(email);
      return { success: true, user: { email: data[i][0], name: data[i][1], role: role } };
    }
  }
  return { success: false, message: 'No account found. Please register first.' };
}

// ══════════════════════════════════════════════════════
//  LOGIN ADMIN
// ══════════════════════════════════════════════════════
function loginAdmin(email, password) {
  email = email.toLowerCase().trim();
  if (email === ADMIN_EMAIL && password === ADMIN_PASSWORD)
    return { success: true, user: { email: ADMIN_EMAIL, name: 'Administrator', role: 'admin' } };
  return { success: false, message: 'Invalid admin credentials' };
}

// ══════════════════════════════════════════════════════
//  CHANGE PASSWORD
// ══════════════════════════════════════════════════════
function changePassword(email, oldPassword, newPassword) {
  email = email.toLowerCase().trim();
  if (!oldPassword || !newPassword)  return { success: false, message: 'Please fill all fields' };
  if (newPassword.length < 6)        return { success: false, message: 'New password must be at least 6 characters' };
  if (oldPassword === newPassword)   return { success: false, message: 'New password must be different' };

  if (email === TEST_EMAIL) {
    if (oldPassword !== TEST_PASSWORD) return { success: false, message: 'Current password is incorrect' };
    return { success: true, message: 'Password changed successfully!' };
  }

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email) {
      if (data[i][3].toString() !== oldPassword.toString())
        return { success: false, message: 'Current password is incorrect' };
      sheet.getRange(i + 1, 4).setValue(newPassword);
      return { success: true, message: 'Password changed successfully!' };
    }
  }
  return { success: false, message: 'Account not found' };
}

// ══════════════════════════════════════════════════════
//  SEND PASSWORD RESET
// ══════════════════════════════════════════════════════
function sendPasswordReset(email) {
  email = email.toLowerCase().trim();
  if (!isValidEmail(email)) return { success: false, message: 'Invalid email domain' };
  if (email === TEST_EMAIL)  return { success: false, message: 'This account uses a fixed password.' };

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email) {
      var verified = data[i][5] ? data[i][5].toString().trim() : '';
      if (verified !== 'verified') return { success: false, message: 'Please verify your email first before resetting password.' };

      var token   = Utilities.getUuid();
      var expires = Date.now() + RESET_TOKEN_EXPIRY_MS;
      sheet.getRange(i + 1, 7).setValue(token);
      sheet.getRange(i + 1, 8).setValue(expires);

      var resetUrl = 'https://mmm244656.github.io/SRU-Consultation-App/?resetToken=' + token;
      GmailApp.sendEmail(
        email,
        'รีเซ็ตรหัสผ่าน - SRU Hub / Password Reset - SRU Hub',
        'สวัสดี ' + data[i][1] + ' / Hello ' + data[i][1] + ',\n\n' +
        'กรุณากดลิงก์ด้านล่างเพื่อตั้งรหัสผ่านใหม่:\n' +
        'Click the link below to set your new password:\n\n' +
        resetUrl + '\n\n' +
        '⏰ ลิงก์นี้จะหมดอายุใน 30 นาที / This link expires in 30 minutes.\n\n' +
        'หากคุณไม่ได้ขอรีเซ็ตรหัสผ่าน กรุณาเพิกเฉยต่ออีเมลนี้\n' +
        'If you did not request this, please ignore this email.\n\n' +
        'SRU Hub - วิทยาลัยนานาชาติการท่องเที่ยว'
      );
      return { success: true };
    }
  }
  return { success: false, message: 'No account found with this email.' };
}

// ══════════════════════════════════════════════════════
//  RESET PASSWORD FROM TOKEN
// ══════════════════════════════════════════════════════
function resetPasswordFromToken(token, password) {
  if (!token || !password) return { success: false, message: 'Missing required fields' };
  if (password.length < 6)  return { success: false, message: 'Password must be at least 6 characters' };

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var storedToken   = data[i][6] ? data[i][6].toString() : '';
    var storedExpires = data[i][7] ? Number(data[i][7]) : 0;
    if (storedToken === token) {
      if (Date.now() > storedExpires) return { success: false, message: 'Reset link has expired. Please request a new one.' };
      // Update password and clear token
      sheet.getRange(i + 1, 4).setValue(password);
      sheet.getRange(i + 1, 7).setValue('');
      sheet.getRange(i + 1, 8).setValue('');
      return { success: true };
    }
  }
  return { success: false, message: 'Invalid or already used reset link.' };
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
      out.push({
        id: data[i][0], category: data[i][2], details: data[i][3],
        status: data[i][4]||'pending', reply: data[i][5]||'',
        timestamp: data[i][6] ? new Date(data[i][6]).toLocaleString() : ''
      });
    }
  }
  return out.reverse();
}

function getAllRequests() {
  var data = getRequestsSheet().getDataRange().getValues();
  var out  = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var email    = data[i][1] ? data[i][1].toString() : '';
      var userInfo = getUserInfo(email);
      out.push({
        rowIndex:      i + 1,
        id:            data[i][0],
        studentEmail:  email,
        studentName:   userInfo.name,
        submitterRole: userInfo.role,
        category:      data[i][2]||'',
        details:       data[i][3]||'',
        status:        data[i][4]||'pending',
        reply:         data[i][5]||'',
        timestamp:     data[i][6] ? new Date(data[i][6]).toLocaleString() : ''
      });
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
