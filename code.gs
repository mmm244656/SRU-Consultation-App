// ══════════════════════════════════════════════════════
//  CONFIGURATION
// ══════════════════════════════════════════════════════
const SHEET_ID             = '1agszlx4U4gFBnzClfrt-fFwSrg5vipJxAzCNcliPSDY';
const ADMIN_EMAIL          = 'admin@rex.com';
const ADMIN_PASSWORD       = 'rex123123';
const STUDENT_EMAIL_DOMAIN = '@student.sru.ac.th';
const SCRIPT_URL           = 'https://script.google.com/macros/s/AKfycbwcd_YywzV6mjkRFGS2bM0cEGdaYMVE4V_wd10N43F61WOfCa6f75rGfPWb1cpCR5m1/exec';
const PORTAL_URL           = 'https://mmm244656.github.io/SRU-Student-Hub/';
const TEST_EMAIL           = 'test@student.sru.ac.th';
const TEST_PASSWORD        = 'test123';

// Magic link token expiry: 30 minutes
const MAGIC_LINK_EXPIRY_MS = 30 * 60 * 1000;

// ══════════════════════════════════════════════════════
//  ENTRY POINTS
// ══════════════════════════════════════════════════════
function doGet(e) {
  if (e && e.parameter) {
    // Magic link clicked in email
    if (e.parameter.magicToken) {
      var token  = e.parameter.magicToken;
      var result = preVerifyMagicToken(token);
      if (result.success) {
        // Redirect to portal set-password page with token in URL
        var redirectUrl = PORTAL_URL + '?token=' + encodeURIComponent(token);
        return HtmlService.createHtmlOutput(
          '<script>window.location.href = "' + redirectUrl + '";<\/script>'
          + '<p style="font-family:sans-serif;">Redirecting… <a href="' + redirectUrl + '">Click here if not redirected</a></p>'
        );
      } else {
        return HtmlService.createHtmlOutput(
          '<h2 style="font-family:sans-serif;color:red;">❌ Link Expired</h2>'
          + '<p style="font-family:sans-serif;">This magic link has expired or already been used.<br>'
          + '<a href="' + PORTAL_URL + '">← Back to Portal</a></p>'
        );
      }
    }
    // Regular API action
    if (e.parameter.action) {
      return handleApiRequest(e.parameter);
    }
  }
  return ContentService.createTextOutput('SRU Portal Backend').setMimeType(ContentService.MimeType.TEXT);
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
      // ── Auth ──
      case 'loginStudent':         result = loginStudent(p.email, p.password);                          break;
      case 'loginAdmin':           result = loginAdmin(p.email, p.password);                            break;
      case 'checkStudentAccount':  result = checkStudentAccount(p.email);                               break;
      case 'sendMagicLink':        result = sendMagicLink(p.email);                                     break;
      case 'verifyMagicToken':     result = verifyMagicToken(p.token);                                  break;
      case 'setPasswordFromToken': result = setPasswordFromToken(p.token, p.email, p.password);         break;
      case 'changePassword':       result = changePassword(p.email, p.oldPassword, p.newPassword);      break;
      case 'forgotPassword':       result = forgotPassword(p.email);                                    break;
      // ── Requests ──
      case 'submitRequest':        result = submitRequest(p.studentEmail, p.category, p.details);       break;
      case 'getMyRequests':        result = getMyRequests(p.studentEmail);                              break;
      case 'getAllRequests':        result = getAllRequests();                                            break;
      case 'getStats':             result = getStats();                                                  break;
      case 'updateRequestStatus':  result = updateRequestStatus(Number(p.rowIndex), p.newStatus);       break;
      case 'replyToRequest':       result = replyToRequest(Number(p.rowIndex), p.replyText);            break;
      case 'deleteRequest':        result = deleteRequest(Number(p.rowIndex));                          break;
      default: result = { success: false, message: 'Unknown action: ' + action };
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
//  Users columns:
//  A(0)=email  B(1)=name  C(2)=role  D(3)=password
//  E(4)=created_at  F(5)=verified  G(6)=magic_token  H(7)=token_expires
// ══════════════════════════════════════════════════════
function getUsersSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['email','name','role','password','created_at','verified','magic_token','token_expires']);
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
    return email.split('@')[0];
  } catch(e) { return email.split('@')[0]; }
}

// ══════════════════════════════════════════════════════
//  AUTH HELPERS
// ══════════════════════════════════════════════════════
function isValidStudentEmail(email) {
  return email.toLowerCase().endsWith(STUDENT_EMAIL_DOMAIN);
}

// ══════════════════════════════════════════════════════
//  CHECK STUDENT ACCOUNT
//  Returns: { exists, hasPassword }
// ══════════════════════════════════════════════════════
function checkStudentAccount(email) {
  email = email.toLowerCase().trim();
  if (!isValidStudentEmail(email)) return { success: false, message: 'Invalid email domain' };

  // Test email — always has password
  if (email === TEST_EMAIL) return { exists: true, hasPassword: true };

  var data = getUsersSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email) {
      var pw = data[i][3] ? data[i][3].toString().trim() : '';
      return { exists: true, hasPassword: pw !== '' };
    }
  }
  return { exists: false, hasPassword: false };
}

// ══════════════════════════════════════════════════════
//  SEND MAGIC LINK
// ══════════════════════════════════════════════════════
function sendMagicLink(email) {
  email = email.toLowerCase().trim();
  if (!isValidStudentEmail(email)) return { success: false, message: 'Email must end with ' + STUDENT_EMAIL_DOMAIN };

  // Block test email — it uses password only
  if (email === TEST_EMAIL) return { success: false, message: 'This account uses a password. Please enter your password.' };

  var sheet   = getUsersSheet();
  var data    = sheet.getDataRange().getValues();
  var token   = Utilities.getUuid();
  var expires = new Date(Date.now() + MAGIC_LINK_EXPIRY_MS);
  var magicUrl = SCRIPT_URL + '?magicToken=' + token;

  // Find existing row and update token
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email) {
      sheet.getRange(i + 1, 7).setValue(token);
      sheet.getRange(i + 1, 8).setValue(expires.getTime());
      sendMagicEmail(email, data[i][1], magicUrl);
      return { success: true };
    }
  }

  // New student — create account row (no password yet)
  var name = email.split('@')[0];
  sheet.appendRow([email, name, 'student', '', new Date(), 'pending', token, expires.getTime()]);
  sendMagicEmail(email, name, magicUrl);
  return { success: true };
}

function sendMagicEmail(email, name, magicUrl) {
  GmailApp.sendEmail(
    email,
    'Your SRU Portal Magic Link / ลิงก์เข้าสู่ระบบ SRU Portal',
    'สวัสดี ' + name + ' / Hello ' + name + ',\n\n' +
    'คลิกลิงก์ด้านล่างเพื่อเข้าสู่ระบบ SRU Portal และตั้งรหัสผ่านของคุณ:\n' +
    'Click the link below to sign in to SRU Portal and set your password:\n\n' +
    magicUrl + '\n\n' +
    '⏰ ลิงก์นี้จะหมดอายุใน 30 นาที / This link expires in 30 minutes.\n\n' +
    'หากคุณไม่ได้ขอลิงก์นี้ ให้เพิกเฉยต่ออีเมลนี้\n' +
    'If you did not request this, please ignore this email.\n\n' +
    'SRU Portal — วิทยาลัยนานาชาติการท่องเที่ยว'
  );
}

// ══════════════════════════════════════════════════════
//  VERIFY MAGIC TOKEN (called by redirect — server side)
// ══════════════════════════════════════════════════════
function preVerifyMagicToken(token) {
  if (!token) return { success: false };
  var data = getUsersSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var stored  = data[i][6] ? data[i][6].toString() : '';
    var expires = data[i][7] ? Number(data[i][7]) : 0;
    if (stored === token) {
      if (Date.now() > expires) return { success: false, message: 'Token expired' };
      return { success: true, email: data[i][0] };
    }
  }
  return { success: false, message: 'Token not found' };
}

// ══════════════════════════════════════════════════════
//  VERIFY MAGIC TOKEN (called by frontend JS)
// ══════════════════════════════════════════════════════
function verifyMagicToken(token) {
  if (!token) return { success: false, message: 'No token provided' };
  var data = getUsersSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var stored  = data[i][6] ? data[i][6].toString() : '';
    var expires = data[i][7] ? Number(data[i][7]) : 0;
    if (stored === token) {
      if (Date.now() > expires) return { success: false, message: 'Link expired. Please request a new one.' };
      return { success: true, email: data[i][0] };
    }
  }
  return { success: false, message: 'Invalid or already used link.' };
}

// ══════════════════════════════════════════════════════
//  SET PASSWORD FROM TOKEN (first-time setup)
// ══════════════════════════════════════════════════════
function setPasswordFromToken(token, email, password) {
  if (!token || !email || !password) return { success: false, message: 'Missing required fields' };
  if (password.length < 6) return { success: false, message: 'Password must be at least 6 characters' };
  email = email.toLowerCase().trim();

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var stored  = data[i][6] ? data[i][6].toString() : '';
    var expires = data[i][7] ? Number(data[i][7]) : 0;
    if (stored === token && data[i][0].toString().toLowerCase() === email) {
      if (Date.now() > expires) return { success: false, message: 'Link expired. Please request a new one.' };
      // Set password, mark verified, clear token
      sheet.getRange(i + 1, 4).setValue(password);
      sheet.getRange(i + 1, 6).setValue('verified');
      sheet.getRange(i + 1, 7).setValue('');
      sheet.getRange(i + 1, 8).setValue('');
      // Update name if it was just the email prefix
      var name = data[i][1] ? data[i][1].toString() : email.split('@')[0];
      return { success: true, user: { email: data[i][0], name: name, role: 'student' } };
    }
  }
  return { success: false, message: 'Invalid token or email mismatch.' };
}

// ══════════════════════════════════════════════════════
//  LOGIN
// ══════════════════════════════════════════════════════
function loginStudent(email, password) {
  email = email.toLowerCase().trim();
  if (!isValidStudentEmail(email)) return { success: false, code: 'INVALID_EMAIL', message: 'Email must end with ' + STUDENT_EMAIL_DOMAIN };

  // Test email shortcut
  if (email === TEST_EMAIL) {
    if (password === TEST_PASSWORD) return { success: true, user: { email: email, name: 'Test Student', role: 'student' } };
    return { success: false, code: 'WRONG_PASSWORD', message: 'Incorrect password' };
  }

  var data = getUsersSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email && data[i][2] === 'student') {
      var pw = data[i][3] ? data[i][3].toString() : '';
      if (pw === '') return { success: false, code: 'NO_PASSWORD', message: 'Please use your magic link to set a password first.' };
      if (pw !== password.toString()) return { success: false, code: 'WRONG_PASSWORD', message: 'Incorrect password' };
      var verified = data[i][5] ? data[i][5].toString().trim() : '';
      if (verified !== 'verified') return { success: false, code: 'NOT_VERIFIED', message: 'Account not yet verified.' };
      return { success: true, user: { email: data[i][0], name: data[i][1], role: 'student' } };
    }
  }
  return { success: false, code: 'NO_ACCOUNT', message: 'No account found with this email.' };
}

function loginAdmin(email, password) {
  email = email.toLowerCase().trim();
  if (email === ADMIN_EMAIL && password === ADMIN_PASSWORD)
    return { success: true, user: { email: ADMIN_EMAIL, name: 'Administrator', role: 'admin' } };
  return { success: false, message: 'Invalid admin credentials' };
}

// ══════════════════════════════════════════════════════
//  FORGOT PASSWORD (kept as fallback)
// ══════════════════════════════════════════════════════
function forgotPassword(email) {
  email = email.toLowerCase().trim();
  if (!isValidStudentEmail(email)) return { success: false, message: 'Email must end with ' + STUDENT_EMAIL_DOMAIN };
  // For new system, just send a new magic link
  return sendMagicLink(email);
}

// ══════════════════════════════════════════════════════
//  CHANGE PASSWORD
// ══════════════════════════════════════════════════════
function changePassword(email, oldPassword, newPassword) {
  email = email.toLowerCase().trim();
  if (!isValidStudentEmail(email))  return { success: false, message: 'Invalid student email' };
  if (!oldPassword || !newPassword) return { success: false, message: 'Please fill all fields' };
  if (newPassword.length < 6)       return { success: false, message: 'New password must be at least 6 characters' };
  if (oldPassword === newPassword)  return { success: false, message: 'New password must be different' };

  // Test email
  if (email === TEST_EMAIL) {
    if (oldPassword !== TEST_PASSWORD) return { success: false, message: 'Current password is incorrect' };
    return { success: true, message: 'Password changed successfully!' };
  }

  var sheet = getUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email && data[i][2] === 'student') {
      if (data[i][3].toString() !== oldPassword.toString()) return { success: false, message: 'Current password is incorrect' };
      sheet.getRange(i + 1, 4).setValue(newPassword);
      return { success: true, message: 'Password changed successfully!' };
    }
  }
  return { success: false, message: 'Account not found' };
}

// ══════════════════════════════════════════════════════
//  REQUESTS (unchanged)
// ══════════════════════════════════════════════════════
function submitRequest(studentEmail, category, details) {
  if (!details || details.trim() === '') return { success: false, message: 'Please enter details' };
  var sheet = getRequestsSheet();
  sheet.appendRow([new Date().getTime(), studentEmail, category, details, 'pending', '', new Date()]);
  return { success: true };
}

function getMyRequests(studentEmail) {
  var data     = getRequestsSheet().getDataRange().getValues();
  var requests = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().toLowerCase() === studentEmail.toLowerCase()) {
      requests.push({ id: data[i][0], category: data[i][2], details: data[i][3], status: data[i][4]||'pending', reply: data[i][5]||'', timestamp: data[i][6] ? new Date(data[i][6]).toLocaleString() : '' });
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
      requests.push({ rowIndex: i+1, id: data[i][0], studentEmail: email, studentName: getStudentName(email), category: data[i][2]||'', details: data[i][3]||'', status: data[i][4]||'pending', reply: data[i][5]||'', timestamp: data[i][6] ? new Date(data[i][6]).toLocaleString() : '' });
    }
  }
  return requests.reverse();
}

function getStats() {
  var data = getRequestsSheet().getDataRange().getValues();
  var total=0, pending=0, completed=0, rejected=0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      total++;
      var s = (data[i][4]||'pending').toString().toLowerCase();
      if (s==='pending') pending++;
      else if (s==='completed') completed++;
      else if (s==='rejected') rejected++;
    }
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
