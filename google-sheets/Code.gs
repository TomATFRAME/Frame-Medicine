// ============================================
// FRAME Medicine — Google Apps Script Backend
// ============================================
// RULES: var only, no const/let, no arrows, no template literals, no ES6
// Google Sheets is the database. This is the API layer.
// ============================================

// ---- SPREADSHEET ID ----
var SHEET_ID = "YOUR_SHEET_ID_HERE";

// ---- TWILIO CREDENTIALS ----
var TWILIO_SID = "YOUR_TWILIO_SID";
var TWILIO_TOKEN = "YOUR_TWILIO_TOKEN";
var TWILIO_NUMBER = "YOUR_TWILIO_NUMBER";
var TWILIO_VERIFY_SID = "YOUR_TWILIO_VERIFY_SID";

// ---- VAPID KEYS (generate your own for production) ----
var VAPID_PUBLIC_KEY = "";
var VAPID_PRIVATE_KEY = "";

// ---- DEPLOY URL ----
var DEPLOY_URL = "https://script.google.com/macros/s/AKfycbw_jAIyCz1bR8kg_MhAucH2X1yRZPi--N7skj4MAVie7gTXqzCC3vISGFZQiuIOpDpe/exec";

// ---- APP URLS ----
var APP_URL = "https://app.framemedicine.com";
// NOTE: Never commit real credentials to this repo. Paste them directly in Apps Script only.
var ADMIN_URL = "https://framemedicine.com/admin-app";

// ============================================
// COLUMN INDEX CONSTANTS
// ============================================

// Patients Tab
var P_NAME = 0, P_PREFERRED = 1, P_DOB = 2, P_PHONE = 3, P_EMAIL = 4;
var P_ADDR = 5, P_CITY = 6, P_STATE = 7, P_ZIP = 8, P_SINCE = 9;
var P_MED = 10, P_PLAN = 11, P_RATE = 12, P_TERM = 13, P_MEMSTART = 14;
var P_CONTEND = 15, P_CYCLES = 16, P_OUTSTANDING = 17;
var P_CIDAY = 18, P_CITIME = 19, P_GLPDAY = 20, P_GLPTIME = 21;
var P_STATUS = 22, P_FOLLOWUP = 23, P_NOTES = 24;
var P_PUSH = 25, P_PUSHSUB = 26, P_REFSOURCE = 27, P_REFBY = 28, P_BIOTOKEN = 29, P_BIOTOKEN_DATE = 30, P_COMMPREF = 31;

// Billing Tab
var S_PATIENT = 0, S_PLAN = 1, S_RATE = 2, S_TERM = 3, S_MEMSTART = 4;
var S_LASTPAY = 5, S_CONTEND = 6, S_CYCLES = 7, S_OUTSTANDING = 8;
var S_STATUS = 9, S_LASTSHIP = 10, S_NEXTSHIP = 11, S_NEXTPAYDUE = 12, S_NOTES = 13;

// Medications Tab (orders start at row 15, 0-indexed row 14)
var M_ORDERDATE = 0, M_PATIENT = 1, M_PHONE = 2, M_MED = 3, M_FORM = 4;
var M_DOSE = 5, M_VIALS = 6, M_DAYS = 7, M_SHIPDATE = 8, M_NEXTDUE = 9;
var M_VIALCOST = 10, M_SUPPLY = 11, M_SHIPPING = 12, M_TOTAL = 13;
var M_MONTHLY = 14, M_NOTES = 15;
var MED_DATA_START_ROW = 15; // 1-indexed, row 15 is where order data begins

// Labs Tab
var L_PATIENT = 0, L_ENROLL = 1, L_INIT_DATE = 2, L_INIT_DONE = 3;
var L_90_DUE = 4, L_90_DONE = 5, L_180_DUE = 6, L_180_DONE = 7;
var L_ANN_DUE = 8, L_ANN_DONE = 9, L_NEXT_DUE = 10, L_STATUS = 11, L_NOTES = 12;

// Leads Tab
var LD_NAME = 0, LD_PHONE = 1, LD_EMAIL = 2, LD_SOURCE = 3, LD_DATE = 4;
var LD_INTEREST = 5, LD_STAGE = 6, LD_ASSIGNED = 7, LD_LASTCONTACT = 8;
var LD_NEXTFOLLOWUP = 9, LD_NOTES = 10, LD_CONVERTED = 11;
var LD_CONVERTEDDATE = 12, LD_PATIENTNAME = 13;

// Messages Tab
var MSG_TIMESTAMP = 0, MSG_PATIENT = 1, MSG_PHONE = 2, MSG_DIRECTION = 3;
var MSG_TEXT = 4, MSG_READ = 5, MSG_SOURCE = 6, MSG_CONTACTTYPE = 7;

// Check-In Responses Tab
var CI_DATE = 0, CI_PATIENT = 1, CI_PHONE = 2, CI_MED = 3;
var CI_SYMPTOMS = 4, CI_RATING = 5, CI_NOTES = 6;
var CI_RESPONSE_REQ = 7, CI_RESPONDED = 8;

// Check-Ins (Schedule) Tab
var CIS_PATIENT = 0, CIS_PHONE = 1, CIS_MED = 2, CIS_DAY = 3;
var CIS_TIME = 4, CIS_LASTSENT = 5, CIS_RESPONSE = 6;
var CIS_RESPONSEDATE = 7, CIS_STATUS = 8;

// Weight Log Tab
var W_DATE = 0, W_PATIENT = 1, W_MED = 2, W_BLANK = 3;
var W_WEIGHT = 4, W_CHANGE = 5, W_AVG = 6, W_TOTAL = 7;
var W_START = 8, W_WEEKS = 9, W_SOURCE = 10;

// Dose History Tab
var DH_DATE = 0, DH_PATIENT = 1, DH_MED = 2, DH_OLDDOSE = 3;
var DH_NEWDOSE = 4, DH_CHANGEDBY = 5, DH_REASON = 6;

// Finance Tab
var FIN_MONTH = 0, FIN_YEAR = 1, FIN_MONTHNUM = 2, FIN_REVENUE = 3;
var FIN_MEDCOSTS = 4, FIN_OVERHEAD = 5, FIN_NET = 6, FIN_TOM = 7;
var FIN_COLIN = 8, FIN_LOCKED = 9;

// Overhead Items Tab
var OH_MONTH = 0, OH_YEAR = 1, OH_DESC = 2, OH_AMOUNT = 3;

// Refill Log Tab
var RL_TIMESTAMP = 0, RL_PATIENT = 1, RL_MED = 2, RL_ACTION = 3;
var RL_METHOD = 4, RL_NOTES = 5;

// Sales Tab — matches JaneApp export
var SALE_INVOICE = 0, SALE_PATIENT = 1, SALE_ITEM = 2, SALE_DATE = 3;
var SALE_TOTAL = 4, SALE_STATUS = 5;


// ============================================
// UTILITY FUNCTIONS
// ============================================

function getSheet(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(name);
}

function getSheetData(name) {
  var sheet = getSheet(name);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1); // skip header
}

function findRowByValue(sheetName, colIndex, value, startRow) {
  var sheet = getSheet(sheetName);
  if (!sheet) return -1;
  var data = sheet.getDataRange().getValues();
  var start = startRow || 1;
  for (var i = start; i < data.length; i++) {
    if (String(data[i][colIndex]).trim().toLowerCase() === String(value).trim().toLowerCase()) {
      return i + 1; // 1-indexed row number
    }
  }
  return -1;
}

function findRowByPhone(sheetName, colIndex, phone) {
  var normalized = formatPhone(phone);
  var sheet = getSheet(sheetName);
  if (!sheet) return -1;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (formatPhone(String(data[i][colIndex])) === normalized) {
      return i + 1;
    }
  }
  return -1;
}

function formatPhone(phone) {
  if (!phone) return "";
  var digits = String(phone).replace(/[^0-9]/g, "");
  if (digits.length === 10) return "+1" + digits;
  if (digits.length === 11 && digits.charAt(0) === "1") return "+" + digits;
  if (String(phone).charAt(0) === "+") return String(phone);
  return "+1" + digits;
}

function safeString(val) {
  if (val === null || val === undefined) return "";
  return String(val);
}

function safeNumber(val) {
  var n = Number(val);
  if (isNaN(n)) return 0;
  return n;
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  var d = new Date(val);
  if (isNaN(d.getTime())) return null;
  return d;
}

function formatDateStr(d) {
  if (!d) return "";
  if (!(d instanceof Date)) d = new Date(d);
  if (isNaN(d.getTime())) return "";
  var m = d.getMonth() + 1;
  var day = d.getDate();
  var y = d.getFullYear();
  return (m < 10 ? "0" + m : m) + "/" + (day < 10 ? "0" + day : day) + "/" + y;
}

function formatDateISO(d) {
  if (!d) return "";
  if (!(d instanceof Date)) d = new Date(d);
  if (isNaN(d.getTime())) return "";
  return d.toISOString().split("T")[0];
}

function daysBetween(d1, d2) {
  if (!d1 || !d2) return 0;
  var a = new Date(d1);
  var b = new Date(d2);
  a.setHours(0, 0, 0, 0);
  b.setHours(0, 0, 0, 0);
  return Math.round((b - a) / 86400000);
}

function addDays(d, n) {
  var result = new Date(d);
  result.setDate(result.getDate() + n);
  return result;
}

function addMonths(d, n) {
  var result = new Date(d);
  result.setMonth(result.getMonth() + n);
  return result;
}

function getSettingValue(key) {
  var sheet = getSheet("Settings");
  if (!sheet) return "";
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (safeString(data[i][0]).trim() === key) {
      return data[i][1];
    }
  }
  return "";
}

function setSettingValue(key, value) {
  var sheet = getSheet("Settings");
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (safeString(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function successResponse(data) {
  var resp = { success: true };
  if (data) {
    var keys = Object.keys(data);
    for (var i = 0; i < keys.length; i++) {
      resp[keys[i]] = data[keys[i]];
    }
  }
  return jsonResponse(resp);
}

function errorResponse(msg) {
  return jsonResponse({ success: false, error: msg });
}


// ============================================
// TWILIO FUNCTIONS
// ============================================

function sendTwilioSMS(to, body) {
  var url = "https://api.twilio.com/2010-04-01/Accounts/" + TWILIO_SID + "/Messages.json";
  var payload = {
    To: formatPhone(to),
    From: TWILIO_NUMBER,
    Body: body
  };
  var options = {
    method: "post",
    payload: payload,
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(TWILIO_SID + ":" + TWILIO_TOKEN)
    },
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function sendTwilioVerify(phone) {
  var url = "https://verify.twilio.com/v2/Services/" + TWILIO_VERIFY_SID + "/Verifications";
  var payload = {
    To: formatPhone(phone),
    Channel: "sms"
  };
  var options = {
    method: "post",
    payload: payload,
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(TWILIO_SID + ":" + TWILIO_TOKEN)
    },
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function checkTwilioVerify(phone, code) {
  var url = "https://verify.twilio.com/v2/Services/" + TWILIO_VERIFY_SID + "/VerificationCheck";
  var payload = {
    To: formatPhone(phone),
    Code: code
  };
  var options = {
    method: "post",
    payload: payload,
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(TWILIO_SID + ":" + TWILIO_TOKEN)
    },
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}


// ============================================
// EMAIL FUNCTIONS
// ============================================

function patientLink(name) {
  return ADMIN_URL + "#patient/" + encodeURIComponent(name);
}

function patientLinkHtml(name) {
  return '<a href="' + patientLink(name) + '" style="color:#E8891A;text-decoration:none;font-weight:600;">' + name + '</a>';
}

function sendBrandedEmail(to, subject, htmlBody) {
  var tomEmail = safeString(getSettingValue("Tom Email"));
  var colinEmail = safeString(getSettingValue("Colin Email"));
  var recipients = [];
  if (to === "both") {
    if (tomEmail) recipients.push(tomEmail);
    if (colinEmail) recipients.push(colinEmail);
  } else if (to === "tom") {
    if (tomEmail) recipients.push(tomEmail);
  } else if (to === "colin") {
    if (colinEmail) recipients.push(colinEmail);
  } else {
    recipients.push(to);
  }
  if (recipients.length === 0) return;
  var fullHtml = "<div style='font-family:Arial,sans-serif;max-width:600px;margin:0 auto;'>"
    + "<div style='background:#080808;padding:20px;text-align:center;'>"
    + "<img src=\"https://framemedicine.com/wp-content/uploads/2025/08/Untitled-design.png\" style=\"height:40px;\" />"
    + "</div>"
    + "<div style='background:#131313;color:#ffffff;padding:20px;'>"
    + htmlBody
    + "</div>"
    + "<div style='background:#080808;padding:15px;text-align:center;color:rgba(255,255,255,0.35);font-size:12px;'>"
    + "FRAME Medicine"
    + "</div></div>";
  MailApp.sendEmail({
    to: recipients.join(","),
    subject: "FRAME: " + subject,
    htmlBody: fullHtml
  });
}


// ============================================
// PUSH NOTIFICATION FUNCTIONS
// ============================================

function sendPushNotification(patientName, title, body) {
  var row = findRowByValue("Patients", P_NAME, patientName);
  if (row === -1) return false;
  var sheet = getSheet("Patients");
  var data = sheet.getRange(row, 1, 1, 32).getValues()[0];
  if (safeString(data[P_PUSH]) !== "YES") return false;
  var subJson = safeString(data[P_PUSHSUB]);
  if (!subJson) return false;
  // Push notification sending requires a server with web-push library
  // For Apps Script, we log the intent and fall back to SMS
  // In production, this would call an external push service endpoint
  return false;
}

// Communication preferences: "all" (default), "email", "sms", "app", "email,sms", "none"
// Returns { sms: true/false, email: true/false, app: true/false }
function getCommPref(patientName) {
  var defaults = { sms: true, email: true, app: true };
  if (!patientName) return defaults;
  var pRow = findRowByValue("Patients", P_NAME, patientName);
  if (pRow === -1) return defaults;
  var pref = safeString(getSheet("Patients").getRange(pRow, P_COMMPREF + 1).getValue()).toLowerCase().trim();
  if (!pref || pref === "all") return defaults;
  // Parse comma-separated: "email,sms" or single: "sms"
  var channels = pref.split(",");
  var result = {
    sms: channels.indexOf("sms") !== -1 || channels.indexOf("text") !== -1,
    email: channels.indexOf("email") !== -1,
    app: channels.indexOf("app") !== -1 || channels.indexOf("push") !== -1
  };
  // Must have at least one channel — fall back to SMS if they somehow opt out of everything
  if (!result.sms && !result.email && !result.app) result.sms = true;
  return result;
}

function smartSendMessage(patientName, phone, title, body, isCritical) {
  var prefs = getCommPref(patientName);
  // Critical messages always send via SMS
  if (!isCritical && !prefs.sms && !prefs.app) return false;

  var pushSent = false;
  if (patientName && prefs.app) {
    pushSent = sendPushNotification(patientName, title, body);
  }
  if (prefs.sms && (!pushSent || isCritical)) {
    sendTwilioSMS(phone, body);
  }
  return true;
}

// Check if we should send email notifications for this patient
function shouldEmailForPatient(patientName) {
  if (!patientName) return true; // default yes for unknown
  var prefs = getCommPref(patientName);
  return prefs.email;
}


// ============================================
// doGet — READ OPERATIONS
// ============================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
  try {
    if (action === "login") return handleLogin(e.parameter);
    if (action === "verifyOtp") return handleVerifyOtp(e.parameter);
    if (action === "getPatient") return handleGetPatient(e.parameter);
    if (action === "getPatientDashboard") return handleGetPatientDashboard(e.parameter);
    if (action === "getMessages") return handleGetMessages(e.parameter);
    if (action === "getWeightLog") return handleGetWeightLog(e.parameter);
    if (action === "getCheckIns") return handleGetCheckIns(e.parameter);
    if (action === "getLabStatus") return handleGetLabStatus(e.parameter);
    if (action === "verifyBiometric") return handleVerifyBiometric(e.parameter);
    if (action === "adminLogin") return handleAdminLogin(e.parameter);
    if (action === "getDashboard") return handleGetDashboard(e.parameter);
    if (action === "getPatientDetail") return handleGetPatientDetail(e.parameter);
    if (action === "getLeads") return handleGetLeads(e.parameter);
    if (action === "getLeadDetail") return handleGetLeadDetail(e.parameter);
    if (action === "getInbox") return handleGetInbox(e.parameter);
    if (action === "getConversation") return handleGetConversation(e.parameter);
    if (action === "getPnl") return handleGetPnl(e.parameter);
    if (action === "getSettings") return handleGetSettings(e.parameter);
    if (action === "getLabsDashboard") return handleGetLabsDashboard(e.parameter);
    if (action === "getBillingDue") return handleGetBillingDue(e.parameter);
    if (action === "getCatalog") return handleGetCatalog(e.parameter);
    if (action === "getRefillLog") return handleGetRefillLog(e.parameter);
    if (action === "getDoseHistory") return handleGetDoseHistory(e.parameter);
    return errorResponse("Unknown action: " + action);
  } catch (err) {
    return errorResponse("doGet error: " + err.message);
  }
}


// ============================================
// doPost — WRITE OPERATIONS + TWILIO INBOUND
// ============================================

function doPost(e) {
  try {
    var body = e.postData ? e.postData.contents : "";
    // Check if JSON (from our apps) or form-encoded (from Twilio)
    var firstChar = body.charAt(0);
    if (firstChar === "{" || firstChar === "[") {
      // JSON from our apps
      var data = JSON.parse(body);
      var action = data.action || "";
      if (action === "submitCheckIn") return handleSubmitCheckIn(data);
      if (action === "logWeight") return handleLogWeight(data);
      if (action === "confirmRefill") return handleConfirmRefill(data);
      if (action === "declineRefill") return handleDeclineRefill(data);
      if (action === "sendMessage") return handleSendMessage(data);
      if (action === "markRead") return handleMarkRead(data);
      // ---- Admin write actions (require admin token) ----
      var adminWriteActions = ["savePatient","updateStatus","markPaid","editBilling",
        "newOrder","markLabDone","sendLabReminder","saveLead","convertLead",
        "updateLeadStage","importPatients","importSales","updateSettings",
        "lockMonth","updateOverhead","addOverheadItem","removeOverheadItem",
        "clearFollowUp","saveNotes","logDoseChange","diffPatients","diffSales"];
      if (adminWriteActions.indexOf(action) !== -1) {
        var adminWriteErr = requireAdmin(data);
        if (adminWriteErr) return adminWriteErr;
      }
      if (action === "savePatient") return handleSavePatient(data);
      if (action === "updateStatus") return handleUpdateStatus(data);
      if (action === "markPaid") return handleMarkPaid(data);
      if (action === "editBilling") return handleEditBilling(data);
      if (action === "newOrder") return handleNewOrder(data);
      if (action === "markLabDone") return handleMarkLabDone(data);
      if (action === "sendLabReminder") return handleSendLabReminder(data);
      if (action === "saveLead") return handleSaveLead(data);
      if (action === "convertLead") return handleConvertLead(data);
      if (action === "updateLeadStage") return handleUpdateLeadStage(data);
      if (action === "importPatients") return handleImportPatients(data);
      if (action === "importSales") return handleImportSales(data);
      if (action === "updateSettings") return handleUpdateSettings(data);
      if (action === "lockMonth") return handleLockMonth(data);
      if (action === "updateOverhead") return handleUpdateOverhead(data);
      if (action === "addOverheadItem") return handleAddOverheadItem(data);
      if (action === "removeOverheadItem") return handleRemoveOverheadItem(data);
      if (action === "clearFollowUp") return handleClearFollowUp(data);
      if (action === "saveNotes") return handleSaveNotes(data);
      if (action === "savePushSubscription") return handleSavePushSubscription(data);
      if (action === "logDoseChange") return handleLogDoseChange(data);
      if (action === "saveBiometric") return handleSaveBiometric(data);
      if (action === "revokeBiometric") return handleRevokeBiometric(data);
      if (action === "diffPatients") return handleDiffPatients(data);
      if (action === "diffSales") return handleDiffSales(data);
      if (action === "importMemberships") return handleImportMemberships(data);

      // ---- Patient app reads (no admin auth needed) ----
      if (action === "login") return handleLogin(data);
      if (action === "verifyOtp") return handleVerifyOtp(data);
      if (action === "verifyBiometric") return handleVerifyBiometric(data);
      if (action === "getPatient") return handleGetPatient(data);
      if (action === "getPatientDashboard") return handleGetPatientDashboard(data);
      if (action === "getMessages") return handleGetMessages(data);
      if (action === "getWeightLog") return handleGetWeightLog(data);
      if (action === "getCheckIns") return handleGetCheckIns(data);
      if (action === "getLabStatus") return handleGetLabStatus(data);

      // ---- Admin reads (require admin token) ----
      if (action === "adminLogin") return handleAdminLogin(data);
      var adminActions = ["getDashboard","getPatientDetail","getLeads","getLeadDetail",
        "getInbox","getConversation","getPnl","getSettings","getLabsDashboard",
        "getBillingDue","getCatalog","getRefillLog","getDoseHistory"];
      if (adminActions.indexOf(action) !== -1) {
        var authErr = requireAdmin(data);
        if (authErr) return authErr;
      }
      if (action === "getDashboard") return handleGetDashboard(data);
      if (action === "getPatientDetail") return handleGetPatientDetail(data);
      if (action === "getLeads") return handleGetLeads(data);
      if (action === "getLeadDetail") return handleGetLeadDetail(data);
      if (action === "getInbox") return handleGetInbox(data);
      if (action === "getConversation") return handleGetConversation(data);
      if (action === "getPnl") return handleGetPnl(data);
      if (action === "getSettings") return handleGetSettings(data);
      if (action === "getLabsDashboard") return handleGetLabsDashboard(data);
      if (action === "getBillingDue") return handleGetBillingDue(data);
      if (action === "getCatalog") return handleGetCatalog(data);
      if (action === "getRefillLog") return handleGetRefillLog(data);
      if (action === "getDoseHistory") return handleGetDoseHistory(data);

      return errorResponse("Unknown action: " + action);
    } else {
      // Twilio inbound SMS (form-encoded)
      return handleTwilioInbound(e);
    }
  } catch (err) {
    return errorResponse("doPost error: " + err.message);
  }
}


// ============================================
// PATIENT APP — GET HANDLERS
// ============================================

function handleLogin(params) {
  var phone = formatPhone(params.phone || "");
  if (!phone) return errorResponse("Phone number required");
  // Check if patient exists
  var row = findRowByPhone("Patients", P_PHONE, phone);
  if (row === -1) return errorResponse("No account found for this number");
  // Send OTP
  var result = sendTwilioVerify(phone);
  if (result.status === "pending") {
    return successResponse({ message: "OTP sent" });
  }
  return errorResponse("Failed to send OTP: " + (result.message || "unknown error"));
}

function handleVerifyOtp(params) {
  var phone = formatPhone(params.phone || "");
  var code = String(params.code || "").replace(/\D/g, "");
  if (!phone || !code) return errorResponse("Phone and code required");
  var result = checkTwilioVerify(phone, code);
  Logger.log("Verify result for " + phone + ": " + JSON.stringify(result));
  if (result.status === "approved" || result.valid === true) {
    // Get patient data
    var row = findRowByPhone("Patients", P_PHONE, phone);
    if (row === -1) return errorResponse("Patient not found");
    var sheet = getSheet("Patients");
    var data = sheet.getRange(row, 1, 1, 32).getValues()[0];
    return successResponse({
      verified: true,
      patient: buildPatientObj(data)
    });
  }
  return errorResponse("Invalid code — " + (result.status || result.message || "verification failed"));
}

function handleSaveBiometric(data) {
  var phone = formatPhone(data.phone || "");
  var token = data.token || "";
  if (!phone || !token) return errorResponse("Phone and token required");
  var row = findRowByPhone("Patients", P_PHONE, phone);
  if (row === -1) return errorResponse("Patient not found");
  var sheet = getSheet("Patients");
  sheet.getRange(row, P_BIOTOKEN + 1).setValue(token);
  sheet.getRange(row, P_BIOTOKEN_DATE + 1).setValue(new Date());
  return successResponse({ message: "Biometric token saved" });
}

function handleRevokeBiometric(data) {
  var phone = formatPhone(data.phone || "");
  if (!phone) return errorResponse("Phone required");
  var row = findRowByPhone("Patients", P_PHONE, phone);
  if (row === -1) return errorResponse("Patient not found");
  var sheet = getSheet("Patients");
  sheet.getRange(row, P_BIOTOKEN + 1).setValue("");
  sheet.getRange(row, P_BIOTOKEN_DATE + 1).setValue("");
  return successResponse({ message: "Biometric token revoked" });
}

function handleVerifyBiometric(params) {
  var token = params.token || "";
  if (!token) return errorResponse("Token required");
  var patients = getSheetData("Patients");
  for (var i = 0; i < patients.length; i++) {
    if (safeString(patients[i][P_BIOTOKEN]) === token) {
      // Check 90-day expiry
      var created = parseDate(patients[i][P_BIOTOKEN_DATE]);
      if (created && daysBetween(created, new Date()) > 90) {
        var s = getSheet("Patients");
        s.getRange(i + 2, P_BIOTOKEN + 1).setValue("");
        s.getRange(i + 2, P_BIOTOKEN_DATE + 1).setValue("");
        return errorResponse("Token expired. Please sign in with your phone number.");
      }
      var sheet = getSheet("Patients");
      var data = sheet.getRange(i + 2, 1, 1, 32).getValues()[0];
      return successResponse({ verified: true, patient: buildPatientObj(data) });
    }
  }
  return errorResponse("Invalid biometric token");
}

function handleGetPatient(params) {
  var phone = formatPhone(params.phone || "");
  if (!phone) return errorResponse("Phone required");
  var row = findRowByPhone("Patients", P_PHONE, phone);
  if (row === -1) return errorResponse("Patient not found");
  var sheet = getSheet("Patients");
  var data = sheet.getRange(row, 1, 1, 32).getValues()[0];
  return successResponse({ patient: buildPatientObj(data) });
}

function handleGetPatientDashboard(params) {
  var phone = formatPhone(params.phone || "");
  if (!phone) return errorResponse("Phone required");
  var row = findRowByPhone("Patients", P_PHONE, phone);
  if (row === -1) return errorResponse("Patient not found");
  var sheet = getSheet("Patients");
  var pData = sheet.getRange(row, 1, 1, 32).getValues()[0];
  var patient = buildPatientObj(pData);
  var name = safeString(pData[P_NAME]);

  // Get latest orders for supply info
  var orders = getPatientOrders(name);
  var latestOrder = orders.length > 0 ? orders[0] : null;
  var daysLeft = 0;
  var totalDays = 0;
  if (latestOrder) {
    var nextDue = parseDate(latestOrder.nextDue);
    if (nextDue) {
      daysLeft = daysBetween(new Date(), nextDue);
      totalDays = safeNumber(latestOrder.daysCovered);
    }
  }

  // Get last 8 weights
  var weights = getPatientWeights(name, 8);

  // Get last check-in
  var checkIns = getPatientCheckIns(name, 1);

  // Get lab status
  var labs = getPatientLabs(name);

  // Get unread message count
  var unreadCount = getUnreadCount(phone);

  return successResponse({
    patient: patient,
    supply: {
      daysLeft: daysLeft,
      totalDays: totalDays,
      lastOrder: latestOrder
    },
    weights: weights,
    lastCheckIn: checkIns.length > 0 ? checkIns[0] : null,
    labs: labs,
    unreadMessages: unreadCount
  });
}

function handleGetMessages(params) {
  var phone = formatPhone(params.phone || "");
  if (!phone) return errorResponse("Phone required");
  var allMessages = getSheetData("Messages");
  var messages = [];
  for (var i = 0; i < allMessages.length; i++) {
    if (formatPhone(safeString(allMessages[i][MSG_PHONE])) === phone) {
      messages.push({
        timestamp: formatDateStr(allMessages[i][MSG_TIMESTAMP]),
        timestampISO: formatDateISO(allMessages[i][MSG_TIMESTAMP]),
        patient: safeString(allMessages[i][MSG_PATIENT]),
        direction: safeString(allMessages[i][MSG_DIRECTION]),
        text: safeString(allMessages[i][MSG_TEXT]),
        read: safeString(allMessages[i][MSG_READ]),
        source: safeString(allMessages[i][MSG_SOURCE])
      });
    }
  }
  messages.sort(function(a, b) {
    return new Date(a.timestampISO) - new Date(b.timestampISO);
  });
  return successResponse({ messages: messages });
}

function handleGetWeightLog(params) {
  var name = params.name || "";
  var limit = safeNumber(params.limit) || 50;
  if (!name) return errorResponse("Patient name required");
  var weights = getPatientWeights(name, limit);
  return successResponse({ weights: weights });
}

function handleGetCheckIns(params) {
  var name = params.name || "";
  var limit = safeNumber(params.limit) || 50;
  if (!name) return errorResponse("Patient name required");
  var checkIns = getPatientCheckIns(name, limit);
  return successResponse({ checkIns: checkIns });
}

function handleGetLabStatus(params) {
  var name = params.name || "";
  if (!name) return errorResponse("Patient name required");
  var labs = getPatientLabs(name);
  return successResponse({ labs: labs });
}


// ============================================
// ADMIN APP — GET HANDLERS
// ============================================

function handleAdminLogin(params) {
  var email = safeString(params.email).trim().toLowerCase();
  var password = safeString(params.password);
  if (!email || !password) return errorResponse("Email and password required");

  var tomEmail = safeString(getSettingValue("Tom Email")).trim().toLowerCase();
  var colinEmail = safeString(getSettingValue("Colin Email")).trim().toLowerCase();
  var tomPw = safeString(getSettingValue("Tom Password"));
  var colinPw = safeString(getSettingValue("Colin Password"));

  var role = "";
  var name = "";
  if (email === tomEmail && password === tomPw) {
    role = "tom"; name = "Tom";
  } else if (email === colinEmail && password === colinPw) {
    role = "colin"; name = "Dr. Sheffield";
  } else {
    return errorResponse("Invalid credentials");
  }

  // Generate session token
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  cache.put("admin_session_" + token, JSON.stringify({ role: role, name: name, email: email }), 28800); // 8 hours
  return successResponse({ role: role, name: name, email: email, adminToken: token });
}

function verifyAdminToken(params) {
  var token = params.adminToken || "";
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var session = cache.get("admin_session_" + token);
  if (!session) return null;
  return JSON.parse(session);
}

function requireAdmin(params) {
  var session = verifyAdminToken(params);
  if (!session) return errorResponse("Unauthorized — please log in");
  return null; // null means authorized
}

function handleGetDashboard(params) {
  var patients = getSheetData("Patients");
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var stats = {
    active: 0,
    overdue: 0,
    dueSoon: 0,
    followUp: 0,
    labsDue: 0,
    unpaid: 0
  };

  var allOrders = getAllPatientOrders();

  var patientList = [];
  for (var i = 0; i < patients.length; i++) {
    var p = patients[i];
    var status = safeString(p[P_STATUS]);
    if (status === "INACTIVE" || status === "Staff") continue;

    var name = safeString(p[P_NAME]);
    var med = safeString(p[P_MED]);
    var phone = safeString(p[P_PHONE]);

    // Get latest order for supply info (from pre-loaded map)
    var orderEntry = allOrders[name.toLowerCase()];
    var orders = orderEntry ? orderEntry.orders : [];
    var daysLeft = null;
    var nextDue = null;
    if (orders.length > 0) {
      var nd = parseDate(orders[0].nextDue);
      if (nd) {
        daysLeft = daysBetween(today, nd);
        nextDue = formatDateStr(nd);
      }
    }

    // Count stats
    if (status === "Active") stats.active++;
    if (daysLeft !== null && daysLeft < 0) stats.overdue++;
    if (daysLeft !== null && daysLeft >= 0 && daysLeft <= 14) stats.dueSoon++;
    if (safeString(p[P_FOLLOWUP]) === "YES") stats.followUp++;
    if (safeNumber(p[P_OUTSTANDING]) > 0) stats.unpaid++;

    var badge = "on-track";
    if (daysLeft !== null && daysLeft < 0) badge = "overdue";
    else if (daysLeft !== null && daysLeft <= 14) badge = "due-soon";

    patientList.push({
      name: name,
      preferredName: safeString(p[P_PREFERRED]),
      medication: med,
      phone: phone,
      status: status,
      daysLeft: daysLeft,
      nextDue: nextDue,
      badge: badge,
      followUp: safeString(p[P_FOLLOWUP]) === "YES",
      outstanding: safeNumber(p[P_OUTSTANDING]),
      plan: safeString(p[P_PLAN]),
      rate: safeNumber(p[P_RATE])
    });
  }

  // Check labs due
  var labs = getSheetData("Labs");
  for (var j = 0; j < labs.length; j++) {
    var labNextDue = parseDate(labs[j][L_NEXT_DUE]);
    if (labNextDue) {
      var labDays = daysBetween(today, labNextDue);
      if (labDays <= 30) stats.labsDue++;
    }
  }

  // Sort: overdue first, then due-soon, then on-track
  patientList.sort(function(a, b) {
    var order = { overdue: 0, "due-soon": 1, "on-track": 2 };
    var aOrder = order[a.badge] !== undefined ? order[a.badge] : 2;
    var bOrder = order[b.badge] !== undefined ? order[b.badge] : 2;
    if (aOrder !== bOrder) return aOrder - bOrder;
    if (a.daysLeft === null && b.daysLeft === null) return 0;
    if (a.daysLeft === null) return 1;
    if (b.daysLeft === null) return -1;
    return a.daysLeft - b.daysLeft;
  });

  return successResponse({
    stats: stats,
    patients: patientList
  });
}

function handleGetPatientDetail(params) {
  var name = params.name || "";
  if (!name) return errorResponse("Patient name required");

  var row = findRowByValue("Patients", P_NAME, name);
  if (row === -1) return errorResponse("Patient not found");

  var sheet = getSheet("Patients");
  var pData = sheet.getRange(row, 1, 1, 32).getValues()[0];
  var patient = buildPatientObj(pData);

  // Billing
  var billing = getPatientBilling(name);

  // Orders
  var orders = getPatientOrders(name);

  // Check-ins
  var checkIns = getPatientCheckIns(name, 20);

  // Weights
  var weights = getPatientWeights(name, 20);

  // Messages
  var phone = formatPhone(safeString(pData[P_PHONE]));
  var messages = [];
  var allMsg = getSheetData("Messages");
  for (var i = 0; i < allMsg.length; i++) {
    if (formatPhone(safeString(allMsg[i][MSG_PHONE])) === phone) {
      messages.push({
        timestamp: formatDateStr(allMsg[i][MSG_TIMESTAMP]),
        timestampISO: formatDateISO(allMsg[i][MSG_TIMESTAMP]),
        direction: safeString(allMsg[i][MSG_DIRECTION]),
        text: safeString(allMsg[i][MSG_TEXT]),
        read: safeString(allMsg[i][MSG_READ]),
        source: safeString(allMsg[i][MSG_SOURCE])
      });
    }
  }
  messages.sort(function(a, b) {
    return new Date(a.timestampISO) - new Date(b.timestampISO);
  });

  // Dose history
  var doseHistory = getPatientDoseHistory(name);

  // Refill log (last 20)
  var refillLog = getPatientRefillLog(name, 20);

  // Labs
  var labs = getPatientLabs(name);

  return successResponse({
    patient: patient,
    billing: billing,
    orders: orders,
    checkIns: checkIns,
    weights: weights,
    messages: messages,
    doseHistory: doseHistory,
    refillLog: refillLog,
    labs: labs
  });
}

function handleGetLeads(params) {
  var data = getSheetData("Leads");
  var leads = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (safeString(row[LD_CONVERTED]) === "YES") continue; // skip converted by default
    var inquiryDate = parseDate(row[LD_DATE]);
    var daysSince = inquiryDate ? daysBetween(inquiryDate, new Date()) : 0;
    leads.push({
      name: safeString(row[LD_NAME]),
      phone: safeString(row[LD_PHONE]),
      email: safeString(row[LD_EMAIL]),
      source: safeString(row[LD_SOURCE]),
      inquiryDate: formatDateStr(row[LD_DATE]),
      interest: safeString(row[LD_INTEREST]),
      stage: safeString(row[LD_STAGE]),
      assignedTo: safeString(row[LD_ASSIGNED]),
      lastContact: formatDateStr(row[LD_LASTCONTACT]),
      nextFollowUp: formatDateStr(row[LD_NEXTFOLLOWUP]),
      notes: safeString(row[LD_NOTES]),
      daysSinceInquiry: daysSince,
      converted: safeString(row[LD_CONVERTED])
    });
  }
  if (params && params.includeConverted === "true") {
    // Re-fetch including converted
    leads = [];
    for (var j = 0; j < data.length; j++) {
      var r = data[j];
      var iDate = parseDate(r[LD_DATE]);
      var ds = iDate ? daysBetween(iDate, new Date()) : 0;
      leads.push({
        name: safeString(r[LD_NAME]),
        phone: safeString(r[LD_PHONE]),
        email: safeString(r[LD_EMAIL]),
        source: safeString(r[LD_SOURCE]),
        inquiryDate: formatDateStr(r[LD_DATE]),
        interest: safeString(r[LD_INTEREST]),
        stage: safeString(r[LD_STAGE]),
        assignedTo: safeString(r[LD_ASSIGNED]),
        lastContact: formatDateStr(r[LD_LASTCONTACT]),
        nextFollowUp: formatDateStr(r[LD_NEXTFOLLOWUP]),
        notes: safeString(r[LD_NOTES]),
        daysSinceInquiry: ds,
        converted: safeString(r[LD_CONVERTED])
      });
    }
  }
  return successResponse({ leads: leads });
}

function handleGetLeadDetail(params) {
  var name = params.name || "";
  if (!name) return errorResponse("Lead name required");
  var row = findRowByValue("Leads", LD_NAME, name);
  if (row === -1) return errorResponse("Lead not found");
  var sheet = getSheet("Leads");
  var d = sheet.getRange(row, 1, 1, 14).getValues()[0];
  var lead = {
    name: safeString(d[LD_NAME]),
    phone: safeString(d[LD_PHONE]),
    email: safeString(d[LD_EMAIL]),
    source: safeString(d[LD_SOURCE]),
    inquiryDate: formatDateStr(d[LD_DATE]),
    interest: safeString(d[LD_INTEREST]),
    stage: safeString(d[LD_STAGE]),
    assignedTo: safeString(d[LD_ASSIGNED]),
    lastContact: formatDateStr(d[LD_LASTCONTACT]),
    nextFollowUp: formatDateStr(d[LD_NEXTFOLLOWUP]),
    notes: safeString(d[LD_NOTES]),
    converted: safeString(d[LD_CONVERTED]),
    convertedDate: formatDateStr(d[LD_CONVERTEDDATE]),
    patientName: safeString(d[LD_PATIENTNAME])
  };
  // Get messages for this lead
  var phone = formatPhone(safeString(d[LD_PHONE]));
  var messages = [];
  var allMsg = getSheetData("Messages");
  for (var i = 0; i < allMsg.length; i++) {
    if (formatPhone(safeString(allMsg[i][MSG_PHONE])) === phone) {
      messages.push({
        timestamp: formatDateStr(allMsg[i][MSG_TIMESTAMP]),
        timestampISO: formatDateISO(allMsg[i][MSG_TIMESTAMP]),
        direction: safeString(allMsg[i][MSG_DIRECTION]),
        text: safeString(allMsg[i][MSG_TEXT]),
        read: safeString(allMsg[i][MSG_READ]),
        source: safeString(allMsg[i][MSG_SOURCE])
      });
    }
  }
  messages.sort(function(a, b) {
    return new Date(a.timestampISO) - new Date(b.timestampISO);
  });
  return successResponse({ lead: lead, messages: messages });
}

function handleGetInbox(params) {
  var allMsg = getSheetData("Messages");
  var conversations = {};
  for (var i = 0; i < allMsg.length; i++) {
    var phone = formatPhone(safeString(allMsg[i][MSG_PHONE]));
    if (!phone) continue;
    if (!conversations[phone]) {
      conversations[phone] = {
        name: safeString(allMsg[i][MSG_PATIENT]),
        phone: phone,
        contactType: safeString(allMsg[i][MSG_CONTACTTYPE]) || "patient",
        lastMessage: "",
        lastTimestamp: "",
        lastTimestampISO: "",
        unread: 0,
        direction: ""
      };
    }
    var ts = allMsg[i][MSG_TIMESTAMP];
    var tsISO = formatDateISO(ts);
    var existing = conversations[phone].lastTimestampISO;
    if (!existing || tsISO > existing) {
      conversations[phone].lastMessage = safeString(allMsg[i][MSG_TEXT]);
      conversations[phone].lastTimestamp = formatDateStr(ts);
      conversations[phone].lastTimestampISO = tsISO;
      conversations[phone].direction = safeString(allMsg[i][MSG_DIRECTION]);
    }
    if (safeString(allMsg[i][MSG_DIRECTION]) === "inbound" && safeString(allMsg[i][MSG_READ]) !== "Yes") {
      conversations[phone].unread++;
    }
  }
  var list = [];
  var keys = Object.keys(conversations);
  for (var j = 0; j < keys.length; j++) {
    list.push(conversations[keys[j]]);
  }
  list.sort(function(a, b) {
    return (b.lastTimestampISO || "").localeCompare(a.lastTimestampISO || "");
  });
  return successResponse({ conversations: list });
}

function handleGetConversation(params) {
  var phone = formatPhone(params.phone || "");
  if (!phone) return errorResponse("Phone required");
  var allMsg = getSheetData("Messages");
  var messages = [];
  for (var i = 0; i < allMsg.length; i++) {
    if (formatPhone(safeString(allMsg[i][MSG_PHONE])) === phone) {
      messages.push({
        timestamp: formatDateStr(allMsg[i][MSG_TIMESTAMP]),
        timestampISO: formatDateISO(allMsg[i][MSG_TIMESTAMP]),
        patient: safeString(allMsg[i][MSG_PATIENT]),
        direction: safeString(allMsg[i][MSG_DIRECTION]),
        text: safeString(allMsg[i][MSG_TEXT]),
        read: safeString(allMsg[i][MSG_READ]),
        source: safeString(allMsg[i][MSG_SOURCE]),
        contactType: safeString(allMsg[i][MSG_CONTACTTYPE])
      });
    }
  }
  messages.sort(function(a, b) {
    return new Date(a.timestampISO) - new Date(b.timestampISO);
  });
  return successResponse({ messages: messages });
}

function handleGetPnl(params) {
  var today = new Date();
  var requestedMonth = safeNumber(params.month);
  var requestedYear = safeNumber(params.year);

  // Default to previous month if in first 5 days
  if (!requestedMonth && !requestedYear) {
    if (today.getDate() <= 5) {
      var prev = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      requestedMonth = prev.getMonth(); // 0-indexed
      requestedYear = prev.getFullYear();
    } else {
      requestedMonth = today.getMonth();
      requestedYear = today.getFullYear();
    }
  }

  var monthNames = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];
  var monthLabel = monthNames[requestedMonth] + " " + requestedYear;

  // Check Finance tab for locked data
  var financeData = getSheetData("Finance");
  var lockedRow = null;
  for (var i = 0; i < financeData.length; i++) {
    if (safeNumber(financeData[i][FIN_YEAR]) === requestedYear &&
        safeNumber(financeData[i][FIN_MONTHNUM]) === requestedMonth) {
      lockedRow = financeData[i];
      break;
    }
  }

  if (lockedRow && safeString(lockedRow[FIN_LOCKED]) === "YES") {
    // Return locked data
    return successResponse({
      month: requestedMonth,
      year: requestedYear,
      monthLabel: monthLabel,
      locked: true,
      revenue: safeNumber(lockedRow[FIN_REVENUE]),
      medCosts: safeNumber(lockedRow[FIN_MEDCOSTS]),
      overhead: safeNumber(lockedRow[FIN_OVERHEAD]),
      netProfit: safeNumber(lockedRow[FIN_NET]),
      tomShare: safeNumber(lockedRow[FIN_TOM]),
      colinShare: safeNumber(lockedRow[FIN_COLIN]),
      overheadItems: getOverheadItems(requestedMonth, requestedYear),
      orderBreakdown: [],
      revenueBreakdown: []
    });
  }

  // Calculate live data
  var revenue = calculateMonthlyRevenue(requestedMonth, requestedYear);
  var medCosts = calculateMonthlyMedCosts(requestedMonth, requestedYear);
  var overheadResult = calculateMonthlyOverhead(requestedMonth, requestedYear);
  var overhead = overheadResult.total;
  var netProfit = revenue.total - medCosts.total - overhead;

  var tomSplit = safeNumber(getSettingValue("Tom Split")) || 0.25;
  var colinSplit = safeNumber(getSettingValue("Colin Split")) || 0.75;

  // Calculate YTD totals (Jan through current month of requested year)
  var ytdRevenue = 0, ytdMedCosts = 0, ytdOverhead = 0;
  var ytdMonths = [];
  for (var m = 0; m <= requestedMonth; m++) {
    var mRev = calculateMonthlyRevenue(m, requestedYear);
    var mCost = calculateMonthlyMedCosts(m, requestedYear);
    var mOver = calculateMonthlyOverhead(m, requestedYear);
    ytdRevenue += mRev.total;
    ytdMedCosts += mCost.total;
    ytdOverhead += mOver.total;
    ytdMonths.push({
      month: monthNames[m].substring(0, 3),
      revenue: mRev.total,
      costs: mCost.total + mOver.total,
      net: mRev.total - mCost.total - mOver.total
    });
  }
  var ytdNet = ytdRevenue - ytdMedCosts - ytdOverhead;

  return successResponse({
    month: requestedMonth,
    year: requestedYear,
    monthLabel: monthLabel,
    locked: false,
    revenue: revenue.total,
    medCosts: medCosts.total,
    overhead: overhead,
    netProfit: netProfit,
    tomShare: Math.round(netProfit * tomSplit * 100) / 100,
    colinShare: Math.round(netProfit * colinSplit * 100) / 100,
    tomSplit: tomSplit,
    colinSplit: colinSplit,
    overheadItems: overheadResult.items,
    overheadBase: overheadResult.base,
    orderBreakdown: medCosts.breakdown,
    revenueBreakdown: revenue.breakdown,
    ytd: {
      revenue: Math.round(ytdRevenue * 100) / 100,
      medCosts: Math.round(ytdMedCosts * 100) / 100,
      overhead: Math.round(ytdOverhead * 100) / 100,
      netProfit: Math.round(ytdNet * 100) / 100,
      months: ytdMonths
    }
  });
}

function handleGetSettings(params) {
  var sheet = getSheet("Settings");
  if (!sheet) return errorResponse("Settings tab not found");
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 0; i < data.length; i++) {
    var key = safeString(data[i][0]).trim();
    if (key && key.indexOf("(section)") === -1) {
      settings[key] = data[i][1];
    }
  }
  return successResponse({ settings: settings });
}

function handleGetLabsDashboard(params) {
  var labs = getSheetData("Labs");
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var result = [];
  for (var i = 0; i < labs.length; i++) {
    var patient = safeString(labs[i][L_PATIENT]);
    var nextDue = parseDate(labs[i][L_NEXT_DUE]);
    if (!nextDue) continue;
    var daysUntil = daysBetween(today, nextDue);
    var labType = "";
    if (!labs[i][L_INIT_DONE]) labType = "Initial";
    else if (!labs[i][L_90_DONE]) labType = "90-Day";
    else if (!labs[i][L_180_DONE]) labType = "180-Day";
    else if (!labs[i][L_ANN_DONE]) labType = "Annual";
    else labType = "Complete";

    var badge = "on-track";
    if (daysUntil < 0) badge = "overdue";
    else if (daysUntil <= 30) badge = "due-soon";

    result.push({
      patient: patient,
      labType: labType,
      dueDate: formatDateStr(nextDue),
      daysUntil: daysUntil,
      badge: badge,
      status: safeString(labs[i][L_STATUS]),
      notes: safeString(labs[i][L_NOTES])
    });
  }
  result.sort(function(a, b) {
    return a.daysUntil - b.daysUntil;
  });
  return successResponse({ labs: result });
}

function handleGetBillingDue(params) {
  var billing = getSheetData("Billing");
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var due = [];
  for (var i = 0; i < billing.length; i++) {
    var status = safeString(billing[i][S_STATUS]);
    if (status !== "Active" && status !== "Past Due") continue;
    var nextPayDue = parseDate(billing[i][S_NEXTPAYDUE]);
    if (!nextPayDue) continue;
    var daysUntil = daysBetween(today, nextPayDue);
    if (daysUntil <= 7) {
      due.push({
        patient: safeString(billing[i][S_PATIENT]),
        plan: safeString(billing[i][S_PLAN]),
        rate: safeNumber(billing[i][S_RATE]),
        nextPayDue: formatDateStr(nextPayDue),
        daysUntil: daysUntil,
        outstanding: safeNumber(billing[i][S_OUTSTANDING]),
        cyclesLeft: safeNumber(billing[i][S_CYCLES]),
        status: status
      });
    }
  }
  due.sort(function(a, b) { return a.daysUntil - b.daysUntil; });
  return successResponse({ billingDue: due });
}

function handleGetCatalog(params) {
  var data = getSheetData("Catalog");
  var catalog = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    catalog.push({
      productCode: safeString(data[i][0]),
      name: safeString(data[i][1]),
      strength: safeString(data[i][2]),
      form: safeString(data[i][3]),
      category: safeString(data[i][4]),
      unit: safeString(data[i][5]),
      cost: safeNumber(data[i][6])
    });
  }
  return successResponse({ catalog: catalog });
}

function handleGetRefillLog(params) {
  var name = params.name || "";
  var data = getSheetData("Refill Log");
  var log = [];
  for (var i = 0; i < data.length; i++) {
    if (name && safeString(data[i][RL_PATIENT]).toLowerCase() !== name.toLowerCase()) continue;
    log.push({
      timestamp: formatDateStr(data[i][RL_TIMESTAMP]),
      patient: safeString(data[i][RL_PATIENT]),
      medication: safeString(data[i][RL_MED]),
      action: safeString(data[i][RL_ACTION]),
      method: safeString(data[i][RL_METHOD]),
      notes: safeString(data[i][RL_NOTES])
    });
  }
  log.reverse(); // newest first
  if (params.limit) log = log.slice(0, safeNumber(params.limit));
  return successResponse({ log: log });
}

function handleGetDoseHistory(params) {
  var name = params.name || "";
  if (!name) return errorResponse("Patient name required");
  var history = getPatientDoseHistory(name);
  return successResponse({ doseHistory: history });
}


// ============================================
// PATIENT APP — POST HANDLERS
// ============================================

function handleSubmitCheckIn(data) {
  var sheet = getSheet("Check-In Responses");
  if (!sheet) return errorResponse("Check-In Responses tab not found");
  sheet.appendRow([
    new Date(),
    data.patientName || "",
    formatPhone(data.phone || ""),
    data.medication || "",
    data.symptoms || "",
    data.rating || "",
    data.notes || "",
    data.responseRequested ? "Yes" : "No",
    ""
  ]);
  // Update Check-Ins schedule tab
  var ciRow = findRowByValue("Check-Ins", CIS_PATIENT, data.patientName);
  if (ciRow !== -1) {
    var ciSheet = getSheet("Check-Ins");
    ciSheet.getRange(ciRow, CIS_RESPONSE + 1).setValue(data.notes || data.symptoms || "Submitted");
    ciSheet.getRange(ciRow, CIS_RESPONSEDATE + 1).setValue(new Date());
  }
  // Email providers
  var accent = data.responseRequested ? "#ff4444" : "#E8891A";
  var responseFlag = data.responseRequested
    ? "<div style='background:rgba(255,68,68,0.15);border:1px solid rgba(255,68,68,0.4);border-radius:6px;padding:10px;text-align:center;margin-bottom:12px;'><span style='color:#ff4444;font-weight:bold;font-size:13px;letter-spacing:0.1em;text-transform:uppercase;'>RESPONSE REQUESTED</span></div>"
    : "";
  sendBrandedEmail("both",
    (data.responseRequested ? "URGENT: " : "") + "Check-In: " + data.patientName,
    responseFlag
    + "<h2 style='color:" + accent + ";margin:0 0 12px 0;'>" + patientLinkHtml(data.patientName) + " — Check-In</h2>"
    + "<table style='width:100%;border-collapse:collapse;'>"
    + "<tr><td style='padding:8px 0;color:rgba(255,255,255,0.5);width:100px;'>Medication</td><td style='padding:8px 0;color:#fff;'>" + (data.medication || "N/A") + "</td></tr>"
    + "<tr><td style='padding:8px 0;color:rgba(255,255,255,0.5);'>Rating</td><td style='padding:8px 0;color:#fff;'>" + (data.rating || "N/A") + "</td></tr>"
    + "<tr><td style='padding:8px 0;color:rgba(255,255,255,0.5);'>Symptoms</td><td style='padding:8px 0;color:#fff;'>" + (data.symptoms || "None reported") + "</td></tr>"
    + "<tr><td style='padding:8px 0;color:rgba(255,255,255,0.5);'>Notes</td><td style='padding:8px 0;color:#fff;'>" + (data.notes || "None") + "</td></tr>"
    + "</table>"
  );
  // Log to refill log
  appendRefillLog(data.patientName, data.medication, "Check-in submitted", "App", data.notes || "");
  return successResponse({ message: "Check-in submitted" });
}

function handleLogWeight(data) {
  var sheet = getSheet("Weight Log");
  if (!sheet) return errorResponse("Weight Log tab not found");
  var name = data.patientName || "";
  var weight = safeNumber(data.weight);
  if (!name || !weight) return errorResponse("Patient name and weight required");

  // Get previous weights for calculations
  var allWeights = getPatientWeights(name, 999);
  var lastWeight = allWeights.length > 0 ? allWeights[0].weight : weight;
  var startWeight = allWeights.length > 0 ? allWeights[allWeights.length - 1].weight : weight;
  var change = Math.round((weight - lastWeight) * 10) / 10;
  var totalChange = Math.round((weight - startWeight) * 10) / 10;

  // Calculate weeks since start
  var firstDate = allWeights.length > 0 ? new Date(allWeights[allWeights.length - 1].dateISO) : new Date();
  var weeks = Math.max(1, Math.round(daysBetween(firstDate, new Date()) / 7));
  var avgPerWeek = weeks > 0 ? Math.round((totalChange / weeks) * 10) / 10 : 0;

  var logDate = parseDate(data.date) || new Date();

  sheet.appendRow([
    logDate,
    name,
    data.medication || "",
    "",
    weight,
    change,
    avgPerWeek,
    totalChange,
    startWeight,
    weeks,
    data.source || "app"
  ]);
  return successResponse({
    message: "Weight logged",
    change: change,
    totalChange: totalChange
  });
}

function handleConfirmRefill(data) {
  var name = data.patientName || "";
  if (!name) return errorResponse("Patient name required");
  // Log refill
  appendRefillLog(name, data.medication || "", "Refill confirmed", "App", "");
  // Email providers
  sendBrandedEmail("both",
    "Refill Confirmed: " + name,
    "<div style='border-left:4px solid #4CAF50;padding:12px 16px;background:rgba(76,175,80,0.08);border-radius:0 8px 8px 0;'>"
    + "<h2 style='color:#4CAF50;margin:0 0 8px 0;'>" + patientLinkHtml(name) + " — Refill Confirmed</h2>"
    + "<p style='color:rgba(255,255,255,0.6);margin:0;'>" + (data.medication || "N/A") + " — confirmed via app</p>"
    + "</div>"
  );
  return successResponse({ message: "Refill confirmed" });
}

function handleDeclineRefill(data) {
  var name = data.patientName || "";
  if (!name) return errorResponse("Patient name required");
  // Set follow-up flag
  var row = findRowByValue("Patients", P_NAME, name);
  if (row !== -1) {
    var sheet = getSheet("Patients");
    sheet.getRange(row, P_FOLLOWUP + 1).setValue("YES");
    sheet.getRange(row, P_STATUS + 1).setValue("Declined Refill");
  }
  appendRefillLog(name, data.medication || "", "Refill declined: " + (data.reason || ""), "App", data.reason || "");
  sendBrandedEmail("both",
    "REFILL DECLINED: " + name,
    "<div style='border-left:4px solid #ff4444;padding:12px 16px;background:rgba(255,68,68,0.08);border-radius:0 8px 8px 0;'>"
    + "<h2 style='color:#ff4444;margin:0 0 8px 0;'>" + patientLinkHtml(name) + " — Refill Declined</h2>"
    + "<p style='color:#fff;margin:0 0 4px 0;'><strong>Medication:</strong> " + (data.medication || "N/A") + "</p>"
    + "<p style='color:#fff;margin:0 0 4px 0;'><strong>Reason:</strong> " + (data.reason || "No reason given") + "</p>"
    + "<p style='color:#ff4444;margin:8px 0 0 0;font-weight:bold;'>Follow-up flag set</p>"
    + "</div>"
  );
  return successResponse({ message: "Decline recorded" });
}


// ============================================
// SHARED — POST HANDLERS (used by both apps)
// ============================================

function handleSendMessage(data) {
  var name = data.name || data.patientName || "";
  var phone = formatPhone(data.phone || "");
  var text = data.text || data.message || "";
  var source = data.source || "admin";
  var contactType = data.contactType || "patient";
  var alsoSms = data.alsoSms;
  if (!phone || !text) return errorResponse("Phone and message required");

  // Log to Messages tab
  var msgSheet = getSheet("Messages");
  if (msgSheet) {
    msgSheet.appendRow([
      new Date(),
      name,
      phone,
      "outbound",
      text,
      "",
      source,
      contactType
    ]);
  }

  // Send via smart routing (or force SMS)
  if (alsoSms || source === "admin") {
    // Admin messages: try push first, also SMS if toggled or no push
    var pushSent = sendPushNotification(name, "FRAME Medicine", text);
    if (alsoSms || !pushSent) {
      sendTwilioSMS(phone, "FRAME Medicine: " + text);
    }
  } else {
    // Patient app messages just get logged (provider will see in inbox)
    // But also notify via email
    sendBrandedEmail("both",
      "New Message from " + name,
      "<div style='border-left:4px solid #E8891A;padding:12px 16px;background:rgba(232,137,26,0.08);border-radius:0 8px 8px 0;'>"
      + "<h2 style='color:#E8891A;margin:0 0 8px 0;'>" + patientLinkHtml(name) + "</h2>"
      + "<p style='color:#fff;margin:0;'>" + text + "</p>"
      + "</div>"
    );
  }
  return successResponse({ message: "Message sent" });
}

function handleMarkRead(data) {
  var phone = formatPhone(data.phone || "");
  if (!phone) return errorResponse("Phone required");
  var sheet = getSheet("Messages");
  if (!sheet) return errorResponse("Messages tab not found");
  var allData = sheet.getDataRange().getValues();
  for (var i = 1; i < allData.length; i++) {
    if (formatPhone(safeString(allData[i][MSG_PHONE])) === phone &&
        safeString(allData[i][MSG_DIRECTION]) === "inbound" &&
        safeString(allData[i][MSG_READ]) !== "Yes") {
      sheet.getRange(i + 1, MSG_READ + 1).setValue("Yes");
    }
  }
  return successResponse({ message: "Messages marked as read" });
}

function handleSavePushSubscription(data) {
  var phone = formatPhone(data.phone || "");
  var subscription = data.subscription || "";
  if (!phone) return errorResponse("Phone required");
  var row = findRowByPhone("Patients", P_PHONE, phone);
  if (row === -1) return errorResponse("Patient not found");
  var sheet = getSheet("Patients");
  sheet.getRange(row, P_PUSH + 1).setValue("YES");
  sheet.getRange(row, P_PUSHSUB + 1).setValue(
    typeof subscription === "string" ? subscription : JSON.stringify(subscription)
  );
  return successResponse({ message: "Push subscription saved" });
}


// ============================================
// ADMIN — POST HANDLERS
// ============================================

function handleSavePatient(data) {
  var sheet = getSheet("Patients");
  if (!sheet) return errorResponse("Patients tab not found");
  var name = data.name || "";
  if (!name) return errorResponse("Patient name required");

  var existingRow = findRowByValue("Patients", P_NAME, name);
  var isNew = existingRow === -1;

  // If editing and name changed, check by original name
  if (!isNew && data.originalName && data.originalName !== name) {
    existingRow = findRowByValue("Patients", P_NAME, data.originalName);
  }

  var phone = formatPhone(data.phone || "");
  var memberStart = parseDate(data.membershipStart);
  var term = safeNumber(data.term);
  var contractEnd = memberStart && term ? addMonths(memberStart, term) : null;
  var cyclesLeft = term;

  var rowData = [
    name,
    data.preferredName || "",
    parseDate(data.dob) || "",
    phone,
    data.email || "",
    data.street || "",
    data.city || "",
    data.state || "",
    data.zip || "",
    isNew ? new Date() : (parseDate(data.memberSince) || new Date()),
    data.medication || "",
    data.plan || "",
    safeNumber(data.rate),
    term,
    memberStart || "",
    contractEnd || "",
    cyclesLeft,
    safeNumber(data.outstanding),
    data.checkInDay || "",
    data.checkInTime || "",
    data.glpWeightDay || "",
    data.glpWeightTime || "",
    data.status || "Active",
    data.followUp || "",
    data.notes || "",
    data.pushEnabled || "",
    data.pushSubscription || "",
    data.referralSource || "",
    data.referredBy || ""
  ];

  if (isNew) {
    sheet.appendRow(rowData);
    // Also create billing row
    createBillingRow(name, data);
    // Also create labs row
    createLabsRow(name);
    // Also create check-in schedule row
    createCheckInRow(name, data);
    auditLog(data.adminToken, name, "Patient Created", "New patient added");
    return successResponse({ message: "Patient created", isNew: true });
  } else {
    // Preserve biometric and comm pref columns (29-31) that aren't in the save form
    var existing = sheet.getRange(existingRow, P_BIOTOKEN + 1, 1, 3).getValues()[0];
    rowData.push(existing[0], existing[1], existing[2]); // BioToken, BioTokenDate, CommPref
    var range = sheet.getRange(existingRow, 1, 1, rowData.length);
    range.setValues([rowData]);
    auditLog(data.adminToken, name, "Patient Updated", "Record modified");
    return successResponse({ message: "Patient updated", isNew: false });
  }
}

function handleUpdateStatus(data) {
  var name = data.name || "";
  var status = data.status || "";
  if (!name || !status) return errorResponse("Name and status required");
  var row = findRowByValue("Patients", P_NAME, name);
  if (row === -1) return errorResponse("Patient not found");
  var sheet = getSheet("Patients");
  sheet.getRange(row, P_STATUS + 1).setValue(status);
  // Also update billing status if INACTIVE
  if (status === "INACTIVE") {
    var bRow = findRowByValue("Billing", S_PATIENT, name);
    if (bRow !== -1) {
      getSheet("Billing").getRange(bRow, S_STATUS + 1).setValue("Cancelled");
    }
  }
  auditLog(data.adminToken, name, "Status Changed", "Set to " + status);
  return successResponse({ message: "Status updated" });
}

function handleMarkPaid(data) {
  var name = data.name || "";
  if (!name) return errorResponse("Patient name required");
  var bRow = findRowByValue("Billing", S_PATIENT, name);
  if (bRow === -1) return errorResponse("Billing record not found");
  var sheet = getSheet("Billing");
  var bData = sheet.getRange(bRow, 1, 1, 14).getValues()[0];
  var cycles = safeNumber(bData[S_CYCLES]);
  var rate = safeNumber(bData[S_RATE]);
  var outstanding = safeNumber(bData[S_OUTSTANDING]);
  var today = new Date();

  // Use custom amount if provided, otherwise use plan rate
  var payAmount = safeNumber(data.amount) || rate;

  // Update last payment date
  sheet.getRange(bRow, S_LASTPAY + 1).setValue(today);
  // Decrement cycles
  if (cycles > 0) {
    sheet.getRange(bRow, S_CYCLES + 1).setValue(cycles - 1);
  }
  // Update next payment due (1 month from now)
  sheet.getRange(bRow, S_NEXTPAYDUE + 1).setValue(addMonths(today, 1));
  // Reduce outstanding by payment amount
  var newOutstanding = Math.max(0, outstanding - payAmount);
  sheet.getRange(bRow, S_OUTSTANDING + 1).setValue(newOutstanding);
  // Update status
  if (cycles - 1 <= 0) {
    sheet.getRange(bRow, S_STATUS + 1).setValue("Expired");
  } else {
    sheet.getRange(bRow, S_STATUS + 1).setValue("Active");
  }
  // Also update patient outstanding
  var pRow = findRowByValue("Patients", P_NAME, name);
  if (pRow !== -1) {
    getSheet("Patients").getRange(pRow, P_OUTSTANDING + 1).setValue(newOutstanding);
  }

  appendRefillLog(name, "", "Payment marked ($" + payAmount + ")", "Admin", "Cycles remaining: " + (cycles - 1));
  auditLog(data.adminToken, name, "Payment Recorded", "$" + payAmount + " — cycles left: " + (cycles - 1));
  return successResponse({
    message: "Payment recorded",
    newOutstanding: newOutstanding,
    cyclesLeft: cycles - 1
  });
}

function handleEditBilling(data) {
  var name = data.name || "";
  if (!name) return errorResponse("Patient name required");
  var bRow = findRowByValue("Billing", S_PATIENT, name);
  if (bRow === -1) return errorResponse("Billing record not found");
  var sheet = getSheet("Billing");
  if (data.rate !== undefined) sheet.getRange(bRow, S_RATE + 1).setValue(safeNumber(data.rate));
  if (data.term !== undefined) sheet.getRange(bRow, S_TERM + 1).setValue(safeNumber(data.term));
  if (data.startDate) {
    var start = parseDate(data.startDate);
    sheet.getRange(bRow, S_MEMSTART + 1).setValue(start);
    var term = safeNumber(data.term) || safeNumber(sheet.getRange(bRow, S_TERM + 1).getValue());
    sheet.getRange(bRow, S_CONTEND + 1).setValue(addMonths(start, term));
  }
  if (data.outstanding !== undefined) sheet.getRange(bRow, S_OUTSTANDING + 1).setValue(safeNumber(data.outstanding));
  if (data.plan) sheet.getRange(bRow, S_PLAN + 1).setValue(data.plan);
  if (data.status) sheet.getRange(bRow, S_STATUS + 1).setValue(data.status);
  auditLog(data.adminToken, data.patient || "", "Billing Updated", "Billing record modified");
  return successResponse({ message: "Billing updated" });
}

function handleNewOrder(data) {
  var sheet = getSheet("Medications");
  if (!sheet) return errorResponse("Medications tab not found");

  var name = data.patient || "";
  var med = data.medication || "";
  var dose = safeNumber(data.dose);
  var vials = safeNumber(data.vials);
  var shipDate = parseDate(data.shipDate) || new Date();
  var phone = data.phone || "";

  // Look up vial cost from catalog
  var vialCost = safeNumber(data.vialCost) || 0;
  var supplies = safeNumber(data.supplies) || 0;
  var shipping = safeNumber(data.shipping) || 0;

  // Calculate days covered (dose-dependent)
  var daysCovered = 0;
  if (dose > 0 && vials > 0) {
    // Standard vial is 200mg/mL, 1mL per vial (adjust per medication)
    var mgPerVial = safeNumber(data.mgPerVial) || 200;
    var totalMg = mgPerVial * vials;
    daysCovered = Math.floor((totalMg / dose) * 7); // dose is mg/week
  }
  if (data.daysCovered) daysCovered = safeNumber(data.daysCovered);

  var nextDue = addDays(shipDate, daysCovered);
  var total = (vialCost * vials) + supplies + shipping;
  var monthlyEst = daysCovered > 0 ? Math.round((total / daysCovered) * 30.44 * 100) / 100 : 0;

  sheet.appendRow([
    new Date(),       // Order Date
    name,             // Patient
    formatPhone(phone), // Phone
    med,              // Medication
    data.formulation || "", // Formulation
    dose,             // Dose mg/wk
    vials,            // Vials
    daysCovered,      // Days Covered
    shipDate,         // Ship Date
    nextDue,          // Next Due
    vialCost,         // Vial Cost
    supplies,         // Supplies
    shipping,         // Shipping
    total,            // Total
    monthlyEst,       // Monthly Est
    data.notes || ""  // Notes
  ]);

  // Update patient last shipped and status
  var pRow = findRowByValue("Patients", P_NAME, name);
  if (pRow !== -1) {
    var pSheet = getSheet("Patients");
    pSheet.getRange(pRow, P_STATUS + 1).setValue("Active");
    pSheet.getRange(pRow, P_FOLLOWUP + 1).setValue("");
  }

  // Update billing last shipped
  var bRow = findRowByValue("Billing", S_PATIENT, name);
  if (bRow !== -1) {
    var bSheet = getSheet("Billing");
    bSheet.getRange(bRow, S_LASTSHIP + 1).setValue(shipDate);
    bSheet.getRange(bRow, S_NEXTSHIP + 1).setValue(nextDue);
  }

  // Check for dose change
  var prevOrders = getPatientOrders(name);
  if (prevOrders.length > 1) {
    var prevDose = safeNumber(prevOrders[1].dose);
    if (prevDose > 0 && prevDose !== dose) {
      logDoseChange(name, med, prevDose, dose, data.changedBy || "Sheffield", "New order");
    }
  }

  appendRefillLog(name, med, "Order logged: " + vials + " vials", "Admin", "Ship: " + formatDateStr(shipDate));
  auditLog(data.adminToken, name, "Order Logged", med + " " + dose + "mg/wk, " + vials + " vials, ship " + formatDateStr(shipDate));

  return successResponse({
    message: "Order logged",
    daysCovered: daysCovered,
    nextDue: formatDateStr(nextDue),
    total: total
  });
}

function handleMarkLabDone(data) {
  var name = data.patient || "";
  var labType = data.labType || "";
  if (!name || !labType) return errorResponse("Patient and lab type required");
  var row = findRowByValue("Labs", L_PATIENT, name);
  if (row === -1) return errorResponse("Lab record not found");
  var sheet = getSheet("Labs");
  var today = new Date();

  if (labType === "Initial") {
    sheet.getRange(row, L_INIT_DONE + 1).setValue(today);
    // Set 90-day due
    sheet.getRange(row, L_90_DUE + 1).setValue(addDays(today, 90));
  } else if (labType === "90-Day") {
    sheet.getRange(row, L_90_DONE + 1).setValue(today);
    sheet.getRange(row, L_180_DUE + 1).setValue(addDays(today, 90));
  } else if (labType === "180-Day") {
    sheet.getRange(row, L_180_DONE + 1).setValue(today);
    sheet.getRange(row, L_ANN_DUE + 1).setValue(addDays(today, 180));
  } else if (labType === "Annual") {
    sheet.getRange(row, L_ANN_DONE + 1).setValue(today);
    // Reset for next annual
    sheet.getRange(row, L_ANN_DUE + 1).setValue(addDays(today, 365));
    sheet.getRange(row, L_ANN_DONE + 1).setValue(today);
  }

  // Update next due and status
  updateLabNextDue(row);
  return successResponse({ message: "Lab marked done" });
}

function handleSendLabReminder(data) {
  var name = data.patient || "";
  if (!name) return errorResponse("Patient name required");
  var pRow = findRowByValue("Patients", P_NAME, name);
  if (pRow === -1) return errorResponse("Patient not found");
  var pSheet = getSheet("Patients");
  var phone = safeString(pSheet.getRange(pRow, P_PHONE + 1).getValue());
  var labType = data.labType || "upcoming";
  var message = "Hi " + name.split(" ")[0] + ", this is FRAME Medicine. "
    + "You have " + labType + " labs due. Please schedule your blood work at your earliest convenience. "
    + "Questions? Reply to this text or message us in the app: " + APP_URL;

  smartSendMessage(name, phone, "Lab Reminder", message, false);
  // Log message
  var msgSheet = getSheet("Messages");
  if (msgSheet) {
    msgSheet.appendRow([new Date(), name, formatPhone(phone), "outbound", message, "", "admin", "patient"]);
  }
  return successResponse({ message: "Lab reminder sent" });
}

function handleSaveLead(data) {
  var sheet = getSheet("Leads");
  if (!sheet) return errorResponse("Leads tab not found");
  var name = data.name || "";
  if (!name) return errorResponse("Lead name required");

  var existingRow = findRowByValue("Leads", LD_NAME, name);
  if (data.originalName && data.originalName !== name) {
    existingRow = findRowByValue("Leads", LD_NAME, data.originalName);
  }
  var isNew = existingRow === -1;

  var rowData = [
    name,
    formatPhone(data.phone || ""),
    data.email || "",
    data.source || "",
    isNew ? new Date() : (parseDate(data.inquiryDate) || new Date()),
    data.interest || "",
    data.stage || "Inquiry",
    data.assignedTo || "",
    new Date(),
    parseDate(data.nextFollowUp) || "",
    data.notes || "",
    data.converted || "",
    parseDate(data.convertedDate) || "",
    data.patientName || ""
  ];

  if (isNew) {
    sheet.appendRow(rowData);
    return successResponse({ message: "Lead created", isNew: true });
  } else {
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    return successResponse({ message: "Lead updated", isNew: false });
  }
}

function handleConvertLead(data) {
  var leadName = data.leadName || "";
  if (!leadName) return errorResponse("Lead name required");

  // Save patient first
  var result = handleSavePatient(data);

  // Mark lead as converted
  var leadRow = findRowByValue("Leads", LD_NAME, leadName);
  if (leadRow !== -1) {
    var sheet = getSheet("Leads");
    sheet.getRange(leadRow, LD_CONVERTED + 1).setValue("YES");
    sheet.getRange(leadRow, LD_CONVERTEDDATE + 1).setValue(new Date());
    sheet.getRange(leadRow, LD_PATIENTNAME + 1).setValue(data.name || leadName);
    sheet.getRange(leadRow, LD_STAGE + 1).setValue("Enrolled");
  }

  return successResponse({ message: "Lead converted to patient" });
}

function handleUpdateLeadStage(data) {
  var name = data.name || "";
  var stage = data.stage || "";
  if (!name || !stage) return errorResponse("Name and stage required");
  var row = findRowByValue("Leads", LD_NAME, name);
  if (row === -1) return errorResponse("Lead not found");
  getSheet("Leads").getRange(row, LD_STAGE + 1).setValue(stage);
  return successResponse({ message: "Stage updated" });
}

function handleClearFollowUp(data) {
  var name = data.name || "";
  if (!name) return errorResponse("Patient name required");
  var row = findRowByValue("Patients", P_NAME, name);
  if (row === -1) return errorResponse("Patient not found");
  getSheet("Patients").getRange(row, P_FOLLOWUP + 1).setValue("");
  return successResponse({ message: "Follow-up cleared" });
}

function handleSaveNotes(data) {
  var name = data.name || "";
  var notes = data.notes || "";
  if (!name) return errorResponse("Name required");

  // Could be patient or lead
  var target = data.type || "patient";
  if (target === "lead") {
    var row = findRowByValue("Leads", LD_NAME, name);
    if (row === -1) return errorResponse("Lead not found");
    getSheet("Leads").getRange(row, LD_NOTES + 1).setValue(notes);
  } else {
    var pRow = findRowByValue("Patients", P_NAME, name);
    if (pRow === -1) return errorResponse("Patient not found");
    getSheet("Patients").getRange(pRow, P_NOTES + 1).setValue(notes);
  }
  return successResponse({ message: "Notes saved" });
}


// ============================================
// CSV IMPORT HANDLERS
// ============================================

function handleDiffPatients(data) {
  var rows = data.rows || [];
  if (rows.length === 0) return errorResponse("No data to import");

  var existingPatients = getSheetData("Patients");
  var existingByPhone = {};
  var existingByName = {};
  for (var i = 0; i < existingPatients.length; i++) {
    var ph = formatPhone(safeString(existingPatients[i][P_PHONE]));
    var nm = safeString(existingPatients[i][P_NAME]).toLowerCase();
    if (ph) existingByPhone[ph] = i;
    if (nm) existingByName[nm] = i;
  }

  var newPatients = [];
  var changedPatients = [];
  var unchangedCount = 0;

  for (var j = 0; j < rows.length; j++) {
    var row = rows[j];
    var importPhone = formatPhone(row.phone || "");
    var importName = safeString(row.name || "").trim();
    if (!importName) continue;

    // Match by phone first, then name
    var matchIdx = -1;
    if (importPhone && existingByPhone[importPhone] !== undefined) {
      matchIdx = existingByPhone[importPhone];
    } else if (existingByName[importName.toLowerCase()] !== undefined) {
      matchIdx = existingByName[importName.toLowerCase()];
    }

    if (matchIdx === -1) {
      newPatients.push(row);
    } else {
      // Check for changes
      var existing = existingPatients[matchIdx];
      var changes = [];
      if (row.email && safeString(existing[P_EMAIL]) !== row.email) changes.push({ field: "email", old: safeString(existing[P_EMAIL]), new: row.email });
      if (row.phone && formatPhone(safeString(existing[P_PHONE])) !== importPhone) changes.push({ field: "phone", old: safeString(existing[P_PHONE]), new: row.phone });
      if (row.street && safeString(existing[P_ADDR]) !== row.street) changes.push({ field: "street", old: safeString(existing[P_ADDR]), new: row.street });
      if (row.city && safeString(existing[P_CITY]) !== row.city) changes.push({ field: "city", old: safeString(existing[P_CITY]), new: row.city });
      if (row.state && safeString(existing[P_STATE]) !== row.state) changes.push({ field: "state", old: safeString(existing[P_STATE]), new: row.state });
      if (row.zip && safeString(existing[P_ZIP]) !== row.zip) changes.push({ field: "zip", old: safeString(existing[P_ZIP]), new: row.zip });
      if (changes.length > 0) {
        changedPatients.push({ name: importName, changes: changes, rowIndex: matchIdx });
      } else {
        unchangedCount++;
      }
    }
  }

  return successResponse({
    newPatients: newPatients,
    changedPatients: changedPatients,
    unchangedCount: unchangedCount,
    totalImported: rows.length
  });
}

function handleImportPatients(data) {
  var newPatients = data.newPatients || [];
  var changedPatients = data.changedPatients || [];
  var sheet = getSheet("Patients");
  if (!sheet) return errorResponse("Patients tab not found");

  var addedCount = 0;
  var updatedCount = 0;

  // Add new patients
  for (var i = 0; i < newPatients.length; i++) {
    var p = newPatients[i];
    handleSavePatient({
      name: p.name || "",
      preferredName: p.preferredName || "",
      phone: p.phone || "",
      email: p.email || "",
      dob: p.dob || "",
      street: p.street || "",
      city: p.city || "",
      state: p.state || "",
      zip: p.zip || "",
      status: "Active - No Med",
      adminToken: data.adminToken
    });
    addedCount++;
  }

  // Apply changes to existing
  var allData = sheet.getDataRange().getValues();
  for (var j = 0; j < changedPatients.length; j++) {
    var cp = changedPatients[j];
    var rowIdx = cp.rowIndex + 1 + 1; // +1 for header, +1 for 1-indexed
    var changes = cp.changes || [];
    for (var k = 0; k < changes.length; k++) {
      var c = changes[k];
      var colIdx = -1;
      if (c.field === "email") colIdx = P_EMAIL;
      else if (c.field === "phone") colIdx = P_PHONE;
      else if (c.field === "street") colIdx = P_ADDR;
      else if (c.field === "city") colIdx = P_CITY;
      else if (c.field === "state") colIdx = P_STATE;
      else if (c.field === "zip") colIdx = P_ZIP;
      if (colIdx !== -1) {
        sheet.getRange(rowIdx, colIdx + 1).setValue(c.new || c["new"]);
      }
    }
    updatedCount++;
  }

  return successResponse({
    message: "Import complete",
    added: addedCount,
    updated: updatedCount
  });
}

function handleDiffSales(data) {
  var rows = data.rows || [];
  if (rows.length === 0) return errorResponse("No data to import");

  var existingSales = getSheetData("Sales");
  var existingInvoices = {};
  for (var i = 0; i < existingSales.length; i++) {
    var inv = safeString(existingSales[i][SALE_INVOICE]);
    if (inv) existingInvoices[inv] = i;
  }

  var newSales = [];
  var existingCount = 0;

  for (var j = 0; j < rows.length; j++) {
    var row = rows[j];
    var invoice = safeString(row.invoice || row.invoiceNumber || "");
    if (!invoice) continue;
    if (existingInvoices[invoice] !== undefined) {
      existingCount++;
    } else {
      newSales.push(row);
    }
  }

  return successResponse({
    newSales: newSales,
    existingCount: existingCount,
    totalImported: rows.length
  });
}

function handleImportSales(data) {
  var newSales = data.newSales || [];
  var sheet = getSheet("Sales");
  if (!sheet) return errorResponse("Sales tab not found");

  for (var i = 0; i < newSales.length; i++) {
    var s = newSales[i];
    sheet.appendRow([
      s.invoice || s.invoiceNumber || "",
      s.patient || "",
      s.item || "",
      parseDate(s.purchaseDate || s.date) || new Date(),
      safeNumber(s.total),
      s.status || ""
    ]);
  }

  return successResponse({
    message: "Sales import complete",
    added: newSales.length
  });
}


function handleImportMemberships(data) {
  var memberships = data.memberships || [];
  if (memberships.length === 0) return errorResponse("No memberships to import");

  var sheet = getSheet("Billing");
  if (!sheet) return errorResponse("Billing tab not found");

  // Load all membership rates from Settings
  var rateMap = getMembershipRates();

  var added = 0;
  var updated = 0;
  var skipped = 0;

  for (var i = 0; i < memberships.length; i++) {
    var m = memberships[i];
    var name = m.patient || "";
    if (!name) { skipped++; continue; }

    // Check if patient exists
    var pRow = findRowByValue("Patients", P_NAME, name);
    if (pRow === -1) { skipped++; continue; }

    var bRow = findRowByValue("Billing", S_PATIENT, name);
    var startDate = parseDate(m.startDate) || parseDate(m.purchaseDate) || new Date();
    var contractEnd = parseDate(m.contractEnd) || "";
    var cyclesLeft = safeNumber(m.cyclesLeft);
    var outstanding = safeNumber(m.outstanding);
    var plan = m.plan || "";
    var status = m.status || "Active";

    if (status === "Expired") status = "Expired";
    else if (status === "Cancelled") status = "Cancelled";
    else status = "Active";

    // Look up rate from Settings
    var rate = lookupRate(plan, rateMap);

    // Calculate term from start date to contract end
    var term = 0;
    if (contractEnd && startDate) {
      var months = (contractEnd.getFullYear() - startDate.getFullYear()) * 12 + (contractEnd.getMonth() - startDate.getMonth());
      if (months > 0) term = months;
    }

    if (bRow === -1) {
      sheet.appendRow([
        name,
        plan,
        rate,
        term,
        startDate,
        "", // last payment
        contractEnd,
        cyclesLeft,
        outstanding,
        status,
        "", // last shipped
        "", // next ship
        "", // next pay due
        ""  // notes
      ]);

      // Also update patient record with plan and rate
      var pSheet = getSheet("Patients");
      pSheet.getRange(pRow, P_PLAN + 1).setValue(plan);
      pSheet.getRange(pRow, P_RATE + 1).setValue(rate);

      added++;
    } else {
      sheet.getRange(bRow, S_PLAN + 1).setValue(plan);
      sheet.getRange(bRow, S_RATE + 1).setValue(rate);
      sheet.getRange(bRow, S_TERM + 1).setValue(term);
      sheet.getRange(bRow, S_MEMSTART + 1).setValue(startDate);
      sheet.getRange(bRow, S_CONTEND + 1).setValue(contractEnd);
      sheet.getRange(bRow, S_CYCLES + 1).setValue(cyclesLeft);
      sheet.getRange(bRow, S_OUTSTANDING + 1).setValue(outstanding);
      sheet.getRange(bRow, S_STATUS + 1).setValue(status);

      // Also update patient record
      var pSheet2 = getSheet("Patients");
      pSheet2.getRange(pRow, P_PLAN + 1).setValue(plan);
      pSheet2.getRange(pRow, P_RATE + 1).setValue(rate);

      updated++;
    }
  }

  return successResponse({
    message: "Memberships import complete",
    added: added,
    updated: updated,
    skipped: skipped
  });
}

// Look up membership rates from Settings tab
function getMembershipRates() {
  var settings = getSheetData("Settings");
  var rates = {};
  for (var i = 0; i < settings.length; i++) {
    var key = safeString(settings[i][0]).trim();
    var val = safeNumber(settings[i][1]);
    if (val > 0) {
      rates[key.toLowerCase()] = val;
    }
  }
  return rates;
}

function lookupRate(plan, rateMap) {
  if (!plan) return 0;
  var planLower = plan.toLowerCase();

  // Direct match first
  if (rateMap[planLower]) return rateMap[planLower];

  // Fuzzy match — check if any setting key is contained in the plan name or vice versa
  var keys = Object.keys(rateMap);
  for (var i = 0; i < keys.length; i++) {
    if (planLower.indexOf(keys[i]) !== -1 || keys[i].indexOf(planLower) !== -1) {
      return rateMap[keys[i]];
    }
  }

  // Check common keywords
  if (planLower.indexOf("family") !== -1) return rateMap["family rate"] || 50;
  if (planLower.indexOf("sponsored") !== -1 && planLower.indexOf("test") !== -1) return rateMap["sponsored testosterone"] || 140;

  return 0;
}


// ============================================
// SETTINGS / P&L HANDLERS
// ============================================

function handleUpdateSettings(data) {
  var updates = data.updates || {};
  var keys = Object.keys(updates);
  for (var i = 0; i < keys.length; i++) {
    setSettingValue(keys[i], updates[keys[i]]);
  }
  return successResponse({ message: "Settings updated" });
}

function handleLockMonth(data) {
  var month = safeNumber(data.month);
  var year = safeNumber(data.year);
  if (!year) return errorResponse("Month and year required");

  // Calculate final numbers
  var revenue = calculateMonthlyRevenue(month, year);
  var medCosts = calculateMonthlyMedCosts(month, year);
  var overheadResult = calculateMonthlyOverhead(month, year);
  var overhead = overheadResult.total;
  var net = revenue.total - medCosts.total - overhead;
  var tomSplit = safeNumber(getSettingValue("Tom Split")) || 0.25;
  var colinSplit = safeNumber(getSettingValue("Colin Split")) || 0.75;

  var monthNames = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];
  var monthLabel = monthNames[month] + " " + year;

  // Check if row exists in Finance tab
  var sheet = getSheet("Finance");
  if (!sheet) return errorResponse("Finance tab not found");
  var data2 = sheet.getDataRange().getValues();
  var existingRow = -1;
  for (var i = 1; i < data2.length; i++) {
    if (safeNumber(data2[i][FIN_YEAR]) === year && safeNumber(data2[i][FIN_MONTHNUM]) === month) {
      existingRow = i + 1;
      break;
    }
  }

  var rowData = [
    monthLabel,
    year,
    month,
    revenue.total,
    medCosts.total,
    overhead,
    net,
    Math.round(net * tomSplit * 100) / 100,
    Math.round(net * colinSplit * 100) / 100,
    "YES"
  ];

  if (existingRow !== -1) {
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return successResponse({ message: "Month locked: " + monthLabel });
}

function handleUpdateOverhead(data) {
  var baseOverhead = data.baseOverhead;
  if (baseOverhead !== undefined) {
    setSettingValue("Monthly Overhead", safeNumber(baseOverhead));
  }
  return successResponse({ message: "Overhead updated" });
}

function handleAddOverheadItem(data) {
  var month = safeNumber(data.month);
  var year = safeNumber(data.year);
  var description = data.description || "";
  var amount = safeNumber(data.amount);
  if (!description) return errorResponse("Description required");

  var sheet = getSheet("Overhead Items");
  if (!sheet) {
    // Create the tab if it doesn't exist
    var ss = SpreadsheetApp.openById(SHEET_ID);
    sheet = ss.insertSheet("Overhead Items");
    sheet.appendRow(["Month", "Year", "Description", "Amount"]);
  }
  sheet.appendRow([month, year, description, amount]);
  return successResponse({ message: "Overhead item added" });
}

function handleRemoveOverheadItem(data) {
  var month = safeNumber(data.month);
  var year = safeNumber(data.year);
  var description = data.description || "";
  var sheet = getSheet("Overhead Items");
  if (!sheet) return errorResponse("Overhead Items tab not found");
  var allData = sheet.getDataRange().getValues();
  for (var i = allData.length - 1; i >= 1; i--) {
    if (safeNumber(allData[i][OH_MONTH]) === month &&
        safeNumber(allData[i][OH_YEAR]) === year &&
        safeString(allData[i][OH_DESC]) === description) {
      sheet.deleteRow(i + 1);
      return successResponse({ message: "Overhead item removed" });
    }
  }
  return errorResponse("Item not found");
}

function handleLogDoseChange(data) {
  logDoseChange(
    data.patient || "",
    data.medication || "",
    safeNumber(data.oldDose),
    safeNumber(data.newDose),
    data.changedBy || "",
    data.reason || ""
  );
  return successResponse({ message: "Dose change logged" });
}


// ============================================
// TWILIO INBOUND SMS HANDLER
// ============================================

function handleTwilioInbound(e) {
  var params = {};
  if (e.parameter) params = e.parameter;
  var from = safeString(params.From || "");
  var body = safeString(params.Body || "");
  var phone = formatPhone(from);

  if (!phone || !body) {
    return ContentService.createTextOutput("<Response></Response>")
      .setMimeType(ContentService.MimeType.XML);
  }

  // Find who this is from (patient or lead)
  var contactName = "";
  var contactType = "patient";
  var pRow = findRowByPhone("Patients", P_PHONE, phone);
  if (pRow !== -1) {
    contactName = safeString(getSheet("Patients").getRange(pRow, P_NAME + 1).getValue());
  } else {
    var lRow = findRowByPhone("Leads", LD_PHONE, phone);
    if (lRow !== -1) {
      contactName = safeString(getSheet("Leads").getRange(lRow, LD_NAME + 1).getValue());
      contactType = "lead";
    } else {
      contactName = phone;
    }
  }

  // Log to Messages tab
  var msgSheet = getSheet("Messages");
  if (msgSheet) {
    msgSheet.appendRow([
      new Date(),
      contactName,
      phone,
      "inbound",
      body,
      "",
      "sms",
      contactType
    ]);
  }

  // Check for keyword responses
  var lowerBody = body.toLowerCase().trim();

  // Refill responses
  if (lowerBody === "yes" || lowerBody === "refill" || lowerBody === "confirm") {
    appendRefillLog(contactName, "", "Refill confirmed via text", "Twilio", body);
    sendBrandedEmail("both",
      "Refill Confirmed (Text): " + contactName,
      "<div style='border-left:4px solid #4CAF50;padding:12px 16px;background:rgba(76,175,80,0.08);border-radius:0 8px 8px 0;'>"
      + "<h2 style='color:#4CAF50;margin:0 0 8px 0;'>" + patientLinkHtml(contactName) + " — Refill Confirmed</h2>"
      + "<p style='color:rgba(255,255,255,0.6);margin:0;'>Via text: \"" + body + "\"</p>"
      + "</div>"
    );
  } else if (lowerBody === "no" || lowerBody === "decline" || lowerBody === "stop") {
    if (pRow !== -1) {
      var pSheet = getSheet("Patients");
      pSheet.getRange(pRow, P_FOLLOWUP + 1).setValue("YES");
    }
    appendRefillLog(contactName, "", "Refill declined via text", "Twilio", body);
    sendBrandedEmail("both",
      "REFILL DECLINED (Text): " + contactName,
      "<div style='border-left:4px solid #ff4444;padding:12px 16px;background:rgba(255,68,68,0.08);border-radius:0 8px 8px 0;'>"
      + "<h2 style='color:#ff4444;margin:0 0 8px 0;'>" + patientLinkHtml(contactName) + " — Refill Declined</h2>"
      + "<p style='color:#fff;margin:0 0 4px 0;'>Via text: \"" + body + "\"</p>"
      + "<p style='color:#ff4444;margin:4px 0 0 0;font-weight:bold;'>Follow-up flag set</p>"
      + "</div>"
    );
  } else {
    // General inbound message — email notification
    sendBrandedEmail("both",
      "New Text from " + contactName,
      "<div style='border-left:4px solid #E8891A;padding:12px 16px;background:rgba(232,137,26,0.08);border-radius:0 8px 8px 0;'>"
      + "<h2 style='color:#E8891A;margin:0 0 8px 0;'>" + patientLinkHtml(contactName) + "</h2>"
      + "<p style='color:#fff;margin:0;'>" + body + "</p>"
      + "</div>"
    );
  }

  // Return empty TwiML
  return ContentService.createTextOutput("<Response></Response>")
    .setMimeType(ContentService.MimeType.XML);
}


// ============================================
// HELPER FUNCTIONS — DATA RETRIEVAL
// ============================================

function buildPatientObj(data) {
  return {
    name: safeString(data[P_NAME]),
    preferredName: safeString(data[P_PREFERRED]),
    dob: formatDateStr(data[P_DOB]),
    phone: safeString(data[P_PHONE]),
    email: safeString(data[P_EMAIL]),
    street: safeString(data[P_ADDR]),
    city: safeString(data[P_CITY]),
    state: safeString(data[P_STATE]),
    zip: safeString(data[P_ZIP]),
    memberSince: formatDateStr(data[P_SINCE]),
    medication: safeString(data[P_MED]),
    plan: safeString(data[P_PLAN]),
    rate: safeNumber(data[P_RATE]),
    term: safeNumber(data[P_TERM]),
    membershipStart: formatDateStr(data[P_MEMSTART]),
    contractEnd: formatDateStr(data[P_CONTEND]),
    cyclesLeft: safeNumber(data[P_CYCLES]),
    outstanding: safeNumber(data[P_OUTSTANDING]),
    checkInDay: safeString(data[P_CIDAY]),
    checkInTime: safeString(data[P_CITIME]),
    glpWeightDay: safeString(data[P_GLPDAY]),
    glpWeightTime: safeString(data[P_GLPTIME]),
    status: safeString(data[P_STATUS]),
    followUp: safeString(data[P_FOLLOWUP]),
    notes: safeString(data[P_NOTES]),
    pushEnabled: safeString(data[P_PUSH]),
    referralSource: safeString(data[P_REFSOURCE]),
    referredBy: safeString(data[P_REFBY]),
    commPref: safeString(data[P_COMMPREF]) || "all"
  };
}

function getPatientOrders(name) {
  var sheet = getSheet("Medications");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var orders = [];
  for (var i = MED_DATA_START_ROW - 1; i < data.length; i++) {
    if (safeString(data[i][M_PATIENT]).toLowerCase() === name.toLowerCase()) {
      orders.push({
        orderDate: formatDateStr(data[i][M_ORDERDATE]),
        medication: safeString(data[i][M_MED]),
        formulation: safeString(data[i][M_FORM]),
        dose: safeNumber(data[i][M_DOSE]),
        vials: safeNumber(data[i][M_VIALS]),
        daysCovered: safeNumber(data[i][M_DAYS]),
        shipDate: formatDateStr(data[i][M_SHIPDATE]),
        nextDue: formatDateStr(data[i][M_NEXTDUE]),
        vialCost: safeNumber(data[i][M_VIALCOST]),
        supplies: safeNumber(data[i][M_SUPPLY]),
        shipping: safeNumber(data[i][M_SHIPPING]),
        total: safeNumber(data[i][M_TOTAL]),
        monthlyEst: safeNumber(data[i][M_MONTHLY]),
        notes: safeString(data[i][M_NOTES])
      });
    }
  }
  orders.sort(function(a, b) {
    return new Date(b.orderDate) - new Date(a.orderDate);
  });
  return orders;
}

// Bulk load: read Medications sheet ONCE, return map of name → orders
function getAllPatientOrders() {
  var sheet = getSheet("Medications");
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = MED_DATA_START_ROW - 1; i < data.length; i++) {
    var name = safeString(data[i][M_PATIENT]);
    var key = name.toLowerCase();
    if (!key) continue;
    if (!map[key]) map[key] = { name: name, orders: [] };
    map[key].orders.push({
      orderDate: formatDateStr(data[i][M_ORDERDATE]),
      medication: safeString(data[i][M_MED]),
      formulation: safeString(data[i][M_FORM]),
      dose: safeNumber(data[i][M_DOSE]),
      vials: safeNumber(data[i][M_VIALS]),
      daysCovered: safeNumber(data[i][M_DAYS]),
      shipDate: formatDateStr(data[i][M_SHIPDATE]),
      nextDue: formatDateStr(data[i][M_NEXTDUE]),
      vialCost: safeNumber(data[i][M_VIALCOST]),
      supplies: safeNumber(data[i][M_SUPPLY]),
      shipping: safeNumber(data[i][M_SHIPPING]),
      total: safeNumber(data[i][M_TOTAL]),
      monthlyEst: safeNumber(data[i][M_MONTHLY]),
      notes: safeString(data[i][M_NOTES])
    });
  }
  // Sort each patient's orders
  var keys = Object.keys(map);
  for (var j = 0; j < keys.length; j++) {
    map[keys[j]].orders.sort(function(a, b) {
      return new Date(b.orderDate) - new Date(a.orderDate);
    });
  }
  return map;
}

function getPatientBilling(name) {
  var data = getSheetData("Billing");
  for (var i = 0; i < data.length; i++) {
    if (safeString(data[i][S_PATIENT]).toLowerCase() === name.toLowerCase()) {
      return {
        plan: safeString(data[i][S_PLAN]),
        rate: safeNumber(data[i][S_RATE]),
        term: safeNumber(data[i][S_TERM]),
        startDate: formatDateStr(data[i][S_MEMSTART]),
        lastPayment: formatDateStr(data[i][S_LASTPAY]),
        contractEnd: formatDateStr(data[i][S_CONTEND]),
        cyclesLeft: safeNumber(data[i][S_CYCLES]),
        outstanding: safeNumber(data[i][S_OUTSTANDING]),
        status: safeString(data[i][S_STATUS]),
        lastShipped: formatDateStr(data[i][S_LASTSHIP]),
        nextShipment: formatDateStr(data[i][S_NEXTSHIP]),
        nextPayDue: formatDateStr(data[i][S_NEXTPAYDUE]),
        notes: safeString(data[i][S_NOTES])
      };
    }
  }
  return null;
}

function getPatientWeights(name, limit) {
  var data = getSheetData("Weight Log");
  var weights = [];
  for (var i = 0; i < data.length; i++) {
    if (safeString(data[i][W_PATIENT]).toLowerCase() === name.toLowerCase()) {
      weights.push({
        date: formatDateStr(data[i][W_DATE]),
        dateISO: formatDateISO(data[i][W_DATE]),
        weight: safeNumber(data[i][W_WEIGHT]),
        change: safeNumber(data[i][W_CHANGE]),
        avgPerWeek: safeNumber(data[i][W_AVG]),
        totalChange: safeNumber(data[i][W_TOTAL]),
        startWeight: safeNumber(data[i][W_START]),
        source: safeString(data[i][W_SOURCE])
      });
    }
  }
  weights.sort(function(a, b) {
    return new Date(b.dateISO) - new Date(a.dateISO);
  });
  return weights.slice(0, limit || 50);
}

function getPatientCheckIns(name, limit) {
  var data = getSheetData("Check-In Responses");
  var checkIns = [];
  for (var i = 0; i < data.length; i++) {
    if (safeString(data[i][CI_PATIENT]).toLowerCase() === name.toLowerCase()) {
      checkIns.push({
        date: formatDateStr(data[i][CI_DATE]),
        dateISO: formatDateISO(data[i][CI_DATE]),
        medication: safeString(data[i][CI_MED]),
        symptoms: safeString(data[i][CI_SYMPTOMS]),
        rating: safeString(data[i][CI_RATING]),
        notes: safeString(data[i][CI_NOTES]),
        responseRequested: safeString(data[i][CI_RESPONSE_REQ]),
        responded: safeString(data[i][CI_RESPONDED])
      });
    }
  }
  checkIns.sort(function(a, b) {
    return new Date(b.dateISO) - new Date(a.dateISO);
  });
  return checkIns.slice(0, limit || 50);
}

function getPatientLabs(name) {
  var data = getSheetData("Labs");
  for (var i = 0; i < data.length; i++) {
    if (safeString(data[i][L_PATIENT]).toLowerCase() === name.toLowerCase()) {
      return {
        enrollDate: formatDateStr(data[i][L_ENROLL]),
        initialDate: formatDateStr(data[i][L_INIT_DATE]),
        initialDone: formatDateStr(data[i][L_INIT_DONE]),
        ninetyDue: formatDateStr(data[i][L_90_DUE]),
        ninetyDone: formatDateStr(data[i][L_90_DONE]),
        oneEightyDue: formatDateStr(data[i][L_180_DUE]),
        oneEightyDone: formatDateStr(data[i][L_180_DONE]),
        annualDue: formatDateStr(data[i][L_ANN_DUE]),
        annualDone: formatDateStr(data[i][L_ANN_DONE]),
        nextDue: formatDateStr(data[i][L_NEXT_DUE]),
        status: safeString(data[i][L_STATUS]),
        notes: safeString(data[i][L_NOTES])
      };
    }
  }
  return null;
}

function getPatientDoseHistory(name) {
  var data = getSheetData("Dose History");
  var history = [];
  for (var i = 0; i < data.length; i++) {
    if (safeString(data[i][DH_PATIENT]).toLowerCase() === name.toLowerCase()) {
      history.push({
        date: formatDateStr(data[i][DH_DATE]),
        dateISO: formatDateISO(data[i][DH_DATE]),
        medication: safeString(data[i][DH_MED]),
        oldDose: safeNumber(data[i][DH_OLDDOSE]),
        newDose: safeNumber(data[i][DH_NEWDOSE]),
        changedBy: safeString(data[i][DH_CHANGEDBY]),
        reason: safeString(data[i][DH_REASON])
      });
    }
  }
  history.sort(function(a, b) {
    return new Date(b.dateISO) - new Date(a.dateISO);
  });
  return history;
}

function getPatientRefillLog(name, limit) {
  var data = getSheetData("Refill Log");
  var log = [];
  for (var i = 0; i < data.length; i++) {
    if (safeString(data[i][RL_PATIENT]).toLowerCase() === name.toLowerCase()) {
      log.push({
        timestamp: formatDateStr(data[i][RL_TIMESTAMP]),
        medication: safeString(data[i][RL_MED]),
        action: safeString(data[i][RL_ACTION]),
        method: safeString(data[i][RL_METHOD]),
        notes: safeString(data[i][RL_NOTES])
      });
    }
  }
  log.reverse();
  return log.slice(0, limit || 20);
}

function getUnreadCount(phone) {
  var normalized = formatPhone(phone);
  var data = getSheetData("Messages");
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    if (formatPhone(safeString(data[i][MSG_PHONE])) === normalized &&
        safeString(data[i][MSG_DIRECTION]) === "inbound" &&
        safeString(data[i][MSG_READ]) !== "Yes") {
      count++;
    }
  }
  return count;
}


// ============================================
// HELPER FUNCTIONS — DATA CREATION
// ============================================

function createBillingRow(name, data) {
  var sheet = getSheet("Billing");
  if (!sheet) return;
  var start = parseDate(data.membershipStart) || new Date();
  var term = safeNumber(data.term);
  var rate = safeNumber(data.rate);
  sheet.appendRow([
    name,
    data.plan || "",
    rate,
    term,
    start,
    "",
    term ? addMonths(start, term) : "",
    term,
    rate * term,
    "Active",
    "",
    "",
    addMonths(start, 1),
    ""
  ]);
}

function createLabsRow(name) {
  var sheet = getSheet("Labs");
  if (!sheet) return;
  var today = new Date();
  sheet.appendRow([
    name,
    today,
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    today, // Next due = now (initial labs)
    "Awaiting",
    ""
  ]);
}

function createCheckInRow(name, data) {
  var sheet = getSheet("Check-Ins");
  if (!sheet) return;
  sheet.appendRow([
    name,
    formatPhone(data.phone || ""),
    data.medication || "",
    data.checkInDay || "",
    data.checkInTime || "",
    "",
    "",
    "",
    "Active"
  ]);
}

function appendRefillLog(patient, medication, action, method, notes) {
  var sheet = getSheet("Refill Log");
  if (!sheet) return;
  sheet.appendRow([new Date(), patient, medication, action, method, notes]);
}

function logDoseChange(patient, medication, oldDose, newDose, changedBy, reason) {
  var sheet = getSheet("Dose History");
  if (!sheet) return;
  sheet.appendRow([new Date(), patient, medication, oldDose, newDose, changedBy, reason]);
}

// Audit trail: logs who changed what on which patient record
function auditLog(adminToken, patient, action, details) {
  var sheet = getSheet("Audit Log");
  if (!sheet) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.insertSheet("Audit Log");
      sheet.appendRow(["Timestamp", "Admin", "Role", "Patient", "Action", "Details"]);
    } catch (e) {
      Logger.log("Failed to create Audit Log: " + e.message);
      return;
    }
  }
  var who = "System";
  var role = "";
  if (adminToken) {
    var session = verifyAdminToken({ adminToken: adminToken });
    if (session) { who = session.name; role = session.role; }
  }
  sheet.appendRow([new Date(), who, role, patient, action, details]);
}

function updateLabNextDue(sheetRow) {
  var sheet = getSheet("Labs");
  var data = sheet.getRange(sheetRow, 1, 1, 13).getValues()[0];
  var nextDue = null;
  var status = "Current";

  if (!data[L_INIT_DONE]) {
    nextDue = data[L_INIT_DATE] || data[L_ENROLL];
    status = "Awaiting";
  } else if (!data[L_90_DONE] && data[L_90_DUE]) {
    nextDue = data[L_90_DUE];
  } else if (!data[L_180_DONE] && data[L_180_DUE]) {
    nextDue = data[L_180_DUE];
  } else if (!data[L_ANN_DONE] && data[L_ANN_DUE]) {
    nextDue = data[L_ANN_DUE];
  }

  if (nextDue) {
    var daysUntil = daysBetween(new Date(), nextDue);
    if (daysUntil < 0) status = "Overdue";
  }

  sheet.getRange(sheetRow, L_NEXT_DUE + 1).setValue(nextDue || "");
  sheet.getRange(sheetRow, L_STATUS + 1).setValue(status);
}


// ============================================
// FINANCIAL CALCULATIONS
// ============================================

function calculateMonthlyRevenue(month, year) {
  var billing = getSheetData("Billing");
  var total = 0;
  var breakdown = [];
  var monthStart = new Date(year, month, 1);
  var monthEnd = new Date(year, month + 1, 0);

  // 1) Recurring membership revenue from the Billing tab
  for (var i = 0; i < billing.length; i++) {
    var status = safeString(billing[i][S_STATUS]);
    if (status !== "Active" && status !== "Past Due") continue;

    var start = parseDate(billing[i][S_MEMSTART]);
    var end = parseDate(billing[i][S_CONTEND]);

    // Check if billing was active during this month
    if (start && start <= monthEnd) {
      if (!end || end >= monthStart) {
        var rate = safeNumber(billing[i][S_RATE]);
        total += rate;
        breakdown.push({
          patient: safeString(billing[i][S_PATIENT]),
          plan: safeString(billing[i][S_PLAN]),
          rate: rate,
          source: "membership"
        });
      }
    }
  }

  // 2) One-off sales from the Sales tab (JaneApp imports) — Paid only, in-month
  // Sales columns: 0=Invoice, 1=Patient, 2=Item, 3=Purchase Date, 4=Total, 5=Status
  var sales = getSheetData("Sales");
  for (var j = 0; j < sales.length; j++) {
    var saleDate = parseDate(sales[j][3]);
    if (!saleDate) continue;
    if (saleDate < monthStart || saleDate > monthEnd) continue;
    var saleStatus = safeString(sales[j][5]).toLowerCase();
    if (saleStatus.indexOf("paid") === -1 && saleStatus !== "applied" && saleStatus !== "completed") continue;
    var saleAmount = safeNumber(sales[j][4]);
    if (saleAmount <= 0) continue;
    total += saleAmount;
    breakdown.push({
      patient: safeString(sales[j][1]),
      plan: safeString(sales[j][2]),
      rate: saleAmount,
      source: "sale"
    });
  }

  return { total: total, breakdown: breakdown };
}

function calculateMonthlyMedCosts(month, year) {
  var sheet = getSheet("Medications");
  if (!sheet) return { total: 0, breakdown: [] };
  var data = sheet.getDataRange().getValues();
  var total = 0;
  var breakdown = [];
  var monthStart = new Date(year, month, 1);
  var monthEnd = new Date(year, month + 1, 0);

  for (var i = MED_DATA_START_ROW - 1; i < data.length; i++) {
    var shipDate = parseDate(data[i][M_SHIPDATE]);
    if (!shipDate) continue;
    if (shipDate >= monthStart && shipDate <= monthEnd) {
      var cost = safeNumber(data[i][M_TOTAL]);
      total += cost;
      breakdown.push({
        patient: safeString(data[i][M_PATIENT]),
        medication: safeString(data[i][M_MED]),
        total: cost,
        shipDate: formatDateStr(shipDate)
      });
    }
  }
  return { total: total, breakdown: breakdown };
}

function calculateMonthlyOverhead(month, year) {
  var base = safeNumber(getSettingValue("Monthly Overhead")) || 623;
  var items = getOverheadItems(month, year);
  var itemsTotal = 0;
  for (var i = 0; i < items.length; i++) {
    itemsTotal += safeNumber(items[i].amount);
  }
  return {
    base: base,
    items: items,
    total: base + itemsTotal
  };
}

function getOverheadItems(month, year) {
  var sheet = getSheet("Overhead Items");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (safeNumber(data[i][OH_MONTH]) === month && safeNumber(data[i][OH_YEAR]) === year) {
      items.push({
        description: safeString(data[i][OH_DESC]),
        amount: safeNumber(data[i][OH_AMOUNT])
      });
    }
  }
  return items;
}


// ============================================
// AUTOMATED TRIGGERS
// ============================================

function setupSheetHeaders() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var tabs = {
    "Patients": ["Name","Preferred Name","DOB","Phone","Email","Street","City","State","Zip","Member Since","Medication","Plan","Rate","Term","Membership Start","Contract End","Cycles Left","Outstanding","Check-In Day","Check-In Time","GLP Weight Day","GLP Weight Time","Status","Follow-Up","Notes","Push Enabled","Push Subscription","Referral Source","Referred By","BioToken","BioTokenDate","CommPref"],
    "Billing": ["Patient","Plan","Rate","Term","Membership Start","Last Payment","Contract End","Cycles Left","Outstanding","Status","Last Shipped","Next Ship","Next Pay Due","Notes"],
    "Medications": ["Order Date","Patient","Phone","Medication","Formulation","Dose (mg/wk)","Vials","Days Covered","Ship Date","Next Due","Vial Cost","Supplies","Shipping","Total","Monthly Est","Notes"],
    "Labs": ["Patient","Enrolled","Initial Date","Initial Done","90-Day Due","90-Day Done","180-Day Due","180-Day Done","Annual Due","Annual Done","Next Due","Status","Notes"],
    "Leads": ["Name","Phone","Email","Source","Date","Interest","Stage","Assigned","Last Contact","Next Follow-Up","Notes","Converted","Converted Date","Patient Name"],
    "Messages": ["Timestamp","Patient","Phone","Direction","Text","Read","Source","Contact Type"],
    "Check-Ins": ["Patient","Phone","Medication","Preferred Day","Status","Last Sent","Response","Response Date"],
    "Check-In Responses": ["Timestamp","Patient","Phone","Medication","Symptoms","Rating","Notes","Response Requested","Admin Notes"],
    "Weight Log": ["Date","Patient","Medication","Week","Weight","Change","Avg/Week","Total Change","Start Weight","Weeks","Source"],
    "Refill Log": ["Timestamp","Patient","Medication","Action","Method","Notes"],
    "Dose History": ["Date","Patient","Medication","Old Dose","New Dose","Changed By","Reason"],
    "Finance": ["Month","Year","Month Num","Revenue","Med Costs","Overhead","Net Profit","Tom Split","Colin Split","Locked"],
    "Overhead Items": ["Month","Year","Description","Amount"],
    "Sales": ["Invoice","Patient","Item","Purchase Date","Total","Status"],
    "Catalog": ["Product Code","Name","Strength","Form","Category","Unit","Negotiated Price"],
    "Settings": ["Key","Value"],
    "Audit Log": ["Timestamp","Admin","Role","Patient","Action","Details"]
  };

  var tabNames = Object.keys(tabs);
  for (var i = 0; i < tabNames.length; i++) {
    var name = tabNames[i];
    var existing = null;
    try { existing = ss.getSheetByName(name); } catch (e) {}
    if (!existing) {
      var newSheet = ss.insertSheet(name);
      newSheet.appendRow(tabs[name]);
    } else {
      // Check if header row exists, add if empty
      var firstRow = existing.getRange(1, 1, 1, tabs[name].length).getValues()[0];
      var isEmpty = !firstRow[0] || firstRow[0] === "";
      if (isEmpty) {
        existing.getRange(1, 1, 1, tabs[name].length).setValues([tabs[name]]);
      } else {
        // Extend headers if new columns were added (e.g., BioToken, CommPref)
        var currentCols = existing.getLastColumn();
        var expectedCols = tabs[name].length;
        if (currentCols < expectedCols) {
          for (var c = currentCols; c < expectedCols; c++) {
            existing.getRange(1, c + 1).setValue(tabs[name][c]);
          }
        }
      }
    }
  }
}

// Generate Billing + Labs + Check-In rows from existing Patient data
function syncBillingFromPatients() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var patients = getSheetData("Patients");
  var billingSheet = getSheet("Billing");
  var labsSheet = getSheet("Labs");
  var ciSheet = getSheet("Check-Ins");
  if (!billingSheet || !labsSheet || !ciSheet) return;

  // Get existing billing/labs/checkin names
  var existingBilling = {};
  var bData = getSheetData("Billing");
  for (var b = 0; b < bData.length; b++) {
    existingBilling[safeString(bData[b][S_PATIENT]).toLowerCase()] = true;
  }
  var existingLabs = {};
  var lData = getSheetData("Labs");
  for (var l = 0; l < lData.length; l++) {
    existingLabs[safeString(lData[l][L_PATIENT]).toLowerCase()] = true;
  }
  var existingCI = {};
  var cData = getSheetData("Check-Ins");
  for (var c = 0; c < cData.length; c++) {
    existingCI[safeString(cData[c][0]).toLowerCase()] = true;
  }

  var added = 0;
  for (var i = 0; i < patients.length; i++) {
    var p = patients[i];
    var name = safeString(p[P_NAME]);
    var status = safeString(p[P_STATUS]);
    if (!name || status === "INACTIVE" || status === "Staff") continue;

    var plan = safeString(p[P_PLAN]);
    var rate = safeNumber(p[P_RATE]);
    var term = safeNumber(p[P_TERM]);
    var memStart = parseDate(p[P_MEMSTART]);
    var contEnd = parseDate(p[P_CONTEND]);
    var cycles = safeNumber(p[P_CYCLES]);
    var outstanding = safeNumber(p[P_OUTSTANDING]);
    var phone = safeString(p[P_PHONE]);
    var med = safeString(p[P_MED]);
    var ciDay = safeString(p[P_CIDAY]);

    // Create or update billing row
    if (!existingBilling[name.toLowerCase()] && plan) {
      billingSheet.appendRow([
        name, plan, rate, term,
        memStart || new Date(), "", // last payment
        contEnd || "", cycles, outstanding,
        rate > 0 ? "Active" : "Individual",
        "", "", // last shipped, next ship
        memStart ? addMonths(memStart, 1) : "", // next pay due
        ""
      ]);
      added++;
    } else if (existingBilling[name.toLowerCase()] && rate > 0) {
      // Update existing billing row if rate is missing
      var bRow = findRowByValue("Billing", S_PATIENT, name);
      if (bRow !== -1) {
        var existingRate = safeNumber(billingSheet.getRange(bRow, S_RATE + 1).getValue());
        if (existingRate === 0) {
          billingSheet.getRange(bRow, S_RATE + 1).setValue(rate);
          billingSheet.getRange(bRow, S_PLAN + 1).setValue(plan);
          billingSheet.getRange(bRow, S_TERM + 1).setValue(term);
          if (memStart) billingSheet.getRange(bRow, S_MEMSTART + 1).setValue(memStart);
          if (contEnd) billingSheet.getRange(bRow, S_CONTEND + 1).setValue(contEnd);
          if (cycles) billingSheet.getRange(bRow, S_CYCLES + 1).setValue(cycles);
          billingSheet.getRange(bRow, S_STATUS + 1).setValue("Active");
          added++;
        }
      }
    }

    // Create labs row if missing
    if (!existingLabs[name.toLowerCase()]) {
      var enrollDate = parseDate(p[P_SINCE]) || new Date();
      labsSheet.appendRow([
        name, enrollDate, "", "",
        addDays(enrollDate, 90), "",
        addDays(enrollDate, 180), "",
        addDays(enrollDate, 365), "",
        addDays(enrollDate, 90),
        "Awaiting", ""
      ]);
    }

    // Create check-in row if missing
    if (!existingCI[name.toLowerCase()] && med) {
      ciSheet.appendRow([
        name, formatPhone(phone), med,
        ciDay || "Monday", "", "", "", "", "Active"
      ]);
    }
  }
  return added;
}

function setupTriggers() {
  // Set up all sheet headers and sync billing
  setupSheetHeaders();
  syncBillingFromPatients();

  // Clear existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // Daily 9 AM — Digest email
  ScriptApp.newTrigger("dailyDigest")
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  // Daily 12 PM — Patient texts (refill alerts, check-ins, weight)
  ScriptApp.newTrigger("dailyRefillAlerts")
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger("dailyCheckIns")
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger("dailyWeightTexts")
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();

  // Last day of month — CSV import reminder
  ScriptApp.newTrigger("monthlyImportReminder")
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
}


// ---- DAILY DIGEST (9 AM) — ONE email with everything ----
function dailyDigest() {
  var patients = getSheetData("Patients");
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // ---- Collect data ----
  var allOrders = getAllPatientOrders();
  var overdue = [];
  var dueSoon = [];
  var followUps = [];

  for (var i = 0; i < patients.length; i++) {
    var status = safeString(patients[i][P_STATUS]);
    if (status === "INACTIVE" || status === "Staff") continue;
    var name = safeString(patients[i][P_NAME]);

    var orderEntry = allOrders[name.toLowerCase()];
    var orders = orderEntry ? orderEntry.orders : [];
    if (orders.length > 0) {
      var nextDue = parseDate(orders[0].nextDue);
      if (nextDue) {
        var days = daysBetween(today, nextDue);
        if (days < 0) overdue.push({ name: name, detail: Math.abs(days) + "d overdue", med: safeString(patients[i][P_MED]) });
        else if (days <= 14) dueSoon.push({ name: name, detail: days + "d left", med: safeString(patients[i][P_MED]) });
      }
    }
    if (safeString(patients[i][P_FOLLOWUP]) === "YES") followUps.push(name);
  }

  // Labs (minimal — just counts, detail demoted)
  var labs = getSheetData("Labs");
  var labsOverdueCount = 0;
  var labsDueSoonCount = 0;
  for (var j = 0; j < labs.length; j++) {
    var labNext = parseDate(labs[j][L_NEXT_DUE]);
    if (!labNext) continue;
    var labDays = daysBetween(today, labNext);
    if (labDays < 0) {
      labsOverdueCount++;
      getSheet("Labs").getRange(j + 2, L_STATUS + 1).setValue("Overdue");
    } else if (labDays <= 30) {
      labsDueSoonCount++;
    }
  }

  // Billing
  var billing = getSheetData("Billing");
  var billingDue = [];
  var billingPastDue = [];
  for (var k = 0; k < billing.length; k++) {
    var bStatus = safeString(billing[k][S_STATUS]);
    if (bStatus !== "Active" && bStatus !== "Past Due") continue;
    var bPatient = safeString(billing[k][S_PATIENT]);
    var nextPay = parseDate(billing[k][S_NEXTPAYDUE]);
    if (!nextPay) continue;
    var bDays = daysBetween(today, nextPay);
    if (bDays < 0) {
      billingPastDue.push({ name: bPatient, detail: "$" + safeNumber(billing[k][S_OUTSTANDING]) + " • " + Math.abs(bDays) + "d overdue" });
      getSheet("Billing").getRange(k + 2, S_STATUS + 1).setValue("Past Due");
    } else if (bDays <= 7) {
      billingDue.push({ name: bPatient, detail: "$" + safeNumber(billing[k][S_RATE]) + " in " + bDays + "d" });
    }
  }

  // Escalation (unread messages)
  var escalationHours = safeNumber(getSettingValue("No-Response Escalation (hrs)")) || 48;
  var messages = getSheetData("Messages");
  var now = new Date();
  var unreadEscalation = [];
  for (var m = 0; m < messages.length; m++) {
    if (safeString(messages[m][MSG_DIRECTION]) !== "inbound") continue;
    if (safeString(messages[m][MSG_READ]) === "Yes") continue;
    var ts = parseDate(messages[m][MSG_TIMESTAMP]);
    if (!ts) continue;
    var hoursSince = (now - ts) / 3600000;
    if (hoursSince >= escalationHours) {
      var msgName = safeString(messages[m][MSG_PATIENT]);
      unreadEscalation.push({ name: msgName, detail: Math.round(hoursSince) + "h unread" });
    }
  }

  // Lead follow-ups
  var leads = getSheetData("Leads");
  var leadFollowUps = [];
  for (var n = 0; n < leads.length; n++) {
    if (safeString(leads[n][LD_CONVERTED]) === "YES") continue;
    var nextFU = parseDate(leads[n][LD_NEXTFOLLOWUP]);
    if (!nextFU) continue;
    if (daysBetween(nextFU, today) >= 0) {
      leadFollowUps.push({ name: safeString(leads[n][LD_NAME]), detail: daysBetween(nextFU, today) + "d overdue" });
    }
  }

  // ---- Build HTML (card/box layout, matches patient homescreen language) ----
  // Color palette (same as patient app):
  //   bg:      #080808 (wrapper) / #131313 (card)
  //   orange:  #E8891A (accent, upcoming)
  //   red:     #ff4444 (urgent)
  //   green:   #4CAF50 (ok)
  //   text:    #ffffff / rgba(255,255,255,0.6) muted

  function card(accentColor, eyebrow, body) {
    // table-based so it renders in Outlook as well as webmail
    var s = "";
    s += "<table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='margin:0 0 14px 0;background:#131313;border-radius:8px;border:1px solid rgba(255,255,255,0.06);overflow:hidden;'>";
    s += "<tr><td style='height:3px;background:" + accentColor + ";font-size:1px;line-height:3px;'>&nbsp;</td></tr>";
    s += "<tr><td style='padding:14px 16px 16px 16px;'>";
    s += "<div style='font-family:Arial,sans-serif;font-size:10px;font-weight:700;letter-spacing:0.22em;text-transform:uppercase;color:" + accentColor + ";margin:0 0 10px 0;'>" + eyebrow + "</div>";
    s += body;
    s += "</td></tr>";
    s += "</table>";
    return s;
  }

  function itemRow(name, detail, med) {
    var s = "";
    s += "<table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='border-bottom:1px solid rgba(255,255,255,0.06);'>";
    s += "<tr>";
    s += "<td style='padding:8px 0;font-family:Arial,sans-serif;font-size:14px;color:#ffffff;font-weight:600;'>" + patientLinkHtml(name) + "</td>";
    s += "<td style='padding:8px 0;font-family:Arial,sans-serif;font-size:12px;color:rgba(255,255,255,0.5);text-align:right;'>";
    if (med) s += "<span style='margin-right:10px;'>" + med + "</span>";
    s += detail;
    s += "</td>";
    s += "</tr>";
    s += "</table>";
    return s;
  }

  function itemsList(items, showMed) {
    var b = "";
    for (var x = 0; x < items.length; x++) {
      b += itemRow(items[x].name, items[x].detail, showMed ? (items[x].med || "") : "");
    }
    // Strip final border (last row)
    return b.replace(/border-bottom:1px solid rgba\(255,255,255,0\.06\);'>(?![\s\S]*border-bottom)/, "'>");
  }

  function nameList(names) {
    var items = [];
    for (var x = 0; x < names.length; x++) items.push({ name: names[x], detail: "flagged" });
    return itemsList(items, false);
  }

  var html = "";

  // ---- HEADER ----
  html += "<div style='font-family:Arial,sans-serif;text-align:center;padding:4px 0 20px 0;'>";
  html += "<div style='font-size:10px;font-weight:700;letter-spacing:0.24em;text-transform:uppercase;color:#E8891A;margin-bottom:6px;'>Daily Digest</div>";
  html += "<div style='font-size:26px;font-weight:800;color:#ffffff;letter-spacing:-0.01em;'>" + formatDateStr(today) + "</div>";
  html += "</div>";

  // ---- SUMMARY TILES (3 big stat cards side-by-side) ----
  var totalUrgent = overdue.length + billingPastDue.length + unreadEscalation.length;
  var totalWarning = dueSoon.length + billingDue.length + followUps.length + leadFollowUps.length;
  var activePatients = 0;
  for (var pi = 0; pi < patients.length; pi++) {
    var ps = safeString(patients[pi][P_STATUS]);
    if (ps && ps !== "INACTIVE" && ps !== "Staff") activePatients++;
  }

  html += "<table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='margin:0 0 14px 0;'>";
  html += "<tr>";
  html += "<td width='33%' style='padding-right:6px;'>";
  html += "<table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='background:#131313;border-radius:8px;border:1px solid rgba(255,255,255,0.06);'>";
  html += "<tr><td style='padding:14px 10px;text-align:center;'>";
  html += "<div style='font-family:Arial,sans-serif;font-size:9px;font-weight:700;letter-spacing:0.2em;text-transform:uppercase;color:rgba(255,255,255,0.5);margin-bottom:4px;'>Urgent</div>";
  html += "<div style='font-family:Arial,sans-serif;font-size:30px;font-weight:800;color:" + (totalUrgent > 0 ? "#ff4444" : "#4CAF50") + ";line-height:1;'>" + totalUrgent + "</div>";
  html += "</td></tr></table></td>";
  html += "<td width='33%' style='padding:0 3px;'>";
  html += "<table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='background:#131313;border-radius:8px;border:1px solid rgba(255,255,255,0.06);'>";
  html += "<tr><td style='padding:14px 10px;text-align:center;'>";
  html += "<div style='font-family:Arial,sans-serif;font-size:9px;font-weight:700;letter-spacing:0.2em;text-transform:uppercase;color:rgba(255,255,255,0.5);margin-bottom:4px;'>Upcoming</div>";
  html += "<div style='font-family:Arial,sans-serif;font-size:30px;font-weight:800;color:" + (totalWarning > 0 ? "#E8891A" : "#4CAF50") + ";line-height:1;'>" + totalWarning + "</div>";
  html += "</td></tr></table></td>";
  html += "<td width='34%' style='padding-left:6px;'>";
  html += "<table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='background:#131313;border-radius:8px;border:1px solid rgba(255,255,255,0.06);'>";
  html += "<tr><td style='padding:14px 10px;text-align:center;'>";
  html += "<div style='font-family:Arial,sans-serif;font-size:9px;font-weight:700;letter-spacing:0.2em;text-transform:uppercase;color:rgba(255,255,255,0.5);margin-bottom:4px;'>Patients</div>";
  html += "<div style='font-family:Arial,sans-serif;font-size:30px;font-weight:800;color:#ffffff;line-height:1;'>" + activePatients + "</div>";
  html += "</td></tr></table></td>";
  html += "</tr></table>";

  // ---- URGENT CARDS (red accent) ----
  if (overdue.length > 0) {
    html += card("#ff4444", "Overdue Refills (" + overdue.length + ")", itemsList(overdue, true));
  }
  if (billingPastDue.length > 0) {
    html += card("#ff4444", "Billing Past Due (" + billingPastDue.length + ")", itemsList(billingPastDue, false));
  }
  if (unreadEscalation.length > 0) {
    html += card("#ff4444", "Unread Messages (" + unreadEscalation.length + ")", itemsList(unreadEscalation, false));
  }

  // ---- UPCOMING CARDS (orange accent) ----
  if (dueSoon.length > 0) {
    html += card("#E8891A", "Refills Due Soon (" + dueSoon.length + ")", itemsList(dueSoon, true));
  }
  if (billingDue.length > 0) {
    html += card("#E8891A", "Billing Due (" + billingDue.length + ")", itemsList(billingDue, false));
  }
  if (followUps.length > 0) {
    html += card("#E8891A", "Follow-Ups (" + followUps.length + ")", nameList(followUps));
  }

  // ---- LEADS (blue accent) ----
  if (leadFollowUps.length > 0) {
    html += card("#6B8AFF", "Lead Follow-Ups (" + leadFollowUps.length + ")", itemsList(leadFollowUps, false));
  }

  // ---- LABS (minimized — single condensed row at the bottom) ----
  if (labsOverdueCount > 0 || labsDueSoonCount > 0) {
    var labsBody = "<div style='font-family:Arial,sans-serif;font-size:13px;color:rgba(255,255,255,0.75);'>";
    if (labsOverdueCount > 0) labsBody += "<span style='color:#ff4444;font-weight:700;'>" + labsOverdueCount + " overdue</span>";
    if (labsOverdueCount > 0 && labsDueSoonCount > 0) labsBody += "<span style='color:rgba(255,255,255,0.25);margin:0 8px;'>•</span>";
    if (labsDueSoonCount > 0) labsBody += "<span style='color:#E8891A;font-weight:700;'>" + labsDueSoonCount + " due within 30 days</span>";
    labsBody += "<div style='font-size:11px;color:rgba(255,255,255,0.4);margin-top:4px;'>Open dashboard for details</div>";
    labsBody += "</div>";
    html += card("rgba(255,255,255,0.15)", "Labs", labsBody);
  }

  // ---- ALL-CLEAR STATE ----
  if (totalUrgent === 0 && totalWarning === 0) {
    html += "<table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='margin:0 0 14px 0;background:rgba(76,175,80,0.12);border-radius:8px;border:1px solid rgba(76,175,80,0.3);'>";
    html += "<tr><td style='padding:24px 16px;text-align:center;'>";
    html += "<div style='font-family:Arial,sans-serif;font-size:18px;font-weight:800;color:#4CAF50;'>All clear.</div>";
    html += "<div style='font-family:Arial,sans-serif;font-size:13px;color:rgba(255,255,255,0.6);margin-top:4px;'>No action items today.</div>";
    html += "</td></tr></table>";
  }

  // ---- CTA BUTTON ----
  html += "<table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='margin:8px 0 4px 0;'>";
  html += "<tr><td align='center' style='padding:10px 0;'>";
  html += "<a href='" + ADMIN_URL + "' style='display:inline-block;padding:14px 36px;background:#E8891A;color:#ffffff;text-decoration:none;border-radius:6px;font-family:Arial,sans-serif;font-weight:700;font-size:13px;letter-spacing:0.12em;text-transform:uppercase;'>Open Dashboard</a>";
  html += "</td></tr></table>";

  sendBrandedEmail("both", "Daily Digest — " + formatDateStr(today), html);
}


// ---- DAILY REFILL ALERTS (12 PM) ----
function dailyRefillAlerts() {
  var patients = getSheetData("Patients");
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var alertDays = safeNumber(getSettingValue("Refill Alert Days Before Due")) || 14;
  var allOrders = getAllPatientOrders();

  for (var i = 0; i < patients.length; i++) {
    var status = safeString(patients[i][P_STATUS]);
    if (status !== "Active") continue;
    var name = safeString(patients[i][P_NAME]);
    var phone = safeString(patients[i][P_PHONE]);
    var preferred = safeString(patients[i][P_PREFERRED]) || name.split(" ")[0];

    var orderEntry = allOrders[name.toLowerCase()];
    var orders = orderEntry ? orderEntry.orders : [];
    if (orders.length === 0) continue;
    var nextDue = parseDate(orders[0].nextDue);
    if (!nextDue) continue;

    var daysLeft = daysBetween(today, nextDue);

    if (daysLeft === alertDays || daysLeft === 7 || daysLeft === 3) {
      var msg = "Hi " + preferred + ", your " + safeString(patients[i][P_MED])
        + " supply is running low (" + daysLeft + " days remaining). "
        + "Reply YES to confirm your refill or visit " + APP_URL;
      smartSendMessage(name, phone, "Refill Reminder", msg, false);
      appendRefillLog(name, safeString(patients[i][P_MED]),
        "Refill alert sent (" + daysLeft + " days)", "Auto", "");
    } else if (daysLeft < 0 && (Math.abs(daysLeft) % 3 === 0)) {
      var overdueMsg = "Hi " + preferred + ", you are " + Math.abs(daysLeft)
        + " days past your " + safeString(patients[i][P_MED])
        + " refill date. Please confirm your refill: " + APP_URL
        + " or reply YES to this text.";
      smartSendMessage(name, phone, "Refill Overdue", overdueMsg, true);
      appendRefillLog(name, safeString(patients[i][P_MED]),
        "Overdue alert sent (" + Math.abs(daysLeft) + " days)", "Auto", "");
    }
  }
}


// ---- DAILY CHECK-INS (12 PM) ----
function dailyCheckIns() {
  var schedule = getSheetData("Check-Ins");
  var today = new Date();
  var dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  var todayDay = dayNames[today.getDay()];

  for (var i = 0; i < schedule.length; i++) {
    var preferredDay = safeString(schedule[i][CIS_DAY]);
    if (preferredDay.toLowerCase() !== todayDay.toLowerCase()) continue;
    if (safeString(schedule[i][CIS_STATUS]) !== "Active") continue;

    var patient = safeString(schedule[i][CIS_PATIENT]);
    var phone = safeString(schedule[i][CIS_PHONE]);
    var med = safeString(schedule[i][CIS_MED]);

    // Check if already sent this month
    var lastSent = parseDate(schedule[i][CIS_LASTSENT]);
    if (lastSent) {
      var daysSince = daysBetween(lastSent, today);
      if (daysSince < 25) continue; // Don't send more than once a month
    }

    var msg = "Hi " + patient.split(" ")[0] + ", it's time for your monthly "
      + med + " check-in. How are you feeling? "
      + "Tap here to complete it: " + APP_URL;
    smartSendMessage(patient, phone, "Monthly Check-In", msg, false);

    // Update last sent
    var sheet = getSheet("Check-Ins");
    sheet.getRange(i + 2, CIS_LASTSENT + 1).setValue(today);
  }
}


// ---- DAILY WEIGHT TEXTS (12 PM, GLP-1 only) ----
function dailyWeightTexts() {
  var patients = getSheetData("Patients");
  var today = new Date();
  var dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  var todayDay = dayNames[today.getDay()];

  for (var i = 0; i < patients.length; i++) {
    var glpDay = safeString(patients[i][P_GLPDAY]);
    if (!glpDay || glpDay.toLowerCase() !== todayDay.toLowerCase()) continue;
    var status = safeString(patients[i][P_STATUS]);
    if (status !== "Active") continue;

    var name = safeString(patients[i][P_NAME]);
    var phone = safeString(patients[i][P_PHONE]);
    var preferred = safeString(patients[i][P_PREFERRED]) || name.split(" ")[0];

    var msg = "Hi " + preferred + ", time for your weekly weigh-in! "
      + "Log your weight in the app: " + APP_URL;
    smartSendMessage(name, phone, "Weekly Weigh-In", msg, false);
  }
}


// Escalation, billing status, lead follow-ups, and lab compliance
// are now all handled inside dailyDigest() above.
// No separate emails — one digest covers everything.


// ---- MONTHLY CSV IMPORT REMINDER (runs daily, fires on last day of month) ----
function monthlyImportReminder() {
  var today = new Date();
  var tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);

  // Check if tomorrow is the 1st (meaning today is the last day of the month)
  if (tomorrow.getDate() !== 1) return;

  var monthNames = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];
  var monthLabel = monthNames[today.getMonth()] + " " + today.getFullYear();
  var adminUrl = "https://framemedicine.com/admin";

  var html = "<h2 style='color:#E8891A;'>Monthly CSV Import Reminder</h2>"
    + "<p>It's the last day of <strong>" + monthLabel + "</strong>.</p>"
    + "<p>Time to export and import your monthly data:</p>"
    + "<ol>"
    + "<li><strong>Sales CSV</strong> - Export from JaneApp and import in the admin app</li>"
    + "<li><strong>Patients CSV</strong> - Export from JaneApp and import in the admin app</li>"
    + "</ol>"
    + "<p><a href='" + adminUrl + "' style='color:#E8891A;font-weight:bold;'>Open Admin App to Import</a></p>";

  sendBrandedEmail("tom", "Monthly CSV Import Reminder - " + monthLabel, html);
}
