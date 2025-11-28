/**
 * Lab Staff Check-In/Check-Out System
 * SIMPLIFIED DAILY QR CODE VERSION
 * - Auto-generates at 7 AM daily
 * - Manual generation option
 * - Simple QR Display (no security)
 * - All dates FORCED to plain text
 */

// Configuration Constants
const STAFF_SHEET = "Staff";
const CHECKIN_LOG_SHEET = "Check-In Logs";
const ADMIN_SHEET = "Admin";
const QR_CODE_SHEET = "QR Codes";
const DATE_TIME_FORMAT = "yyyy-MM-dd HH:mm:ss";
const DATE_FORMAT = "yyyy-MM-dd";

/**
 * Handles GET requests and routes to appropriate pages
 */
function doGet(e) {
  try {
    const page = e.parameter.page || 'main';
    const qrToken = e.parameter.token || '';
    
    if (page === 'admin') {
      return HtmlService.createTemplateFromFile('AdminLogin')
        .evaluate()
        .setTitle('Lab Admin Login')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
    if (page === 'dashboard') {
      return HtmlService.createTemplateFromFile('AdminDashboard')
        .evaluate()
        .setTitle('Lab Admin Dashboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
 
    // QR Display Page
  if (page === 'qrdisplay') {
    return HtmlService.createTemplateFromFile('QRDisplay')
    .evaluate()
    .setTitle('Lab QR Code Display')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
    
    // Main check-in page
    const template = HtmlService.createTemplateFromFile('CheckInPage');
    template.qrToken = qrToken;
    return template.evaluate()
      .setTitle('Lab Check-In System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  } catch (error) {
    return createErrorPage('Error: ' + error.message);
  }
}

/**
 * Include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Create error page
 */
function createErrorPage(message) {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
      </head>
      <body style="font-family: Arial; padding: 20px; background: #f8f9fa;">
        <div style="background: #dc3545; color: white; padding: 20px; border-radius: 10px;">
          <h2>⚠️ Error</h2>
          <p>${message}</p>
        </div>
      </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html);
}

// ============================================
// ADMIN AUTHENTICATION FUNCTIONS
// ============================================

function validateAdmin(username, password) {
  try {
    const ss = SpreadsheetApp.getActive();
    const adminSheet = ss.getSheetByName(ADMIN_SHEET);
    
    if (!adminSheet) {
      return { success: false, message: "Admin sheet not found. Please set up admin credentials." };
    }
    
    const data = adminSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const storedUsername = data[i][0]?.toString().trim();
      const storedPassword = data[i][1]?.toString().trim();
      const name = data[i][2] || 'Admin';
      
      if (storedUsername === username && storedPassword === password) {
        return { 
          success: true, 
          message: "Login successful",
          adminName: name
        };
      }
    }
    
    return { success: false, message: "Invalid username or password" };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

// ============================================
// DAILY QR CODE MANAGEMENT
// ============================================

/**
 * Auto-generate daily QR code at 7 AM
 */
function automaticDailyQRGeneration() {
  try {
    Logger.log("Starting automatic daily QR code generation...");
    const result = generateDailyQRCode();
    
    if (result.success) {
      Logger.log("SUCCESS: " + result.message);
      Logger.log("Token: " + result.token);
    } else {
      Logger.log("ERROR: " + result.message);
    }
    
    return result;
  } catch (error) {
    Logger.log("FATAL ERROR: " + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * Test function to check if QR display is working
 */
function testQRDisplay() {
  try {
    const result = getCurrentQRCode();
    Logger.log('QR Test Result: ' + JSON.stringify(result));
    return result;
  } catch (error) {
    Logger.log('QR Test Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Setup daily trigger at 7 AM - RUN THIS ONCE
 */
function setupDailyQRTrigger() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'automaticDailyQRGeneration') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // Create 7 AM daily trigger
    ScriptApp.newTrigger('automaticDailyQRGeneration')
      .timeBased()
      .atHour(7)
      .everyDays(1)
      .create();
    
    Logger.log("Daily QR trigger created! Runs at 7:00 AM daily.");
    
    // Generate first QR immediately
    const firstQR = generateDailyQRCode();
    
    return {
      success: true,
      message: "Daily trigger set up at 7:00 AM. First QR code generated.",
      firstQR: firstQR
    };
    
  } catch (error) {
    Logger.log("Error setting up trigger: " + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}

/**
 * Generate daily QR code with BULLETPROOF plain text enforcement
 */
function generateDailyQRCode() {
  try {
    const ss = SpreadsheetApp.getActive();
    let qrSheet = ss.getSheetByName(QR_CODE_SHEET);
    
    // Create sheet if doesn't exist
    if (!qrSheet) {
      qrSheet = ss.insertSheet(QR_CODE_SHEET);
      qrSheet.appendRow(['Date', 'Token', 'QR Code URL', 'Generated By', 'Timestamp']);
      qrSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
    }
    
    // FORCE plain text format (multiple layers)
    qrSheet.getRange('A:A').setNumberFormat('@STRING@');
    qrSheet.getRange('E:E').setNumberFormat('@STRING@');
    
    const now = new Date();
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), DATE_FORMAT);
    const token = generateToken();
    const webAppUrl = ScriptApp.getService().getUrl();
    const qrCodeUrl = `${webAppUrl}?token=${token}`;
    const qrImageUrl = `https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encodeURIComponent(qrCodeUrl)}`;
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
    
    // Prepend quote to FORCE plain text
    const dateText = "'" + today;
    const timestampText = "'" + timestamp;
    
    // Check if today's QR exists
    const data = qrSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let sheetDate = data[i][0];
      
      if (sheetDate instanceof Date) {
        sheetDate = Utilities.formatDate(sheetDate, Session.getScriptTimeZone(), DATE_FORMAT);
      } else {
        sheetDate = String(sheetDate).replace(/^'/, '').trim();
      }
      
      if (sheetDate === today) {
        // Update existing
        qrSheet.getRange(i + 1, 2).setValue(token);
        qrSheet.getRange(i + 1, 3).setValue(qrCodeUrl);
        qrSheet.getRange(i + 1, 4).setValue('System');
        qrSheet.getRange(i + 1, 5).setValue(timestampText);
        
        return {
          success: true,
          token: token,
          qrCodeUrl: qrCodeUrl,
          qrImageUrl: qrImageUrl,
          date: today,
          timestamp: timestamp,
          message: "Today's QR code updated"
        };
      }
    }
    
    // Add new row
    const newRow = qrSheet.getLastRow() + 1;
    qrSheet.getRange(newRow, 1).setValue(dateText);
    qrSheet.getRange(newRow, 2).setValue(token);
    qrSheet.getRange(newRow, 3).setValue(qrCodeUrl);
    qrSheet.getRange(newRow, 4).setValue('System');
    qrSheet.getRange(newRow, 5).setValue(timestampText);
    
    return {
      success: true,
      token: token,
      qrCodeUrl: qrCodeUrl,
      qrImageUrl: qrImageUrl,
      date: today,
      timestamp: timestamp,
      message: "New daily QR code generated"
    };
    
  } catch (error) {
    Logger.log("Error in generateDailyQRCode: " + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}

/**
 * Generate random token
 */
function generateToken() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let token = '';
  for (let i = 0; i < 32; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

/**
 * Validate QR token (daily)
 */
function validateQRToken(token) {
  try {
    if (!token) {
      return { valid: false, message: "No QR code scanned." };
    }
    
    const ss = SpreadsheetApp.getActive();
    const qrSheet = ss.getSheetByName(QR_CODE_SHEET);
    
    if (!qrSheet) {
      return { valid: false, message: "QR system not initialized." };
    }
    
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
    const data = qrSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      let sheetDate = data[i][0];
      
      if (sheetDate instanceof Date) {
        sheetDate = Utilities.formatDate(sheetDate, Session.getScriptTimeZone(), DATE_FORMAT);
      } else {
        sheetDate = String(sheetDate).replace(/^'/, '').trim();
      }
      
      const sheetToken = String(data[i][1]).trim();
      const inputToken = String(token).trim();
      
      if (sheetDate === today && sheetToken === inputToken) {
        return { valid: true, message: "QR code verified" };
      }
    }
    
    return { valid: false, message: "Invalid or expired QR code." };
    
  } catch (error) {
    return { valid: false, message: "Error: " + error.message };
  }
}

/**
 * Get current QR code
 */
function getCurrentQRCode() {
  try {
    const ss = SpreadsheetApp.getActive();
    let qrSheet = ss.getSheetByName(QR_CODE_SHEET);
    
    if (!qrSheet) {
      Logger.log('No QR sheet, generating first one...');
      return generateDailyQRCode();
    }
    
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
    const data = qrSheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      let sheetDate = data[i][0];
      
      if (sheetDate instanceof Date) {
        sheetDate = Utilities.formatDate(sheetDate, Session.getScriptTimeZone(), DATE_FORMAT);
      } else {
        sheetDate = String(sheetDate).replace(/^'/, '').trim();
      }
      
      if (sheetDate === today) {
        const qrCodeUrl = data[i][2];
        const qrImageUrl = `https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encodeURIComponent(qrCodeUrl)}`;
        
        return {
          success: true,
          date: today,
          token: data[i][1],
          qrCodeUrl: qrCodeUrl,
          qrImageUrl: qrImageUrl,
          generatedBy: data[i][3],
          timestamp: String(data[i][4]).replace(/^'/, '')
        };
      }
    }
    
    Logger.log('No QR for today, generating...');
    return generateDailyQRCode();
    
  } catch (error) {
    Logger.log('Error in getCurrentQRCode: ' + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}

// ============================================
// STAFF MANAGEMENT
// ============================================

function getAllStaff() {
  try {
    const ss = SpreadsheetApp.getActive();
    const staffSheet = ss.getSheetByName(STAFF_SHEET);
    
    if (!staffSheet) {
      return { success: false, message: "Staff sheet not found" };
    }
    
    const data = staffSheet.getDataRange().getValues();
    const staff = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        staff.push({
          id: data[i][0],
          name: data[i][1],
          email: data[i][2],
          role: data[i][3] || 'Staff',
          status: data[i][4] || 'Active'
        });
      }
    }
    
    return { success: true, staff: staff };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function getStaffById(staffId) {
  try {
    const ss = SpreadsheetApp.getActive();
    const staffSheet = ss.getSheetByName(STAFF_SHEET);
    
    if (!staffSheet) {
      return { success: false, message: "Staff sheet not found" };
    }
    
    const data = staffSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]?.toString().trim() === staffId.toString().trim()) {
        return {
          success: true,
          staff: {
            id: data[i][0],
            name: data[i][1],
            email: data[i][2],
            role: data[i][3] || 'Staff',
            status: data[i][4] || 'Active'
          }
        };
      }
    }
    
    return { success: false, message: "Staff member not found" };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function addStaff(staffData) {
  try {
    const ss = SpreadsheetApp.getActive();
    let staffSheet = ss.getSheetByName(STAFF_SHEET);
    
    if (!staffSheet) {
      staffSheet = ss.insertSheet(STAFF_SHEET);
      staffSheet.appendRow(['Staff ID', 'Name', 'Email', 'Role', 'Status', 'Date Added']);
      staffSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#2196F3').setFontColor('white');
    }
    
    const data = staffSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]?.toString().trim() === staffData.id.toString().trim()) {
        return { success: false, message: "Staff ID already exists" };
      }
    }
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
    staffSheet.appendRow([
      staffData.id,
      staffData.name,
      staffData.email,
      staffData.role || 'Staff',
      staffData.status || 'Active',
      timestamp
    ]);
    
    return { success: true, message: "Staff member added successfully" };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function updateStaff(staffData) {
  try {
    const ss = SpreadsheetApp.getActive();
    const staffSheet = ss.getSheetByName(STAFF_SHEET);
    
    if (!staffSheet) {
      return { success: false, message: "Staff sheet not found" };
    }
    
    const data = staffSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]?.toString().trim() === staffData.id.toString().trim()) {
        staffSheet.getRange(i + 1, 2, 1, 4).setValues([[
          staffData.name,
          staffData.email,
          staffData.role,
          staffData.status
        ]]);
        return { success: true, message: "Staff member updated successfully" };
      }
    }
    
    return { success: false, message: "Staff member not found" };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function deleteStaff(staffId) {
  try {
    const ss = SpreadsheetApp.getActive();
    const staffSheet = ss.getSheetByName(STAFF_SHEET);
    
    if (!staffSheet) {
      return { success: false, message: "Staff sheet not found" };
    }
    
    const data = staffSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]?.toString().trim() === staffId.toString().trim()) {
        staffSheet.deleteRow(i + 1);
        return { success: true, message: "Staff member deleted successfully" };
      }
    }
    
    return { success: false, message: "Staff member not found" };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function verifyStaffAndGetStatus(staffId) {
  try {
    const staffResult = getStaffById(staffId);
    if (!staffResult.success) {
      return { success: false, message: "Invalid Staff ID." };
    }
    
    const staff = staffResult.staff;
    
    if (staff.status !== 'Active') {
      return { success: false, message: "Staff account not active." };
    }
    
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      return {
        success: true,
        staff: staff,
        checkedInToday: false,
        checkInTime: null,
        isCheckedIn: false
      };
    }
    
    const data = logSheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      const logStaffId = String(data[i][1]).trim();
      const logDate = data[i][4];
      const logStatus = data[i][8];
      
      let dateStr = logDate;
      if (logDate instanceof Date) {
        dateStr = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
      } else {
        dateStr = String(logDate).replace(/^'/, '').trim();
      }
      
      if (logStaffId === String(staffId).trim() && dateStr === today) {
        if (logStatus === 'Checked In') {
          let checkInTimeValue = data[i][5];
          
          if (typeof checkInTimeValue === 'string') {
            checkInTimeValue = checkInTimeValue.replace(/^'/, '');
          }
          if (checkInTimeValue instanceof Date) {
            checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
          } else {
            checkInTimeValue = String(checkInTimeValue).trim();
          }
          
          return {
            success: true,
            staff: staff,
            checkedInToday: true,
            checkInTime: checkInTimeValue,
            currentlyCheckedIn: true,
            isCheckedIn: true
          };
        } else if (logStatus === 'Checked Out') {
          let checkInTimeValue = data[i][5];
          let checkOutTimeValue = data[i][6];
          
          if (typeof checkInTimeValue === 'string') {
            checkInTimeValue = checkInTimeValue.replace(/^'/, '');
          }
          if (checkInTimeValue instanceof Date) {
            checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
          } else {
            checkInTimeValue = String(checkInTimeValue).trim();
          }
          
          if (typeof checkOutTimeValue === 'string') {
            checkOutTimeValue = checkOutTimeValue.replace(/^'/, '');
          }
          if (checkOutTimeValue instanceof Date) {
            checkOutTimeValue = Utilities.formatDate(checkOutTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
          } else {
            checkOutTimeValue = String(checkOutTimeValue).trim();
          }
          
          return {
            success: true,
            staff: staff,
            checkedInToday: false,
            checkInTime: checkInTimeValue,
            checkOutTime: checkOutTimeValue,
            currentlyCheckedIn: false,
            alreadyCompletedToday: true,
            isCheckedIn: false
          };
        }
      }
    }
    
    return {
      success: true,
      staff: staff,
      checkedInToday: false,
      checkInTime: null,
      isCheckedIn: false
    };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

// ============================================
// CHECK-IN/CHECK-OUT
// ============================================

function checkIn(staffId, qrToken) {
  try {
    const qrValidation = validateQRToken(qrToken);
    if (!qrValidation.valid) {
      return { success: false, message: qrValidation.message };
    }
    
    const staffResult = getStaffById(staffId);
    if (!staffResult.success) {
      return { success: false, message: "Invalid Staff ID" };
    }
    
    const staff = staffResult.staff;
    
    if (staff.status !== 'Active') {
      return { success: false, message: "Staff account not active" };
    }
    
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (logSheet) {
      const data = logSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const logStaffId = String(data[i][1]).trim();
        const logDate = data[i][4];
        
        let dateStr = logDate;
        if (logDate instanceof Date) {
          dateStr = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
        } else {
          dateStr = String(logDate).replace(/^'/, '').trim();
        }
        
        if (logStaffId === String(staffId).trim() && dateStr === today) {
          return {
            success: false,
            message: "Already checked in today."
          };
        }
      }
    }
    
    let finalLogSheet = logSheet;
    if (!finalLogSheet) {
      finalLogSheet = ss.insertSheet(CHECKIN_LOG_SHEET);
      finalLogSheet.appendRow([
        'Log ID', 'Staff ID', 'Staff Name', 'Action', 'Date', 'Check-In Time', 
        'Check-Out Time', 'Duration (Hours)', 'Status', 'Notes'
      ]);
      finalLogSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#FF9800').setFontColor('white');
    }
    
    // FORCE plain text
    finalLogSheet.getRange('E:E').setNumberFormat('@STRING@');
    finalLogSheet.getRange('F:F').setNumberFormat('@STRING@');
    finalLogSheet.getRange('G:G').setNumberFormat('@STRING@');
    
    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), DATE_FORMAT);
    const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
    const logId = 'LOG' + Date.now();
    
    const dateText = "'" + dateStr;
    const timeText = "'" + timeStr;
    
    const newRow = finalLogSheet.getLastRow() + 1;
    finalLogSheet.getRange(newRow, 1).setValue(logId);
    finalLogSheet.getRange(newRow, 2).setValue(staffId);
    finalLogSheet.getRange(newRow, 3).setValue(staff.name);
    finalLogSheet.getRange(newRow, 4).setValue('Check-In');
    finalLogSheet.getRange(newRow, 5).setValue(dateText);
    finalLogSheet.getRange(newRow, 6).setValue(timeText);
    finalLogSheet.getRange(newRow, 7).setValue('');
    finalLogSheet.getRange(newRow, 8).setValue('');
    finalLogSheet.getRange(newRow, 9).setValue('Checked In');
    finalLogSheet.getRange(newRow, 10).setValue('Checked in via QR code');
    
    return { 
      success: true, 
      message: `Welcome ${staff.name}! Checked in successfully.`,
      staff: staff,
      checkInTime: timeStr
    };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function checkOut(staffId) {
  try {
    const staffResult = getStaffById(staffId);
    if (!staffResult.success) {
      return { success: false, message: "Invalid Staff ID" };
    }
    
    const staff = staffResult.staff;
    const checkInStatus = getCheckInStatus(staffId);
    
    if (!checkInStatus.checkedIn) {
      return { success: false, message: "Not currently checked in" };
    }
    
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    const data = logSheet.getDataRange().getValues();
    
    const timestamp = new Date();
    const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
    
    logSheet.getRange('G:G').setNumberFormat('@STRING@');
    
    for (let i = data.length - 1; i >= 1; i--) {
      const logStaffId = String(data[i][1]).trim();
      const logStatus = data[i][8];
      
      if (logStaffId === String(staffId).trim() && logStatus === 'Checked In') {
        let checkInTimeValue = data[i][5];
        
        if (typeof checkInTimeValue === 'string') {
          checkInTimeValue = checkInTimeValue.replace(/^'/, '');
        }
        if (checkInTimeValue instanceof Date) {
          checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        
        const checkInTime = new Date(checkInTimeValue);
        const duration = (timestamp - checkInTime) / (1000 * 60 * 60);
        
        const timeText = "'" + timeStr;
        
        logSheet.getRange(i + 1, 7).setValue(timeText);
        logSheet.getRange(i + 1, 8).setValue(duration.toFixed(2));
        logSheet.getRange(i + 1, 9).setValue('Checked Out');
        
        return { 
          success: true, 
          message: `Goodbye ${staff.name}! Duration: ${duration.toFixed(2)} hours`,
          staff: staff,
          checkOutTime: timeStr,
          duration: duration.toFixed(2)
        };
      }
    }
    
    return { success: false, message: "No active check-in found" };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function getCheckInStatus(staffId) {
  try {
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      return { checkedIn: false };
    }
    
    const data = logSheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      const logStaffId = String(data[i][1]).trim();
      const logStatus = data[i][8];
      
      if (logStaffId === String(staffId).trim()) {
        if (logStatus === 'Checked In') {
          let checkInTimeValue = data[i][5];
          if (typeof checkInTimeValue === 'string') {
            checkInTimeValue = checkInTimeValue.replace(/^'/, '');
          }
          if (checkInTimeValue instanceof Date) {
            checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
          } else {
            checkInTimeValue = String(checkInTimeValue).trim();
          }
          
          return {
            checkedIn: true,
            checkInTime: checkInTimeValue,
            logId: data[i][0]
          };
        } else {
          return { checkedIn: false };
        }
      }
    }
    
    return { checkedIn: false };
    
  } catch (error) {
    return { checkedIn: false };
  }
}

// ============================================
// DASHBOARD & REPORTS
// ============================================

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    const staffSheet = ss.getSheetByName(STAFF_SHEET);
    
    if (!logSheet || !staffSheet) {
      return { success: false, message: "Required sheets not found" };
    }
    
    const logData = logSheet.getDataRange().getValues();
    const staffData = staffSheet.getDataRange().getValues();
    
    const checkedInStaff = [];
    const staffCheckInMap = {};
    
    for (let i = logData.length - 1; i >= 1; i--) {
      const staffId = logData[i][1];
      if (!staffCheckInMap[staffId]) {
        if (logData[i][8] === 'Checked In') {
          let checkInTimeValue = logData[i][5];
          if (typeof checkInTimeValue === 'string') {
            checkInTimeValue = checkInTimeValue.replace(/^'/, '');
          }
          if (checkInTimeValue instanceof Date) {
            checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
          }
          
          checkedInStaff.push({
            staffId: staffId,
            name: logData[i][2],
            checkInTime: checkInTimeValue,
            duration: calculateCurrentDuration(checkInTimeValue)
          });
        }
        staffCheckInMap[staffId] = true;
      }
    }
    
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
    let todayCheckIns = 0;
    let todayCheckOuts = 0;
    let totalHoursToday = 0;
    
    for (let i = 1; i < logData.length; i++) {
      let logDate = logData[i][4];
      
      if (logDate instanceof Date) {
        logDate = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
      } else {
        logDate = String(logDate).replace(/^'/, '').trim();
      }
      
      if (logDate === today) {
        if (logData[i][3] === 'Check-In') todayCheckIns++;
        if (logData[i][8] === 'Checked Out') {
          todayCheckOuts++;
          totalHoursToday += parseFloat(logData[i][7] || 0);
        }
      }
    }
    
    const totalStaff = staffData.length - 1;
    const activeStaff = staffData.filter((row, idx) => idx > 0 && row[4] === 'Active').length;
    
    let todayGuests = 0;
    for (let i = 1; i < logData.length; i++) {
      let logDate = logData[i][4];
      const staffId = logData[i][1];
      
      if (logDate instanceof Date) {
        logDate = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
      } else {
        logDate = String(logDate).replace(/^'/, '').trim();
      }
      
      if (logDate === today && String(staffId).startsWith('GUEST-')) {
        todayGuests++;
      }
    }

    return {
      success: true,
      currentlyCheckedIn: checkedInStaff.length,
      checkedInStaff: checkedInStaff,
      todayCheckIns: todayCheckIns,
      todayCheckOuts: todayCheckOuts,
      totalHoursToday: totalHoursToday.toFixed(2),
      totalStaff: totalStaff,
      activeStaff: activeStaff,
      todayGuests: todayGuests,  // ← ADD THIS
      date: today
    };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function calculateCurrentDuration(checkInTime) {
  try {
    const checkIn = new Date(checkInTime);
    const now = new Date();
    const duration = (now - checkIn) / (1000 * 60 * 60);
    return duration.toFixed(2);
  } catch (error) {
    return "0.00";
  }
}

function getStaffTimeReport(staffId, startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      return { success: false, message: "Log sheet not found" };
    }
    
    const data = logSheet.getDataRange().getValues();
    const logs = [];
    let totalHours = 0;
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    for (let i = 1; i < data.length; i++) {
      let logDate = data[i][4];
      const logStaffId = data[i][1]?.toString().trim();
      
      if (!(logDate instanceof Date)) {
        logDate = new Date(String(logDate).replace(/^'/, ''));
      }
      
      if ((!staffId || logStaffId === staffId.toString().trim()) && 
          logDate >= start && logDate <= end) {
        
        const duration = parseFloat(data[i][7]) || 0;
        totalHours += duration;
        
        let checkInTimeValue = data[i][5];
        let checkOutTimeValue = data[i][6];
        
        if (typeof checkInTimeValue === 'string') {
          checkInTimeValue = checkInTimeValue.replace(/^'/, '');
        }
        if (checkInTimeValue instanceof Date) {
          checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        
        if (typeof checkOutTimeValue === 'string') {
          checkOutTimeValue = checkOutTimeValue.replace(/^'/, '');
        }
        if (checkOutTimeValue instanceof Date) {
          checkOutTimeValue = Utilities.formatDate(checkOutTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        
        logs.push({
          logId: data[i][0],
          staffId: data[i][1],
          staffName: data[i][2],
          action: data[i][3],
          date: data[i][4],
          checkInTime: checkInTimeValue,
          checkOutTime: checkOutTimeValue,
          duration: duration,
          status: data[i][8]
        });
      }
    }
    
    return {
      success: true,
      logs: logs,
      totalHours: totalHours.toFixed(2),
      totalDays: logs.filter(log => log.status === 'Checked Out').length,
      averageHoursPerDay: logs.length > 0 ? (totalHours / logs.length).toFixed(2) : "0.00"
    };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function forceCheckOutAll() {
  try {
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      return { success: false, message: "Log sheet not found" };
    }
    
    logSheet.getRange('G:G').setNumberFormat('@STRING@');
    
    const data = logSheet.getDataRange().getValues();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_TIME_FORMAT);
    const timeText = "'" + timestamp;
    let count = 0;
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][8] === 'Checked In') {
        let checkInTimeValue = data[i][5];
        if (typeof checkInTimeValue === 'string') {
          checkInTimeValue = checkInTimeValue.replace(/^'/, '');
        }
        if (checkInTimeValue instanceof Date) {
          checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        
        const checkInTime = new Date(checkInTimeValue);
        const checkOutTime = new Date();
        const duration = (checkOutTime - checkInTime) / (1000 * 60 * 60);
        
        logSheet.getRange(i + 1, 7).setValue(timeText);
        logSheet.getRange(i + 1, 8).setValue(duration.toFixed(2));
        logSheet.getRange(i + 1, 9).setValue('Checked Out');
        logSheet.getRange(i + 1, 10).setValue('Force checked out by admin');
        count++;
      }
    }
    
    return { 
      success: true, 
      message: `Checked out ${count} staff member(s)`,
      count: count
    };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function manualCheckIn(staffId, dateTime) {
  try {
    const staffResult = getStaffById(staffId);
    if (!staffResult.success) {
      return { success: false, message: "Invalid Staff ID" };
    }
    
    const staff = staffResult.staff;
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    logSheet.getRange('E:E').setNumberFormat('@STRING@');
    logSheet.getRange('F:F').setNumberFormat('@STRING@');
    logSheet.getRange('G:G').setNumberFormat('@STRING@');
    
    const timestamp = new Date(dateTime);
    const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), DATE_FORMAT);
    const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
    const logId = 'LOG' + Date.now();
    
    const dateText = "'" + dateStr;
    const timeText = "'" + timeStr;
    
    const newRow = logSheet.getLastRow() + 1;
    logSheet.getRange(newRow, 1).setValue(logId);
    logSheet.getRange(newRow, 2).setValue(staffId);
    logSheet.getRange(newRow, 3).setValue(staff.name);
    logSheet.getRange(newRow, 4).setValue('Check-In');
    logSheet.getRange(newRow, 5).setValue(dateText);
    logSheet.getRange(newRow, 6).setValue(timeText);
    logSheet.getRange(newRow, 7).setValue('');
    logSheet.getRange(newRow, 8).setValue('');
    logSheet.getRange(newRow, 9).setValue('Checked In');
    logSheet.getRange(newRow, 10).setValue('Manual check-in by admin');
    
    return { success: true, message: "Manual check-in recorded" };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

// ============================================
// REPORTING & LOGS FUNCTIONS (COMPLETE)
// ============================================

/**
 * Get recent check-in logs (last N entries)
 */
function getRecentLogs(limit) {
  try {
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      return { success: false, message: "Log sheet not found", logs: [] };
    }
    
    const data = logSheet.getDataRange().getValues();
    const logs = [];
    
    const maxLimit = limit || 20;
    const startRow = Math.max(1, data.length - maxLimit);
    
    for (let i = data.length - 1; i >= startRow; i--) {
      if (data[i][0]) { // Has Log ID
        let checkInTimeValue = data[i][5];
        let checkOutTimeValue = data[i][6];
        let dateValue = data[i][4];
        
        // Handle quoted text
        if (typeof checkInTimeValue === 'string') {
          checkInTimeValue = checkInTimeValue.replace(/^'/, '');
        }
        if (typeof checkOutTimeValue === 'string') {
          checkOutTimeValue = checkOutTimeValue.replace(/^'/, '');
        }
        if (typeof dateValue === 'string') {
          dateValue = dateValue.replace(/^'/, '');
        }
        
        // Format dates
        if (checkInTimeValue instanceof Date) {
          checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        if (checkOutTimeValue instanceof Date) {
          checkOutTimeValue = Utilities.formatDate(checkOutTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        if (dateValue instanceof Date) {
          dateValue = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), DATE_FORMAT);
        }
        
        logs.push({
          logId: data[i][0],
          staffId: data[i][1],
          staffName: data[i][2],
          action: data[i][3],
          date: dateValue,
          checkInTime: checkInTimeValue,
          checkOutTime: checkOutTimeValue,
          duration: data[i][7] || '',
          status: data[i][8],
          notes: data[i][9] || ''
        });
      }
    }
    
    return { success: true, logs: logs };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message, logs: [] };
  }
}

/**
 * Get logs by date
 */
/**
 * Get logs by date - FIXED TIMEZONE ISSUE
 */
function getLogsByDate(date) {
  try {
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      return { success: false, message: "Log sheet not found", logs: [] };
    }
    
    // Parse the input date in the script's timezone
    const inputDate = new Date(date + 'T00:00:00');
    const targetDate = Utilities.formatDate(inputDate, Session.getScriptTimeZone(), DATE_FORMAT);
    
    Logger.log('Daily Report - Input date: ' + date);
    Logger.log('Daily Report - Target date: ' + targetDate);
    
    const data = logSheet.getDataRange().getValues();
    const logs = [];
    let totalHours = 0;
    let completedSessions = 0;
    
    for (let i = 1; i < data.length; i++) {
      let logDate = data[i][4];
      
      if (logDate instanceof Date) {
        logDate = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
      } else {
        logDate = String(logDate).replace(/^'/, '').trim();
      }
      
      Logger.log('Comparing: Sheet date [' + logDate + '] vs Target [' + targetDate + ']');
      
      if (logDate === targetDate) {
        let checkInTimeValue = data[i][5];
        let checkOutTimeValue = data[i][6];
        const duration = parseFloat(data[i][7]) || 0;
        
        if (typeof checkInTimeValue === 'string') {
          checkInTimeValue = checkInTimeValue.replace(/^'/, '');
        }
        if (typeof checkOutTimeValue === 'string') {
          checkOutTimeValue = checkOutTimeValue.replace(/^'/, '');
        }
        
        if (checkInTimeValue instanceof Date) {
          checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        if (checkOutTimeValue instanceof Date) {
          checkOutTimeValue = Utilities.formatDate(checkOutTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        
        if (data[i][8] === 'Checked Out') {
          totalHours += duration;
          completedSessions++;
        }
        
        logs.push({
          logId: data[i][0],
          staffId: data[i][1],
          staffName: data[i][2],
          action: data[i][3],
          date: logDate,
          checkInTime: checkInTimeValue,
          checkOutTime: checkOutTimeValue,
          duration: duration,
          status: data[i][8],
          notes: data[i][9] || ''
        });
      }
    }
    
    Logger.log('Daily Report - Total logs found: ' + logs.length);
    
    return { 
      success: true, 
      logs: logs, 
      date: targetDate,
      totalLogs: logs.length,
      totalHours: totalHours.toFixed(2),
      completedSessions: completedSessions,
      averageHours: completedSessions > 0 ? (totalHours / completedSessions).toFixed(2) : "0.00"
    };
    
  } catch (error) {
    Logger.log('Error in getLogsByDate: ' + error.message);
    return { success: false, message: "Error: " + error.message, logs: [] };
  }
}
/**
 * Get logs by date range
 */
function getLogsByDateRange(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      return { success: false, message: "Log sheet not found", logs: [] };
    }
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    end.setHours(23, 59, 59, 999); // Include entire end date
    
    const data = logSheet.getDataRange().getValues();
    const logs = [];
    let totalHours = 0;
    let completedSessions = 0;
    const staffMap = {};
    
    for (let i = 1; i < data.length; i++) {
      let logDate = data[i][4];
      
      if (!(logDate instanceof Date)) {
        logDate = new Date(String(logDate).replace(/^'/, ''));
      }
      
      if (logDate >= start && logDate <= end) {
        let checkInTimeValue = data[i][5];
        let checkOutTimeValue = data[i][6];
        const duration = parseFloat(data[i][7]) || 0;
        const staffId = data[i][1];
        
        if (typeof checkInTimeValue === 'string') {
          checkInTimeValue = checkInTimeValue.replace(/^'/, '');
        }
        if (typeof checkOutTimeValue === 'string') {
          checkOutTimeValue = checkOutTimeValue.replace(/^'/, '');
        }
        
        if (checkInTimeValue instanceof Date) {
          checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        if (checkOutTimeValue instanceof Date) {
          checkOutTimeValue = Utilities.formatDate(checkOutTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        
        let formattedDate = logDate;
        if (logDate instanceof Date) {
          formattedDate = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
        }
        
        if (data[i][8] === 'Checked Out') {
          totalHours += duration;
          completedSessions++;
          
          // Track per staff
          if (!staffMap[staffId]) {
            staffMap[staffId] = { name: data[i][2], sessions: 0, hours: 0 };
          }
          staffMap[staffId].sessions++;
          staffMap[staffId].hours += duration;
        }
        
        logs.push({
          logId: data[i][0],
          staffId: staffId,
          staffName: data[i][2],
          action: data[i][3],
          date: formattedDate,
          checkInTime: checkInTimeValue,
          checkOutTime: checkOutTimeValue,
          duration: duration,
          status: data[i][8],
          notes: data[i][9] || ''
        });
      }
    }
    
    // Convert staffMap to array
    const staffStats = Object.keys(staffMap).map(id => ({
      staffId: id,
      staffName: staffMap[id].name,
      sessions: staffMap[id].sessions,
      hours: staffMap[id].hours.toFixed(2)
    })).sort((a, b) => b.hours - a.hours);
    
    return { 
      success: true, 
      logs: logs,
      startDate: Utilities.formatDate(start, Session.getScriptTimeZone(), DATE_FORMAT),
      endDate: Utilities.formatDate(end, Session.getScriptTimeZone(), DATE_FORMAT),
      totalLogs: logs.length,
      totalHours: totalHours.toFixed(2),
      completedSessions: completedSessions,
      averageHours: completedSessions > 0 ? (totalHours / completedSessions).toFixed(2) : "0.00",
      uniqueStaff: Object.keys(staffMap).length,
      staffStats: staffStats
    };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message, logs: [] };
  }
}

/**
 * Get logs by staff ID
 */
/**
 * Get logs by staff ID - FIXED VERSION
 */
/**
 * Get logs by staff ID - WITH DETAILED LOGGING
 */
function getLogsByStaff(staffId, startDate, endDate) {
  try {
    Logger.log('=== getLogsByStaff START ===');
    Logger.log('Input - staffId: ' + staffId);
    Logger.log('Input - startDate: ' + startDate);
    Logger.log('Input - endDate: ' + endDate);
    
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      Logger.log('ERROR: Log sheet not found');
      return { success: false, message: "Log sheet not found", logs: [] };
    }
    
    const data = logSheet.getDataRange().getValues();
    Logger.log('Total rows in sheet: ' + data.length);
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    end.setHours(23, 59, 59, 999);
    
    Logger.log('Parsed start date: ' + start);
    Logger.log('Parsed end date: ' + end);
    
    const logs = [];
    let totalHours = 0;
    let completedSessions = 0;
    let matchCount = 0;
    
    // Log first few staff IDs to see format
    Logger.log('Sample staff IDs from sheet:');
    for (let i = 1; i < Math.min(5, data.length); i++) {
      Logger.log('  Row ' + i + ': [' + data[i][1] + '] (type: ' + typeof data[i][1] + ')');
    }
    Logger.log('Looking for staff ID: [' + staffId + '] (type: ' + typeof staffId + ')');
    
    for (let i = 1; i < data.length; i++) {
      const logStaffId = String(data[i][1]).trim();
      const searchStaffId = String(staffId).trim();
      
      // Check if this log matches the staff ID
      if (logStaffId === searchStaffId) {
        matchCount++;
        
        let logDate = data[i][4];
        
        // Parse date
        if (!(logDate instanceof Date)) {
          const dateStr = String(logDate).replace(/^'/, '').trim();
          logDate = new Date(dateStr);
        }
        
        Logger.log('Match #' + matchCount + ' - Date: ' + logDate + ', Status: ' + data[i][8]);
        
        if (logDate >= start && logDate <= end) {
          let checkInTimeValue = data[i][5];
          let checkOutTimeValue = data[i][6];
          const duration = parseFloat(data[i][7]) || 0;
          
          if (typeof checkInTimeValue === 'string') {
            checkInTimeValue = checkInTimeValue.replace(/^'/, '');
          }
          if (typeof checkOutTimeValue === 'string') {
            checkOutTimeValue = checkOutTimeValue.replace(/^'/, '');
          }
          
          if (checkInTimeValue instanceof Date) {
            checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
          }
          if (checkOutTimeValue instanceof Date) {
            checkOutTimeValue = Utilities.formatDate(checkOutTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
          }
          
          let formattedDate = logDate;
          if (logDate instanceof Date) {
            formattedDate = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
          }
          
          if (data[i][8] === 'Checked Out') {
            totalHours += duration;
            completedSessions++;
          }
          
          logs.push({
            logId: data[i][0],
            staffId: data[i][1],
            staffName: data[i][2],
            action: data[i][3],
            date: formattedDate,
            checkInTime: checkInTimeValue,
            checkOutTime: checkOutTimeValue,
            duration: duration,
            status: data[i][8],
            notes: data[i][9] || ''
          });
        } else {
          Logger.log('  Date out of range');
        }
      }
    }
    
    Logger.log('Total matches found: ' + matchCount);
    Logger.log('Logs in date range: ' + logs.length);
    Logger.log('=== getLogsByStaff END ===');
    
    // Get staff name
    let staffName = staffId;
    const staffResult = getStaffById(staffId);
    if (staffResult.success) {
      staffName = staffResult.staff.name;
      Logger.log('Staff name: ' + staffName);
    }
    
    return { 
      success: true, 
      logs: logs,
      staffId: staffId,
      staffName: staffName,
      startDate: Utilities.formatDate(start, Session.getScriptTimeZone(), DATE_FORMAT),
      endDate: Utilities.formatDate(end, Session.getScriptTimeZone(), DATE_FORMAT),
      totalLogs: logs.length,
      totalHours: totalHours.toFixed(2),
      completedSessions: completedSessions,
      averageHours: completedSessions > 0 ? (totalHours / completedSessions).toFixed(2) : "0.00"
    };
    
  } catch (error) {
    Logger.log('ERROR in getLogsByStaff: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return { success: false, message: "Error: " + error.message, logs: [] };
  }
}


/**
 * Auto check-out all staff at 7 PM daily
 */
function autoCheckOutAt7PM() {
  try {
    Logger.log('Auto check-out triggered at: ' + new Date());
    
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      Logger.log('ERROR: Log sheet not found');
      return { success: false, message: "Log sheet not found" };
    }
    
    logSheet.getRange('G:G').setNumberFormat('@STRING@');
    
    const data = logSheet.getDataRange().getValues();
    const now = new Date();
    
    // Set checkout time to exactly 7:00 PM today
    const checkoutTime = new Date();
    checkoutTime.setHours(19, 0, 0, 0); // 7 PM
    
    const timeStr = Utilities.formatDate(checkoutTime, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
    const timeText = "'" + timeStr;
    
    let count = 0;
    const checkedOutStaff = [];
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][8] === 'Checked In') {
        let checkInTimeValue = data[i][5];
        if (typeof checkInTimeValue === 'string') {
          checkInTimeValue = checkInTimeValue.replace(/^'/, '');
        }
        if (checkInTimeValue instanceof Date) {
          checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        
        const checkInTime = new Date(checkInTimeValue);
        const duration = (checkoutTime - checkInTime) / (1000 * 60 * 60);
        
        // Update the row
        logSheet.getRange(i + 1, 7).setValue(timeText);
        logSheet.getRange(i + 1, 8).setValue(duration.toFixed(2));
        logSheet.getRange(i + 1, 9).setValue('Checked Out');
        logSheet.getRange(i + 1, 10).setValue('Auto checked out at 7 PM');
        
        count++;
        checkedOutStaff.push({
          name: data[i][2],
          id: data[i][1],
          checkInTime: checkInTimeValue,
          duration: duration.toFixed(2)
        });
        
        Logger.log('Auto checked out: ' + data[i][2] + ' (' + data[i][1] + ')');
      }
    }
    
    Logger.log('Auto check-out complete. Total: ' + count + ' staff members');
    
    return { 
      success: true, 
      message: `Auto checked out ${count} staff member(s) at 7:00 PM`,
      count: count,
      staff: checkedOutStaff,
      checkoutTime: timeStr
    };
    
  } catch (error) {
    Logger.log('ERROR in autoCheckOutAt7PM: ' + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}

/**
 * Setup 7 PM auto check-out trigger - RUN THIS ONCE
 */
function setup7PMAutoCheckOut() {
  try {
    // Delete existing auto-checkout triggers
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'autoCheckOutAt7PM') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // Create 7 PM daily trigger
    ScriptApp.newTrigger('autoCheckOutAt7PM')
      .timeBased()
      .atHour(19) // 7 PM (24-hour format)
      .everyDays(1)
      .create();
    
    Logger.log("7 PM auto check-out trigger created successfully!");
    
    return {
      success: true,
      message: "7 PM auto check-out trigger set up successfully. Will run daily at 7:00 PM."
    };
    
  } catch (error) {
    Logger.log("Error setting up 7 PM trigger: " + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}

/**
 * Generate comprehensive report based on type - FIXED VERSION
 */
/**
 * Generate comprehensive report based on type - FIXED STAFF CASE
 */
function generateComprehensiveReport(reportType, staffId, startDate, endDate, singleDate) {
  try {
    let result;
    const timezone = Session.getScriptTimeZone();
    
    switch(reportType) {
      case 'daily':
        result = getLogsByDate(singleDate);
        break;
        
      case 'weekly':
        const today = new Date();
        const dayOfWeek = today.getDay();
        
        const monday = new Date(today);
        const daysToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
        monday.setDate(today.getDate() + daysToMonday);
        monday.setHours(0, 0, 0, 0);
        
        const sunday = new Date(monday);
        sunday.setDate(monday.getDate() + 6);
        sunday.setHours(23, 59, 59, 999);
        
        const weekStart = Utilities.formatDate(monday, timezone, DATE_FORMAT);
        const weekEnd = Utilities.formatDate(sunday, timezone, DATE_FORMAT);
        
        Logger.log('Weekly Report - Start: ' + weekStart + ', End: ' + weekEnd);
        
        result = getLogsByDateRange(weekStart, weekEnd);
        result.periodDescription = 'This Week (Monday to Sunday)';
        break;
        
      case 'monthly':
        const now = new Date();
        const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
        firstDay.setHours(0, 0, 0, 0);
        
        const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        lastDay.setHours(23, 59, 59, 999);
        
        const monthStart = Utilities.formatDate(firstDay, timezone, DATE_FORMAT);
        const monthEnd = Utilities.formatDate(lastDay, timezone, DATE_FORMAT);
        
        Logger.log('Monthly Report - Start: ' + monthStart + ', End: ' + monthEnd);
        
        result = getLogsByDateRange(monthStart, monthEnd);
        result.periodDescription = now.toLocaleString('default', { month: 'long', year: 'numeric' });
        break;
        
      case 'custom':
        result = getLogsByDateRange(startDate, endDate);
        result.periodDescription = 'Custom Range';
        break;
        
      case 'staff':
        // SIMPLIFIED: Always get ALL records for staff
        Logger.log('Staff Report - Getting ALL records for staff: ' + staffId);
        
        // Use very wide date range to get everything
        const allTimeStart = '2000-01-01';
        const allTimeEnd = Utilities.formatDate(new Date(), timezone, DATE_FORMAT);
        
        Logger.log('Date range: ' + allTimeStart + ' to ' + allTimeEnd);
        
        result = getLogsByStaff(staffId, allTimeStart, allTimeEnd);
        result.periodDescription = 'All Time Records';
        break;
        
      default:
        return { success: false, message: "Invalid report type" };
    }
    
    if (result.success) {
      result.reportType = reportType;
      result.generatedAt = Utilities.formatDate(new Date(), timezone, DATE_TIME_FORMAT);
      Logger.log('Report generated successfully. Total logs: ' + result.logs.length);
    } else {
      Logger.log('Report generation failed: ' + result.message);
    }
    
    return result;
    
  } catch (error) {
    Logger.log('Error in generateComprehensiveReport: ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
    return { success: false, message: "Error generating report: " + error.message };
  }
}


function manualCheckOut(staffId, dateTime) {
  try {
    const staffResult = getStaffById(staffId);
    if (!staffResult.success) {
      return { success: false, message: "Invalid Staff ID" };
    }
    
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    const data = logSheet.getDataRange().getValues();
    
    logSheet.getRange('G:G').setNumberFormat('@STRING@');
    
    const timestamp = new Date(dateTime);
    const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
    const timeText = "'" + timeStr;
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === staffId && data[i][8] === 'Checked In') {
        let checkInTimeValue = data[i][5];
        if (typeof checkInTimeValue === 'string') {
          checkInTimeValue = checkInTimeValue.replace(/^'/, '');
        }
        if (checkInTimeValue instanceof Date) {
          checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        
        const checkInTime = new Date(checkInTimeValue);
        const duration = (timestamp - checkInTime) / (1000 * 60 * 60);
        
        logSheet.getRange(i + 1, 7).setValue(timeText);
        logSheet.getRange(i + 1, 8).setValue(duration.toFixed(2));
        logSheet.getRange(i + 1, 9).setValue('Checked Out');
        logSheet.getRange(i + 1, 10).setValue('Manual check-out by admin');
        
        return { success: true, message: "Manual check-out recorded" };
      }
    }
    
    return { success: false, message: "No active check-in found" };
    
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

// ============================================
// GUEST CHECK-IN/CHECK-OUT FUNCTIONS
// ============================================

/**
 * Generate unique guest ID for the day
 */
function generateGuestId() {
  try {
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      return 'GUEST-001';
    }
    
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
    const data = logSheet.getDataRange().getValues();
    
    let maxGuestNumber = 0;
    
    // Find highest guest number for today
    for (let i = 1; i < data.length; i++) {
      const staffId = String(data[i][1]).trim();
      let logDate = data[i][4];
      
      if (logDate instanceof Date) {
        logDate = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
      } else {
        logDate = String(logDate).replace(/^'/, '').trim();
      }
      
      if (logDate === today && staffId.startsWith('GUEST-')) {
        const parts = staffId.split('-');
        if (parts.length >= 3) {
          const guestNum = parseInt(parts[2]) || 0;
          if (guestNum > maxGuestNumber) {
            maxGuestNumber = guestNum;
          }
        }
      }
    }
    
    const newGuestNumber = maxGuestNumber + 1;
    const guestId = 'GUEST-' + today.replace(/-/g, '') + '-' + String(newGuestNumber).padStart(3, '0');
    
    Logger.log('Generated Guest ID: ' + guestId);
    return guestId;
    
  } catch (error) {
    Logger.log('Error generating guest ID: ' + error.message);
    return 'GUEST-' + Date.now();
  }
}

/**
 * Verify guest by name (case-insensitive) and get status
 */
function verifyGuestAndGetStatus(firstName, lastName) {
  try {
    // Normalize names (trim and lowercase for comparison)
    const searchFirstName = String(firstName).trim().toLowerCase();
    const searchLastName = String(lastName).trim().toLowerCase();
    const fullName = firstName.trim() + ' ' + lastName.trim();
    
    Logger.log('Looking for guest: ' + fullName);
    
    if (!searchFirstName || !searchLastName) {
      return { success: false, message: "Please enter both first and last name" };
    }
    
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      return {
        success: true,
        isGuest: true,
        guestName: fullName,
        checkedInToday: false,
        isCheckedIn: false
      };
    }
    
    const data = logSheet.getDataRange().getValues();
    
    // Search for guest by name (case-insensitive) for today
    for (let i = data.length - 1; i >= 1; i--) {
      const logStaffId = String(data[i][1]).trim();
      const logName = String(data[i][2]).trim().toLowerCase();
      let logDate = data[i][4];
      const logStatus = data[i][8];
      
      if (logDate instanceof Date) {
        logDate = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
      } else {
        logDate = String(logDate).replace(/^'/, '').trim();
      }
      
      // Check if this is a guest entry for today with matching name
      if (logStaffId.startsWith('GUEST-') && logDate === today) {
        if (logName === searchFirstName + ' ' + searchLastName) {
          Logger.log('Found matching guest: ' + data[i][2] + ' with status: ' + logStatus);
          
          if (logStatus === 'Checked In') {
            let checkInTimeValue = data[i][5];
            
            if (typeof checkInTimeValue === 'string') {
              checkInTimeValue = checkInTimeValue.replace(/^'/, '');
            }
            if (checkInTimeValue instanceof Date) {
              checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
            } else {
              checkInTimeValue = String(checkInTimeValue).trim();
            }
            
            return {
              success: true,
              isGuest: true,
              guestId: logStaffId,
              guestName: data[i][2],
              checkedInToday: true,
              checkInTime: checkInTimeValue,
              currentlyCheckedIn: true,
              isCheckedIn: true
            };
          } else if (logStatus === 'Checked Out') {
            let checkInTimeValue = data[i][5];
            let checkOutTimeValue = data[i][6];
            
            if (typeof checkInTimeValue === 'string') {
              checkInTimeValue = checkInTimeValue.replace(/^'/, '');
            }
            if (checkInTimeValue instanceof Date) {
              checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
            } else {
              checkInTimeValue = String(checkInTimeValue).trim();
            }
            
            if (typeof checkOutTimeValue === 'string') {
              checkOutTimeValue = checkOutTimeValue.replace(/^'/, '');
            }
            if (checkOutTimeValue instanceof Date) {
              checkOutTimeValue = Utilities.formatDate(checkOutTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
            } else {
              checkOutTimeValue = String(checkOutTimeValue).trim();
            }
            
            return {
              success: true,
              isGuest: true,
              guestId: logStaffId,
              guestName: data[i][2],
              checkedInToday: false,
              checkInTime: checkInTimeValue,
              checkOutTime: checkOutTimeValue,
              currentlyCheckedIn: false,
              alreadyCompletedToday: true,
              isCheckedIn: false
            };
          }
        }
      }
    }
    
    // Guest not found - new guest
    return {
      success: true,
      isGuest: true,
      guestName: fullName,
      checkedInToday: false,
      isCheckedIn: false
    };
    
  } catch (error) {
    Logger.log('Error in verifyGuestAndGetStatus: ' + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}

/**
 * Guest check-in
 */
function guestCheckIn(firstName, lastName, qrToken) {
  try {
    const qrValidation = validateQRToken(qrToken);
    if (!qrValidation.valid) {
      return { success: false, message: qrValidation.message };
    }
    
    // Normalize names
    const fullName = firstName.trim() + ' ' + lastName.trim();
    const searchName = fullName.toLowerCase();
    
    if (!firstName.trim() || !lastName.trim()) {
      return { success: false, message: "Please enter both first and last name" };
    }
    
    Logger.log('Guest check-in: ' + fullName);
    
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    // Check if guest already checked in today
    if (logSheet) {
      const data = logSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const logStaffId = String(data[i][1]).trim();
        const logName = String(data[i][2]).trim().toLowerCase();
        let logDate = data[i][4];
        
        if (logDate instanceof Date) {
          logDate = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
        } else {
          logDate = String(logDate).replace(/^'/, '').trim();
        }
        
        if (logStaffId.startsWith('GUEST-') && logName === searchName && logDate === today) {
          return {
            success: false,
            message: "You are already checked in today."
          };
        }
      }
    }
    
    let finalLogSheet = logSheet;
    if (!finalLogSheet) {
      finalLogSheet = ss.insertSheet(CHECKIN_LOG_SHEET);
      finalLogSheet.appendRow([
        'Log ID', 'Staff ID', 'Staff Name', 'Action', 'Date', 'Check-In Time', 
        'Check-Out Time', 'Duration (Hours)', 'Status', 'Notes'
      ]);
      finalLogSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#FF9800').setFontColor('white');
    }
    
    // FORCE plain text
    finalLogSheet.getRange('E:E').setNumberFormat('@STRING@');
    finalLogSheet.getRange('F:F').setNumberFormat('@STRING@');
    finalLogSheet.getRange('G:G').setNumberFormat('@STRING@');
    
    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), DATE_FORMAT);
    const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
    const guestId = generateGuestId();
    const logId = 'LOG' + Date.now();
    
    const dateText = "'" + dateStr;
    const timeText = "'" + timeStr;
    
    const newRow = finalLogSheet.getLastRow() + 1;
    finalLogSheet.getRange(newRow, 1).setValue(logId);
    finalLogSheet.getRange(newRow, 2).setValue(guestId);
    finalLogSheet.getRange(newRow, 3).setValue(fullName);
    finalLogSheet.getRange(newRow, 4).setValue('Check-In');
    finalLogSheet.getRange(newRow, 5).setValue(dateText);
    finalLogSheet.getRange(newRow, 6).setValue(timeText);
    finalLogSheet.getRange(newRow, 7).setValue('');
    finalLogSheet.getRange(newRow, 8).setValue('');
    finalLogSheet.getRange(newRow, 9).setValue('Checked In');
    finalLogSheet.getRange(newRow, 10).setValue('Guest check-in via QR code');
    
    Logger.log('Guest checked in successfully: ' + fullName + ' (' + guestId + ')');
    
    return { 
      success: true, 
      message: `Welcome ${fullName}! Checked in successfully as guest.`,
      guestId: guestId,
      guestName: fullName,
      checkInTime: timeStr,
      isGuest: true
    };
    
  } catch (error) {
    Logger.log('Error in guestCheckIn: ' + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}

/**
 * Guest check-out
 */
function guestCheckOut(firstName, lastName) {
  try {
    // Normalize names
    const fullName = firstName.trim() + ' ' + lastName.trim();
    const searchName = fullName.toLowerCase();
    
    if (!firstName.trim() || !lastName.trim()) {
      return { success: false, message: "Please enter both first and last name" };
    }
    
    Logger.log('Guest check-out: ' + fullName);
    
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMAT);
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    
    if (!logSheet) {
      return { success: false, message: "No check-in record found" };
    }
    
    const data = logSheet.getDataRange().getValues();
    const timestamp = new Date();
    const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
    
    logSheet.getRange('G:G').setNumberFormat('@STRING@');
    
    // Find most recent checked-in guest with matching name
    for (let i = data.length - 1; i >= 1; i--) {
      const logStaffId = String(data[i][1]).trim();
      const logName = String(data[i][2]).trim().toLowerCase();
      const logStatus = data[i][8];
      let logDate = data[i][4];
      
      if (logDate instanceof Date) {
        logDate = Utilities.formatDate(logDate, Session.getScriptTimeZone(), DATE_FORMAT);
      } else {
        logDate = String(logDate).replace(/^'/, '').trim();
      }
      
      if (logStaffId.startsWith('GUEST-') && logName === searchName && logDate === today && logStatus === 'Checked In') {
        let checkInTimeValue = data[i][5];
        
        if (typeof checkInTimeValue === 'string') {
          checkInTimeValue = checkInTimeValue.replace(/^'/, '');
        }
        if (checkInTimeValue instanceof Date) {
          checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }
        
        const checkInTime = new Date(checkInTimeValue);
        const duration = (timestamp - checkInTime) / (1000 * 60 * 60);
        
        const timeText = "'" + timeStr;
        
        logSheet.getRange(i + 1, 7).setValue(timeText);
        logSheet.getRange(i + 1, 8).setValue(duration.toFixed(2));
        logSheet.getRange(i + 1, 9).setValue('Checked Out');
        
        Logger.log('Guest checked out successfully: ' + data[i][2] + ' (' + logStaffId + ')');
        
        return { 
          success: true, 
          message: `Goodbye ${fullName}! Duration: ${duration.toFixed(2)} hours`,
          guestId: logStaffId,
          guestName: data[i][2],
          checkOutTime: timeStr,
          duration: duration.toFixed(2),
          isGuest: true
        };
      }
    }
    
    return { success: false, message: "No active check-in found for " + fullName + " today. Please verify your name." };
    
  } catch (error) {
    Logger.log('Error in guestCheckOut: ' + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}

/**
 * Automatically check out guests who have been checked in for 2+ hours
 * This function runs periodically (every 30 minutes) via a time-based trigger
 */
function autoCheckOutGuests() {
  try {
    Logger.log('Auto guest check-out triggered at: ' + new Date());

    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);

    if (!logSheet) {
      Logger.log('ERROR: Log sheet not found');
      return { success: false, message: "Log sheet not found" };
    }

    logSheet.getRange('G:G').setNumberFormat('@STRING@');

    const data = logSheet.getDataRange().getValues();
    const now = new Date();
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
    const timeText = "'" + timeStr;

    let count = 0;
    const checkedOutGuests = [];
    const twoHoursInMs = 2 * 60 * 60 * 1000; // 2 hours in milliseconds

    // Iterate through logs from bottom to top (most recent first)
    for (let i = data.length - 1; i >= 1; i--) {
      const logStaffId = String(data[i][1]).trim();
      const logStatus = data[i][8];

      // Only process guests who are currently checked in
      if (logStaffId.startsWith('GUEST-') && logStatus === 'Checked In') {
        let checkInTimeValue = data[i][5];

        if (typeof checkInTimeValue === 'string') {
          checkInTimeValue = checkInTimeValue.replace(/^'/, '');
        }
        if (checkInTimeValue instanceof Date) {
          checkInTimeValue = Utilities.formatDate(checkInTimeValue, Session.getScriptTimeZone(), DATE_TIME_FORMAT);
        }

        const checkInTime = new Date(checkInTimeValue);
        const timeDiff = now - checkInTime;

        // Check if guest has been checked in for 2+ hours
        if (timeDiff >= twoHoursInMs) {
          const duration = timeDiff / (1000 * 60 * 60); // Convert to hours

          // Update the row with checkout information
          logSheet.getRange(i + 1, 7).setValue(timeText); // Check-Out Time
          logSheet.getRange(i + 1, 8).setValue(duration.toFixed(2)); // Duration
          logSheet.getRange(i + 1, 9).setValue('Checked Out'); // Status
          logSheet.getRange(i + 1, 10).setValue('Auto checked out after 2 hours'); // Notes

          count++;
          checkedOutGuests.push({
            name: data[i][2],
            id: logStaffId,
            checkInTime: checkInTimeValue,
            checkOutTime: timeStr,
            duration: duration.toFixed(2)
          });

          Logger.log('Auto checked out guest: ' + data[i][2] + ' (' + logStaffId + ') after ' + duration.toFixed(2) + ' hours');
        }
      }
    }

    Logger.log('Auto guest check-out complete. Total: ' + count + ' guests');

    return {
      success: true,
      message: `Auto checked out ${count} guest(s) who exceeded 2 hours`,
      count: count,
      guests: checkedOutGuests,
      checkoutTime: timeStr
    };

  } catch (error) {
    Logger.log('ERROR in autoCheckOutGuests: ' + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}

/**
 * Setup automatic guest check-out trigger - RUN THIS ONCE
 * Runs every 30 minutes to check and auto-checkout guests who exceeded 2 hours
 */
function setupAutoGuestCheckOut() {
  try {
    // Delete existing auto-guest-checkout triggers
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'autoCheckOutGuests') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }

    // Create trigger that runs every 30 minutes
    ScriptApp.newTrigger('autoCheckOutGuests')
      .timeBased()
      .everyMinutes(30)
      .create();

    Logger.log("Auto guest check-out trigger created successfully! Runs every 30 minutes.");

    return {
      success: true,
      message: "Auto guest check-out trigger set up successfully. Will run every 30 minutes to check guests."
    };

  } catch (error) {
    Logger.log("Error setting up auto guest check-out trigger: " + error.message);
    return { success: false, message: "Error: " + error.message };
  }
}