/**
 * 1. PORTAL PAGE LOADER
 * Serves the primary user interface (Index.html) when someone opens the web app link.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Career Education Unit Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. LOGIN SECURITY CHECK
 * Verifies user credentials against the "Users" sheet.
 */
function checkLogin(username, password) {
  // SAFETY SHIELD: Do not allow blank entries
  if (!username || !password) {
    return { success: false, message: "Username or Password cannot be blank." };
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, message: "Database error: Active spreadsheet not accessible." };
    }

    var sheet = ss.getSheetByName("Users");
    if (!sheet) {
      return { success: false, message: "Database config error: 'Users' sheet not found." };
    }
    
    // Fetch all values inside the populated area of the sheet
    var data = sheet.getDataRange().getValues();
    
    var inputUser = username.toString().trim().toLowerCase();
    var inputPass = password.toString().trim(); 
    
    // Loop through rows (skipping the header row)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Safety Check: Skip empty rows or cells with formula errors
      if (!row[0] || !row[1] || row[0] instanceof Error || row[1] instanceof Error) {
        continue; 
      }

      var dbUser = row[0].toString().trim().toLowerCase();
      var dbPass = row[1].toString().trim();
      
      // Match found!
      if (inputUser === dbUser && inputPass === dbPass) {
        logUserActivity(username); 
        
        // Returns success along with user role profile metadata (Column C)
        return { 
          success: true, 
          username: dbUser,
          role: row[2] ? row[2].toString().trim() : "User" 
        };
      }
    }
    
    return { success: false, message: "Invalid username or password." };
    
  } catch (error) {
    return { success: false, message: "Backend System Error: " + error.toString() };
  }
}

/**
 * 3. ACTIVITY AUDIT LOGGER
 * Tracks and logs successful logins into a dedicated "Logs" sheet tab.
 */
function logUserActivity(username) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    var logSheet = ss.getSheetByName("Logs");
    
    // AUTO-CREATE "Logs" sheet if it does not exist yet
    if (!logSheet) {
      logSheet = ss.insertSheet("Logs");
      logSheet.appendRow(["Timestamp", "Username", "Activity Description"]);
    }
    
    logSheet.appendRow([new Date(), username, "Logged In Successfully"]);
  } catch(e) {
    Logger.log("Logging failed: " + e.toString());
  }
}

/**
 * 4. SYSTEM STAMP / VERSION CONTROL
 */
const SYSTEM_BUILD_VERSION = "16-07-2026 09:11 AM";

/**
 * Feeds the current build version string to the user interface on load.
 */
function getScriptEditTime() {
  return SYSTEM_BUILD_VERSION;
}

/**
 * 5. DIGITAL SERVICE DESK PIPELINE
 */
function saveInquiry(name, email, phone, category, message) {
  if (!name || !email || !phone || !category || !message) {
    return { success: false, message: "All form fields must be completed." };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, message: "Database error: Active spreadsheet not accessible." };
    }

    var sheet = ss.getSheetByName("Inquiries");
    
    // Auto-create sheet with new headers if missing
    if (!sheet) {
      sheet = ss.insertSheet("Inquiries");
      sheet.appendRow(["Timestamp", "Full Name", "Email Address", "Contact Number", "Inquiry Category", "Detailed Message", "Status"]);
    }
    
    // Append row with contact details included
    sheet.appendRow([
      new Date(), 
      name, 
      email,
      phone,
      category, 
      message,
      "Pending Review"
    ]);
    
    return { success: true };
    
  } catch (error) {
    Logger.log("Form submission failed: " + error.toString());
    return { success: false, message: "Database Error: " + error.toString() };
  }
}