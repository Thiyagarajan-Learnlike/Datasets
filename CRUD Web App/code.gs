/**
 * LIGHTWEIGHT LMS BACKEND CORE ENGINE
 * Specializing in secure Google Workspace micro-databases.
 */

// GLOBAL CONFIGURATION: Replace with the ID printed from createAndInitializeDemoDatabase()
var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
var SHEET_NAME = 'Courses';

/**
 * Initializes a brand new Google Sheet database with demo schema and structural styling.
 * Run this function ONCE from the Apps Script editor toolbar before launching.
 */
function createAndInitializeDemoDatabase() {
  try {
    var ss = SpreadsheetApp.create('LMS_Demo_Database');
    var sheet = ss.getActiveSheet();
    sheet.setName(SHEET_NAME);
    
    // Inject structural headers
    var headers = [["Course ID", "Course Name", "Instructor", "Duration"]];
    var headerRange = sheet.getRange(1, 1, 1, 4);
    headerRange.setValues(headers);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#F3F4F6");
    headerRange.setFontColor("#1F2937");
    
    // Append initial seed data row
    var sampleData = [[101, "Introduction to Full-Stack Apps Script", "Alex Mercer, Principal Dev", "6 Weeks"]];
    sheet.getRange(2, 1, 1, 4).setValues(sampleData);
    
    // Auto-resize columns for clean visualization
    sheet.autoResizeColumns(1, 4);
    
    Logger.log('================================================================');
    Logger.log(' DATABASE INITIALIZATION SUCCESSFUL!');
    Logger.log(' SPREADSHEET ID: ' + ss.getId());
    Logger.log(' -> Copy the ID above and paste it into the SPREADSHEET_ID global variable.');
    Logger.log('================================================================');
    
    return ss.getId();
  } catch (error) {
    Logger.log('Error initializing database: ' + error.toString());
    throw new Error("Initialization failed: " + error.message);
  }
}

/**
 * Serves the frontend single-page HTML application wrapper with cross-iframe permissions.
 */
function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('LMS Core Web Ledger')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper function to open the specific configured Spreadsheet
 */
function _getDbSheet() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error("Invalid Spreadsheet ID. Please run createAndInitializeDemoDatabase() and configure SPREADSHEET_ID.");
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
}

/**
 * READ API: Fetches, maps, and objects-serializes all active course records.
 * @return {Array<Object>} Array of courses
 */
function fetchCoursesFromDatabase() {
  try {
    var sheet = _getDbSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return []; // Return empty array if only headers exist
    }
    
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 4);
    var values = dataRange.getValues();
    
    return values.map(function(row) {
      return {
        courseId: Number(row[0]),
        courseName: String(row[1]),
        instructor: String(row[2]),
        duration: String(row[3])
      };
    });
  } catch (error) {
    throw new Error("Failed to fetch courses: " + error.message);
  }
}

/**
 * CREATE API: Formulates a safe concurrent transactional row appendment with auto-increment tracking IDs.
 */
function addCourseToDatabase(courseName, instructor, duration) {
  var lock = LockService.getScriptLock();
  try {
    // Request an exclusive execution lock block with a 10-second maximum timeout ceiling
    lock.waitLock(10000);
    
    var sheet = _getDbSheet();
    var lastRow = sheet.getLastRow();
    var nextId = 101; // Base initialization sequence default
    
    if (lastRow > 1) {
      var lastIdValue = sheet.getRange(lastRow, 1).getValue();
      if (!isNaN(lastIdValue) && lastIdValue !== "") {
        nextId = Number(lastIdValue) + 1;
      }
    }
    
    // Append transactionally stable record
    sheet.appendRow([nextId, courseName, instructor, duration]);
    return { success: true, generatedId: nextId };
    
  } catch (error) {
    throw new Error("Concurrency collision or system error during append action: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * UPDATE API: Locates a row targeting Column A and overwrites metadata indexes.
 */
function updateCourseInDatabase(courseId, courseName, instructor, duration) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    var sheet = _getDbSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error("No existing records found to modify.");
    
    var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var targetRowIndex = -1;
    
    for (var i = 0; i < idValues.length; i++) {
      if (Number(idValues[i][0]) === Number(courseId)) {
        targetRowIndex = i + 2; // Offset tracking index array layout to real match
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      throw new Error("Target record ID " + courseId + " could not be located in database layer.");
    }
    
    // Write updating attributes across vector block sequence
    sheet.getRange(targetRowIndex, 2, 1, 3).setValues([[courseName, instructor, duration]]);
    return { success: true };
    
  } catch (error) {
    throw new Error("Failed updating course transaction: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * DELETE API: Drops a target index row entirely out of physical sheet grids safely.
 */
function deleteCourseFromDatabase(courseId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    var sheet = _getDbSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error("No data structures available to delete.");
    
    var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var targetRowIndex = -1;
    
    for (var i = 0; i < idValues.length; i++) {
      if (Number(idValues[i][0]) === Number(courseId)) {
        targetRowIndex = i + 2;
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      throw new Error("Target row tracking index not found.");
    }
    
    sheet.deleteRow(targetRowIndex);
    return { success: true };
    
  } catch (error) {
    throw new Error("Failed system deletion workflow: " + error.message);
  } finally {
    lock.releaseLock();
  }
}
