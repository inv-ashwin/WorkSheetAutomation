// ======================================================
// TIMESHEET GENERATOR - MULTI PROJECT VERSION
// ======================================================

// ---------------- CONFIGURATION ----------------

const STATUS_OPTIONS = ["In progress", "Completed", "On hold"];

const ACTIVITY_OPTIONS = [
  "Development",
  "System design",
  "Unit testing",
  "Testing",
  "Bug fix(QA)",
  "Bug fix(Customer)",
  "Verification testing(Customer)",
  "Release",
  "Maintenance Tasks",
  "Meeting",
  "Requirement study / Requirement Analysis",
  "Investigation"
];

const HEADER_BG = "#FEF2CB";
const WEEKEND_BG = "#FF9999";

const CONFIG_FOLDER_NAME = "Timesheet_Configs";

let DESTINATION_FOLDER_ID = "";
let WEBHOOK_URL = "";
let TEST_ENV = false;
let HOLIDAY_DATES = [
  "2026-01-01",
  "2026-01-26",
  "2026-03-20",
  "2026-04-03",
  "2026-04-15",
  "2026-05-01",
  "2026-08-15",
  "2026-10-02",
  "2026-12-25"
];

function syncSettingsFromSheet_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    const sheet = ss.getSheetByName("Manager Sheet");
    if (!sheet) {
      throw new Error("Sheet 'Manager Sheet' not found in Master Config Spreadsheet. Please create it.");
    }
    const values = sheet.getDataRange().getValues();
    const props = PropertiesService.getScriptProperties();
    
    // Map of friendly names / cell labels to script property keys
    const settingMap = {
      "Destination Folder ID": "DESTINATION_FOLDER_ID",
      "Google Chat Webhook": "GOOGLE_CHAT_WEBHOOK_URL",
      "Employee Alert Webhook": "EMPLOYEE_ALERT_WEBHOOK_URL",
      "Test Mode": "TEST_MODE",
      "Holidays": "HOLIDAYS",
      "Rows to Check After Date": "ROWS_TO_CHECK_AFTER_DATE",
      "Sheet Data Range": "SHEET_DATA_RANGE"
    };

    // Initialize/reset optional settings in properties to defaults before reading sheet
    props.setProperty("TEST_MODE", "false");

    // Dynamically find settings column headers
    let keyIdx = 2; // Default to Column C
    let valIdx = 3; // Default to Column D
    if (values.length > 2) {
      const headerRow = values[2]; // Row 3
      for (let c = 0; c < headerRow.length; c++) {
        const cellVal = String(headerRow[c] || "").trim().toLowerCase();
        if (cellVal === "key") keyIdx = c;
        if (cellVal === "value") valIdx = c;
      }
    }
    
    // Settings start at row 4 (index 3) and stop at the "Work Sheet Manager" header
    for (let i = 3; i < values.length; i++) {
      const name = String(values[i][keyIdx] || "").trim();
      const valColB = values[i].length > 1 ? String(values[i][1] || "").trim() : "";
      const valColC = values[i].length > 2 ? String(values[i][2] || "").trim() : "";
      
      // Stop if we reach the header of the Work Sheet Manager table
      if (name.toLowerCase().includes("project name") || 
          name.toLowerCase().includes("work sheet manager") || 
          valColB.toLowerCase().includes("work sheet manager") ||
          valColC.toLowerCase().includes("work sheet manager") ||
          valColB.toLowerCase().includes("sl no") ||
          valColC.toLowerCase().includes("sl no")) {
        break;
      }
      
      const value = String(values[i][valIdx] || "").trim();
      const propKey = settingMap[name];
      if (propKey) {
        if (propKey === "HOLIDAYS") {
          let holidaysArr = [];
          if (value) {
            holidaysArr = value.split(",").map(d => d.trim()).filter(d => d);
          }
          props.setProperty("HOLIDAYS", JSON.stringify(holidaysArr));
        } else {
          props.setProperty(propKey, value);
        }
      }
    }
    
    Logger.log("✓ Settings synced successfully from 'Manager Sheet'.");
  } catch (e) {
    Logger.log("Error syncing settings: " + e);
    throw e;
  }
  
  loadGlobalsFromProperties_();

  // Validate that required properties are set and do not contain placeholder/default text
  const props = PropertiesService.getScriptProperties();
  const dest = props.getProperty("DESTINATION_FOLDER_ID");
  const managerWebhook = props.getProperty("GOOGLE_CHAT_WEBHOOK_URL");
  const employeeWebhook = props.getProperty("EMPLOYEE_ALERT_WEBHOOK_URL");

  if (!dest || dest.trim() === "" || dest === "YOUR_DESTINATION_FOLDER_ID") {
    throw new Error("Missing Destination Folder ID. Please add the folder ID to the settings in your 'Manager Sheet' to generate the sheet.");
  }
  if (!managerWebhook || managerWebhook.trim() === "" || managerWebhook === "YOUR_WEBHOOK_URL") {
    throw new Error("Missing Google Chat Webhook. Please add the webhook URL to the settings in your 'Manager Sheet' to generate the sheet.");
  }
  if (!employeeWebhook || employeeWebhook.trim() === "" || employeeWebhook === "YOUR_ALERT_WEBHOOK_URL") {
    throw new Error("Missing Employee Alert Webhook. Please add the webhook URL to the settings in your 'Manager Sheet' to generate the sheet.");
  }
}

function loadGlobalsFromProperties_() {
  const props = PropertiesService.getScriptProperties();
  DESTINATION_FOLDER_ID = props.getProperty("DESTINATION_FOLDER_ID") || "";
  WEBHOOK_URL = props.getProperty("GOOGLE_CHAT_WEBHOOK_URL") || "";
  
  // Set default values for test mode, ranges and holidays if not already configured
  if (!props.getProperty("TEST_MODE")) {
    props.setProperty("TEST_MODE", "false");
  }
  TEST_ENV = props.getProperty("TEST_MODE").toLowerCase() === "true";
  
  if (!props.getProperty("ROWS_TO_CHECK_AFTER_DATE")) {
    props.setProperty("ROWS_TO_CHECK_AFTER_DATE", "5");
  }
  if (!props.getProperty("SHEET_DATA_RANGE")) {
    props.setProperty("SHEET_DATA_RANGE", "A1:K150");
  }
  
  let holidaysStr = props.getProperty("HOLIDAYS");
  if (!holidaysStr) {
    props.setProperty("HOLIDAYS", JSON.stringify(HOLIDAY_DATES));
    holidaysStr = JSON.stringify(HOLIDAY_DATES);
  }
  
  try {
    HOLIDAY_DATES = JSON.parse(holidaysStr);
  } catch (e) {
    Logger.log("Error parsing HOLIDAYS from properties: " + e);
  }
}

/************************************************************
 * FETCH CONFIGURATION FROM MASTER SPREADSHEET
 ************************************************************/

function fetchProjectsConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Active spreadsheet is not accessible. Make sure the script is bound to the Master Config Spreadsheet.");
  }

  const sheet = ss.getSheetByName("Manager Sheet");
  if (!sheet) {
    throw new Error("Sheet 'Manager Sheet' not found in Master Config Spreadsheet");
  }

  const values = sheet.getDataRange().getValues();
  const projects = {};
  const emails = {};
  const employeeEmails = {};

  // Dynamically find table column headers
  let projIdx = 2;   // Default to Column C
  let memberIdx = 3; // Default to Column D
  let emailIdx = 4;  // Default to Column E
  let activeIdx = 6; // Default to Column G

  if (values.length > 10) {
    const headerRow = values[10]; // Row 11
    for (let c = 0; c < headerRow.length; c++) {
      const cellVal = String(headerRow[c] || "").trim().toLowerCase();
      if (cellVal.includes("project name")) projIdx = c;
      if (cellVal.includes("team member")) memberIdx = c;
      if (cellVal.includes("email")) emailIdx = c;
      if (cellVal.includes("active")) activeIdx = c;
    }
  }

  for (let i = 11; i < values.length; i++) {
    const row = values[i];
    if (row.length <= Math.max(projIdx, memberIdx)) continue;

    const projectName = (row[projIdx] || "").toString().trim();
    const memberName = (row[memberIdx] || "").toString().trim();
    const email = row.length > emailIdx ? (row[emailIdx] || "").toString().trim() : "";

    if (!projectName || !memberName) continue;

    // Check active status - default to true if empty/not provided
    let isActive = true;
    if (row.length > activeIdx && row[activeIdx] !== undefined && row[activeIdx] !== null) {
      const activeVal = row[activeIdx];
      if (activeVal === false) {
        isActive = false;
      } else {
        const activeStr = String(activeVal).trim().toLowerCase();
        if (activeStr === "false" || activeStr === "0" || activeStr === "no" || activeStr === "inactive") {
          isActive = false;
        }
      }
    }

    if (!isActive) {
      Logger.log("Skipping inactive/benched employee: " + memberName + " in project " + projectName);
      continue;
    }

    if (!projects[projectName]) {
      projects[projectName] = [];
    }
    projects[projectName].push(memberName);

    if (email) {
      if (!emails[projectName]) {
        emails[projectName] = [];
      }
      emails[projectName].push(email);

      if (!employeeEmails[projectName]) {
        employeeEmails[projectName] = {};
      }
      employeeEmails[projectName][memberName] = email;
    }
  }

  return {
    projects: projects,
    emails: emails,
    employeeEmails: employeeEmails
  };
}

/************************************************************
 * GRANT ACCESS BASED ON SPREADSHEET EMAILS
 ************************************************************/

function grantProjectAccess_(file, emails) {
  if (!emails || emails.length === 0) {
    Logger.log("⚠ No emails found for sharing.");
    return;
  }

  emails.forEach(email => {
    try {
      file.addEditor(email);   // Change to addViewer() if needed
      Logger.log("✓ Access granted to: " + email);
    } catch (error) {
      Logger.log("❌ Permission error: " + error);
    }
  });
}

// ======================================================
// HELPER FOR DETECTING EXISTING SPREADSHEETS
// ======================================================

function getExistingSpreadsheetId_(projectName, monthName, year) {
  try {
    const configFolder = getOrCreateConfigFolder_();
    const configFileName = projectName + "_timesheet_config.json";
    const files = configFolder.getFilesByName(configFileName);
    if (files.hasNext()) {
      const file = files.next();
      const content = file.getBlob().getDataAsString();
      const parsed = JSON.parse(content || "{}");
      const key = year + "-" + monthName;
      return parsed[key] || null;
    }
  } catch (e) {
    Logger.log("Error getting existing spreadsheet ID: " + e);
  }
  return null;
}

// ======================================================
// MAIN FUNCTION
// ======================================================

function createMonthlyTimesheet(monthName = null, year = null, testMode = true) {
  // Sync properties from Manager Sheet settings first
  syncSettingsFromSheet_();

  if (arguments.length < 3) {
    testMode = TEST_ENV;
  }

  const now = new Date();

  if (!monthName) {
    monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");
  }

  if (!year) {
    year = now.getFullYear();
  }

  const dates = getDatesForMonth_(monthName, year);
  const config = fetchProjectsConfig_();
  const projects = config.projects;
  const projectEmails = config.emails;

  for (const projectName in projects) {

    const employees = projects[projectName];
    const emails = projectEmails[projectName] || [];
    const workbookName =
      projectName + "_work_report_" + monthName + "_" + year;

    let ss = null;
    let isNew = false;
    let existingId = null;

    if (!testMode) {
      existingId = getExistingSpreadsheetId_(projectName, monthName, year);
      if (existingId) {
        try {
          ss = SpreadsheetApp.openById(existingId);
          Logger.log("Found existing spreadsheet: " + workbookName + " (ID: " + existingId + ")");
        } catch (e) {
          Logger.log("Could not open existing spreadsheet by ID, will create a new one. Error: " + e);
        }
      }
    }

    if (!ss) {
      isNew = true;
      if (testMode) {
        Logger.log("TEST MODE → Would create new spreadsheet: " + workbookName);
        employees.forEach(function (emp) {
          Logger.log("TEST MODE → Would create tab: " + emp);
        });
        const fakeId = "TEST_" + Utilities.getUuid();
        updateConfigJson_(projectName, fakeId, monthName, year, true);
        continue;
      }

      ss = SpreadsheetApp.create(workbookName);
      const file = DriveApp.getFileById(ss.getId());

      if (DESTINATION_FOLDER_ID) {
        const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
        file.moveTo(folder);
      }
    }

    const defaultSheet = isNew ? ss.getSheets()[0] : null;

    const newlyAddedEmployees = [];
    employees.forEach(function (emp) {
      const sheet = ss.getSheetByName(emp);
      if (!sheet) {
        Logger.log("Creating new tab for employee: " + emp + " in project " + projectName);
        setupEmployeeSheet_(ss, emp, dates, projectName);
        newlyAddedEmployees.push(emp);
      } else {
        Logger.log("Tab already exists for employee: " + emp + " (skipping)");
      }
    });

    if (isNew) {
      ss.deleteSheet(defaultSheet);
    }

    const url = ss.getUrl();
    const file = DriveApp.getFileById(ss.getId());
    
    // Grant access: if spreadsheet is new, share with all active project members.
    // If spreadsheet already existed, share ONLY with newly added employees.
    let emailsToShare = [];
    if (isNew) {
      emailsToShare = emails;
    } else {
      const employeeEmails = config.employeeEmails[projectName] || {};
      newlyAddedEmployees.forEach(function (emp) {
        const email = employeeEmails[emp];
        if (email) {
          emailsToShare.push(email);
        }
      });
    }

    grantProjectAccess_(file, emailsToShare);
    updateConfigJson_(projectName, ss.getId(), monthName, year);
    
    if (isNew) {
      Logger.log("✓ New spreadsheet '" + workbookName + "' created and shared.");
    } else {
      Logger.log("✓ Existing spreadsheet '" + workbookName + "' updated and shared.");
    }
  }

  Logger.log("✓ All project sheets processed successfully");
}

// ======================================================
// CONFIG FOLDER HANDLING
// ======================================================

function getOrCreateConfigFolder_() {

  const parentFolder = DESTINATION_FOLDER_ID
    ? DriveApp.getFolderById(DESTINATION_FOLDER_ID)
    : DriveApp.getRootFolder();

  const folders = parentFolder.getFoldersByName(CONFIG_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(CONFIG_FOLDER_NAME);
}

function updateConfigJson_(projectName, spreadsheetId, monthName, year, testMode = false) {

  const configFolder = getOrCreateConfigFolder_();
  const configFileName = projectName + "_timesheet_config.json";
  const key = year + "-" + monthName;

  const newEntry = {};
  newEntry[key] = spreadsheetId;

  if (testMode) {
    Logger.log("TEST MODE → JSON Update:");
    Logger.log(JSON.stringify(newEntry, null, 2));
    return;
  }

  const files = configFolder.getFilesByName(configFileName);

  if (files.hasNext()) {
    const file = files.next();
    const content = file.getBlob().getDataAsString();
    const parsed = JSON.parse(content || "{}");

    parsed[key] = spreadsheetId;

    file.setContent(JSON.stringify(parsed, null, 2));
  } else {
    configFolder.createFile(
      configFileName,
      JSON.stringify(newEntry, null, 2),
      MimeType.PLAIN_TEXT
    );
  }
}

// ======================================================
// DATE GENERATION
// ======================================================

function getDatesForMonth_(monthName, year) {
  Logger.log(monthName)

  const monthMap = {
    January: 0, February: 1, March: 2, April: 3,
    May: 4, June: 5, July: 6, August: 7,
    September: 8, October: 9, November: 10, December: 11
  };

  const monthNum = monthMap[monthName];
  const lastDay = new Date(year, monthNum + 1, 0).getDate();
  const monthAbbr = Utilities.formatDate(
    new Date(year, monthNum, 1),
    Session.getScriptTimeZone(),
    "MMM"
  );

  const dates = [];

  for (let day = 1; day <= lastDay; day++) {

    const date = new Date(year, monthNum, day);
    const isWeekend = date.getDay() === 0 || date.getDay() === 6;

    dates.push({
      date: day + "-" + monthAbbr,
      isWeekend: isWeekend,
      dateObj: date
    });
  }

  return dates;
}

function isHoliday_(dateObj) {

  const formatted = Utilities.formatDate(
    dateObj,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  return HOLIDAY_DATES.indexOf(formatted) !== -1;
}

// ======================================================
// SHEET FORMATTER
// ======================================================

/**
 * Create and format an individual employee timesheet
 */
function setupEmployeeSheet_(ss, empName, dates,projectName) {
  // Create new sheet
  const sheet = ss.insertSheet(empName);

  // Set column widths
  sheet.setColumnWidth(1, 30); // A
  sheet.setColumnWidth(2, 80); // B - Date
  sheet.setColumnWidth(3, 400); // C - Module/Area
  sheet.setColumnWidth(4, 400); // D - Task details
  sheet.setColumnWidth(5, 100); // E - Status
  sheet.setColumnWidth(6, 150); // F - Activity Type
  sheet.setColumnWidth(7, 70); // G - Start
  sheet.setColumnWidth(8, 70); // H - End
  sheet.setColumnWidth(9, 70); // I - Task
  sheet.setColumnWidth(10, 70); // J - Total
  sheet.setColumnWidth(11, 420); // K - Remarks

  // Set Calibri font for the whole sheet
  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily("Calibri");

  // Row 2: Title
  sheet.getRange("B2:J2").merge();
  sheet
    .getRange("B2")
    .setValue(projectName+"- " + empName)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(12);

  // Row 4: Headers with proper structure
  // Main headers row
  const mainHeaders = [
    "",
    "Date",
    "Module/Area",
    "Task details/Ticket number",
    "Status",
    "Activity Type",
    "",
    "",
    "",
    "",
    "Remarks",
  ];
  sheet
    .getRange(4, 1, 1, mainHeaders.length)
    .setValues([mainHeaders])
    .setBackground(HEADER_BG)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(12);

  // Merge cells for "Time" header (columns G-H)
  sheet.getRange(4, 7, 1, 2).merge();
  sheet.getRange(4, 7).setValue("Time");

  // Merge cells for "Duration" header (columns I-J)
  sheet.getRange(4, 9, 1, 2).merge();
  sheet.getRange(4, 9).setValue("Duration");

  // Row 5: Sub-headers for Time and Duration
  const subHeaders = [
    "",
    "",
    "",
    "",
    "",
    "",
    "Start",
    "End",
    "Task",
    "Total",
    "",
  ];
  sheet
    .getRange(5, 1, 1, subHeaders.length)
    .setValues([subHeaders])
    .setBackground(HEADER_BG)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(12);

  // Merge Date, Module/Area, Task details, Status, Activity Type, and Remarks vertically (rows 4-5)
  sheet.getRange(4, 2, 2, 1).merge(); // Date
  sheet.getRange(4, 3, 2, 1).merge(); // Module/Area
  sheet.getRange(4, 4, 2, 1).merge(); // Task details
  sheet.getRange(4, 5, 2, 1).merge(); // Status
  sheet.getRange(4, 6, 2, 1).merge(); // Activity Type
  sheet.getRange(4, 11, 2, 1).merge(); // Remarks

  // Data rows: 2 rows per date starting from row 6 (after headers)
  let currentRow = 6;

  dates.forEach(function (dateInfo) {
    const dateStr = dateInfo.date;
    const isWeekend = dateInfo.isWeekend;
    const dateObj = dateInfo.dateObj; // available from getDatesForMonth

    // Merge date cells vertically (2 rows) and center align
    const dateRange = sheet.getRange(currentRow, 2, 2, 1);
    dateRange.merge();
    dateRange
      .setValue(dateStr)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Add task duration formulas for both rows
    for (let offset = 0; offset < 2; offset++) {
      const row = currentRow + offset;
      const taskFormula =
        "=IF(AND(G" +
        row +
        "<TIME(12,30,0),H" +
        row +
        ">TIME(13,0,0)),H" +
        row +
        "-G" +
        row +
        "-TIME(0,30,0),H" +
        row +
        "-G" +
        row +
        ")";
      sheet.getRange(row, 9).setFormula(taskFormula);
    }

    // Add total formula in first row only (sums both task rows) and merge vertically
    const totalFormula = "=SUM(I" + currentRow + ":I" + (currentRow + 1) + ")";
    const totalRange = sheet.getRange(currentRow, 10, 2, 1);
    totalRange.merge();
    totalRange
      .setFormula(totalFormula)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontWeight("bold")
      .setFontSize(11);

    // Format Start and End as clock time
    sheet.getRange(currentRow, 7, 2, 2).setNumberFormat("hh:mm");

    // Format Task and Total as duration hh:mm
    sheet.getRange(currentRow, 9, 2, 2).setNumberFormat("hh:mm");

    // Color weekend rows or holidays
    if (isWeekend || isHoliday_(dateObj)) {
      sheet.getRange(currentRow, 1, 2, 11).setBackground(WEEKEND_BG);
    }

    // Add data validation for Status (column E)
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(STATUS_OPTIONS, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(currentRow, 5, 2, 1).setDataValidation(statusRule);

    // Add data validation for Activity Type (column F)
    const activityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(ACTIVITY_OPTIONS, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(currentRow, 6, 2, 1).setDataValidation(activityRule);

    currentRow += 2;
  });

  const lastDataRow = currentRow - 1;

  // Add "Monthly hours spend" row
  const monthlyRow = currentRow + 1;
  sheet.getRange(monthlyRow, 2, 1, 8).merge();
  sheet
    .getRange(monthlyRow, 2)
    .setValue("Monthly hours spend")
    .setFontWeight("bold")
    .setHorizontalAlignment("right");

  // Monthly total formula
  const monthlyFormula =
    "=TEXT(INT(SUM(J6:J" +
    lastDataRow +
    "))*24+HOUR(SUM(J6:J" +
    lastDataRow +
    ')),"00")&":"&' +
    "TEXT(MINUTE(SUM(J6:J" +
    lastDataRow +
    ')),"00")';
  sheet
    .getRange(monthlyRow, 10)
    .setFormula(monthlyFormula)
    .setFontWeight("bold");

  // Add borders to all cells
  const allDataRange = sheet.getRange(4, 1, monthlyRow - 3, 11);
  allDataRange.setBorder(true, true, true, true, true, true);

  // Set font size 11 for all data rows (from row 6 to last data row)
  sheet
    .getRange(6, 1, sheet.getMaxRows() - 5, sheet.getMaxColumns())
    .setFontSize(11);
  // Center-align text in column I (Task column)
  sheet
    .getRange(6, 7, sheet.getMaxRows() - 5, 3)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  Logger.log("✓ " + empName + " completed (" + dates.length + " dates)");
}




// ======================================================
// CUSTOM MENU AUTOMATION
// ======================================================

/**
 * Automatically runs when the spreadsheet is opened.
 * Adds a custom menu to the spreadsheet UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Timesheet Automation")
    .addItem("Generate Monthly Reports", "btnGenerateMonthlyReports_")
    .addItem("Run Daily Verification", "btnRunDailyVerification_")
    .addToUi();
}

/**
 * Wrapper to run Generate Monthly Reports with verification/production options.
 */
function btnGenerateMonthlyReports_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Generate Monthly Reports",
    "Are you sure you want to generate the monthly reports for the current month?\n\nThis will create and share spreadsheets for all active projects and team members.",
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    ss.toast("Generating monthly reports... Please wait.", "Monthly Reports", -1);
    try {
      createMonthlyTimesheet(null, null, false);
      ss.toast("Monthly reports generated successfully!", "Success", 5);
      ui.alert("Success", "Monthly reports generated successfully!", ui.ButtonSet.OK);
    } catch (e) {
      ss.toast("Failed to generate monthly reports.", "Error", 5);
      ui.alert("Error", "Failed to generate monthly reports: " + e.toString(), ui.ButtonSet.OK);
    }
  }
}

/**
 * Wrapper to run Daily Verification.
 */
function btnRunDailyVerification_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Run Daily Verification",
    "Are you sure you want to run the daily timesheet verification right now?\n\nThis will scan employee timesheets and send any missing summaries/reminders.",
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    ss.toast("Running daily verification... Please wait.", "Daily Verification", -1);
    try {
      runDailyVerification();
      ss.toast("Daily verification run completed!", "Success", 5);
      ui.alert("Success", "Daily verification completed successfully!", ui.ButtonSet.OK);
    } catch (e) {
      ss.toast("Failed to run daily verification.", "Error", 5);
      ui.alert("Error", "Failed to run daily verification: " + e.toString(), ui.ButtonSet.OK);
    }
  }
}
