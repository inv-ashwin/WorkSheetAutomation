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
/************************************************************
 * HYPERLINK FOR SHEET ID
 ************************************************************/
function setSheetLink(range, sheetId) {
  if (!sheetId) return;

  const url = `https://docs.google.com/spreadsheets/d/${sheetId}`;

  const richText = SpreadsheetApp.newRichTextValue()
    .setText(sheetId)
    .setLinkUrl(url)
    .build();

  range.setRichTextValue(richText);
}

function syncSettingsFromSheet_(requireWebhooks = false) {
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

    const seenPropKeys = {};

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
          name.toLowerCase().includes("team member") || 
          name.toLowerCase().includes("name") || 
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
        seenPropKeys[propKey] = true;
        if (propKey === "HOLIDAYS") {
          let holidaysArr = [];
          if (value) {
            holidaysArr = value.split(",").map(d => d.trim()).filter(d => d);
            props.setProperty("HOLIDAYS", JSON.stringify(holidaysArr));
          } else {
            props.deleteProperty("HOLIDAYS");
          }
        } else {
          props.setProperty(propKey, value);
        }
      }
    }

    // Delete any properties from settingMap that were not present in the sheet settings
    const allPropKeys = Object.keys(settingMap).map(k => settingMap[k]);
    allPropKeys.forEach(propKey => {
      if (!seenPropKeys[propKey]) {
        props.deleteProperty(propKey);
      }
    });
    
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
  if (requireWebhooks) {
    if (!managerWebhook || managerWebhook.trim() === "" || managerWebhook === "YOUR_WEBHOOK_URL") {
      throw new Error("Missing Google Chat Webhook. Please add the webhook URL to the settings in your 'Manager Sheet' to run daily verification.");
    }
    if (!employeeWebhook || employeeWebhook.trim() === "" || employeeWebhook === "YOUR_ALERT_WEBHOOK_URL") {
      throw new Error("Missing Employee Alert Webhook. Please add the webhook URL to the settings in your 'Manager Sheet' to run daily verification.");
    }
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

function fetchEmployeesConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Active spreadsheet is not accessible. Make sure the script is bound to the Master Config Spreadsheet.");
  }

  const sheet = ss.getSheetByName("Manager Sheet");
  if (!sheet) {
    throw new Error("Sheet 'Manager Sheet' not found in Master Config Spreadsheet");
  }

  const values = sheet.getDataRange().getValues();
  const employees = [];

  // Dynamically find table column headers
  let projIdx = -1;
  let memberIdx = 2;   // Default to Column C
  let emailIdx = 3;    // Default to Column D
  let chatIdx = 4;     // Default to Column E
  let sheetIdIdx = 5;  // Default to Column F
  let activeIdx = 6;   // Default to Column G
  let templateIdx = -1;

  if (values.length > 10) {
    const headerRow = values[10]; // Row 11
    for (let c = 0; c < headerRow.length; c++) {
      const cellVal = String(headerRow[c] || "").trim().toLowerCase();
      if (cellVal.includes("project name")) projIdx = c;
      if (cellVal.includes("team member") || cellVal.includes("name")) memberIdx = c;
      if (cellVal.includes("email")) emailIdx = c;
      if (cellVal.includes("chat id") || cellVal.includes("chatid")) chatIdx = c;
      if (cellVal.includes("sheet id") || cellVal.includes("sheetid")) sheetIdIdx = c;
      if (cellVal.includes("active")) activeIdx = c;
      if (cellVal.includes("template type") || cellVal.includes("template")) templateIdx = c;
    }
  }

  for (let i = 11; i < values.length; i++) {
    const row = values[i];
    if (row.length <= memberIdx) continue;

    const projectName = projIdx !== -1 && row.length > projIdx ? (row[projIdx] || "").toString().trim() : "";
    const memberName = (row[memberIdx] || "").toString().trim();
    const email = row.length > emailIdx ? (row[emailIdx] || "").toString().trim() : "";
    const chatId = row.length > chatIdx ? (row[chatIdx] || "").toString().trim() : "";
    const sheetId = row.length > sheetIdIdx ? (row[sheetIdIdx] || "").toString().trim() : "";
    const templateTypeRaw = templateIdx !== -1 && row.length > templateIdx ? (row[templateIdx] || "").toString().trim().toLowerCase() : "standard";
    
    let templateType = "Standard";
    if (templateTypeRaw === "alternative" || templateTypeRaw === "alt" || templateTypeRaw === "template 2" || templateTypeRaw === "template2") {
      templateType = "Alternative";
    }

    if (!memberName) continue;

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
      Logger.log("Skipping inactive/benched employee: " + memberName);
      continue;
    }

    employees.push({
      name: memberName,
      email: email,
      chatId: chatId,
      sheetId: sheetId,
      projectName: projectName,
      templateType: templateType,
      rowNum: i + 1, // 1-indexed row number
      sheetIdColNum: sheetIdIdx + 1 // 1-indexed column number
    });
  }

  return employees;
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
  const employees = fetchEmployeesConfig_();
  const tabName = monthName + "-" + year; // Format: month-year (e.g. June-2026)

  const ssMaster = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMaster = ssMaster.getSheetByName("Manager Sheet");

  employees.forEach(function (emp) {
    const workbookName = emp.name + "_work_report_" + year;
    let ss = null;
    let isNew = false;
    let sheetId = emp.sheetId;

    if (!testMode) {
      if (sheetId) {
        try {
          ss = SpreadsheetApp.openById(sheetId);
          Logger.log("Found existing spreadsheet for " + emp.name + " (ID: " + sheetId + ")");
        } catch (e) {
          Logger.log("Could not open existing spreadsheet by ID for " + emp.name + ", will create a new one. Error: " + e);
        }
      }
    }

    if (!ss) {
      isNew = true;
      if (testMode) {
        Logger.log("TEST MODE → Would create new spreadsheet: " + workbookName);
        Logger.log("TEST MODE → Would create tab: " + tabName);
        return;
      }

      ss = SpreadsheetApp.create(workbookName);
      sheetId = ss.getId();

      // Write the new sheet ID back to the Master Config Sheet immediately
      const cell = sheetMaster.getRange(emp.rowNum, emp.sheetIdColNum);

      cell.setValue(sheetId);
      setSheetLink(cell, sheetId);
      Logger.log("✓ Saved sheet ID " + sheetId + " for " + emp.name + " in row " + emp.rowNum);

      const file = DriveApp.getFileById(sheetId);
      if (DESTINATION_FOLDER_ID) {
        const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
        file.moveTo(folder);
      }
    }

    const defaultSheet = isNew ? ss.getSheets()[0] : null;

    const tab = ss.getSheetByName(tabName);
    if (!tab) {
      const templateType = emp.templateType || "Standard";
      const handler = TimesheetTemplates[templateType] || TimesheetTemplates["Standard"];
      Logger.log("Creating new tab: " + tabName + " for employee: " + emp.name + " using template: " + templateType);
      handler.setupSheet(ss, tabName, dates, emp.name, emp.projectName);
    } else {
      Logger.log("Tab " + tabName + " already exists for employee: " + emp.name + " (skipping)");
    }

    if (isNew && defaultSheet) {
      ss.deleteSheet(defaultSheet);
    }

    if (isNew && emp.email) {
      const file = DriveApp.getFileById(ss.getId());
      grantProjectAccess_(file, [emp.email]);
      Logger.log("✓ Spreadsheet '" + workbookName + "' shared with " + emp.email);
    }

    // Sort sheets chronologically to ensure they are ordered correctly
    if (!testMode && ss) {
      sortSheetsChronologically_(ss);
    }
  });

  Logger.log("✓ All employee sheets processed successfully");
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
function setupStandardEmployeeSheet_(ss, tabName, dates, empName, projectName) {
  // Create new sheet
  const sheet = ss.insertSheet(tabName);

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
  const titleValue = projectName ? projectName + "- " + empName : empName;
  sheet
    .getRange("B2")
    .setValue(titleValue)
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

// ======================================================
// CHRONOLOGICAL SHEET SORTING HELPERS
// ======================================================

function sortSheetsChronologically_(ss) {
  const sheets = ss.getSheets();
  const sheetsWithDate = [];
  const otherSheets = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    const date = parseSheetDate_(name);
    if (date) {
      sheetsWithDate.push({ sheet: sheet, date: date });
    } else {
      otherSheets.push(sheet);
    }
  });

  // Sort chronological sheets: oldest to newest
  sheetsWithDate.sort((a, b) => a.date - b.date);

  // Combine them: other sheets (Readme/Templates) first, then chronological sheets
  const sortedSheets = [...otherSheets, ...sheetsWithDate.map(item => item.sheet)];

  // Apply new order to spreadsheet
  for (let i = 0; i < sortedSheets.length; i++) {
    ss.setActiveSheet(sortedSheets[i]);
    ss.moveActiveSheet(i + 1); // 1-indexed in Google Apps Script
  }
}

function parseSheetDate_(name) {
  const parts = name.split("-");
  if (parts.length !== 2) return null;
  const monthMap = {
    january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
    july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
  };
  const month = monthMap[parts[0].toLowerCase()];
  const year = parseInt(parts[1], 10);
  if (month === undefined || isNaN(year)) return null;
  return new Date(year, month, 1);
}

// ======================================================
// TEMPLATE STRATEGY REGISTRY
// ======================================================

const TimesheetTemplates = {
  "Standard": {
    setupSheet: function(ss, tabName, dates, empName, projectName) {
      setupStandardEmployeeSheet_(ss, tabName, dates, empName, projectName);
    }
  },
  "Alternative": {
    setupSheet: function(ss, tabName, dates, empName, projectName) {
      setupAlternativeEmployeeSheet_(ss, tabName, dates, empName, projectName);
    }
  }
};

// ======================================================
// ALTERNATIVE SHEET FORMATTER
// ======================================================

function setupAlternativeEmployeeSheet_(ss, tabName, dates, empName, projectName) {
  // Create new sheet
  const sheet = ss.insertSheet(tabName);

  // Set column widths
  sheet.setColumnWidth(1, 30); // A
  sheet.setColumnWidth(2, 80); // B - Date
  sheet.setColumnWidth(3, 80); // C - Day
  sheet.setColumnWidth(4, 400); // D - Module/Area
  sheet.setColumnWidth(5, 400); // E - Task details
  sheet.setColumnWidth(6, 400); // F - Backlog URL
  sheet.setColumnWidth(7, 100); // G - Status
  sheet.setColumnWidth(8, 70); // H - Start
  sheet.setColumnWidth(9, 70); // I - End
  sheet.setColumnWidth(10, 70); // J - Task
  sheet.setColumnWidth(11, 70); // K - Total
  sheet.setColumnWidth(12, 420); // L - Remarks

  // Set Calibri font for the whole sheet
  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily("Calibri");

  // Calculate working days & period
  let workingDays = 0;
  dates.forEach(function (dateInfo) {
    if (!dateInfo.isWeekend && !isHoliday_(dateInfo.dateObj)) {
      workingDays++;
    }
  });
  const standardHours = workingDays * 8;

  const firstDate = dates[0].dateObj;
  const lastDate = dates[dates.length - 1].dateObj;
  const periodStr = Utilities.formatDate(firstDate, Session.getScriptTimeZone(), "yyyy/MM/dd") + " ~ " + Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

  // Row 2: Title Headers
  sheet.getRange("B2:C2").merge();
  sheet.getRange("B2").setValue("Work Report")
    .setBackground("#1a73e8")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange("D2").setValue("ProjectName:").setBackground("#ffe0b2").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("E2").setValue("Engineer Name").setBackground("#ffe0b2").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("F2").setValue("Work reporting period").setBackground("#ffe0b2").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("G2").setValue("Standard").setBackground("#ffe0b2").setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange("H2:K2").merge();
  sheet.getRange("H2").setValue("Total Work")
    .setBackground("#2e7d32")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Row 3: Title Values
  sheet.getRange("D3").setValue(projectName || "").setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange("E3").setValue(empName).setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange("F3").setValue(periodStr).setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange("G3").setValue(standardHours).setHorizontalAlignment("center").setFontWeight("bold");

  const startRow = 6;
  const endRow = 5 + (dates.length * 2);
  sheet.getRange("H3:K3").merge();
  sheet.getRange("H3").setFormula("=SUM(K" + startRow + ":K" + endRow + ")")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Align Title Row heights
  sheet.setRowHeight(2, 25);
  sheet.setRowHeight(3, 25);

  // Row 4 & 5: Table Headers
  const mainHeaders = [
    "",
    "Date",
    "Day",
    "Module/Area",
    "Task details",
    "Backlog URL",
    "Status",
    "",
    "",
    "",
    "",
    "Remarks",
  ];
  sheet
    .getRange(4, 1, 1, mainHeaders.length)
    .setValues([mainHeaders])
    .setBackground("#ffe0b2")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(12);

  // Merge cells for "Time" header (columns H-I)
  sheet.getRange(4, 8, 1, 2).merge();
  sheet.getRange(4, 8).setValue("Time");

  // Merge cells for "Duration" header (columns J-K)
  sheet.getRange(4, 10, 1, 2).merge();
  sheet.getRange(4, 10).setValue("Duration");

  // Row 5: Sub-headers for Time and Duration
  const subHeaders = [
    "",
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
    .setBackground("#ffe0b2")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(12);

  // Merge non-time/duration columns vertically (rows 4-5)
  [2, 3, 4, 5, 6, 7, 12].forEach(col => {
    sheet.getRange(4, col, 2, 1).merge();
  });

  // Align header row heights
  sheet.setRowHeight(4, 20);
  sheet.setRowHeight(5, 20);

  // Data rows
  let currentRow = 6;

  dates.forEach(function (dateInfo) {
    const dateStr = dateInfo.date;
    const isWeekend = dateInfo.isWeekend;
    const dateObj = dateInfo.dateObj;
    
    // Get Day Name (Mon, Tue, Wed...)
    const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    const dayName = days[dateObj.getDay()];

    // Merge date cells vertically (2 rows) and center align
    const dateRange = sheet.getRange(currentRow, 2, 2, 1);
    dateRange.merge();
    dateRange
      .setValue(dateStr)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Merge day cells vertically (2 rows) and center align
    const dayRange = sheet.getRange(currentRow, 3, 2, 1);
    dayRange.merge();
    dayRange
      .setValue(dayName)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Add task duration formulas for both rows (Col J - index 10)
    for (let offset = 0; offset < 2; offset++) {
      const row = currentRow + offset;
      const taskFormula =
        "=IF(AND(H" +
        row +
        "<TIME(12,30,0),I" +
        row +
        ">TIME(13,0,0)),I" +
        row +
        "-H" +
        row +
        "-TIME(0,30,0),I" +
        row +
        "-H" +
        row +
        ")";
      sheet.getRange(row, 10).setFormula(taskFormula);
      
      // Default Start & End to 00:00
      sheet.getRange(row, 8).setValue("00:00");
      sheet.getRange(row, 9).setValue("00:00");
    }

    // Add total formula in first row only (sums both task rows) and merge vertically (Col K - index 11)
    const totalFormula = "=SUM(J" + currentRow + ":J" + (currentRow + 1) + ")";
    const totalRange = sheet.getRange(currentRow, 11, 2, 1);
    totalRange.merge();
    totalRange
      .setFormula(totalFormula)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontWeight("bold")
      .setFontSize(11);

    // Format Start and End as clock time
    sheet.getRange(currentRow, 8, 2, 2).setNumberFormat("hh:mm");

    // Format Task and Total as duration hh:mm
    sheet.getRange(currentRow, 10, 2, 2).setNumberFormat("hh:mm");

    // Color weekend rows or holidays
    if (isWeekend || isHoliday_(dateObj)) {
      sheet.getRange(currentRow, 1, 2, 12).setBackground(WEEKEND_BG);
    }

    // Add data validation for Status (column G)
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(STATUS_OPTIONS, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(currentRow, 7, 2, 1).setDataValidation(statusRule);

    currentRow += 2;
  });

  const lastDataRow = currentRow - 1;

  // Add "Monthly hours spend" row
  const monthlyRow = currentRow + 1;
  sheet.getRange(monthlyRow, 2, 1, 9).merge();
  sheet
    .getRange(monthlyRow, 2)
    .setValue("Monthly hours spend")
    .setFontWeight("bold")
    .setHorizontalAlignment("right")
    .setVerticalAlignment("middle");

  // Monthly total formula in Column K (Col 11)
  const monthlyFormula =
    "=TEXT(INT(SUM(K6:K" +
    lastDataRow +
    "))*24+HOUR(SUM(K6:K" +
    lastDataRow +
    ')),"00")&":"&' +
    "TEXT(MINUTE(SUM(K6:K" +
    lastDataRow +
    ')),"00")';
  sheet
    .getRange(monthlyRow, 11)
    .setFormula(monthlyFormula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Add borders to all cells
  const allDataRange = sheet.getRange(4, 2, monthlyRow - 3, 11);
  allDataRange.setBorder(true, true, true, true, true, true);

  // Set font size 11 for all data rows (from row 6 to last data row)
  sheet
    .getRange(6, 1, sheet.getMaxRows() - 5, sheet.getMaxColumns())
    .setFontSize(11);
    
  // Center-align text in Start/End/Task/Total columns
  sheet
    .getRange(6, 8, sheet.getMaxRows() - 5, 4)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  Logger.log("✓ Alternative " + empName + " completed (" + dates.length + " dates)");
}
