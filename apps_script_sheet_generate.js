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

const HOLIDAY_DATES = [
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

const HEADER_BG = "#FEF2CB";
const WEEKEND_BG = "#FF9999";

const CONFIG_FOLDER_NAME = "Timesheet_Configs";

const MASTER_CONFIG_SHEET_ID =
  PropertiesService.getScriptProperties().getProperty("MASTER_CONFIG_SHEET_ID");

const DESTINATION_FOLDER_ID =
  PropertiesService.getScriptProperties().getProperty("DESTINATION_FOLDER_ID");

const WEBHOOK_URL =
  PropertiesService.getScriptProperties().getProperty("WEBHOOK_URL");

const TEST_ENV = (
  (PropertiesService.getScriptProperties().getProperty("TEST") || "")
    .toString()
    .toLowerCase() === "true"
);

/************************************************************
 * FETCH CONFIGURATION FROM MASTER SPREADSHEET
 ************************************************************/

function fetchProjectsConfig() {
  if (!MASTER_CONFIG_SHEET_ID) {
    throw new Error("MASTER_CONFIG_SHEET_ID is not configured in Script Properties");
  }

  const ss = SpreadsheetApp.openById(MASTER_CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName("Team");
  if (!sheet) {
    throw new Error("Sheet 'Team' not found in Master Config Spreadsheet");
  }

  const values = sheet.getDataRange().getValues();
  const projects = {};
  const emails = {};
  const employeeEmails = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.length < 3) continue;

    const projectName = (row[0] || "").toString().trim();
    const memberName = (row[1] || "").toString().trim();
    const email = (row[2] || "").toString().trim();

    if (!projectName || !memberName) continue;

    // Check active status (Column E / index 4) - default to true if empty/not provided
    let isActive = true;
    if (row.length > 4 && row[4] !== undefined && row[4] !== null) {
      const activeVal = row[4];
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

function grantProjectAccess(file, emails) {
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

function getExistingSpreadsheetId(projectName, monthName, year) {
  try {
    const configFolder = getOrCreateConfigFolder();
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

  const dates = getDatesForMonth(monthName, year);
  const config = fetchProjectsConfig();
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
      existingId = getExistingSpreadsheetId(projectName, monthName, year);
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
        const fakeUrl = "https://docs.google.com/spreadsheets/d/" + fakeId;
        updateConfigJson(projectName, fakeId, monthName, year, true);
        sendChatNotification(projectName, monthName, year, fakeUrl, true);
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
        setupEmployeeSheet(ss, emp, dates, projectName);
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

    grantProjectAccess(file, emailsToShare);
    updateConfigJson(projectName, ss.getId(), monthName, year);
    
    if (isNew) {
      sendChatNotification(projectName, monthName, year, url);
    } else {
      Logger.log("✓ Existing spreadsheet '" + workbookName + "' updated and shared.");
    }
  }

  Logger.log("✓ All project sheets processed successfully");
}

// ======================================================
// CONFIG FOLDER HANDLING
// ======================================================

function getOrCreateConfigFolder() {

  const parentFolder = DESTINATION_FOLDER_ID
    ? DriveApp.getFolderById(DESTINATION_FOLDER_ID)
    : DriveApp.getRootFolder();

  const folders = parentFolder.getFoldersByName(CONFIG_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(CONFIG_FOLDER_NAME);
}

function updateConfigJson(projectName, spreadsheetId, monthName, year, testMode = false) {

  const configFolder = getOrCreateConfigFolder();
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

function getDatesForMonth(monthName, year) {
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

function isHoliday(dateObj) {

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
function setupEmployeeSheet(ss, empName, dates,projectName) {
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
    if (isWeekend || isHoliday(dateObj)) {
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
// GOOGLE CHAT NOTIFICATION
// ======================================================

function sendChatNotification(projectName, monthName, year, url, testMode = false) {

  if (!WEBHOOK_URL) return;

  if (testMode) {
    Logger.log("TEST MODE → Chat notification for " + projectName);
    return;
  }

  const message = {
    text:
      "<users/all> " +
      projectName +
      " Timesheet for " +
      monthName +
      " " +
      year +
      " generated.\n" +
      url
  };

  UrlFetchApp.fetch(WEBHOOK_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(message)
  });
}

// ======================================================
// MASTER CONFIG SHEET GENERATION UTILITY
// ======================================================

/**
 * Creates a styled Master Config Spreadsheet, initializes it with column headers and sample data,
 * and automatically sets its Spreadsheet ID in the Script Properties.
 */
function createMasterConfigSheet() {
  const ssName = "Master Timesheet Config";
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty("MASTER_CONFIG_SHEET_ID");
  
  if (existingId) {
    try {
      const existingSs = SpreadsheetApp.openById(existingId);
      Logger.log("⚠ A Master Config Sheet already exists at: " + existingSs.getUrl());
      Logger.log("Spreadsheet ID: " + existingId);
      Logger.log("Skipping creation. If you want to force-create a new one, delete the 'MASTER_CONFIG_SHEET_ID' key from Script Properties first.");
      return existingSs.getUrl();
    } catch (e) {
      Logger.log("Could not open existing Master Config Sheet by ID, creating a new one...");
    }
  }
  
  // Create a new spreadsheet
  const ss = SpreadsheetApp.create(ssName);
  const sheet = ss.getSheets()[0];
  sheet.setName("Team");
  
  // Set column widths
  sheet.setColumnWidth(1, 150); // Project Name
  sheet.setColumnWidth(2, 200); // Member Name
  sheet.setColumnWidth(3, 250); // Email
  sheet.setColumnWidth(4, 150); // Chat ID
  sheet.setColumnWidth(5, 100); // Active (Checkbox)
  
  // Apply Calibri font to the whole sheet
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily("Calibri");
  
  // Set headers
  const headers = ["Project Name", "Member Name", "Email", "Chat ID", "Active"];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground("#FEF2CB") // Golden yellow header background
             .setFontWeight("bold")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle")
             .setFontSize(11);
  
  // Add sample rows with checkboxes in Column E
  const sampleData = [
    ["Project 1", "member 1", "email@address.com", "123456789", true],
    ["", "", "", "", false],
    ["", "", "", "", false],
    ["", "", "", "", false],
    ["", "", "", "", false]
  ];
  
  for (let i = 0; i < sampleData.length; i++) {
    const rowNum = i + 2;
    // Set text columns (A-D)
    sheet.getRange(rowNum, 1, 1, 4).setValues([[
      sampleData[i][0],
      sampleData[i][1],
      sampleData[i][2],
      sampleData[i][3]
    ]]);
    
    // Add checkbox for Column E (Active)
    const checkboxCell = sheet.getRange(rowNum, 5);
    checkboxCell.insertCheckboxes();
    checkboxCell.setValue(sampleData[i][4]);
  }
  
  // Apply borders to header and data rows
  const borderRange = sheet.getRange(1, 1, sampleData.length + 1, headers.length);
  borderRange.setBorder(true, true, true, true, true, true);
  
  // Set cell alignments
  sheet.getRange(2, 1, sampleData.length, 4).setHorizontalAlignment("left");
  sheet.getRange(2, 5, sampleData.length, 1).setHorizontalAlignment("center");
  
  // Move sheet to configured destination folder if DESTINATION_FOLDER_ID is set
  const file = DriveApp.getFileById(ss.getId());
  const folderId = props.getProperty("DESTINATION_FOLDER_ID");
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      file.moveTo(folder);
      Logger.log("✓ Moved Master Config Sheet to configured destination folder.");
    } catch (e) {
      Logger.log("Could not move Master Config Sheet to destination folder: " + e);
    }
  }
  
  // Save new ID to Script Properties automatically
  props.setProperty("MASTER_CONFIG_SHEET_ID", ss.getId());
  
  Logger.log("\n✓ Master Config Sheet created successfully!");
  Logger.log("Spreadsheet Name: " + ssName);
  Logger.log("Spreadsheet URL: " + ss.getUrl());
  Logger.log("Spreadsheet ID: " + ss.getId());
  Logger.log("✓ MASTER_CONFIG_SHEET_ID updated in Script Properties.");
  
  return ss.getUrl();
}
