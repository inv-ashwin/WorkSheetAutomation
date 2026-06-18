// ======================================================
// MASTER CONFIG SHEET GENERATION UTILITY
// ======================================================

/**
 * Creates a styled Master Config Spreadsheet, initializes it with column headers and sample data,
 * and automatically sets its Spreadsheet ID in the Script Properties.
 * 
 * Run this function once from the standalone setup script to bootstrap the system.
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
