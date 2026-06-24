/**
 * Timesheet Verification System
 *
 * This script verifies employee timesheet entries based on the configurations
 * defined in a centralized Master Config Spreadsheet and sends daily chat reports/reminders.
 *
 * ============================================================================
 * QUICK START
 * ============================================================================
 *
 * 1. Automatically Create Master Config Sheet (Recommended):
 *    - Select and run createMasterConfigSheet() in Apps Script to automatically
 *      generate a pre-styled Master Config Sheet, set up template columns and checkboxes,
 *      and automatically update your MASTER_CONFIG_SHEET_ID in Script Properties.
 *
 * 2. Set script properties:
 *    - Go to Project Settings -> Script Properties in Apps Script.
 *    - Define the required properties below (such as Webhook URLs).
 *
 * 3. Setup config fallback values (Optional):
 *    - Run setupConfiguration() to quickly set up default/initial property values.
 *
 * ============================================================================
 * CONFIGURATION - Script Properties
 * ============================================================================
 * - MASTER_CONFIG_SHEET_ID: (Required) ID of the Master Config spreadsheet containing projects and member active status
 * - DESTINATION_FOLDER_ID: (Required) Folder ID where timesheet config JSON files are stored
 * - GOOGLE_CHAT_WEBHOOK_URL: (Required) Webhook URL for manager daily summary reports
 * - EMPLOYEE_ALERT_WEBHOOK_URL: (Required) Webhook URL for employee reminders
 * - TEST_MODE: "true" or "false" (default: false)
 * - HOLIDAYS: JSON array of holiday dates in YYYY-MM-DD format
 * - ROWS_TO_CHECK_AFTER_DATE: Number of rows to check after finding a date (default: 5)
 * - CONFIG_FILE_ID: (Optional/Legacy) ID of timesheet_config.json file in Google Drive
 * - MAIN_SPREADSHEET_ID: (Fallback) The main spreadsheet ID - used if config file or master config not found
 * - ENGINEER_NAMES: (Fallback) Comma-separated list of engineer names
 * - EMPLOYEE_CHAT_IDS: (Fallback) JSON object mapping employee names to chat IDs
 *
 * Note: The script automatically fetches the spreadsheet ID from the project's timesheet config
 * JSON based on the current year-month (format: "2026-June"). If no project-specific config is
 * found, it falls back to the legacy/global config file or MAIN_SPREADSHEET_ID.
 * ============================================================================
 */

/**
 * Main class for Sheets Verification
 */
class SheetsVerifier {
  constructor() {
    this.props = PropertiesService.getScriptProperties();
    this.config = this._loadConfig();
  }

  /**
   * Get current month name (helper for config lookup)
   */
  _getCurrentMonthName() {
    const now = new Date();
    const monthNames = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December",
    ];
    return monthNames[now.getMonth()];
  }

  /**
   * Get month name from a Date object
   */
  _getMonthNameFromDate(date) {
    const monthNames = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December",
    ];
    return monthNames[date.getMonth()];
  }
  _getCurrentMonthSpreadsheetId(projectKey = null) {
    try {
      // Delegate to the existing function which accepts an optional date and project key
      return this._getSpreadsheetIdForDate(new Date(), projectKey);
    } catch (e) {
      Logger.log("Error in _getCurrentMonthSpreadsheetId: " + e);
      return null;
    }
  }
  /**
   * Fetch spreadsheet ID from timesheet config JSON based on a specific date's year-month.
   * When projectKey is set, uses {projectKey}_timesheet_config.json in the destination folder.
   * Otherwise uses timesheet_config.json (or CONFIG_FILE_ID) for backward compatibility.
   * Returns the sheet ID for the specified date's month or null if not found
   * @param {Date} targetDate - Optional date to get sheet ID for (defaults to current date)
   * @param {string} projectKey - Optional project key (e.g. "Rapid-Raise"); when set, looks for {projectKey}_timesheet_config.json
   */
  _getSpreadsheetIdForDate(targetDate = null, projectKey = null) {
    try {
      const props = PropertiesService.getScriptProperties();
      const destinationFolderId = props.getProperty("DESTINATION_FOLDER_ID");

      // Config file name: per-project {key}_timesheet_config.json or single timesheet_config.json
      const configFileName = projectKey
        ? projectKey + "_timesheet_config.json"
        : "timesheet_config.json";

      let configFile = null;

      // Try to get config file by ID first (only when not using per-project config)
      if (!projectKey) {
        const configFileId = props.getProperty("CONFIG_FILE_ID");
        if (configFileId) {
          try {
            configFile = DriveApp.getFileById(configFileId);
          } catch (e) {
            Logger.log("Could not find config file by ID, trying by name");
          }
        }
      }

      // Find in Timesheet_Configs subfolder first
      if (!configFile) {
        try {
          const parentFolder = destinationFolderId
            ? DriveApp.getFolderById(destinationFolderId)
            : DriveApp.getRootFolder();
          const subfolders = parentFolder.getFoldersByName("Timesheet_Configs");
          if (subfolders.hasNext()) {
            const configFolder = subfolders.next();
            const files = configFolder.getFilesByName(configFileName);
            if (files.hasNext()) {
              configFile = files.next();
            }
          }
        } catch (e) {
          Logger.log("Could not search in Timesheet_Configs folder: " + e);
        }
      }

      // Find by name in destination folder (fallback)
      if (!configFile && destinationFolderId) {
        try {
          const folder = DriveApp.getFolderById(destinationFolderId);
          const files = folder.getFilesByName(configFileName);
          if (files.hasNext()) {
            configFile = files.next();
          }
        } catch (e) {
          Logger.log("Could not find config file in folder: " + configFileName);
        }
      }

      // If still not found and single config, try root folder (fallback)
      if (!configFile && !projectKey) {
        try {
          const files = DriveApp.getRootFolder().getFilesByName(
            "timesheet_config.json",
          );
          if (files.hasNext()) {
            configFile = files.next();
          }
        } catch (e) {
          Logger.log("Could not find config file in root folder");
        }
      }

      if (!configFile) {
        Logger.log(configFileName + " not found");
        return null;
      }

      // Read and parse config file safely
      const configContent = configFile.getBlob().getDataAsString();
      let config = {};
      try {
        config = JSON.parse(configContent || "{}");
      } catch (e) {
        Logger.log("Warning: Failed to parse config JSON content: " + e + ". Content: '" + configContent + "'");
      }

      // Use targetDate if provided, otherwise use current date
      const dateToUse = targetDate || new Date();
      const targetYear = dateToUse.getFullYear();
      const targetMonth = this._getMonthNameFromDate(dateToUse);
      const configKey = targetYear + "-" + targetMonth;

      // Get spreadsheet ID for target month
      const spreadsheetId = config[configKey];

      if (spreadsheetId) {
        Logger.log(
          "✓ Found spreadsheet ID for " + configKey + ": " + spreadsheetId,
        );
        return spreadsheetId;
      } else {
        Logger.log("No spreadsheet ID found for " + configKey);
        const availableKeys = Object.keys(config).filter(
          (k) => k !== "updated_at" && k !== "latest",
        );
        if (availableKeys.length > 0) {
          Logger.log("Available keys: " + availableKeys.join(", "));
        }
        return null;
      }
    } catch (e) {
      Logger.log("Error fetching spreadsheet ID from config: " + e);
      return null;
    }
  }

  /**
   * Fetch project configuration and engineer chat IDs from Master Config Spreadsheet
   */
  _fetchProjectsFromMasterSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log("WARNING: Active spreadsheet is not accessible.");
      return { projects: {}, employeeChatIds: {}, engineerNames: [] };
    }

    try {
      const sheet = ss.getSheetByName("Manager Sheet");
      if (!sheet) {
        Logger.log("WARNING: Sheet 'Manager Sheet' not found in Master Config Spreadsheet");
        return { projects: {}, employeeChatIds: {}, engineerNames: [] };
      }

      const values = sheet.getDataRange().getValues();
      const projects = {};
      const employeeChatIds = {};
      const seenEngineers = {};
      const engineerNames = [];

      // Dynamically find table column headers
      let projIdx = 2;   // Default to Column C
      let memberIdx = 3; // Default to Column D
      let chatIdx = 5;   // Default to Column F
      let activeIdx = 6; // Default to Column G

      if (values.length > 10) {
        const headerRow = values[10]; // Row 11
        for (let c = 0; c < headerRow.length; c++) {
          const cellVal = String(headerRow[c] || "").trim().toLowerCase();
          if (cellVal.includes("project name")) projIdx = c;
          if (cellVal.includes("team member")) memberIdx = c;
          if (cellVal.includes("chat id")) chatIdx = c;
          if (cellVal.includes("active")) activeIdx = c;
        }
      }

      for (let i = 11; i < values.length; i++) {
        const row = values[i];
        if (row.length <= Math.max(projIdx, memberIdx)) continue;

        const projectName = (row[projIdx] || "").toString().trim();
        const memberName = (row[memberIdx] || "").toString().trim();
        const chatId = row.length > chatIdx ? (row[chatIdx] || "").toString().trim() : "";

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
          projects[projectName] = {};
        }

        // Map memberName -> chatId (keep it empty/falsy if not provided)
        projects[projectName][memberName] = chatId || "";

        if (chatId) {
          employeeChatIds[memberName] = chatId;
        }

        if (!seenEngineers[memberName]) {
          seenEngineers[memberName] = true;
          engineerNames.push(memberName);
        }
      }

      return {
        projects: projects,
        employeeChatIds: employeeChatIds,
        engineerNames: engineerNames
      };
    } catch (e) {
      Logger.log("Error loading config from master sheet: " + e);
      return { projects: {}, employeeChatIds: {}, engineerNames: [] };
    }
  }

  /**
   * Load configuration from Script Properties and Master Sheet
   */
  _loadConfig() {
    const props = PropertiesService.getScriptProperties();

    // Fetch dynamic project config from Master Config Sheet
    const masterConfig = this._fetchProjectsFromMasterSheet();

    // Parse engineer names
    let engineerNames = masterConfig.engineerNames;
    if (engineerNames.length === 0) {
      const engineerNamesStr = props.getProperty("ENGINEER_NAMES");
      if (engineerNamesStr) {
        engineerNames = engineerNamesStr.split(",").map((n) => n.trim());
      }
    }

    // Parse employee chat IDs
    let employeeChatIds = masterConfig.employeeChatIds;
    if (Object.keys(employeeChatIds).length === 0) {
      try {
        const chatIdsStr = props.getProperty("EMPLOYEE_CHAT_IDS");
        if (chatIdsStr) {
          employeeChatIds = JSON.parse(chatIdsStr);
        }
      } catch (e) {
        Logger.log("Error parsing EMPLOYEE_CHAT_IDS: " + e);
      }
    }

    // Parse PROJECTS
    let projects = masterConfig.projects;
    if (Object.keys(projects).length === 0) {
      try {
        const projectsStr = props.getProperty("PROJECTS");
        if (projectsStr) {
          projects = JSON.parse(projectsStr);
          // Merge per-project chat IDs into employeeChatIds
          const projectKeys = Object.keys(projects);
          for (let p = 0; p < projectKeys.length; p++) {
            const pk = projectKeys[p];
            const engineers = projects[pk];
            if (engineers && typeof engineers === "object") {
              const names = Object.keys(engineers);
              for (let n = 0; n < names.length; n++) {
                employeeChatIds[names[n]] = engineers[names[n]];
              }
            }
          }
        }
      } catch (e) {
        Logger.log("Error parsing PROJECTS: " + e);
      }
    }

    // When PROJECTS is set, engineer list for Backlog etc. is union of all project engineers
    if (Object.keys(projects).length > 0) {
      const seen = {};
      const allEngineers = [];
      for (let p = 0; p < Object.keys(projects).length; p++) {
        const engs = projects[Object.keys(projects)[p]];
        if (engs && typeof engs === "object") {
          const names = Object.keys(engs);
          for (let n = 0; n < names.length; n++) {
            if (!seen[names[n]]) {
              seen[names[n]] = true;
              allEngineers.push(names[n]);
            }
          }
        }
      }
      if (allEngineers.length > 0) {
        engineerNames = allEngineers;
      }
    }

    // Parse holidays
    let holidays = [];
    try {
      const holidaysStr = props.getProperty("HOLIDAYS");
      if (holidaysStr && JSON.parse(holidaysStr).length > 0) {
        holidays = JSON.parse(holidaysStr);
      } else {
        // Default holidays for 2025
        holidays = [
          "2026-01-01",
          "2026-01-26",
          "2026-03-20",
          "2026-04-03",
          "2026-04-15",
          "2026-05-01",
          "2026-08-15",
          "2026-08-25",
          "2026-08-26",
          "2026-09-04",
          "2026-10-02",
          "2026-10-20",
          "2026-12-25",
        ];
      }
    } catch (e) {
      Logger.log("Error parsing HOLIDAYS: " + e);
    }

    // Get spreadsheet ID for current month from config file
    // When PROJECTS is set, also try first project's spreadsheet as fallback for iteration
    let mainSpreadsheetId = this._getCurrentMonthSpreadsheetId(null);
    if (!mainSpreadsheetId && Object.keys(projects).length > 0) {
      const firstProjectKey = Object.keys(projects)[0];
      mainSpreadsheetId = this._getCurrentMonthSpreadsheetId(firstProjectKey);
      if (mainSpreadsheetId) {
        Logger.log(
          "Using first project (" +
            firstProjectKey +
            ") spreadsheet as fallback",
        );
      }
    }
    if (!mainSpreadsheetId) {
      mainSpreadsheetId = props.getProperty("MAIN_SPREADSHEET_ID");
      if (mainSpreadsheetId) {
        Logger.log(
          "Using MAIN_SPREADSHEET_ID from Script Properties as fallback",
        );
      }
    }

    return {
      mainSpreadsheetId: mainSpreadsheetId,
      projectKey: null,
      projects: projects,
      engineerNames: engineerNames,
      googleChatWebhookUrl: props.getProperty("GOOGLE_CHAT_WEBHOOK_URL"),
      employeeAlertWebhookUrl: props.getProperty("EMPLOYEE_ALERT_WEBHOOK_URL"),
      testMode: props.getProperty("TEST_MODE") === "true",
      employeeChatIds: employeeChatIds,
      holidays: holidays,
      rowsToCheckAfterDate: parseInt(
        props.getProperty("ROWS_TO_CHECK_AFTER_DATE") || "5",
        10,
      ),
      sheetDataRange: props.getProperty("SHEET_DATA_RANGE") || "A1:K150",
    };
  }

  /**
   * Get current month sheet name (e.g., 'July', 'August')
   */
  getCurrentMonthSheetName() {
    const now = new Date();
    const monthNames = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December",
    ];
    return monthNames[now.getMonth()];
  }

  /**
   * Check if a given date is a holiday
   */
  isHoliday(date) {
    const dateStr = Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    return this.config.holidays.indexOf(dateStr) !== -1;
  }

  /**
   * Get the last working day (excluding weekends and holidays)
   * Returns an object with both the formatted date string and the Date object
   */
  getLastWorkingDay() {
    const today = new Date();
    const currentDay = today.getDay(); // 0=Sunday, 6=Saturday

    // Skip if today is weekend
    if (currentDay === 0 || currentDay === 6) {
      return null;
    }

    // Find the last working day
    let checkDate = new Date(today);
    checkDate.setDate(checkDate.getDate() - 1);

    // Keep going back until we find a non-weekend, non-holiday day
    while (
      checkDate.getDay() === 0 ||
      checkDate.getDay() === 6 ||
      this.isHoliday(checkDate)
    ) {
      checkDate.setDate(checkDate.getDate() - 1);
    }

    return {
      dateString: Utilities.formatDate(
        checkDate,
        Session.getScriptTimeZone(),
        "d-MMM",
      ),
      dateObject: checkDate,
    };
  }

  /**
   * Get current day
   */
  getCurrentDay() {
    const today = new Date();
    const currentDay = today.getDay();

    // Skip if today is weekend or holiday
    if (currentDay === 0 || currentDay === 6 || this.isHoliday(today)) {
      return null;
    }

    return Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM");
  }

  /**
   * Get sheet data from Google Sheets
   */
  getSheetData(spreadsheetId, rangeName) {
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      // Remove quotes if present (for compatibility), but we don't add them anymore
      const sheetName = rangeName.split("!")[0].replace(/^'|'$/g, "");
      const range = rangeName.split("!")[1];

      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log("Sheet not found: " + sheetName);
        return [];
      }

      const dataRange = sheet.getRange(range);
      const values = dataRange.getValues();

      return values;
    } catch (e) {
      Logger.log("Error fetching sheet data: " + e);
      return [];
    }
  }

  /**
   * Get engineer sheet data
   * @param {string} engineerName - Name of the engineer (sheet name)
   * @param {string} monthName - Month name (e.g., "December") - for logging purposes
   * @param {string} spreadsheetId - Optional spreadsheet ID (defaults to config.mainSpreadsheetId)
   */
  getEngineerSheetData(engineerName, monthName, spreadsheetId = null) {
    const sheetRange = engineerName + "!" + this.config.sheetDataRange;
    const targetSpreadsheetId = spreadsheetId || this.config.mainSpreadsheetId;

    try {
      return this.getSheetData(targetSpreadsheetId, sheetRange);
    } catch (e) {
      Logger.log("Error fetching data for engineer " + engineerName + ": " + e);
      return [];
    }
  }

  /**
   * Helper function to pad numbers with leading zeros
   */
  _padZero(num, length) {
    const str = String(num);
    if (str.length >= length) {
      return str;
    }
    // Use Array.join for compatibility
    const zeros = Array(length - str.length + 1).join("0");
    return zeros + str;
  }

  /**
   * Convert time value (Date object or string) to time string format
   * Handles: Date objects (like "Sat Dec 30 1899 12:30:00"), time strings ("12:30"), etc.
   * Returns: "HH:mm" format (e.g., "12:30") or "HH:mm:ss" if seconds present, or empty string if cannot parse
   */
  _normalizeTime(timeValue) {
    if (!timeValue) {
      return "";
    }

    // If it's a Date object (Google Sheets time values come as Date objects)
    if (timeValue instanceof Date) {
      const hours = timeValue.getHours();
      const minutes = timeValue.getMinutes();
      const seconds = timeValue.getSeconds();

      // Format as HH:mm (or HH:mm:ss if seconds are present)
      if (seconds > 0) {
        return (
          this._padZero(hours, 2) +
          ":" +
          this._padZero(minutes, 2) +
          ":" +
          this._padZero(seconds, 2)
        );
      } else {
        return this._padZero(hours, 2) + ":" + this._padZero(minutes, 2);
      }
    }

    const timeStr = String(timeValue).trim();
    if (!timeStr) {
      return "";
    }

    // If it's already in time format (HH:mm or HH:mm:ss)
    const timeRegex = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/;
    const match = timeStr.match(timeRegex);
    if (match) {
      const hours = parseInt(match[1], 10);
      const minutes = parseInt(match[2], 10);
      const seconds = match[3] ? parseInt(match[3], 10) : 0;

      if (seconds > 0) {
        return (
          this._padZero(hours, 2) +
          ":" +
          this._padZero(minutes, 2) +
          ":" +
          this._padZero(seconds, 2)
        );
      } else {
        return this._padZero(hours, 2) + ":" + this._padZero(minutes, 2);
      }
    }

    // Try to parse as Date string (e.g., "Sat Dec 30 1899 12:30:00 GMT+0521")
    try {
      const parsedDate = new Date(timeStr);
      if (!isNaN(parsedDate.getTime())) {
        // Check if it looks like a time-only date (Dec 30, 1899 is the epoch for time values)
        // Also check for Jan 1, 1900 which is another common epoch
        if (
          parsedDate.getFullYear() === 1899 ||
          parsedDate.getFullYear() === 1900 ||
          timeStr.match(/Dec 30 1899/i) ||
          timeStr.match(/Jan 01 1900/i) ||
          timeStr.match(/Jan 1 1900/i)
        ) {
          const hours = parsedDate.getHours();
          const minutes = parsedDate.getMinutes();
          const seconds = parsedDate.getSeconds();

          if (seconds > 0) {
            return (
              this._padZero(hours, 2) +
              ":" +
              this._padZero(minutes, 2) +
              ":" +
              this._padZero(seconds, 2)
            );
          } else {
            return this._padZero(hours, 2) + ":" + this._padZero(minutes, 2);
          }
        }
      }
    } catch (e) {
      // Not a parseable date, continue
    }

    // If we can't parse it, return the original string (might already be formatted)
    return timeStr;
  }

  /**
   * Normalize date from various formats to a standard format for comparison
   * Handles: Date objects, "d-MMM" format, "dd-MMM" format, full date strings, etc.
   * Returns: "d-MMM" format (e.g., "15-Dec") or null if cannot parse
   */
  _normalizeDate(dateValue) {
    if (!dateValue) {
      return null;
    }

    // If it's already a Date object
    if (dateValue instanceof Date) {
      return Utilities.formatDate(
        dateValue,
        Session.getScriptTimeZone(),
        "d-MMM",
      );
    }

    const dateStr = String(dateValue).trim();
    if (!dateStr) {
      return null;
    }

    // Try to parse as Date if it looks like a full date string
    // e.g., "Mon Dec 15 2025 00:00:00 GMT+0530"
    try {
      const parsedDate = new Date(dateStr);
      if (!isNaN(parsedDate.getTime())) {
        // Check if it's a valid date (not just a string that can't be parsed)
        // If the original string contains day names and full month names, it's likely a Date string
        if (
          dateStr.match(/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)/i) ||
          (dateStr.match(/\d{4}/) && dateStr.length > 10)
        ) {
          return Utilities.formatDate(
            parsedDate,
            Session.getScriptTimeZone(),
            "d-MMM",
          );
        }
      }
    } catch (e) {
      // Not a Date object, continue with string parsing
    }

    // Check if it's already in "d-MMM" or "dd-MMM" format (e.g., "15-Dec" or "5-Dec")
    const dMMMRegex = /^(\d{1,2})-([A-Za-z]{3})$/i;
    const match = dateStr.match(dMMMRegex);
    if (match) {
      const day = parseInt(match[1], 10);
      const month = match[2];
      // Return in consistent format (remove leading zero if present)
      return (
        day + "-" + month.charAt(0).toUpperCase() + month.slice(1).toLowerCase()
      );
    }

    // Try other common formats
    // Format: "DD/MM/YYYY" or "MM/DD/YYYY"
    const slashRegex = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
    const slashMatch = dateStr.match(slashRegex);
    if (slashMatch) {
      try {
        // Try DD/MM/YYYY first (more common internationally)
        const day = parseInt(slashMatch[1], 10);
        const month = parseInt(slashMatch[2], 10) - 1; // Month is 0-indexed
        const year = parseInt(slashMatch[3], 10);
        const date = new Date(year, month, day);
        if (date.getDate() === day && date.getMonth() === month) {
          return Utilities.formatDate(
            date,
            Session.getScriptTimeZone(),
            "d-MMM",
          );
        }
      } catch (e) {
        // Try MM/DD/YYYY
        try {
          const month = parseInt(slashMatch[1], 10) - 1;
          const day = parseInt(slashMatch[2], 10);
          const year = parseInt(slashMatch[3], 10);
          const date = new Date(year, month, day);
          if (date.getDate() === day && date.getMonth() === month) {
            return Utilities.formatDate(
              date,
              Session.getScriptTimeZone(),
              "d-MMM",
            );
          }
        } catch (e2) {
          // Ignore
        }
      }
    }

    // Format: "YYYY-MM-DD"
    const isoRegex = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
    const isoMatch = dateStr.match(isoRegex);
    if (isoMatch) {
      try {
        const year = parseInt(isoMatch[1], 10);
        const month = parseInt(isoMatch[2], 10) - 1;
        const day = parseInt(isoMatch[3], 10);
        const date = new Date(year, month, day);
        if (date.getDate() === day && date.getMonth() === month) {
          return Utilities.formatDate(
            date,
            Session.getScriptTimeZone(),
            "d-MMM",
          );
        }
      } catch (e) {
        // Ignore
      }
    }

    // If we can't parse it, return null
    Logger.log("Warning: Could not parse date format: " + dateStr);
    return null;
  }

  /**
   * Parse last working day entries from sheet data
   */
  parseLastWorkingDayEntries(sheetData, targetDate) {
    if (!sheetData || sheetData.length < 2) {
      return [];
    }

    // Normalize target date to "d-MMM" format
    const normalizedTargetDate = this._normalizeDate(targetDate);
    if (!normalizedTargetDate) {
      Logger.log("Error: Could not normalize target date: " + targetDate);
      return [];
    }

    Logger.log("Looking for normalized target date: " + normalizedTargetDate);

    // Find header row (look for 'Date' in column B, index 1)
    let headerRow = null;
    for (let i = 0; i < sheetData.length; i++) {
      const row = sheetData[i];
      if (row.length > 1 && String(row[1]).indexOf("Date") !== -1) {
        headerRow = i;
        break;
      }
    }

    if (headerRow === null) {
      Logger.log("Could not find header row with 'Date' in column B");
      return [];
    }

    // Extract entries for the target date and next configurable rows
    const targetEntries = [];
    let i = headerRow + 1;

    while (i < sheetData.length) {
      const row = sheetData[i];
      if (row.length === 0) {
        i++;
        continue;
      }

      // Check if this row has the target date in column B (index 1)
      const dateCellValue = row.length > 1 ? row[1] : null;
      const normalizedDateCell = this._normalizeDate(dateCellValue);

      // Compare normalized dates
      if (normalizedDateCell && normalizedDateCell === normalizedTargetDate) {
        Logger.log(
          "Found target date '" +
            normalizedTargetDate +
            "' (original: " +
            dateCellValue +
            ") at row " +
            (i + 1) +
            " (column B)",
        );

        // Parse this row and the next configurable rows for all tasks on this date
        const rowsToCheck = this.config.rowsToCheckAfterDate + 1; // +1 to include the current row

        // First, check if the first row has total_duration (merged cell case)
        const firstRow = sheetData[i];
        let mergedTotalDuration = "";
        if (firstRow.length > 9) {
          mergedTotalDuration = this._normalizeTime(firstRow[9]);
        }

        for (let j = i; j < Math.min(i + rowsToCheck, sheetData.length); j++) {
          const taskRow = sheetData[j];
          if (taskRow.length === 0) {
            continue;
          }

          // Skip if this is another date row (unless it's the first one)
          if (j > i && taskRow.length > 1) {
            const nextDateCellValue = taskRow[1];
            const normalizedNextDate = this._normalizeDate(nextDateCellValue);
            // If we can normalize it and it's different from target date, it's a new date
            if (
              normalizedNextDate &&
              normalizedNextDate !== normalizedTargetDate
            ) {
              Logger.log(
                "  Found new date at row " +
                  (j + 1) +
                  " (" +
                  normalizedNextDate +
                  "), stopping task collection",
              );
              break;
            }
            // Also check if it looks like a date string (has numbers)
            const nextDateStr = String(nextDateCellValue).trim();
            if (
              nextDateStr.length > 0 &&
              /[0-9]/.test(nextDateStr) &&
              !normalizedNextDate
            ) {
              // This might be a new date we couldn't parse, but has numbers - be cautious
              Logger.log(
                "  Found potential new date at row " +
                  (j + 1) +
                  ", stopping task collection",
              );
              break;
            }
          }

          // Parse the task data according to correct column structure
          const entry = {
            row_number: j + 1,
            date: normalizedTargetDate,
            module_area: taskRow.length > 2 ? String(taskRow[2]) : "",
            task_details: taskRow.length > 3 ? String(taskRow[3]) : "",
            status: taskRow.length > 4 ? String(taskRow[4]) : "",
            activity_type: taskRow.length > 5 ? String(taskRow[5]) : "",
            start_time:
              taskRow.length > 6 ? this._normalizeTime(taskRow[6]) : "",
            end_time: taskRow.length > 7 ? this._normalizeTime(taskRow[7]) : "",
            total_duration:
              taskRow.length > 9 ? this._normalizeTime(taskRow[9]) : "",
            remarks: taskRow.length > 10 ? String(taskRow[10]) : "",
          };

          // If this entry doesn't have total_duration but we found it in merged cell, use it
          if (!entry.total_duration.trim() && mergedTotalDuration) {
            entry.total_duration = mergedTotalDuration;
          }

          // Only add entries that have meaningful task details
          if (entry.task_details.trim() || entry.module_area.trim()) {
            targetEntries.push(entry);
            Logger.log(
              "  Added task from row " +
                (j + 1) +
                ": " +
                entry.task_details.substring(0, 50) +
                "...",
            );
          } else if (
            entry.start_time.trim() &&
            entry.end_time.trim() &&
            targetEntries.length > 0
          ) {
            // This indicates additional time for the previous task
            const previousEntry = targetEntries[targetEntries.length - 1];
            const additionalEntry = {
              row_number: entry.row_number,
              date: entry.date,
              module_area: previousEntry.module_area,
              task_details: previousEntry.task_details,
              status: previousEntry.status,
              activity_type: previousEntry.activity_type,
              start_time: entry.start_time,
              end_time: entry.end_time,
              total_duration: entry.total_duration,
              remarks: entry.remarks,
            };
            targetEntries.push(additionalEntry);
            Logger.log(
              "  Added additional time entry from row " +
                (j + 1) +
                " for previous task: " +
                previousEntry.task_details.substring(0, 50) +
                "...",
            );
          }
        }

        // Return immediately once entries are found and processed
        Logger.log(
          "Total entries found for " +
            normalizedTargetDate +
            ": " +
            targetEntries.length,
        );
        return targetEntries;
      } else {
        i++;
      }
    }

    Logger.log(
      "Total entries found for " +
        normalizedTargetDate +
        ": " +
        targetEntries.length,
    );
    return targetEntries;
  }

  /**
   * Generate formatted analysis message from timesheet entries
   */
  generateAnalysisMessage(employeeName, lastDayEntries) {
    if (!lastDayEntries || lastDayEntries.length === 0) {
      return "👤 " + employeeName + " (Total: 0h 0m)\n└ No entries found";
    }

    // Get total duration from the first entry (since it's a merged cell)
    const totalDuration = lastDayEntries[0].total_duration || "0:00";

    // Convert duration to hours format
    const durationToHours = function (durationStr) {
      if (!durationStr || durationStr.trim() === "") {
        return 0.0;
      }
      try {
        if (durationStr.indexOf(":") !== -1) {
          const parts = durationStr.split(":");
          return parseFloat(parts[0]) + parseFloat(parts[1]) / 60;
        } else {
          return parseFloat(durationStr);
        }
      } catch (e) {
        return 0.0;
      }
    };

    // Calculate task duration
    const calculateTaskDuration = function (start, end) {
      if (!start || !end) {
        return 0.0;
      }
      try {
        const startParts = start.split(":");
        const endParts = end.split(":");
        const startMinutes =
          parseInt(startParts[0]) * 60 + parseInt(startParts[1]);
        const endMinutes = parseInt(endParts[0]) * 60 + parseInt(endParts[1]);
        const durationMinutes = endMinutes - startMinutes;
        return durationMinutes / 60.0;
      } catch (e) {
        return 0.0;
      }
    };

    // Format duration
    const formatDuration = function (hoursDecimal) {
      if (hoursDecimal === 0) {
        return "0h 0m";
      }
      const hours = Math.floor(hoursDecimal);
      const minutes = Math.floor((hoursDecimal - hours) * 60);
      if (hours === 0) {
        return minutes + "m";
      } else if (minutes === 0) {
        return hours + "h";
      } else {
        return hours + "h " + minutes + "m";
      }
    };

    const totalHours = durationToHours(totalDuration);

    // Group entries by activity type
    const activityGroups = {};

    for (let i = 0; i < lastDayEntries.length; i++) {
      const entry = lastDayEntries[i];
      const activityType = entry.activity_type || "Task";

      if (!activityGroups[activityType]) {
        activityGroups[activityType] = [];
      }

      // Calculate individual task duration
      const startTime = entry.start_time || "";
      const endTime = entry.end_time || "";
      const taskDuration = calculateTaskDuration(startTime, endTime);

      // Get task details
      const taskDetails = entry.task_details || "No details";
      const moduleArea = entry.module_area || "";

      // Prefer module_area over task_details for description
      const description =
        moduleArea && moduleArea.trim() ? moduleArea : taskDetails;

      activityGroups[activityType].push({
        description: description,
        duration: taskDuration,
      });
    }

    // Format the header
    let message =
      "\n👤 " + employeeName + " (Total: " + formatDuration(totalHours) + ")\n";

    // Process grouped activities
    const activityTypes = Object.keys(activityGroups);

    for (let i = 0; i < activityTypes.length; i++) {
      const activityType = activityTypes[i];
      const tasks = activityGroups[activityType];

      // Combine descriptions and calculate total duration for this activity type
      const descriptions = [];
      let totalActivityDuration = 0.0;

      for (let j = 0; j < tasks.length; j++) {
        let desc = tasks[j].description;
        const duration = tasks[j].duration;

        // Truncate individual descriptions if too long
        if (desc.length > 40) {
          desc = desc.substring(0, 37) + "...";
        }

        descriptions.push(desc + " (" + formatDuration(duration) + ")");
        totalActivityDuration += duration;
      }

      // Combine all descriptions for this activity type
      const combinedDesc = descriptions.join(", ");

      // Use appropriate tree symbol
      const symbol = i === activityTypes.length - 1 ? "└" : "├";

      message += symbol + " " + activityType + " – " + combinedDesc + "\n";
    }

    // Add remarks summary if available
    const remarksList = [];
    for (let i = 0; i < lastDayEntries.length; i++) {
      const remarks = (lastDayEntries[i].remarks || "").trim();
      if (remarks) {
        remarksList.push(remarks);
      }
    }

    if (remarksList.length > 0) {
      // Combine all remarks
      let allRemarks = remarksList.join("; ");
      if (allRemarks.length > 100) {
        allRemarks = allRemarks.substring(0, 97) + "...";
      }
      message += "📝 Additional: " + allRemarks + "\n";
    }

    return message.trim();
  }

  /**
   * Verify all employee sheets and return analysis results.
   * When PROJECTS is configured, iterates over each project and uses that project's
   * {projectKey}_timesheet_config.json and engineers; otherwise uses PROJECT_KEY / ENGINEER_NAMES.
   */
  verifyAllEmployees() {
    const results = [];

    // Get last working day (returns object with dateString and dateObject)
    const lastWorkingDayInfo = this.getLastWorkingDay();
    if (!lastWorkingDayInfo) {
      Logger.log("No last working day found");
      return results;
    }

    const lastWorkingDay = lastWorkingDayInfo.dateString;
    const lastWorkingDayDate = lastWorkingDayInfo.dateObject;

    // Get the month of the last working day (not current month!)
    const lastWorkingDayMonth = this._getMonthNameFromDate(lastWorkingDayDate);
    const lastWorkingDayYear = lastWorkingDayDate.getFullYear();

    const projectKeys =
      this.config.projects && Object.keys(this.config.projects).length > 0
        ? Object.keys(this.config.projects)
        : null;

    if (projectKeys && projectKeys.length > 0) {
      // Iterate over each project: use project's config and engineers
      for (let p = 0; p < projectKeys.length; p++) {
        const projectKey = projectKeys[p];
        const projectEngineers = this.config.projects[projectKey];
        if (!projectEngineers || typeof projectEngineers !== "object") {
          continue;
        }
        const engineerNames = Object.keys(projectEngineers);

        const spreadsheetId = this._getSpreadsheetIdForDate(
          lastWorkingDayDate,
          projectKey,
        );
        const effectiveSpreadsheetId =
          spreadsheetId || this.config.mainSpreadsheetId;

        Logger.log(
          "Project " +
            projectKey +
            ": checking " +
            lastWorkingDay +
            ", spreadsheet: " +
            (effectiveSpreadsheetId || "none"),
        );

        for (let i = 0; i < engineerNames.length; i++) {
          const engineerName = engineerNames[i];
          Logger.log("Checking [" + projectKey + "] " + engineerName + "...");

          const sheetData = this.getEngineerSheetData(
            engineerName,
            lastWorkingDayMonth,
            effectiveSpreadsheetId,
          );
          const lastDayEntries = this.parseLastWorkingDayEntries(
            sheetData,
            lastWorkingDay,
          );

          Logger.log(
            "Found " +
              lastDayEntries.length +
              " entries for " +
              engineerName +
              " on " +
              lastWorkingDay,
          );

          if (lastDayEntries.length > 0) {
            const analysis = this.generateAnalysisMessage(
              engineerName,
              lastDayEntries,
            );
            results.push("[" + projectKey + "] " + analysis);
          } else {
            results.push([engineerName, false]);
          }
        }
      }
      return results;
    }

    // Single-project flow: PROJECT_KEY / ENGINEER_NAMES
    const spreadsheetId = this._getSpreadsheetIdForDate(
      lastWorkingDayDate,
      this.config.projectKey,
    );
    if (!spreadsheetId) {
      Logger.log(
        "Warning: Could not find spreadsheet ID for " +
          lastWorkingDayYear +
          "-" +
          lastWorkingDayMonth +
          ", using fallback",
      );
    }

    const effectiveSpreadsheetId =
      spreadsheetId || this.config.mainSpreadsheetId;

    Logger.log("Checking entries for working day: " + lastWorkingDay);
    Logger.log(
      "Last working day is in month: " +
        lastWorkingDayMonth +
        " " +
        lastWorkingDayYear,
    );
    Logger.log("Using spreadsheet ID: " + effectiveSpreadsheetId);

    for (let i = 0; i < this.config.engineerNames.length; i++) {
      const engineerName = this.config.engineerNames[i];
      Logger.log("Checking " + engineerName + "'s sheet...");

      const sheetData = this.getEngineerSheetData(
        engineerName,
        lastWorkingDayMonth,
        effectiveSpreadsheetId,
      );
      const lastDayEntries = this.parseLastWorkingDayEntries(
        sheetData,
        lastWorkingDay,
      );

      Logger.log(
        "Found " +
          lastDayEntries.length +
          " entries for " +
          engineerName +
          " on " +
          lastWorkingDay,
      );

      if (lastDayEntries.length > 0) {
        const analysis = this.generateAnalysisMessage(
          engineerName,
          lastDayEntries,
        );
        results.push(analysis);
      } else {
        results.push([engineerName, false]);
      }
    }

    return results;
  }

  /**
   * Send message to Google Chat space
   */
  sendGoogleChatMessage(message, webhookUrl) {
    if (this.config.testMode) {
      Logger.log("TEST MODE: Would send Google Chat message: " + message);
      return true;
    }

    const url = webhookUrl || this.config.googleChatWebhookUrl;
    if (!url) {
      Logger.log("Google Chat webhook URL not configured");
      return false;
    }

    const payload = { text: message };

    try {
      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      };

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        Logger.log("Message sent to Google Chat successfully");
        return true;
      } else {
        Logger.log("Error sending message to Google Chat: " + responseCode);
        return false;
      }
    } catch (e) {
      Logger.log("Error sending message to Google Chat: " + e);
      return false;
    }
  }

  /**
   * Send reminder message to employees who haven't submitted timesheets
   */
  sendEmployeeReminder(employeesToRemind) {
    if (!this.config.employeeAlertWebhookUrl) {
      return false;
    }

    // Create mentions for employees with chat IDs
    const employeeMentions = [];

    for (let i = 0; i < employeesToRemind.length; i++) {
      const employee = employeesToRemind[i];
      const chatId = this.config.employeeChatIds[employee];
      if (chatId && chatId !== "2541") {
        employeeMentions.push("<users/" + chatId + ">");
      } else {
        Logger.log(
          "Skipping chat reminder for " +
            employee +
            " because no Chat ID is configured.",
        );
      }
    }

    // Build message with proper mentions
    if (employeeMentions.length > 0) {
      let message = "Dear " + employeeMentions.join(", ") + ",\n\n";
      message += "\nPlease update your timesheet for the last working day. ✅";

      // Send message with mentions if available
      return this.sendGoogleChatMessage(
        message,
        this.config.employeeAlertWebhookUrl,
      );
    }

    return false;
  }

  runDailyVerification() {
    Logger.log("Starting daily task verification...");
    const today = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );

    // Check if we should run verification today
    const lastWorkingDayInfo = this.getLastWorkingDay();
    if (!lastWorkingDayInfo) {
      Logger.log(
        "No verification needed - today is weekend or no valid working day to check",
      );
      return [];
    }
    const lastWorkingDay = lastWorkingDayInfo.dateString;

    // Verify all employee sheets
    const results = this.verifyAllEmployees();

    // Find employees who need to be reminded (no data or poor performance)
    const employeesToRemind = [];
    let summaryMessage =
      "📊 Daily Task Report Summary - " + lastWorkingDay + " \n";

    for (let i = 0; i < results.length; i++) {
      const data = results[i];
      if (Array.isArray(data) && !data[1]) {
        employeesToRemind.push(data[0]);
        summaryMessage += data[0] + "❌ Not Added \n";
      } else {
        summaryMessage += data + "\n";
      }
    }

    // Send employee reminders if needed
    if (employeesToRemind.length > 0) {
      Logger.log("Sending reminders to: " + employeesToRemind.join(", "));
      this.sendEmployeeReminder(employeesToRemind);
    } else {
      Logger.log(
        "No employee reminders needed - all timesheets are properly submitted",
      );
    }

    Logger.log("\n" + "=".repeat(50));
    Logger.log("VERIFICATION SUMMARY");
    Logger.log("=".repeat(50));
    Logger.log(summaryMessage);
    Logger.log("=".repeat(50));

    this.sendGoogleChatMessage(summaryMessage);

    return results;
  }
}

/**
 * Main function to run the verification
 * This can be called manually or set up as a time-driven trigger
 */
function runDailyVerification() {
  if (typeof syncSettingsFromSheet_ === "function") {
    syncSettingsFromSheet_();
  }
  const verifier = new SheetsVerifier();
  const results = verifier.runDailyVerification();
  return results;
}

/**
 * Setup function to configure Script Properties
 * Run this once to set up your configuration
 */
function setupConfiguration() {
  const props = PropertiesService.getScriptProperties();
  props.deleteAllProperties();
  props.setProperty("DESTINATION_FOLDER_ID", "YOUR_DESTINATION_FOLDER_ID");
  props.setProperty("GOOGLE_CHAT_WEBHOOK_URL", "YOUR_WEBHOOK_URL");
  props.setProperty("EMPLOYEE_ALERT_WEBHOOK_URL", "YOUR_ALERT_WEBHOOK_URL");
  props.setProperty("TEST_MODE", "false");
  props.setProperty("ROWS_TO_CHECK_AFTER_DATE", "5");
  props.setProperty("SHEET_DATA_RANGE", "A1:K150");

  // Set holidays as JSON array
  const holidays = [
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
  props.setProperty("HOLIDAYS", JSON.stringify(holidays));

  Logger.log("\n✓ Configuration set up successfully!");
  Logger.log("Please define DESTINATION_FOLDER_ID and Webhook URLs in Script Properties, then run verification.");
}

