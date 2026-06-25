/**
 * Timesheet Verification System
 *
 * This script verifies employee timesheet entries based on the configurations
 * defined in the Master Config Spreadsheet ("Manager Sheet") and sends daily chat reports/reminders.
 * It accesses individual employee worksheets directly using the Sheet ID stored in the config.
 *
 * ============================================================================
 * QUICK START
 * ============================================================================
 *
 * 1. Configure Master Config Sheet:
 *    - Ensure your "Manager Sheet" contains the columns: Project Name, Team Member,
 *      Email, Chat ID, Sheet ID, and Active.
 *
 * 2. Set script properties:
 *    - Go to Project Settings -> Script Properties in Apps Script.
 *    - Define the required properties below (such as Webhook URLs).
 *
 * 3. Setup initial properties (Optional):
 *    - Run setupConfiguration() to set initial default property values.
 *
 * ============================================================================
 * CONFIGURATION - Script Properties
 * ============================================================================
 * - DESTINATION_FOLDER_ID: (Required) Folder ID where employee spreadsheets are generated
 * - GOOGLE_CHAT_WEBHOOK_URL: (Required) Webhook URL for manager daily summary reports
 * - EMPLOYEE_ALERT_WEBHOOK_URL: (Required) Webhook URL for employee alerts/reminders
 * - TEST_MODE: "true" or "false" (default: false)
 * - HOLIDAYS: JSON array of holiday dates in YYYY-MM-DD format
 * - ROWS_TO_CHECK_AFTER_DATE: Number of rows to check after finding a date (default: 5)
 * - SHEET_DATA_RANGE: Cell range to scan for inputs (default: "A1:K150")
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
  _fetchEmployeesFromMasterSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log("WARNING: Active spreadsheet is not accessible.");
      return [];
    }

    try {
      const sheet = ss.getSheetByName("Manager Sheet");
      if (!sheet) {
        Logger.log("WARNING: Sheet 'Manager Sheet' not found in Master Config Spreadsheet");
        return [];
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
          templateType: templateType
        });
      }

      return employees;
    } catch (e) {
      Logger.log("Error loading config from master sheet: " + e);
      return [];
    }
  }

  /**
   * Load configuration from Script Properties and Master Sheet
   */
  _loadConfig() {
    const props = PropertiesService.getScriptProperties();

    // Fetch dynamic employee list from Master Config Sheet
    const employees = this._fetchEmployeesFromMasterSheet();

    // Map employee name -> chat ID
    const employeeChatIds = {};
    employees.forEach(emp => {
      if (emp.chatId) {
        employeeChatIds[emp.name] = emp.chatId;
      }
    });

    // Parse holidays
    let holidays = [];
    try {
      const holidaysStr = props.getProperty("HOLIDAYS");
      if (holidaysStr && JSON.parse(holidaysStr).length > 0) {
        holidays = JSON.parse(holidaysStr);
      } else {
        // Default holidays for 2026
        holidays = [
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
      }
    } catch (e) {
      Logger.log("Error parsing HOLIDAYS: " + e);
    }

    return {
      employees: employees,
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
   * Parse last working day entries from sheet data for Alternative Template
   */
  parseAlternativeLastWorkingDayEntries(sheetData, targetDate) {
    if (!sheetData || sheetData.length < 2) {
      return [];
    }

    // Normalize target date to "d-MMM" format
    const normalizedTargetDate = this._normalizeDate(targetDate);
    if (!normalizedTargetDate) {
      Logger.log("Error: Could not normalize target date: " + targetDate);
      return [];
    }

    Logger.log("[Alternative] Looking for normalized target date: " + normalizedTargetDate);

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
      Logger.log("[Alternative] Could not find header row with 'Date' in column B");
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
          "[Alternative] Found target date '" +
            normalizedTargetDate +
            "' (original: " +
            dateCellValue +
            ") at row " +
            (i + 1) +
            " (column B)",
        );

        // Parse this row and the next configurable rows for all tasks on this date
        const rowsToCheck = this.config.rowsToCheckAfterDate + 1; // +1 to include the current row

        // First, check if the first row has total_duration (merged cell case in Col K, index 10)
        const firstRow = sheetData[i];
        let mergedTotalDuration = "";
        if (firstRow.length > 10) {
          mergedTotalDuration = this._normalizeTime(firstRow[10]);
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
            if (
              normalizedNextDate &&
              normalizedNextDate !== normalizedTargetDate
            ) {
              Logger.log(
                "  [Alternative] Found new date at row " +
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
              Logger.log(
                "  [Alternative] Found potential new date at row " +
                  (j + 1) +
                  ", stopping task collection",
              );
              break;
            }
          }

          // Parse the task data according to alternative column structure
          const entry = {
            row_number: j + 1,
            date: normalizedTargetDate,
            module_area: taskRow.length > 3 ? String(taskRow[3]) : "",
            task_details: taskRow.length > 4 ? String(taskRow[4]) : "",
            status: taskRow.length > 6 ? String(taskRow[6]) : "",
            activity_type: "Task", // Alternative doesn't have Activity Type, default to Task
            start_time:
              taskRow.length > 7 ? this._normalizeTime(taskRow[7]) : "",
            end_time: taskRow.length > 8 ? this._normalizeTime(taskRow[8]) : "",
            total_duration:
              taskRow.length > 10 ? this._normalizeTime(taskRow[10]) : "",
            remarks: taskRow.length > 11 ? String(taskRow[11]) : "",
          };

          // If this entry doesn't have total_duration but we found it in merged cell, use it
          if (!entry.total_duration.trim() && mergedTotalDuration) {
            entry.total_duration = mergedTotalDuration;
          }

          // Only add entries that have meaningful task details
          if (entry.task_details.trim() || entry.module_area.trim()) {
            targetEntries.push(entry);
            Logger.log(
              "  [Alternative] Added task from row " +
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
              "  [Alternative] Added additional time entry from row " +
                (j + 1) +
                " for previous task: " +
                previousEntry.task_details.substring(0, 50) +
                "...",
            );
          }
        }
        break; // found and parsed the date, stop search
      }
      i++;
    }

    Logger.log(
      "[Alternative] Total entries found for " +
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

    const employees = this.config.employees || [];
    const tabName = lastWorkingDayMonth + "-" + lastWorkingDayYear; // e.g. June-2026

    for (let i = 0; i < employees.length; i++) {
      const emp = employees[i];
      Logger.log("Checking " + emp.name + "'s sheet...");

      const spreadsheetId = emp.sheetId;
      if (!spreadsheetId) {
        Logger.log("Warning: No Sheet ID found for " + emp.name + ", skipping verification.");
        results.push([emp.name, false]);
        continue;
      }

      const sheetData = this.getEngineerSheetData(
        tabName,
        lastWorkingDayMonth,
        spreadsheetId,
      );
      const templateType = emp.templateType || "Standard";
      let lastDayEntries = [];
      if (templateType === "Alternative") {
        lastDayEntries = this.parseAlternativeLastWorkingDayEntries(
          sheetData,
          lastWorkingDay,
        );
      } else {
        lastDayEntries = this.parseLastWorkingDayEntries(
          sheetData,
          lastWorkingDay,
        );
      }

      Logger.log(
        "Found " +
          lastDayEntries.length +
          " entries for " +
          emp.name +
          " on " +
          lastWorkingDay,
      );

      if (lastDayEntries.length > 0) {
        const analysis = this.generateAnalysisMessage(
          emp.name,
          lastDayEntries,
        );
        const prefix = emp.projectName ? "[" + emp.projectName + "] " : "";
        results.push(prefix + analysis);
      } else {
        results.push([emp.name, false]);
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

