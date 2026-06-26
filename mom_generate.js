/**
 * Generates a Minutes of Meeting (MoM) sheet based on a professional template.
 * The sheet is named after the current date (e.g. YYYY-MM-DD).
 */
function createMinutesOfMeetingSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Format today's date as YYYY-MM-DD
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    const sheetName = `${day}-${month}-${year}`;

    // Check if a sheet with this name already exists
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
        // Delete the existing sheet to start fresh
        ss.deleteSheet(sheet);
    }

    // Find the correct insertion index to keep the sheets in date order (ascending)
    const sheets = ss.getSheets();
    let insertIndex = sheets.length; // Default to end of spreadsheet

    function dateToKey(name) {
        const parts = name.split('-');
        if (parts.length === 3) {
            const d = parts[0];
            const m = parts[1];
            const y = parts[2];
            // Match exactly DD-MM-YYYY format
            if (d.length === 2 && m.length === 2 && y.length === 4 && !isNaN(Number(y + m + d))) {
                return Number(y + m + d);
            }
        }
        return null;
    }

    const newKey = dateToKey(sheetName);
    if (newKey !== null) {
        for (let i = 0; i < sheets.length; i++) {
            const currentKey = dateToKey(sheets[i].getName());
            if (currentKey !== null && currentKey > newKey) {
                insertIndex = i;
                break;
            }
        }
    }

    // Insert a new sheet at the sorted position
    sheet = ss.insertSheet(sheetName, insertIndex);

    // Ensure the sheet is active
    sheet.activate();

    // Set grid lines visible
    sheet.setHiddenGridlines(false);

    // Define Colors (Matching the clean, corporate blue aesthetic)
    const COLOR_HEADER_BG = '#8faadc'; // Soft Steel/Cornflower Blue
    const COLOR_BORDER = '#000000';    // Black for borders (as per screenshot)
    const COLOR_TEXT_DARK = '#000000'; // Dark/Black text for high contrast

    // Define Attendance Dropdown Options (Edit these values to change the dropdown list)
    const ATTENDANCE_OPTIONS = ['Present', 'Absent', 'Excused'];

    // Define Prepared By Dropdown Options (Edit these values to change the dropdown list)
    const PREPARED_BY_OPTIONS = ['ashwin', 'abhiram'];

    // Define Column Widths (A is an empty spacer column, B-G are data columns)
    const columnWidths = {
        1: 100,   // Col A: Empty spacer column (very narrow)
        2: 62,   // Col B: Sl. No. / SI No. (narrow)
        3: 99,  // Col C: Invitees / Action Item (wide)
        4: 158,  // Col D: Attendance / Responsibility (medium)
        5: 100,  // Col E: Remarks (first table) / Target Date of Closure (second table) (medium)
        6: 100,   // Col F: Empty (first table) / Part of Remarks (second table) (narrow)
        7: 882   // Col G: Empty (first table) / Part of Remarks (second table) (very wide)
    };

    for (let col in columnWidths) {
        sheet.setColumnWidth(Number(col), columnWidths[col]);
    }

    // Helper to apply borders to a range (thin black borders matching the exact design)
    function applyBorders(range) {
        range.setBorder(true, true, true, true, true, true, COLOR_BORDER, SpreadsheetApp.BorderStyle.SOLID);
    }

    // Helper to style a range
    function styleRange(range, bg, isBold, fontSize, align, fontColor) {
        if (bg) range.setBackground(bg);
        if (isBold !== undefined) range.setFontWeight(isBold ? 'bold' : 'normal');
        if (fontSize) range.setFontSize(fontSize);
        if (align) range.setHorizontalAlignment(align);
        if (fontColor) range.setFontColor(fontColor);
        range.setVerticalAlignment('middle');
    }

    // --- ROW HEIGHTS INITIALIZATION ---
    sheet.setRowHeight(1, 15);  // Spacer row 1
    sheet.setRowHeight(2, 40);  // Title row 2
    sheet.setRowHeight(3, 15);  // Spacer row 3
    sheet.setRowHeight(4, 25);  // Meta row 4
    sheet.setRowHeight(5, 25);  // Meta row 5
    sheet.setRowHeight(6, 25);  // Meta row 6 (Prepared by)
    sheet.setRowHeight(7, 15);  // Spacer row 7 (was 6)
    sheet.setRowHeight(8, 25);  // First table header row 8 (was 7)

    // First table data (Rows 9-16)
    for (let r = 9; r <= 16; r++) {
        sheet.setRowHeight(r, 20);
    }

    // Spacers between tables (Rows 17-20)
    sheet.setRowHeight(17, 15);
    sheet.setRowHeight(18, 15);
    sheet.setRowHeight(19, 15);
    sheet.setRowHeight(20, 15);

    sheet.setRowHeight(21, 21);  // Second table header row 21 (was 20)

    // Second table data (Rows 22-39)
    for (let r = 22; r <= 39; r++) {
        sheet.setRowHeight(r, 20);
    }

    // --- 1. MAIN HEADER ---
    const titleRange = sheet.getRange("B2:G2");
    titleRange.merge();
    titleRange.setValue("Minutes of Meeting");
    styleRange(titleRange, COLOR_HEADER_BG, true, 14, "center", COLOR_TEXT_DARK);
    applyBorders(titleRange);

    // --- 2. MEETING INFO BLOCK ---
    // Row 4
    const cellB4C4 = sheet.getRange("B4:C4");
    cellB4C4.merge();
    cellB4C4.setValue("Meeting ID");
    styleRange(cellB4C4, COLOR_HEADER_BG, true, 10, "center", COLOR_TEXT_DARK);

    const cellD4 = sheet.getRange("D4"); // Input value, empty by default
    styleRange(cellD4, null, false, 10, "center", COLOR_TEXT_DARK);

    const cellE4F4 = sheet.getRange("E4:F4");
    cellE4F4.merge();
    cellE4F4.setValue("Project/ Department Name");
    styleRange(cellE4F4, COLOR_HEADER_BG, true, 10, "center", COLOR_TEXT_DARK);

    const cellG4 = sheet.getRange("G4");
    cellG4.setValue("Interview Platform");
    styleRange(cellG4, null, false, 10, "left", COLOR_TEXT_DARK);

    // Row 5
    const cellB5C5 = sheet.getRange("B5:C5");
    cellB5C5.merge();
    cellB5C5.setValue("Date of Meeting");
    styleRange(cellB5C5, COLOR_HEADER_BG, true, 10, "center", COLOR_TEXT_DARK);

    const cellD5 = sheet.getRange("D5");
    cellD5.setValue(`${month}/${day}/${year}`);
    styleRange(cellD5, null, false, 10, "center", COLOR_TEXT_DARK);

    const cellE5F5 = sheet.getRange("E5:F5");
    cellE5F5.merge();
    cellE5F5.setValue("Duration");
    styleRange(cellE5F5, COLOR_HEADER_BG, true, 10, "center", COLOR_TEXT_DARK);

    const cellG5 = sheet.getRange("G5");
    cellG5.setValue("45 minutes");
    styleRange(cellG5, null, false, 10, "left", COLOR_TEXT_DARK);

    // Row 6
    const cellB6C6 = sheet.getRange("B6:C6");
    cellB6C6.merge();
    cellB6C6.setValue("Prepared by");
    styleRange(cellB6C6, COLOR_HEADER_BG, true, 10, "center", COLOR_TEXT_DARK);

    const cellD6 = sheet.getRange("D6");
    const preparedByRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(PREPARED_BY_OPTIONS, true)
        .setAllowInvalid(false)
        .build();
    cellD6.setDataValidation(preparedByRule);
    cellD6.setValue(PREPARED_BY_OPTIONS[0]);
    styleRange(cellD6, null, false, 10, "center", COLOR_TEXT_DARK);

    const cellE6F6 = sheet.getRange("E6:F6");
    cellE6F6.merge();
    styleRange(cellE6F6, null, false, 10, "center", COLOR_TEXT_DARK);

    const cellG6 = sheet.getRange("G6");
    styleRange(cellG6, null, false, 10, "left", COLOR_TEXT_DARK);

    // Set borders for Info Block
    applyBorders(sheet.getRange("B4:G6"));

    // --- 3. ATTENDANCE TABLE ---
    // Table Header (Row 8)
    const attHeader = sheet.getRange("B8:G8");
    sheet.getRange("B8").setValue("Sl. No.");
    sheet.getRange("C8").setValue("Invitees");
    sheet.getRange("D8").setValue("Attendance");
    sheet.getRange("E8").setValue("Remarks");
    // F8 and G8 are kept blank but styled as part of the header row
    sheet.getRange("F8").setValue("");
    sheet.getRange("G8").setValue("");

    // Style table headers
    styleRange(attHeader, COLOR_HEADER_BG, true, 10, "center", COLOR_TEXT_DARK);
    applyBorders(attHeader);

    // Table Data (Rows 9 to 16 - 8 rows)
    const validationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(ATTENDANCE_OPTIONS, true)
        .setAllowInvalid(false)
        .build();

    for (let i = 0; i < 8; i++) {
        const row = 9 + i;

        // Sl. No.
        const cellSl = sheet.getRange(row, 2);
        cellSl.setValue(i + 1);
        styleRange(cellSl, null, false, 10, "center");

        // Invitee Name (Left blank as requested)
        const cellName = sheet.getRange(row, 3);
        cellName.setValue("");
        styleRange(cellName, null, false, 10, "left");

        // Attendance dropdown validation
        const cellAttendance = sheet.getRange(row, 4);
        cellAttendance.setDataValidation(validationRule);
        cellAttendance.setValue(ATTENDANCE_OPTIONS[0]);
        styleRange(cellAttendance, null, false, 10, "center");

        // Remarks (Col E)
        styleRange(sheet.getRange(row, 5), null, false, 10, "left");

        // Col F and G (empty)
        styleRange(sheet.getRange(row, 6), null, false, 10, "left");
        styleRange(sheet.getRange(row, 7), null, false, 10, "left");
    }

    // Apply borders for the attendance table data
    applyBorders(sheet.getRange("B9:G16"));

    // --- 4. ACTION ITEMS TABLE ---
    // Table Header (Row 21)
    const actionHeader = sheet.getRange("B21:G21");
    sheet.getRange("B21").setValue("SI No.");
    
    // Action Item in second table merges C:D
    const cellActionHeader = sheet.getRange("C21:D21");
    cellActionHeader.merge();
    cellActionHeader.setValue("Action Item");
    
    sheet.getRange("E21").setValue("Responsibility");
    
    const cellTargetDate = sheet.getRange("F21");
    cellTargetDate.setValue("Target Date of\nClosure");
    cellTargetDate.setWrap(true);
    
    sheet.getRange("G21").setValue("Remarks");

    // Style table headers
    styleRange(actionHeader, COLOR_HEADER_BG, true, 10, "center", COLOR_TEXT_DARK);
    applyBorders(actionHeader);

    // Table Data (Rows 22 to 39 - 18 rows)
    for (let i = 0; i < 18; i++) {
        const row = 22 + i;

        // SI No.
        const cellSl = sheet.getRange(row, 2);
        cellSl.setValue(i + 1);
        styleRange(cellSl, null, false, 10, "center");

        // Action Item in second table merges C:D for each row
        const cellAction = sheet.getRange(row, 3, 1, 2);
        cellAction.merge();
        styleRange(cellAction, null, false, 10, "left");

        // Responsibility, Target Date, Remarks are blank inputs
        styleRange(sheet.getRange(row, 5), null, false, 10, "left");
        styleRange(sheet.getRange(row, 6), null, false, 10, "center");
        styleRange(sheet.getRange(row, 7), null, false, 10, "left");
    }

    // Apply borders for the action items table data
    applyBorders(sheet.getRange("B22:G39"));
}
