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

    // Insert a new sheet
    sheet = ss.insertSheet(sheetName);

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
    sheet.setRowHeight(6, 15);  // Spacer row 6
    sheet.setRowHeight(7, 25);  // First table header row 7

    // First table data (Rows 8-15)
    for (let r = 8; r <= 15; r++) {
        sheet.setRowHeight(r, 20);
    }

    // Spacers between tables (Rows 16-19)
    sheet.setRowHeight(16, 15);
    sheet.setRowHeight(17, 15);
    sheet.setRowHeight(18, 15);
    sheet.setRowHeight(19, 15);

    sheet.setRowHeight(20, 21);  // Second table header row 20

    // Second table data (Rows 21-38)
    for (let r = 21; r <= 38; r++) {
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

    // Set borders for Info Block
    applyBorders(sheet.getRange("B4:G5"));

    // --- 3. ATTENDANCE TABLE ---
    // Table Header (Row 7)
    const attHeader = sheet.getRange("B7:G7");
    sheet.getRange("B7").setValue("Sl. No.");
    sheet.getRange("C7").setValue("Invitees");
    sheet.getRange("D7").setValue("Attendance");
    sheet.getRange("E7").setValue("Remarks");
    // F7 and G7 are kept blank but styled as part of the header row
    sheet.getRange("F7").setValue("");
    sheet.getRange("G7").setValue("");

    // Style table headers
    styleRange(attHeader, COLOR_HEADER_BG, true, 10, "center", COLOR_TEXT_DARK);
    applyBorders(attHeader);

    // Table Data (Rows 8 to 15 - 8 rows)
    const validationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(ATTENDANCE_OPTIONS, true)
        .setAllowInvalid(false)
        .build();

    for (let i = 0; i < 8; i++) {
        const row = 8 + i;

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
        styleRange(cellAttendance, null, false, 10, "center");

        // Remarks (Col E)
        styleRange(sheet.getRange(row, 5), null, false, 10, "left");

        // Col F and G (empty)
        styleRange(sheet.getRange(row, 6), null, false, 10, "left");
        styleRange(sheet.getRange(row, 7), null, false, 10, "left");
    }

    // Apply borders for the attendance table data
    applyBorders(sheet.getRange("B8:G15"));

    // --- 4. ACTION ITEMS TABLE ---
    // Table Header (Row 20)
    const actionHeader = sheet.getRange("B20:G20");
    sheet.getRange("B20").setValue("SI No.");
    
    // Action Item in second table merges C:D
    const cellActionHeader = sheet.getRange("C20:D20");
    cellActionHeader.merge();
    cellActionHeader.setValue("Action Item");
    
    sheet.getRange("E20").setValue("Responsibility");
    
    const cellTargetDate = sheet.getRange("F20");
    cellTargetDate.setValue("Target Date of\nClosure");
    cellTargetDate.setWrap(true);
    
    sheet.getRange("G20").setValue("Remarks");

    // Style table headers
    styleRange(actionHeader, COLOR_HEADER_BG, true, 10, "center", COLOR_TEXT_DARK);
    applyBorders(actionHeader);

    // Table Data (Rows 21 to 38 - 18 rows)
    for (let i = 0; i < 18; i++) {
        const row = 21 + i;

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
    applyBorders(sheet.getRange("B21:G38"));
}
