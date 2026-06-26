/**
 * Sends an email notification to the user designated in the "Prepared by" field (D6) of the active sheet.
 * The email includes a direct link to the spreadsheet and asks them to fill out the tab.
 */
function alert_member() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  // The 'Prepared by' value is in cell D6
  const preparedByNameCell = sheet.getRange("D6");
  const preparedByName = preparedByNameCell.getValue() ? preparedByNameCell.getValue().toString().trim().toLowerCase() : "";

  if (!preparedByName) {
    Logger.log("No member selected in the 'Prepared by' field (D6).");
    SpreadsheetApp.getUi().alert("Error: No member selected in the 'Prepared by' field (D6).");
    return;
  }

  let recipientEmail = null;
  let matchedKeyName = "";
  const keys = Object.keys(PREPARED_BY_MEMBERS);
  for (let i = 0; i < keys.length; i++) {
    if (keys[i].trim().toLowerCase() === preparedByName) {
      recipientEmail = PREPARED_BY_MEMBERS[keys[i]];
      matchedKeyName = keys[i];
      break;
    }
  }

  if (!recipientEmail) {
    Logger.log("No email mapping found for member: '" + preparedByName + "'");
    Logger.log("Configured keys in PREPARED_BY_MEMBERS: " + JSON.stringify(Object.keys(PREPARED_BY_MEMBERS)));
    SpreadsheetApp.getUi().alert("Error: No email address found for the member '" + preparedByName + "'.\n\n" +
                                 "Configured members in script: " + JSON.stringify(Object.keys(PREPARED_BY_MEMBERS)) + "\n\n" +
                                 "Please ensure you have saved both files in the Google Apps Script editor.");
    return;
  }

  const ssUrl = ss.getUrl();
  // Use the matching key casing for presentation (e.g. "Ashwin Sanalkumar")
  const formattedName = matchedKeyName;

  const subject = `Action Required: Fill MoM Sheet - ${sheetName}`;

  // Beautiful HTML body matching the corporate blue aesthetic
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #d9d9d9; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
      <div style="background-color: #8faadc; padding: 20px; text-align: center;">
        <h2 style="color: #000000; margin: 0; font-size: 20px;">Minutes of Meeting (MoM) Notification</h2>
      </div>
      <div style="padding: 24px; color: #333333; line-height: 1.6;">
        <p style="margin-top: 0;">Hi <strong>${formattedName}</strong>,</p>
        <p>The Minutes of Meeting sheet for the tab <strong>${sheetName}</strong> has been generated and you are assigned as the author/preparer.</p>
        <p>Please click the button below to open the spreadsheet and fill out the necessary details in the tab:</p>
        <div style="text-align: center; margin: 30px 0;">
          <a href="${ssUrl}" style="background-color: #8faadc; color: #000000; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">Open Spreadsheet</a>
        </div>
        <p style="font-size: 13px; color: #666666; margin-bottom: 0;">Link fallback: <a href="${ssUrl}" style="color: #4a90e2;">${ssUrl}</a></p>
      </div>
      <div style="background-color: #f2f2f2; padding: 12px; text-align: center; font-size: 11px; color: #777777; border-top: 1px solid #e0e0e0;">
        This is an automated notification from Google Sheets.
      </div>
    </div>
  `;

  try {
    // Share the spreadsheet with the recipient as an editor so they can fill it out
    try {
      ss.addEditor(recipientEmail);
      Logger.log("Successfully shared spreadsheet with " + recipientEmail + " as an editor.");
    } catch (shareError) {
      Logger.log("Warning: Could not automatically share spreadsheet with " + recipientEmail + ": " + shareError.toString());
    }

    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log("Email successfully sent to " + formattedName + " (" + recipientEmail + ").");
    SpreadsheetApp.getUi().alert("Success: Spreadsheet shared and email alert successfully sent to " + formattedName + " (" + recipientEmail + ")!");
  } catch (error) {
    Logger.log("Failed to send email: " + error.toString());
    SpreadsheetApp.getUi().alert("Error: Failed to send email. " + error.toString());
  }
}
