/**
 * Slide Remote — Purge Old Commands
 *
 * Deletes rows older than `cutoffMinutes` from the Form's response sheet,
 * so the sheet doesn't grow forever. Meant to be run on a time-driven
 * trigger (see README.md, Part 5) — not called directly.
 *
 * Setup:
 *   1. Open the linked Sheet, then Extensions -> Apps Script.
 *   2. Delete any placeholder code and paste in the contents of this file.
 *   3. Update SHEET_NAME below if your response tab isn't "Form Responses 1".
 *   4. Save, then add a time-driven trigger for purgeOldCommands
 *      (see README.md, Part 5, for the exact trigger settings).
 */

function purgeOldCommands() {
  const SHEET_NAME = 'Form Responses 1'; // match your actual tab name
  const cutoffMinutes = 15;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = data.length - 1; i >= 1; i--) { // bottom-up, skip header row
    const rowAge = (now - new Date(data[i][0])) / 60000; // column A = Timestamp
    if (rowAge > cutoffMinutes) sheet.deleteRow(i + 1);
  }
}
