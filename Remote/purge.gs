/**
 * Slide Remote — Sheet purge.
 *
 * Deletes any row in the Form's response sheet whose Timestamp (column A)
 * is older than `cutoffMinutes`. Bound to the Sheet via Extensions > Apps
 * Script, run on a time-driven trigger (see webSetup_slideRemote.md Part 5)
 * every 15 minutes.
 *
 * Kept comfortably ahead of the receiver's own ~15-second staleness
 * window (see Receiver/README.md) — this purge is about not letting the
 * Sheet grow forever, not about the receiver's own dedup/staleness logic,
 * which has already discarded a row as too old to act on long before this
 * ever runs.
 */
function purgeOldCommands() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1'); // match your actual tab name
  const data = sheet.getDataRange().getValues();
  const cutoffMinutes = 15;
  const now = new Date();

  for (let i = data.length - 1; i >= 1; i--) { // bottom-up, skip header row
    const rowAge = (now - new Date(data[i][0])) / 60000; // column A = Timestamp
    if (rowAge > cutoffMinutes) sheet.deleteRow(i + 1);
  }
}
