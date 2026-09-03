/**
 * Google Apps Script
 * 1. Create/open a Google Sheet.
 * 2. Extensions → Apps Script.
 * 3. Paste this code.
 * 4. Change SHEET_NAME if required.
 * 5. Deploy → New deployment → Web app.
 *    Execute as: Me
 *    Who has access: Anyone
 * 6. Copy the Web app URL into APPS_SCRIPT_URL in index.html.
 */

const SHEET_NAME = "Leads";

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ok:true, message:"Booking endpoint is live"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const sheet = getSheet_();
    const raw = e && e.postData ? e.postData.contents : "{}";
    const data = JSON.parse(raw);

    const headers = [
      "Timestamp", "Name", "Phone", "Age", "City",
      "PCOS Status", "Primary Concern", "Message", "Source"
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }

    sheet.appendRow([
      new Date(),
      clean_(data.name),
      clean_(data.phone),
      clean_(data.age),
      clean_(data.city),
      clean_(data.pcosStatus),
      clean_(data.primaryConcern),
      clean_(data.message),
      clean_(data.source)
    ]);

    return json_({ok:true});
  } catch (err) {
    return json_({ok:false, error:String(err)});
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function clean_(value) {
  return value == null ? "" : String(value).trim().slice(0, 2000);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
