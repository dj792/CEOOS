// ============================================================
// 1to100 CRM - Google Apps Script Backend
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Go to https://script.google.com and click "New Project"
// 2. Delete any existing code and paste this entire file
// 3. Click "Deploy" → "New Deployment"
// 4. Type: Web App
// 5. Description: "1to100 CRM API"
// 6. Execute as: Me
// 7. Who has access: Anyone
// 8. Click Deploy → Authorize → Copy the Web App URL
// 9. Paste that URL into the setup screen of 1to100_CRM.html
//
// NOTE: If you update this code later, click Deploy → Manage Deployments
//       → edit your existing deployment and click Deploy again (don't create
//       a new one, or you'll need to update the URL in the CRM).
// ============================================================

const SPREADSHEET_ID = '1guZq-gKmp7OmI9Q3Dh8ZlDcp650Ak047Eapyr681mOk';
const SHEET_NAME = 'Master List';

// ============================================================
// ENTRY POINT — handles all GET requests
// Supports both plain JSON and JSONP (for file:// origins)
// ============================================================
function doGet(e) {
  var callback = e.parameter.callback || '';
  var result;

  try {
    var action  = e.parameter.action  || '';
    var payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};

    switch (action) {
      case 'getData':   result = getData();                                    break;
      case 'addRow':    result = addRow(payload.row);                         break;
      case 'updateRow': result = updateRow(payload.rowIndex, payload.row);    break;
      case 'deleteRow': result = deleteRow(payload.rowIndex);                 break;
      default:          result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  var json = JSON.stringify(result);

  // JSONP mode — wraps response in callback(…) so it works from file:// URLs
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Standard JSON mode (for localhost / hosted origins)
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// HANDLERS
// ============================================================

/** Returns all rows (including header row) as a 2D array */
function getData() {
  var sheet = getSheet();
  var values = sheet.getDataRange().getValues();
  return { success: true, data: values };
}

/** Appends a new row at the end of the sheet */
function addRow(row) {
  if (!row || !Array.isArray(row)) return { error: 'Invalid row data' };
  getSheet().appendRow(row);
  return { success: true };
}

/** Updates the row at rowIndex (1-based Google Sheets row number, row 1 = header) */
function updateRow(rowIndex, row) {
  if (!rowIndex || !row || !Array.isArray(row)) return { error: 'Invalid update data' };
  var sheet = getSheet();
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  return { success: true };
}

/** Deletes the row at rowIndex (1-based Google Sheets row number) */
function deleteRow(rowIndex) {
  if (!rowIndex) return { error: 'Invalid row index' };
  getSheet().deleteRow(rowIndex);
  return { success: true };
}

// ============================================================
// HELPERS
// ============================================================
function getSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found in spreadsheet');
  return sheet;
}
