function doGet() {
  // Main entry point; always serve index.html (the login page)
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Login')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processLogin(name, regno, password) {
  var SHEET_ID = '1AU0SKyW4gtRSry-9q1DfGSFeYATTXtwVDgpotflE5fw';
  var TAB_NAME = 'Password';
  var regExp = /^[a-zA-Z]\d{3}\/\d{4}[a-zA-Z]\/\d{2}$/;
  var regnoTrimmed = regno.trim();
  if (!regExp.test(regnoTrimmed)) {
    return { success: false, message: "Registration number format invalid. Use e.g. E103/1234G/21" };
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(TAB_NAME);
  if (!ws) return {success: false, message: "Sheet/tab named 'Password' not found."};
  var data = ws.getDataRange().getValues();
  if (data.length <= 1) return {success: false, message: "No login data found."};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var sheetName = String(row[0] || "").trim().toLowerCase();
    var sheetRegNo = String(row[1] || "").trim().toLowerCase();
    var sheetPassword = String(row[2] || "").trim();
    if (
      name.trim().toLowerCase() === sheetName &&
      regnoTrimmed.toLowerCase() === sheetRegNo &&
      password.trim() === sheetPassword
    ) {
      return { success: true, message: "Login successful! Welcome, " + row[0] + "." };
    }
  }
  return { success: false, message: "Invalid details. Please check your Name, Reg No, or Password." };
}

function getCategoriesPage() {
  return HtmlService.createHtmlOutputFromFile('categories').getContent();
}
function getExpenditurePage() {
  return HtmlService.createHtmlOutputFromFile('expenditure').getContent();
}

// Expenditure backend
const SHEET_ID = '1AU0SKyW4gtRSry-9q1DfGSFeYATTXtwVDgpotflE5fw';
const TAB_NAME = 'Expenditure';

function submitExpenditureForm(formData) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) return {success: false, message: "Sheet not found!"};
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (!formData.date) {
    return {success: false, message: "Date is required."};
  }

  // 1. Parse dynamic Other item details
  let dynamicOtherDetails = [];
  if (formData.otherDetails) {
    try {
      dynamicOtherDetails = JSON.parse(formData.otherDetails); // array of {title, value}
    } catch(e) {}
  }

  // 2. Add new columns for custom/dynamic fields if not already present (case-insensitive)
  let headerLC = headers.map(h => h.toLowerCase());
  let newCols = [];
  dynamicOtherDetails.forEach(obj => {
    if (obj && obj.title && headerLC.indexOf(String(obj.title).toLowerCase()) === -1) {
      headers.push(obj.title);
      newCols.push(obj.title);
    }
  });
  if (newCols.length > 0) {
    sheet.insertColumnsAfter(headers.length - newCols.length, newCols.length);
    sheet.getRange(1, headers.length - newCols.length + 1, 1, newCols.length).setValues([newCols]);
  }

  // Re-fetch headers (after alteration)
  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headerLC = headers.map(h => h.toLowerCase());

  // Only use: DATE, AMOUNT, and dynamic item titles as columns (remove all static expense default columns)
  // Build row in headers order
  let row = Array(headers.length).fill('');
  let amountSum = 0;

  // Set DATE cell
  let dateIdx = headerLC.indexOf('date');
  if (dateIdx !== -1) row[dateIdx] = formData.date;

  // Calculate total amount from dynamicOtherDetails
  dynamicOtherDetails.forEach(obj => {
    if (obj && obj.title && (typeof obj.value !== "undefined")) {
      let idx = headerLC.indexOf(obj.title.toLowerCase());
      if (idx !== -1) {
        row[idx] = obj.value;
        amountSum += Number(obj.value) || 0;
      }
    }
  });

  // AMOUNT cell
  let amountIdx = headerLC.indexOf('amount');
  if (amountIdx !== -1) row[amountIdx] = amountSum;

  sheet.appendRow(row);

  // Still call submitDepartmentsForm if your logic needs it
  submitDepartmentsForm(formData);

  return {success: true, message: "Successfully submitted! Total: " + amountSum.toFixed(2)};
}

function submitExpenditureTotals(formData) {
  const EXP_TOTAL_TAB = 'Expenditure Totals';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(EXP_TOTAL_TAB);
  if (!sheet) return { success: false, message: 'Expenditure Totals tab not found!' };

  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let headerLC = headers.map(h => h.toLowerCase());

  // Items from formData.otherDetails
  let items = [];
  if (formData.otherDetails) {
    try { items = JSON.parse(formData.otherDetails); } catch(e) {}
  }

  // Track if new types are needed
  let newTypeCols = [];
  items.forEach(item => {
    if (item.type) {
      // If type not present, add column for it
      if (headerLC.indexOf(item.type.toLowerCase()) === -1) {
        headers.push(item.type);
        newTypeCols.push(item.type);
      }
    }
  });
  if (newTypeCols.length > 0) {
    sheet.insertColumnsAfter(headers.length - newTypeCols.length, newTypeCols.length);
    sheet.getRange(1, headers.length - newTypeCols.length + 1, 1, newTypeCols.length).setValues([newTypeCols]);
    // Re-fetch headers to update current sheet structure
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headerLC = headers.map(h => h.toLowerCase());
  }

  // Build output row
  let row = Array(headers.length).fill('');

  // Set DATE
  const dateIdx = headerLC.indexOf('date');
  if (dateIdx !== -1) row[dateIdx] = formData.date || "";

  // Set PARTICULARS (all item titles, joined by "; ")
  const particulars = items.map(it => it.title ? it.title : "").filter(Boolean).join('; ');
  const particularsIdx = headerLC.indexOf('particulars');
  if (particularsIdx !== -1) row[particularsIdx] = particulars;

  // For each item, output its value in the correct Type column
  items.forEach(item => {
    if (item.type) {
      const typeIdx = headerLC.indexOf(item.type.toLowerCase());
      if (typeIdx !== -1) {
        row[typeIdx] = Number(item.value) || "";
      }
    }
  });

  sheet.appendRow(row);

  return { success: true, message: "Expenditure Totals updated!" };
}





function downloadExpenditureExcel() {
  var sheetId = SHEET_ID;
  var ss = SpreadsheetApp.openById(sheetId);
  var url = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/export?'
    + 'format=xlsx'
    + '&gid=' + ss.getSheetByName(TAB_NAME).getSheetId();

  var token = ScriptApp.getOAuthToken();
  var options = {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  };
  var response = UrlFetchApp.fetch(url, options);
  var excelBlob = response.getBlob().setName(TAB_NAME + '.xlsx');
  var file = DriveApp.createFile(excelBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}


/**
 * 100% working export for a custom duration.
 *
 * Rules:
 * - Start: pick the first date row ON or AFTER the entered start date.
 * - End: pick the last date row ON or BEFORE the entered end date,
 *        and include ALL continuation rows under that date (blank date cells)
 *        up to the next date row block (or end of data).
 *
 * Accepts user inputs in either "dd/MM/yyyy" (e.g., 20/09/2024) or "yyyy-MM-dd" (e.g., 2024-11-18).
 * Column A in "Expenditure" is the DATE column. Some rows may have blank DATE to continue the same date block.
 *
 * Implementation detail for reliability:
 * - We copy the selected rows into a brand-new temporary spreadsheet, flush,
 *   export that spreadsheet to XLSX, then delete the temporary spreadsheet.
 *   This avoids timing/flush/gid issues that can lead to empty exports.
 */
function downloadExpenditureExcelForDuration(userStart, userEnd) {
  // Reuse your existing constants if already declared in your project:
  // const SHEET_ID = '...';
  // const TAB_NAME = 'Expenditure';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + TAB_NAME);

  // 1) Normalize incoming dates to ISO (yyyy-MM-dd)
  const isoStart = normalizeToIso(userStart);
  const isoEnd   = normalizeToIso(userEnd);
  if (!isoStart || !isoEnd) {
    throw new Error('Invalid date format. Use dd/MM/yyyy (e.g., 20/09/2024) or yyyy-MM-dd.');
  }
  if (isoStart > isoEnd) {
    throw new Error('Start date is after end date.');
  }

  // 2) Read sheet data (values + display values)
  const range   = sheet.getDataRange();
  const data    = range.getValues();           // raw values (includes header)
  const display = range.getDisplayValues();    // what UI shows (e.g., 2024-11-18) (includes header)
  if (data.length < 2) throw new Error('No data in sheet.');

  const headers = data[0];
  const firstDataRowIdx = 1; // zero-based index; row 1 is the first data row

  // 3) Build "anchor" list: rows that start a date block (Column A shows yyyy-MM-dd)
  //    We rely on display values to match the visible date format in your sheet.
  const anchors = [];
  for (let r = firstDataRowIdx; r < display.length; r++) {
    const shown = String(display[r][0] || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(shown)) {
      anchors.push({ rowIdx: r, iso: shown }); // rowIdx is zero-based into data/display
    }
  }
  if (anchors.length === 0) {
    throw new Error('No date anchors found in column A (expected yyyy-MM-dd).');
  }

  // Ensure anchors are sorted by date (sheet typically is, but we guard anyway)
  anchors.sort((a, b) => (a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : a.rowIdx - b.rowIdx));

  // 4) Locate start anchor (first on/after isoStart)
  let startAnchorIdx = -1;
  for (let i = 0; i < anchors.length; i++) {
    if (anchors[i].iso >= isoStart) {
      startAnchorIdx = i;
      break;
    }
  }
  if (startAnchorIdx === -1) {
    // No date on/after start -> no rows match the rule
    throw new Error('No data on or after the selected start date (' + isoStart + ').');
  }

  // 5) Locate end anchor (last on/before isoEnd)
  let endAnchorIdx = -1;
  for (let i = 0; i < anchors.length; i++) {
    if (anchors[i].iso <= isoEnd) endAnchorIdx = i;
    else break; // because anchors sorted ascending
  }
  if (endAnchorIdx === -1) {
    // All anchors are after isoEnd -> no rows match the rule
    throw new Error('No data on or before the selected end date (' + isoEnd + ').');
  }

  // Validate window
  if (startAnchorIdx > endAnchorIdx) {
    throw new Error('No records in the selected date range (' + isoStart + ' to ' + isoEnd + ').');
  }

  // 6) Compute slice bounds in data array
  // Start at the start anchor row (inclusive)
  const startRowIdxInData = anchors[startAnchorIdx].rowIdx;

  // End at the end of the end anchor's block:
  // i.e., the row before the next anchor OR end of data
  const endExclusiveRowIdxInData = (endAnchorIdx < anchors.length - 1)
    ? anchors[endAnchorIdx + 1].rowIdx
    : data.length;

  // 7) Build rows to export: headers + selected body rows
  const bodyRows = data.slice(startRowIdxInData, endExclusiveRowIdxInData)
                       .map(r => r.slice(0, headers.length)); // ensure same width as headers
  const rowsToExport = [headers].concat(bodyRows);

  if (rowsToExport.length <= 1) {
    throw new Error('The selected range has no rows to export.');
  }

  // 8) Write to a brand-new temporary spreadsheet to ensure export reliability
  const tmpName = 'CustomExport_' + new Date().toISOString().replace(/[:.]/g, '-');
  const tmpSS = SpreadsheetApp.create(tmpName);
  const tmpSheet = tmpSS.getSheets()[0];
  tmpSheet.setName('Export');

  // Resize and write
  tmpSheet.clear();
  if (tmpSheet.getMaxRows() < rowsToExport.length) {
    tmpSheet.insertRowsAfter(1, rowsToExport.length - tmpSheet.getMaxRows());
  }
  if (tmpSheet.getMaxColumns() < headers.length) {
    tmpSheet.insertColumnsAfter(1, headers.length - tmpSheet.getMaxColumns());
  }
  tmpSheet.getRange(1, 1, rowsToExport.length, headers.length).setValues(rowsToExport);

  // Optional: make header bold
  tmpSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  SpreadsheetApp.flush(); // ensure all writes are committed before export

  // 9) Export temp spreadsheet as XLSX
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tmpSS.getId()
                  + '/export?format=xlsx&gid=' + tmpSheet.getSheetId();
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + token } });
  const blob = resp.getBlob().setName('Expenditure-' + isoStart + '_to_' + isoEnd + '.xlsx');

  // 10) Save file to Drive and share
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 11) Clean up the temporary spreadsheet
  DriveApp.getFileById(tmpSS.getId()).setTrashed(true);

  return file.getUrl();
}

/**
 * Normalize date strings to ISO yyyy-MM-dd.
 * Accepts either yyyy-MM-dd or dd/MM/yyyy.
 * Returns "" if invalid.
 */
function normalizeToIso(input) {
  if (!input) return '';
  const s = String(input).trim();

  // yyyy-MM-dd
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const [_, y, mm, dd] = m;
    return isValidYMD(y, mm, dd) ? `${y}-${mm}-${dd}` : '';
  }

  // dd/MM/yyyy
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    let [_, d, mth, y] = m;
    d = d.padStart(2, '0');
    mth = mth.padStart(2, '0');
    return isValidYMD(y, mth, d) ? `${y}-${mth}-${d}` : '';
  }

  return '';
}

/** Basic Y-M-D validity check */
function isValidYMD(y, m, d) {
  const yy = +y, mm = +m, dd = +d;
  if (!yy || mm < 1 || mm > 12 || dd < 1 || dd > 31) return false;
  const dt = new Date(`${y}-${m}-${d}T00:00:00Z`);
  return dt instanceof Date && !isNaN(dt) &&
         dt.getUTCFullYear() === yy &&
         dt.getUTCMonth() + 1 === mm &&
         dt.getUTCDate() === dd;
}







// --- Income functions for code.gs ---
// Paste below this section into your existing code.gs

const INCOME_TAB = 'Income';

// Submits new income entry, dynamic fields allowed, Amount calculated and stored
function submitIncomeForm(formData) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(INCOME_TAB);
  if (!sheet) return { success: false, message: "Income sheet not found!" };
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!formData.date || !formData.particulars) {
    return { success: false, message: "Date and Particulars are required." };
  }

  // Parse dynamic Other item details
  let dynamicOtherDetails = [];
  if (formData.otherDetails) {
    try {
      dynamicOtherDetails = JSON.parse(formData.otherDetails); // array of {title, value}
    } catch (e) {}
  }

  // Add any custom/dynamic fields as new columns if needed (case-insensitive)
  let headerLC = headers.map(h => h.toLowerCase());
  let newCols = [];
  dynamicOtherDetails.forEach(obj => {
    if (obj && obj.title && headerLC.indexOf(String(obj.title).toLowerCase()) === -1) {
      headers.push(obj.title);
      newCols.push(obj.title);
    }
  });
  if (newCols.length > 0) {
    sheet.insertColumnsAfter(headers.length - newCols.length, newCols.length);
    sheet.getRange(1, headers.length - newCols.length + 1, 1, newCols.length).setValues([newCols]);
  }
  // Re-fetch headers (after alteration)
  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Key map (excluding "AMOUNT" from below as it is summed up separately)
  const headerMap = {
    "DATE": "date",
    "PARTICULARS": "particulars",
    "CASH": "cash",
    "MPESA": "mpesa",
    "BANK ACCOUNT": "bank_account",
    "WELFARE": "welfare",
    "HATUA": "hatua",
    "S.D.C": "sdc",
    "CREATIVE AND ARTS": "creative_and_arts",
    "ASSETS": "assets",
    "ICT DOCKET": "ict_docket",
    "CAPTURE MOMENT": "capture_moment",        
    "TRANSPORT": "transport",
    "BIBLE STUDY": "bible_study",
    "REFUND": "refund"
  };

  let row = [];
  let amountSum = 0;
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (header === "AMOUNT") {
      row.push(""); 
      continue;
    }
    const stdKey = headerMap[header];
    if (stdKey) {
      if (stdKey === "date" || stdKey === "particulars") {
        row.push(formData[stdKey] || "");
      } else {
        let val = Number(formData[stdKey]);
        if (!isNaN(val) && val !== 0 && formData[stdKey] !== "" && formData[stdKey] != null) {
          row.push(val);
          amountSum += val;
        } else {
          row.push("");
        }
      }
    } else {
      // check for value from dynamicOtherDetails
      let idx = dynamicOtherDetails.findIndex(obj => String(obj.title).toLowerCase() === String(header).toLowerCase());
      if (idx !== -1) {
        let val = Number(dynamicOtherDetails[idx].value);
        row.push(isNaN(val) ? "" : val);
        if (!isNaN(val)) amountSum += val;
      } else {
        row.push("");
      }
    }
  }
  // Insert calculated amount as correct column (AMOUNT is always col 3 if default)
  let amountIndex = headers.findIndex(h => h === "AMOUNT");
  if (amountIndex === -1) amountIndex = 2; // fallback
  row[amountIndex] = amountSum;

  sheet.appendRow(row);
  return { success: true, message: "Successfully submitted! Total: " + amountSum.toFixed(2) };
}

// Download the entire "Income" tab as Excel
function downloadIncomeExcel() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(INCOME_TAB);
  var url = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/export?'
    + 'format=xlsx'
    + '&gid=' + sheet.getSheetId();

  var token = ScriptApp.getOAuthToken();
  var options = {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  };
  var response = UrlFetchApp.fetch(url, options);
  var excelBlob = response.getBlob().setName(INCOME_TAB + '.xlsx');
  var file = DriveApp.createFile(excelBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// Download filtered timeline export for Income tab
function downloadIncomeExcelForDuration(userStart, userEnd) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(INCOME_TAB);
  if (!sheet) throw new Error("Sheet not found: " + INCOME_TAB);

  // Normalize incoming dates to ISO yyyy-MM-dd
  const isoStart = normalizeToIso(userStart);
  const isoEnd   = normalizeToIso(userEnd);
  if (!isoStart || !isoEnd) throw new Error('Invalid date(s). Use yyyy-MM-dd or dd/MM/yyyy.');
  if (isoStart > isoEnd)   throw new Error('Start date is after end date.');

  // Read display values for row matching
  const range = sheet.getDataRange();
  const data = range.getValues();           // includes header row
  const display = range.getDisplayValues(); // what you see in UI (dates like 2024-11-18)
  if (data.length < 2) throw new Error('No data in sheet.');

  const headers = data[0];
  // Build anchor rows (DATE in col A, yyyy-MM-dd)
  const anchors = [];
  for (let r = 1; r < display.length; r++) {
    const shown = (display[r][0] || '').toString().trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(shown)) {
      anchors.push({ rowIdx: r, iso: shown });
    }
  }
  if (anchors.length === 0) throw new Error('No date anchors found in column A (expected yyyy-MM-dd).');
  anchors.sort((a, b) => (a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : a.rowIdx - b.rowIdx));

  // Start: first anchor with date >= start
  let startAnchorIdx = -1;
  for (let i = 0; i < anchors.length; i++) {
    if (anchors[i].iso >= isoStart) {
      startAnchorIdx = i;
      break;
    }
  }
  if (startAnchorIdx === -1) throw new Error('No data on or after the selected start date (' + isoStart + ').');
  // End: last anchor with date <= end
  let endAnchorIdx = -1;
  for (let i = 0; i < anchors.length; i++) {
    if (anchors[i].iso <= isoEnd) endAnchorIdx = i;
    else break;
  }
  if (endAnchorIdx === -1) throw new Error('No data on or before the selected end date (' + isoEnd + ').');
  if (startAnchorIdx > endAnchorIdx) throw new Error('No records in the selected date range.');

  // Rows: from start anchor thru entire date block of end anchor
  const startRowIdxInData = anchors[startAnchorIdx].rowIdx;
  const endExclusiveRowIdxInData = (endAnchorIdx < anchors.length - 1)
    ? anchors[endAnchorIdx + 1].rowIdx
    : data.length;
  const bodyRows = data.slice(startRowIdxInData, endExclusiveRowIdxInData).map(r => r.slice(0, headers.length));
  const rowsToExport = [headers].concat(bodyRows);
  if (rowsToExport.length <= 1) throw new Error('No rows to export in selection.');

  // Write to new temp spreadsheet to avoid GID/export issues
  const tmpName = 'IncomeExport_' + new Date().toISOString().replace(/[:.]/g, '-');
  const tmpSS = SpreadsheetApp.create(tmpName);
  const tmpSheet = tmpSS.getSheets()[0];
  tmpSheet.setName('Export');
  if (tmpSheet.getMaxRows() < rowsToExport.length) {
    tmpSheet.insertRowsAfter(1, rowsToExport.length - tmpSheet.getMaxRows());
  }
  if (tmpSheet.getMaxColumns() < headers.length) {
    tmpSheet.insertColumnsAfter(1, headers.length - tmpSheet.getMaxColumns());
  }
  tmpSheet.getRange(1, 1, rowsToExport.length, headers.length).setValues(rowsToExport);
  tmpSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  SpreadsheetApp.flush();

  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tmpSS.getId()
                  + '/export?format=xlsx&gid=' + tmpSheet.getSheetId();
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + token } });
  const blob = resp.getBlob().setName('Income-' + isoStart + '_to_' + isoEnd + '.xlsx');
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getFileById(tmpSS.getId()).setTrashed(true);
  return file.getUrl();
}

// Utility function (include once in file)
function normalizeToIso(input) {
  if (!input) return '';
  const s = String(input).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const [_, y, mm, dd] = m;
    return isValidYMD(y, mm, dd) ? `${y}-${mm}-${dd}` : '';
  }
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    let [_, d, mth, y] = m;
    d = d.padStart(2, '0');
    mth = mth.padStart(2, '0');
    return isValidYMD(y, mth, d) ? `${y}-${mth}-${d}` : '';
  }
  return '';
}
function isValidYMD(y, m, d) {
  const yy = +y, mm = +m, dd = +d;
  if (!yy || mm < 1 || mm > 12 || dd < 1 || dd > 31) return false;
  const dt = new Date(`${y}-${m}-${d}T00:00:00Z`);
  return dt instanceof Date && !isNaN(dt) &&
         dt.getUTCFullYear() === yy &&
         dt.getUTCMonth() + 1 === mm &&
         dt.getUTCDate() === dd;
}

function downloadIncomeExcel() {
  var sheetId = SHEET_ID;
  var ss = SpreadsheetApp.openById(sheetId);
  var url = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/export?'
    + 'format=xlsx'
    + '&gid=' + ss.getSheetByName('Income').getSheetId();

  var token = ScriptApp.getOAuthToken();
  var options = {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  };
  var response = UrlFetchApp.fetch(url, options);
  var excelBlob = response.getBlob().setName('Income.xlsx');
  var file = DriveApp.createFile(excelBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

/**
 * Income custom duration XLSX download.
 * Follows same logic as downloadExpenditureExcelForDuration for anchor block slicing.
 */
function downloadIncomeExcelForDuration(userStart, userEnd) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Income');
  if (!sheet) throw new Error('Sheet not found: Income');

  // 1) Normalize incoming dates to ISO (yyyy-MM-dd)
  const isoStart = normalizeToIso(userStart);
  const isoEnd   = normalizeToIso(userEnd);
  if (!isoStart || !isoEnd) {
    throw new Error('Invalid date format. Use dd/MM/yyyy (e.g., 20/09/2024) or yyyy-MM-dd.');
  }
  if (isoStart > isoEnd) {
    throw new Error('Start date is after end date.');
  }

  // 2) Read sheet data (values + display values)
  const range   = sheet.getDataRange();
  const data    = range.getValues();           // raw values (includes header)
  const display = range.getDisplayValues();    // what UI shows (e.g., 2024-11-18) (includes header)
  if (data.length < 2) throw new Error('No data in sheet.');

  const headers = data[0];
  const firstDataRowIdx = 1; // zero-based index; row 1 is the first data row

  // 3) Build "anchor" list: rows that start a date block (Column A shows yyyy-MM-dd)
  const anchors = [];
  for (let r = firstDataRowIdx; r < display.length; r++) {
    const shown = String(display[r][0] || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(shown)) {
      anchors.push({ rowIdx: r, iso: shown });
    }
  }
  if (anchors.length === 0) {
    throw new Error('No date anchors found in column A (expected yyyy-MM-dd).');
  }
  anchors.sort((a, b) => (a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : a.rowIdx - b.rowIdx));

  // 4) Locate start anchor (first on/after isoStart)
  let startAnchorIdx = -1;
  for (let i = 0; i < anchors.length; i++) {
    if (anchors[i].iso >= isoStart) {
      startAnchorIdx = i;
      break;
    }
  }
  if (startAnchorIdx === -1) {
    throw new Error('No data on or after the selected start date (' + isoStart + ').');
  }

  // 5) Locate end anchor (last on/before isoEnd)
  let endAnchorIdx = -1;
  for (let i = 0; i < anchors.length; i++) {
    if (anchors[i].iso <= isoEnd) endAnchorIdx = i;
    else break;
  }
  if (endAnchorIdx === -1) {
    throw new Error('No data on or before the selected end date (' + isoEnd + ').');
  }
  if (startAnchorIdx > endAnchorIdx) {
    throw new Error('No records in the selected date range (' + isoStart + ' to ' + isoEnd + ').');
  }

  // 6) Compute slice bounds in data array
  const startRowIdxInData = anchors[startAnchorIdx].rowIdx;
  const endExclusiveRowIdxInData = (endAnchorIdx < anchors.length - 1)
    ? anchors[endAnchorIdx + 1].rowIdx
    : data.length;

  // 7) Build rows to export: headers + selected body rows
  const bodyRows = data.slice(startRowIdxInData, endExclusiveRowIdxInData)
                       .map(r => r.slice(0, headers.length)); // ensure same width as headers
  const rowsToExport = [headers].concat(bodyRows);

  if (rowsToExport.length <= 1) {
    throw new Error('The selected range has no rows to export.');
  }

  // 8) Temporary spreadsheet for export reliability
  const tmpName = 'IncomeExport_' + new Date().toISOString().replace(/[:.]/g, '-');
  const tmpSS = SpreadsheetApp.create(tmpName);
  const tmpSheet = tmpSS.getSheets()[0];
  tmpSheet.setName('Export');

  // Resize and write
  tmpSheet.clear();
  if (tmpSheet.getMaxRows() < rowsToExport.length) {
    tmpSheet.insertRowsAfter(1, rowsToExport.length - tmpSheet.getMaxRows());
  }
  if (tmpSheet.getMaxColumns() < headers.length) {
    tmpSheet.insertColumnsAfter(1, headers.length - tmpSheet.getMaxColumns());
  }
  tmpSheet.getRange(1, 1, rowsToExport.length, headers.length).setValues(rowsToExport);
  tmpSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  SpreadsheetApp.flush();

  // 9) Export temp spreadsheet as XLSX
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tmpSS.getId()
                  + '/export?format=xlsx&gid=' + tmpSheet.getSheetId();
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + token } });
  const blob = resp.getBlob().setName('Income-' + isoStart + '_to_' + isoEnd + '.xlsx');

  // 10) Save file to Drive and share
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 11) Clean up temp sheet
  DriveApp.getFileById(tmpSS.getId()).setTrashed(true);

  return file.getUrl();
}

function getIncomePage() {
  return HtmlService.createHtmlOutputFromFile('incomes').getContent();
}



const INCOME_EVAL_TAB = 'Income Evaluation';

// Submits new income evaluation entry, including dynamic fields, with totals recorded in AMOUNT column
function submitIncomeEvalForm(formData) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(INCOME_EVAL_TAB);
  if (!sheet) return { success: false, message: "Income Evaluation sheet not found!" };

  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!formData.date || !formData.service) {
    return { success: false, message: "Date and Service are required." };
  }

  let dynamicOtherDetails = [];
  if (formData.otherDetails) {
    try { dynamicOtherDetails = JSON.parse(formData.otherDetails); } catch(e) {}
  }
  let headerLC = headers.map(h => String(h).toLowerCase().trim());
  let newCols = [];
  dynamicOtherDetails.forEach(obj => {
    if (obj && obj.title && headerLC.indexOf(String(obj.title).toLowerCase().trim()) === -1) {
      headers.push(obj.title);
      newCols.push(obj.title);
    }
  });
  if (newCols.length > 0) {
    sheet.insertColumnsAfter(headers.length - newCols.length, newCols.length);
    sheet.getRange(1, headers.length - newCols.length + 1, 1, newCols.length).setValues([newCols]);
  }
  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headerLC = headers.map(h => String(h).toLowerCase().trim());

  // Normalize header for reliable comparisons
  function norm(str) { return String(str).toLowerCase().replace(/[\s\-_]/g, ""); }

  // Accept only these fields from formData for sum
  const sumFields = ["cash", "bank", "mpesa"];

  let row = [];
  let amountSum = 0;
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    const normHeader = norm(header);

    // Totals column -- fill last
    if (normHeader === "amount" || normHeader === "totals") {
      row.push(""); continue;
    }
    if (normHeader === "date" || normHeader === "service") {
      row.push(formData[normHeader] || "");
    } else if (sumFields.includes(normHeader)) {
      let val = Number(formData[normHeader]);
      if (!isNaN(val) && val !== 0 && val !== "" && val != null) {
        row.push(val);
        amountSum += val;
      } else {
        row.push("");
      }
    } else {
      let idx = dynamicOtherDetails.findIndex(obj => norm(obj.title) === normHeader);
      if (idx !== -1) {
        let val = Number(dynamicOtherDetails[idx].value);
        row.push(isNaN(val) ? "" : val);
      } else {
        row.push("");
      }
    }
  }

  // Set total amount in the correct column
  let amountIndex = headerLC.indexOf("amount");
  if (amountIndex === -1) amountIndex = headerLC.indexOf("totals");
  if (amountIndex === -1) amountIndex = 2; // fallback
  row[amountIndex] = amountSum;

  sheet.appendRow(row);

  return {
    success: true,
    message: "Successfully submitted! Total: " + amountSum.toFixed(2)
  };
}

function getIncomeEvalPage() {
  return HtmlService.createHtmlOutputFromFile('income-evaluation').getContent();
}



// --- Expenditure Evaluation functions for code.gs ---
// Paste below this section into your existing code.gs

const EXPENDITURE_EVAL_TAB = 'Expenditure Evaluation';

// Submits new expenditure evaluation entry, with Balance and Refund recorded
function submitExpenditureEvalForm(formData) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(EXPENDITURE_EVAL_TAB);
  if (!sheet) return { success: false, message: "Expenditure Evaluation sheet not found!" };
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!formData.date || !formData.particulars) {
    return { success: false, message: "Date and Particulars are required." };
  }

  // No dynamic fields for this form, but you could adapt here if needed in the future

  // Key map (excluding dynamic fields)
  const headerMap = {
    "DATE": "date",
    "PARTICULARS": "particulars",
    "BUDGETED": "budgeted",
    "USED": "used",
    "BALANCE": "balance",
    "REFUND TO ACCOUNT": "refund_to_account"
  };

  let row = [];
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    const stdKey = headerMap[header];
    if (stdKey) {
      row.push(formData[stdKey] || "");
    } else {
      row.push("");
    }
  }

  sheet.appendRow(row);
  return { success: true, message: "Successfully submitted!" };
}

// Download the entire "Expenditure Evaluation" tab as Excel
function downloadExpenditureEvalExcel() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(EXPENDITURE_EVAL_TAB);
  var url = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/export?'
    + 'format=xlsx'
    + '&gid=' + sheet.getSheetId();

  var token = ScriptApp.getOAuthToken();
  var options = {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  };
  var response = UrlFetchApp.fetch(url, options);
  var excelBlob = response.getBlob().setName(EXPENDITURE_EVAL_TAB + '.xlsx');
  var file = DriveApp.createFile(excelBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// Download filtered timeline export for Expenditure Evaluation tab
function downloadExpenditureEvalExcelForDuration(userStart, userEnd) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(EXPENDITURE_EVAL_TAB);
  if (!sheet) throw new Error("Sheet not found: " + EXPENDITURE_EVAL_TAB);

  // Normalize incoming dates to ISO yyyy-MM-dd
  const isoStart = normalizeToIso(userStart);
  const isoEnd   = normalizeToIso(userEnd);
  if (!isoStart || !isoEnd) throw new Error('Invalid date(s). Use yyyy-MM-dd or dd/MM/yyyy.');
  if (isoStart > isoEnd)   throw new Error('Start date is after end date.');

  // Read display values for row matching
  const range = sheet.getDataRange();
  const data = range.getValues();           // includes header row
  const display = range.getDisplayValues(); // what you see in UI (dates like 2024-11-18)
  if (data.length < 2) throw new Error('No data in sheet.');

  const headers = data[0];
  const anchors = [];
  for (let r = 1; r < display.length; r++) {
    const shown = (display[r][0] || '').toString().trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(shown)) {
      anchors.push({ rowIdx: r, iso: shown });
    }
  }
  if (anchors.length === 0) throw new Error('No date anchors found in column A (expected yyyy-MM-dd).');
  anchors.sort((a, b) => (a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : a.rowIdx - b.rowIdx));

  // Start: first anchor with date >= start
  let startAnchorIdx = -1;
  for (let i = 0; i < anchors.length; i++) {
    if (anchors[i].iso >= isoStart) {
      startAnchorIdx = i;
      break;
    }
  }
  if (startAnchorIdx === -1) throw new Error('No data on or after the selected start date (' + isoStart + ').');
  // End: last anchor with date <= end
  let endAnchorIdx = -1;
  for (let i = 0; i < anchors.length; i++) {
    if (anchors[i].iso <= isoEnd) endAnchorIdx = i;
    else break;
  }
  if (endAnchorIdx === -1) throw new Error('No data on or before the selected end date (' + isoEnd + ').');
  if (startAnchorIdx > endAnchorIdx) throw new Error('No records in the selected date range.');

  const startRowIdxInData = anchors[startAnchorIdx].rowIdx;
  const endExclusiveRowIdxInData = (endAnchorIdx < anchors.length - 1)
    ? anchors[endAnchorIdx + 1].rowIdx
    : data.length;
  const bodyRows = data.slice(startRowIdxInData, endExclusiveRowIdxInData).map(r => r.slice(0, headers.length));
  const rowsToExport = [headers].concat(bodyRows);
  if (rowsToExport.length <= 1) throw new Error('No rows to export in selection.');

  // Write to new temp spreadsheet to avoid GID/export issues
  const tmpName = 'ExpenditureEvalExport_' + new Date().toISOString().replace(/[:.]/g, '-');
  const tmpSS = SpreadsheetApp.create(tmpName);
  const tmpSheet = tmpSS.getSheets()[0];
  tmpSheet.setName('Export');
  if (tmpSheet.getMaxRows() < rowsToExport.length) {
    tmpSheet.insertRowsAfter(1, rowsToExport.length - tmpSheet.getMaxRows());
  }
  if (tmpSheet.getMaxColumns() < headers.length) {
    tmpSheet.insertColumnsAfter(1, headers.length - tmpSheet.getMaxColumns());
  }
  tmpSheet.getRange(1, 1, rowsToExport.length, headers.length).setValues(rowsToExport);
  tmpSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  SpreadsheetApp.flush();

  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tmpSS.getId()
                  + '/export?format=xlsx&gid=' + tmpSheet.getSheetId();
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + token } });
  const blob = resp.getBlob().setName('ExpenditureEval-' + isoStart + '_to_' + isoEnd + '.xlsx');
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getFileById(tmpSS.getId()).setTrashed(true);
  return file.getUrl();
}
function getExpenditureEvalPage() {
  return HtmlService.createHtmlOutputFromFile('expenditure-evaluation').getContent();
}







function submitDepartmentsForm(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var deptSheet = ss.getSheetByName("Departments");

  var DEPARTMENTS = [
    "Prayer Ministry","Ushering","Annual general Meeting(AGM)","Associates","Anza Fyt",
    "Fyt 2nd Years","Fyt 3rd Years","Vuka Fyt","Bible Study and Discipleship","Creative Art and Design",
    "Music Department","Ladies and Gents","Adhoc and Social Weekend","Mission","Media Department",
    "Secretary and Chairperson","Treasury","Focus","Savings"
  ];
  var deptHeaders = deptSheet.getRange(1,1,1,deptSheet.getLastColumn()).getValues()[0];
  var deptRecord = Array(deptHeaders.length).fill("");
  deptRecord[0] = data["date"];

  // For each department, sum all matching values from static and other
  DEPARTMENTS.forEach(function(dept, i) {
    let sum = 0;
    // Static fields: look for all amount fields with dropdown value == dept
    for (var key in data) {
      if (key.startsWith("category_") && data[key] === dept) {
        var fieldName = key.replace("category_", "");
        var amt = Number(data[fieldName] || 0);
        if (amt) sum += amt;
      }
    }
    // Dynamic "Other" items: sum any other row with category == dept
    if (data.otherDetails) {
      try {
        var othersArr = JSON.parse(data.otherDetails);
        othersArr.forEach(function(item) {
          if (item.category === dept && item.value) {
            sum += Number(item.value);
          }
        });
      } catch(e){}
    }
    if (sum > 0) deptRecord[i+1] = sum;
  });

  deptSheet.appendRow(deptRecord);
  return {message: "Recorded to Departments tab for " + data["date"]};
}






function submitControlBudgetForm(data) {
  var SHEET_ID = '1AU0SKyW4gtRSry-9q1DfGSFeYATTXtwVDgpotflE5fw';
  var TAB_NAME = 'Targets';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) return {success: false, message: "Targets tab not found!"};

  var HEADERS = [
  "Start Date", "End Date", "Total Christian Union Expenditure",
  "Prayer Ministry", "Ushering", "Annual general Meeting(AGM)", "Associates", "Anza Fyt",
  "Fyt 2nd Years", "Fyt 3rd Years", "Vuka Fyt", "Bible Study and Discipleship", "Creative Art and Design",
  "Music Department", "Ladies and Gents", "Adhoc and Social Weekend", "Mission", "Media Department",
  "Secretary and Chairperson", "Treasury", "Focus", "Savings"
];

  // Build row in sheet order
  var row = [];
  HEADERS.forEach(function(head, i) {
    if (i === 0) row.push(data.start_date || "");
    else if (i === 1) row.push(data.end_date || "");
    else if (i === 2) row.push(data.main_target || "");
    else {
      var key = "cat_" + head.replace(/\s+/g,'_').replace(/[()]/g,'');
      row.push(data[key] !== undefined && data[key] !== null ? data[key] : "");
    }
  });
  sheet.appendRow(row);
  return {success: true, message: "Budget targets saved!"};
}

function getControlBudgetPage() {
  return HtmlService.createHtmlOutputFromFile('control-budget').getContent();
}

// Fetches the latest non-blank start_date row and returns fields in order for the form
function fetchCurrentBudgetTargetsData() {
  var SHEET_ID = '1AU0SKyW4gtRSry-9q1DfGSFeYATTXtwVDgpotflE5fw';
  var TAB_NAME = 'Targets';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME);
  var data = sheet.getDataRange().getValues(); // First row: header
  var headers = data[0];
  var rowIdx = data.length - 1;
  // Find last row with a non-empty Start Date
  while (rowIdx > 0 && (!data[rowIdx][0] || String(data[rowIdx][0]).trim() === "")) rowIdx--;
  if (rowIdx === 0) return {success:false,message:"No targets found."};
  var row = data[rowIdx];
  // Build output for form fields, mapping column order directly
  var out = {};
  headers.forEach(function(head, i) {
    if(i === 0) out.start_date = row[i];
    else if(i === 1) out.end_date = row[i];
    else if(i === 2) out.main_target = row[i];
    else {
      var key = "cat_" + head.replace(/\s+/g,'_').replace(/[()]/g,'');
      out[key] = row[i];
    }
  });
  return {success:true,data:out};
}

// Lists all target rows for the Load Targets list
function fetchOtherBudgetTargetsList() {
  var SHEET_ID = '1AU0SKyW4gtRSry-9q1DfGSFeYATTXtwVDgpotflE5fw';
  var TAB_NAME = 'Targets';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME);
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for(var i=1; i<data.length; i++) {
    var r = data[i];
    if(r[0] && r[1]) {
      rows.push({ index: i, start_date: r[0], end_date: r[1] });
    }
  }
  return rows;
}

// Fetches a row by index, returning fields for form population
function fetchBudgetTargetsForRow(rowIdx) {
  var SHEET_ID = '1AU0SKyW4gtRSry-9q1DfGSFeYATTXtwVDgpotflE5fw';
  var TAB_NAME = 'Targets';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME);
  var data = sheet.getDataRange().getValues();
  if(rowIdx<1 || rowIdx>=data.length) return {success:false,message:"Row out of bounds"};
  var headers = data[0];
  var row = data[rowIdx];
  var out = {};
  headers.forEach(function(head, i) {
    if(i === 0) out.start_date = row[i];
    else if(i === 1) out.end_date = row[i];
    else if(i === 2) out.main_target = row[i];
    else {
      var key = "cat_" + head.replace(/\s+/g,'_').replace(/[()]/g,'');
      out[key] = row[i];
    }
  });
  return {success:true,data:out};
}





function getDepartmentAnalysisData() {
  var SHEET_ID = '1AU0SKyW4gtRSry-9q1DfGSFeYATTXtwVDgpotflE5fw';
  var TARGETS_TAB = 'Targets';
  var ANALYSIS_TAB = 'Analysis';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var targetsSheet = ss.getSheetByName(TARGETS_TAB);
  var analysisSheet = ss.getSheetByName(ANALYSIS_TAB);
  if (!targetsSheet || !analysisSheet) return { error: "Missing tab(s)" };

  var targetsData = targetsSheet.getDataRange().getValues();
  var analysisData = analysisSheet.getDataRange().getValues();
  var headers = targetsData[0]; // Row 1: headers
  var targetsRow = targetsData[1];
  var analysisRow = analysisData[1];

  // Make sure this list matches the column names in your spreadsheet
  var departmentsList = [
    "Total Christian Union Expenditure",
    "Prayer Ministry","Ushering","Annual general Meeting(AGM)","Associates","Anza Fyt",
    "Fyt 2nd Years","Fyt 3rd Years","Vuka Fyt","Bible Study and Discipleship","Creative Art and Design",
    "Music Department","Ladies and Gents","Adhoc and Social Weekend","Mission","Media Department",
    "Secretary and Chairperson","Treasury","Focus","Savings"
  ];

  function percent(spent, target) {
    spent = Number(spent);
    target = Number(target);
    return target > 0 ? ((spent / target) * 100).toFixed(2) : "0";
  }

  var departments = [];
  departmentsList.forEach(function(deptName){
    // Find column index in header, handle missing columns gracefully
    var colIdx = headers.indexOf(deptName);
    var target = colIdx > -1 ? Number(targetsRow[colIdx]) || 0 : 0;
    var spent = colIdx > -1 ? Number(analysisRow[colIdx]) || 0 : 0;
    departments.push({
      name: deptName,
      target: target,
      spent: spent,
      percentage: percent(spent, target)
    });
  });

  return { departments: departments };
}

function getAnalysisPage() {
  return HtmlService.createHtmlOutputFromFile('analysis').getContent();
}



function downloadWholeSpreadsheetExcel() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  const blob = response.getBlob().setName(ss.getName() + '.xlsx');
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl(); // THIS IS A GOOGLE DRIVE LINK, always prompts for download!
}