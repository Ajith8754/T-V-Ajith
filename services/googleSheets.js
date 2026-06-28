// ============================================================
// googleSheets.js — Google Sheets Integration Service
// Reads and writes data to your team's Google Sheet.
// Uses Google Service Account for authentication.
//
// Features:
//  1. readFromGoogleSheet()    — reads ALL sheets in the spreadsheet
//  2. writeToGoogleSheet()     — appends a new row to Sheet1
//  3. deleteFromGoogleSheet()  — deletes a row by report_number
// ============================================================

const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

// -------------------------------------------------------
// Month Name → Numeric string mapping
// -------------------------------------------------------
const monthMap = {
  jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
  jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12'
};

// -------------------------------------------------------
// Standardize Date Strings to YYYY-MM-DD
// Handles YYYY-MM-DD, DD-MM-YYYY, and DD-MMM-YY formats
// -------------------------------------------------------
function parseDateStringToYYYYMMDD(str) {
  if (!str) return null;
  const s = String(str).trim();

  // 1. Matches YYYY-MM-DD or YYYY/MM/DD
  if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(s)) {
    return s.replace(/\//g, '-');
  }

  // 2. Matches DD-MM-YYYY or DD/MM/YYYY
  const matchDMY = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (matchDMY) {
    const [_, d, m, y] = matchDMY;
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }

  // 3. Matches DD-MMM-YY or DD/MMM/YY (e.g. 29-Jan-25)
  const matchDMMMYY = s.match(/^(\d{1,2})[-/]([A-Za-z]{3})[-/](\d{2})$/);
  if (matchDMMMYY) {
    const [_, d, mName, yShort] = matchDMMMYY;
    const m = monthMap[mName.toLowerCase()];
    if (m) {
      const y = parseInt(yShort) < 50 ? `20${yShort}` : `19${yShort}`;
      return `${y}-${m}-${d.padStart(2, '0')}`;
    }
  }

  // 4. Matches DD-MMM-YYYY or DD/MMM/YYYY (e.g. 29-Jan-2025)
  const matchDMMMYYYY = s.match(/^(\d{1,2})[-/]([A-Za-z]{3})[-/](\d{4})$/);
  if (matchDMMMYYYY) {
    const [_, d, mName, y] = matchDMMMYYYY;
    const m = monthMap[mName.toLowerCase()];
    if (m) {
      return `${y}-${m}-${d.padStart(2, '0')}`;
    }
  }

  return s;
}

// -------------------------------------------------------
// Excel serial date → JavaScript Date converter
// Fix timezone shifts by formatting in local time
// -------------------------------------------------------
function excelDateToString(serial) {
  if (!serial || typeof serial !== 'number') return null;
  const excelEpoch = new Date(1899, 11, 30);
  const date = new Date(excelEpoch.getTime() + serial * 86400000);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// -------------------------------------------------------
// Get authenticated Google Sheets client
// -------------------------------------------------------
function getAuthClient() {
  let credentials;

  // Option 1: Parse from environment variable (ideal for production platforms like Render)
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try {
      credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
    } catch (e) {
      throw new Error(`Failed to parse GOOGLE_CREDENTIALS_JSON environment variable: ${e.message}`);
    }
  } else {
    // Option 2: Fallback to local credentials file path
    const credentialsPath = path.resolve(process.env.GOOGLE_CREDENTIALS_PATH || './google-credentials.json');
    if (!fs.existsSync(credentialsPath)) {
      throw new Error(
        `Google credentials file not found at: ${credentialsPath}\n` +
        'Please follow SETUP_GUIDE.md to create your Google Service Account credentials.'
      );
    }
    credentials = JSON.parse(fs.readFileSync(credentialsPath, 'utf8'));
  }

  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

// -------------------------------------------------------
// Parse a raw 2D array of rows into DB-ready objects
// Handles auto-detection of header row
// -------------------------------------------------------
function parseSheetRows(rows, sheetName = '') {
  if (!rows || rows.length === 0) return [];

  // Detect header row
  const originalHeaders = rows[0].map(h => String(h || '').trim());
  const headersUpper = originalHeaders.map(h => h.toUpperCase());
  const looksLikeHeader = headersUpper.some(h => ['SL NO', 'SL', 'TESTING GROUP', 'TEST GROUP', 'REPORT NUMBER', 'REPORT NO', 'TEST NAME'].includes(h));
  const dataRows = rows.slice(1); // Row 0 is always headers

  let colIndices = {
    sl_no: -1,
    testing_group: -1,
    test_component: -1,
    vehicle_model: -1,
    vin_number: -1,
    report_number: -1,
    test_name: -1,
    test_description: -1,
    requested_by: -1,
    test_location: -1,
    start_date: -1,
    end_date: -1,
    report_date: -1,
    test_engineer: -1,
    test_decision: -1,
    remark: -1,
    test_data: -1,
    category_first: -1,
    category_last: -1
  };

  const findIndex = (names) => headersUpper.findIndex(h => names.some(n => h === n || h.replace(/[^A-Z0-9]/g, '') === n.replace(/[^A-Z0-9]/g, '')));
  const findLastIndex = (names) => {
    for (let i = headersUpper.length - 1; i >= 0; i--) {
      const h = headersUpper[i];
      if (names.some(n => h === n || h.replace(/[^A-Z0-9]/g, '') === n.replace(/[^A-Z0-9]/g, ''))) {
        return i;
      }
    }
    return -1;
  };

  const sl = findIndex(['SL NO', 'SLNO', 'SL']);
  if (sl !== -1) colIndices.sl_no = sl;

  const group = findIndex(['TESTING GROUP', 'TEST GROUP']);
  if (group !== -1) colIndices.testing_group = group;

  const comp = findIndex(['TEST COMPONENT', 'COMPONENT']);
  if (comp !== -1) colIndices.test_component = comp;

  const model = findIndex(['VEHICLE MODEL', 'VEHICLE']);
  if (model !== -1) colIndices.vehicle_model = model;

  const vin = findIndex(['VIN', 'VIN NUMBER', 'VIN NO']);
  if (vin !== -1) colIndices.vin_number = vin;

  const report = findIndex(['REPORT NUMBER', 'REPORT NO', 'REPORT NO.']);
  if (report !== -1) colIndices.report_number = report;

  const name = findIndex(['TEST NAME', 'TESTNAME']);
  if (name !== -1) colIndices.test_name = name;

  const desc = findIndex(['TEST DESCRIPTION', 'DESCRIPTION']);
  if (desc !== -1) colIndices.test_description = desc;

  const req = findIndex(['REQUESTED BY', 'REQUESTEDBY']);
  if (req !== -1) colIndices.requested_by = req;

  const loc = findIndex(['TEST LOCATION', 'LOCATION']);
  if (loc !== -1) colIndices.test_location = loc;

  const start = findIndex(['START DATE', 'STARTDATE']);
  if (start !== -1) colIndices.start_date = start;

  const end = findIndex(['END DATE', 'ENDDATE']);
  if (end !== -1) colIndices.end_date = end;

  const repDate = findIndex(['REPORT DATE', 'REPORTDATE', 'DATE']);
  if (repDate !== -1) colIndices.report_date = repDate;

  const eng = findIndex(['TEST ENGINEER', 'TEST ENGINNER', 'ENGINEER']);
  if (eng !== -1) colIndices.test_engineer = eng;

  const dec = findIndex(['TEST DECISION', 'DECISION']);
  if (dec !== -1) colIndices.test_decision = dec;

  const rem = findIndex(['REMARK', 'REMARKS']);
  if (rem !== -1) colIndices.remark = rem;

  const dataIdx = findIndex(['TEST DATA', 'DATA']);
  if (dataIdx !== -1) colIndices.test_data = dataIdx;

  const catFirst = findIndex(['CATEGORY']);
  if (catFirst !== -1) colIndices.category_first = catFirst;

  const catLast = findLastIndex(['CATEGORY']);
  if (catLast !== -1) colIndices.category_last = catLast;

  return dataRows
    .filter(row => {
      // Filter out completely empty rows
      return row && row.some(cell => String(cell || '').trim() !== '');
    })
    .map((row, idx) => {
      const parseDateVal = (val) => {
        if (!val) return null;
        const s = String(val).trim();
        if (/^\d+(\.\d+)?$/.test(s)) {
          return excelDateToString(parseFloat(s));
        }
        return parseDateStringToYYYYMMDD(s);
      };

      const getVal = (field) => {
        const index = colIndices[field];
        return index !== undefined && index !== -1 ? (row[index] || '') : '';
      };

      // 1. Build raw key-value object
      const rawObj = {};
      originalHeaders.forEach((header, colIdx) => {
        if (header) {
          let val = row[colIdx];
          if (val === undefined || val === null) val = '';
          if (typeof val === 'number' && header.toLowerCase().includes('date') && val > 20000 && val < 60000) {
            val = excelDateToString(val) || val;
          } else if (val instanceof Date) {
            val = val.toISOString().split('T')[0];
          }
          rawObj[header] = val;
        }
      });

      // 2. Fallbacks for report_number & sl_no if they are missing/empty
      let reportNum = String(getVal('report_number')).trim();
      if (!reportNum) {
        reportNum = `SR-${sheetName.replace(/[^A-Za-z0-9]/g, '-')}-${idx + 1}`;
      }

      let slNo = parseInt(getVal('sl_no')) || (idx + 1);

      const categoryLastVal = String(getVal('category_last')).trim().toUpperCase();
      const categoryFirstVal = String(getVal('category_first')).trim().toUpperCase();
      const reportNumVal = String(getVal('report_number')).trim().toUpperCase();

      // Determine category (sub-category)
      let categoryVal = categoryLastVal && categoryLastVal !== 'ALL' ? categoryLastVal : (categoryFirstVal || 'ALL');

      let testingGroupVal = String(getVal('testing_group')).trim().toUpperCase();

      if (!testingGroupVal || testingGroupVal === 'ALL' || testingGroupVal === '') {
        // Primary division comes from Category 1 (Col 5), with fallback to report number prefix
        let primaryCat = categoryFirstVal && categoryFirstVal !== 'ALL' ? categoryFirstVal : '';
        if (!primaryCat) {
          if (reportNumVal.includes('/DUR/')) {
            primaryCat = 'DURABILITY';
          } else if (reportNumVal.includes('/REL/') || reportNumVal.includes('/RLD/')) {
            primaryCat = 'RELIABILITY';
          } else if (reportNumVal.includes('/ORD/')) {
            primaryCat = 'ORD';
          } else if (reportNumVal.includes('/PERF/') || reportNumVal.includes('/VEPER/') || reportNumVal.includes('/MOT/')) {
            primaryCat = 'PERFORMANCE';
          } else if (reportNumVal.includes('/NVH/')) {
            primaryCat = 'NVH';
          }
        }

        if (primaryCat === 'RLDA' || primaryCat === 'RELIABILITY') {
          testingGroupVal = 'RELIABILITY';
        } else if (primaryCat === 'PERFORMANCE' || primaryCat === 'MOTOR PERFORMANCE') {
          testingGroupVal = 'MOTOR PERFORMANCE';
        } else if (primaryCat === 'DURABILITY') {
          testingGroupVal = 'DURABILITY';
        } else if (primaryCat === 'ORD') {
          testingGroupVal = 'ORD';
        } else if (primaryCat === 'NVH') {
          testingGroupVal = 'NVH';
        } else {
          testingGroupVal = primaryCat || 'ALL';
        }
      }

      return {
        sl_no: slNo,
        testing_group: testingGroupVal,
        test_component: String(getVal('test_component')).trim(),
        vehicle_model: String(getVal('vehicle_model')).trim().toUpperCase() || 'ALL',
        vin_number: String(getVal('vin_number')).trim(),
        report_number: reportNum,
        test_name: String(getVal('test_name')).trim(),
        test_description: String(getVal('test_description')).trim(),
        requested_by: String(getVal('requested_by')).trim(),
        test_location: String(getVal('test_location')).trim().toUpperCase(),
        start_date: parseDateVal(getVal('start_date')),
        end_date: parseDateVal(getVal('end_date')),
        report_date: parseDateVal(getVal('report_date')),
        test_engineer: String(getVal('test_engineer')).trim().toUpperCase(),
        test_decision: String(getVal('test_decision')).trim().toUpperCase() || 'NO DECISION',
        remark: String(getVal('remark')).trim(),
        test_data: String(getVal('test_data')).trim(),
        category: categoryVal || 'ALL',
        raw_data: JSON.stringify(rawObj),
        source: `google_sheets:${sheetName}`,
      };
    });
}

// -------------------------------------------------------
// READ — reads ALL sheet tabs in the spreadsheet
// Fix 2: supports Sheet1, Sheet2, Sheet3 etc. automatically
// Handles rename tracking via persistent sheetId mappings
// -------------------------------------------------------
// -------------------------------------------------------
// SYNC METADATA — fetches active sheets structure from Google Sheets
// updates sheet_mappings.json and performs database renames/deletions cleanups
// -------------------------------------------------------
async function syncSheetMetadata() {
  const isConfigured = process.env.GOOGLE_SHEET_ID && 
                       process.env.GOOGLE_SHEET_ID !== 'your_google_sheet_id_here';

  if (!isConfigured) {
    return [];
  }

  const auth = getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = process.env.GOOGLE_SHEET_ID;

  // Step 1: Get all tab names and internal sheet IDs in the spreadsheet
  const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  const sheetsMetadata = meta.data.sheets.map(s => ({
    title: s.properties.title,
    sheetId: String(s.properties.sheetId)
  }));
  const allTabs = sheetsMetadata.map(s => s.title);

  // Step 1.5: Track and handle sheet renaming/deletion using sheet_mappings.json
  const mappingPath = path.resolve(__dirname, '..', 'sheet_mappings.json');
  let oldMappings = {};
  if (fs.existsSync(mappingPath)) {
    try {
      oldMappings = JSON.parse(fs.readFileSync(mappingPath, 'utf8'));
    } catch (e) {
      console.error('Failed to parse sheet_mappings.json:', e);
    }
  }

  // Lazy load TestReport model to prevent circular dependency
  const { TestReport } = require('../db');
  const { Op } = require('sequelize');

  const newMappings = {};
  for (const tab of sheetsMetadata) {
    const { title, sheetId: currentSheetId } = tab;
    const oldTitle = oldMappings[currentSheetId];

    if (oldTitle && oldTitle !== title) {
      const oldSource = `google_sheets:${oldTitle}`;
      const newSource = `google_sheets:${title}`;
      try {
        // Renamed tab: Update source column in SQLite database
        await TestReport.update(
          { source: newSource },
          { where: { source: oldSource } }
        );
        console.log(`🔄 Google Sheet Tab renamed: "${oldTitle}" -> "${title}". Updated all database records in-place.`);
      } catch (dbErr) {
        console.error(`⚠️ Failed to update source in DB for renamed tab: ${dbErr.message}`);
      }
    }
    newMappings[currentSheetId] = title;
  }

  // Handle deleted tabs: Clean up database records of sheets that no longer exist
  for (const oldSheetId of Object.keys(oldMappings)) {
    if (!newMappings[oldSheetId]) {
      const deletedTitle = oldMappings[oldSheetId];
      const deletedSource = `google_sheets:${deletedTitle}`;
      try {
        await TestReport.destroy({
          where: { source: deletedSource }
        });
        console.log(`🗑️ Google Sheet Tab deleted: "${deletedTitle}". Cleaned up corresponding database records.`);
      } catch (dbErr) {
        console.error(`⚠️ Failed to clean up DB records for deleted tab "${deletedTitle}": ${dbErr.message}`);
      }
    }
  }

  // Robust fallback cleanup: remove any database records for google_sheets sources not currently present in the spreadsheet
  const validSources = allTabs.map(tab => `google_sheets:${tab}`);
  validSources.push('google_sheets:upload data'); // Keep upload data source safe
  try {
    const deletedCount = await TestReport.destroy({
      where: {
        source: {
          [Op.like]: 'google_sheets:%',
          [Op.notIn]: validSources
        }
      }
    });
    if (deletedCount > 0) {
      console.log(`🗑️ Cleaned up ${deletedCount} orphaned database records for deleted/inactive Google Sheet tabs.`);
    }
  } catch (dbErr) {
    console.error(`⚠️ Failed to clean up orphaned DB records: ${dbErr.message}`);
  }

  // Save new mappings
  try {
    fs.writeFileSync(mappingPath, JSON.stringify(newMappings, null, 2), 'utf8');
  } catch (fsErr) {
    console.error('Failed to write sheet_mappings.json:', fsErr.message);
  }

  return allTabs;
}

// -------------------------------------------------------
// READ — reads ALL sheet tabs in the spreadsheet
// Fix 2: supports Sheet1, Sheet2, Sheet3 etc. automatically
// Handles rename tracking via persistent sheetId mappings
// -------------------------------------------------------
async function readFromGoogleSheet() {
  const allTabs = await syncSheetMetadata();
  const auth = getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = process.env.GOOGLE_SHEET_ID;

  // Step 2: Read every tab and combine all rows
  let allRecords = [];

  for (const tabName of allTabs) {
    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range: `${tabName}!A:V`,
      });
      const rows = response.data.values || [];
      const records = parseSheetRows(rows, tabName);
      if (records.length > 0) {
        console.log(`  ↳ ${tabName}: ${records.length} rows`);
        allRecords = allRecords.concat(records);
      }
    } catch (err) {
      console.warn(`  ⚠️  Could not read sheet "${tabName}": ${err.message}`);
    }
  }

  return allRecords;
}

// -------------------------------------------------------
// WRITE — append a new row to the 'upload data' sheet tab
// -------------------------------------------------------
async function writeToGoogleSheet(rowData) {
  const auth = getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = process.env.GOOGLE_SHEET_ID;
  const sheetName = 'upload data';

  // Check if 'upload data' tab exists in spreadsheet
  const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  const allTabs = meta.data.sheets.map(s => s.properties.title);
  if (!allTabs.includes(sheetName)) {
    console.log(`Tab "${sheetName}" not found. Creating it...`);
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: sheetName,
              },
            },
          },
        ],
      },
    });
    console.log(`✅ Tab "${sheetName}" created in Google Sheets`);
  }

  const rowArray = [
    rowData.sl_no || '',
    rowData.testing_group || '',
    rowData.test_component || '',
    rowData.vehicle_model || '',
    rowData.vin_number || '',
    rowData.report_number || '',
    rowData.test_name || '',
    rowData.test_description || '',
    rowData.requested_by || '',
    rowData.test_location || '',
    rowData.start_date || '',
    rowData.end_date || '',
    rowData.report_date || '',
    rowData.test_engineer || '',
    rowData.test_decision || '',
    rowData.remark || '',
    rowData.test_data || '',
  ];

  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: `${sheetName}!A:Q`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [rowArray] },
  });

  console.log(`✅ Row appended to Google Sheets (${sheetName}): ${rowData.report_number}`);
}

// -------------------------------------------------------
// DELETE — removes a row from Google Sheets by report_number
// Fix 1: called when user deletes a record from the website
// Searches ALL sheets to find and delete the matching row
// -------------------------------------------------------
async function deleteFromGoogleSheet(reportNumber, targetSheetName = null) {
  if (!reportNumber) return;

  const auth = getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = process.env.GOOGLE_SHEET_ID;

  let allTabs = [];
  if (targetSheetName) {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
    const found = meta.data.sheets.find(s => s.properties.title === targetSheetName);
    if (found) {
      allTabs = [{
        title: found.properties.title,
        sheetId: found.properties.sheetId
      }];
    }
  } else {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
    allTabs = meta.data.sheets.map(s => ({
      title: s.properties.title,
      sheetId: s.properties.sheetId,
    }));
  }

  let deletedAny = false;
  for (const tab of allTabs) {
    try {
      // Read all rows from this tab
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range: `${tab.title}!A:V`,
      });

      const rows = response.data.values || [];

      // Find the 0-based row index where col F (index 5) matches the report number
      const rowIndex = rows.findIndex(
        row => String(row[5] || '').trim() === String(reportNumber).trim()
      );

      if (rowIndex === -1) continue; // Not in this sheet, try next

      // Delete that row using the Sheets API batchUpdate
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: {
          requests: [{
            deleteDimension: {
              range: {
                sheetId: tab.sheetId,
                dimension: 'ROWS',
                startIndex: rowIndex,      // 0-based, inclusive
                endIndex: rowIndex + 1,    // exclusive
              },
            },
          }],
        },
      });

      console.log(`🗑️  Deleted row ${rowIndex + 1} from sheet "${tab.title}" (report: ${reportNumber})`);
      deletedAny = true;
    } catch (err) {
      console.warn(`⚠️  Error searching/deleting in "${tab.title}": ${err.message}`);
    }
  }

  if (!deletedAny) {
    console.warn(`⚠️  Report "${reportNumber}" not found in any sheet — nothing deleted from Sheets`);
  }
}

// -------------------------------------------------------
// UPDATE — updates a row in Google Sheets by old report_number
// Searches ALL sheets to find and update the matching row
// -------------------------------------------------------
async function updateInGoogleSheet(oldReportNumber, updatedRowData) {
  if (!oldReportNumber) return;

  const auth = getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = process.env.GOOGLE_SHEET_ID;

  // Get all tab names + their internal sheet IDs
  const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  const allTabs = meta.data.sheets.map(s => ({
    title: s.properties.title,
    sheetId: s.properties.sheetId,
  }));

  const rowArray = [
    updatedRowData.sl_no || '',
    updatedRowData.testing_group || '',
    updatedRowData.test_component || '',
    updatedRowData.vehicle_model || '',
    updatedRowData.vin_number || '',
    updatedRowData.report_number || '',
    updatedRowData.test_name || '',
    updatedRowData.test_description || '',
    updatedRowData.requested_by || '',
    updatedRowData.test_location || '',
    updatedRowData.start_date || '',
    updatedRowData.end_date || '',
    updatedRowData.report_date || '',
    updatedRowData.test_engineer || '',
    updatedRowData.test_decision || '',
    updatedRowData.remark || '',
    updatedRowData.test_data || '',
  ];

  let updatedAny = false;
  for (const tab of allTabs) {
    try {
      // Read all rows from this tab
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range: `${tab.title}!A:Q`,
      });

      const rows = response.data.values || [];

      // Find the 0-based row index where col F (index 5) matches the old report number
      const rowIndex = rows.findIndex(
        row => String(row[5] || '').trim() === String(oldReportNumber).trim()
      );

      if (rowIndex === -1) continue; // Not in this sheet, try next

      // Update that row using sheets.spreadsheets.values.update
      await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: `${tab.title}!A${rowIndex + 1}:V${rowIndex + 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [rowArray] },
      });

      console.log(`📝 Updated row ${rowIndex + 1} in sheet "${tab.title}" (report: ${oldReportNumber})`);
      updatedAny = true;
    } catch (err) {
      console.warn(`⚠️  Error searching/updating in "${tab.title}": ${err.message}`);
    }
  }

  if (!updatedAny) {
    console.warn(`⚠️  Report "${oldReportNumber}" not found in any sheet — appending as new row instead`);
    // If not found, append to primary sheet
    try {
      await writeToGoogleSheet(updatedRowData);
    } catch (err) {
      console.warn(`⚠️  Failed to append updated report to Google Sheets: ${err.message}`);
    }
  }
}

// -------------------------------------------------------
// CLEAR UPLOAD DATA — clears the 'upload data' sheet tab in Google Sheets
// -------------------------------------------------------
async function clearUploadDataInGoogleSheet() {
  const auth = getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = process.env.GOOGLE_SHEET_ID;
  const sheetName = 'upload data';

  try {
    await sheets.spreadsheets.values.clear({
      spreadsheetId: sheetId,
      range: `${sheetName}!A:Q`,
    });
    console.log(`🗑️  Cleared Google Sheets tab "${sheetName}"`);
  } catch (err) {
    console.warn(`⚠️  Could not clear Google Sheets tab "${sheetName}":`, err.message);
  }
}

// -------------------------------------------------------
// WRITE MULTIPLE — clears and writes multiple rows to the 'upload data' tab
// -------------------------------------------------------
async function writeMultipleToGoogleSheet(records) {
  const auth = getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = process.env.GOOGLE_SHEET_ID;
  const sheetName = 'upload data';

  if (!records || records.length === 0) return;

  const valueRange = records.map(rowData => [
    rowData.sl_no || '',
    rowData.testing_group || '',
    rowData.test_component || '',
    rowData.vehicle_model || '',
    rowData.vin_number || '',
    rowData.report_number || '',
    rowData.test_name || '',
    rowData.test_description || '',
    rowData.requested_by || '',
    rowData.test_location || '',
    rowData.start_date || '',
    rowData.end_date || '',
    rowData.report_date || '',
    rowData.test_engineer || '',
    rowData.test_decision || '',
    rowData.remark || '',
    rowData.test_data || '',
  ]);

  // Headers matching the structure
  const headers = [
    'SL NO', 'Test Group', 'Test Component', 'Vehicle Model', 'VIN Number', 'Report Number',
    'Test Name', 'Test Description', 'Requested By', 'Test Location', 'Start Date',
    'End Date', 'Report Date', 'Test Engineer', 'Test Decision', 'Remark', 'Test Data'
  ];

  try {
    // Check if 'upload data' tab exists in spreadsheet
    const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
    const allTabs = meta.data.sheets.map(s => s.properties.title);
    if (!allTabs.includes(sheetName)) {
      console.log(`Tab "${sheetName}" not found. Creating it...`);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: sheetName,
                },
              },
            },
          ],
        },
      });
      console.log(`✅ Tab "${sheetName}" created in Google Sheets`);
    }

    // 1. Clear existing content in that sheet range
    await sheets.spreadsheets.values.clear({
      spreadsheetId: sheetId,
      range: `${sheetName}!A:Q`,
    });

    // 2. Write headers + data rows
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `${sheetName}!A1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [headers, ...valueRange] },
    });

    console.log(`✅ ${records.length} rows written to Google Sheets (${sheetName})`);
  } catch (err) {
    console.warn(`⚠️ Could not write multiple rows to Google Sheets tab "${sheetName}":`, err.message);
  }
}

module.exports = {
  readFromGoogleSheet,
  writeToGoogleSheet,
  writeMultipleToGoogleSheet,
  deleteFromGoogleSheet,
  updateInGoogleSheet,
  excelDateToString,
  syncSheetMetadata,
  clearUploadDataInGoogleSheet,
};

