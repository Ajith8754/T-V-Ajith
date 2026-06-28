// ============================================================
// seed.js — Seed database with your Excel data
// Run this ONCE with: node seed.js
// It imports all 146 rows from your Excel file into SQLite
// ============================================================

require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const XLSX = require('xlsx');
const path = require('path');
const { initDB, TestReport } = require('./db');

// Helper: Convert Excel serial date to YYYY-MM-DD string
function excelDateToString(serial) {
  if (!serial || typeof serial !== 'number') return null;
  const excelEpoch = new Date(1899, 11, 30);
  const date = new Date(excelEpoch.getTime() + serial * 86400000);
  return date.toISOString().split('T')[0];
}

async function seedDatabase() {
  console.log('🌱 Seeding database from Excel file...');

  // Initialize DB tables
  await initDB();

  // Read Excel file
  const xlsxPath = path.join(__dirname, '..', 'Test Reports Summary wef 2025_Master.xlsx');
  let workbook;
  try {
    workbook = XLSX.readFile(xlsxPath);
  } catch (err) {
    console.error('❌ Could not read Excel file:', err.message);
    console.log('   Make sure "Test Reports Summary wef 2025_Master.xlsx" is in the Ajith folder');
    process.exit(1);
  }

  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  const headers = rows[0].map(h => String(h || '').trim().toUpperCase());
  const findIndex = (names) => headers.findIndex(h => names.some(n => h === n || h.replace(/[^A-Z0-9]/g, '') === n.replace(/[^A-Z0-9]/g, '')));
  const findLastIndex = (names) => {
    for (let i = headers.length - 1; i >= 0; i--) {
      const h = headers[i];
      if (names.some(n => h === n || h.replace(/[^A-Z0-9]/g, '') === n.replace(/[^A-Z0-9]/g, ''))) {
        return i;
      }
    }
    return -1;
  };

  let colIndices = {
    sl_no: findIndex(['SL NO', 'SLNO', 'SL']),
    testing_group: findIndex(['TESTING GROUP', 'TEST GROUP']),
    test_component: findIndex(['TEST COMPONENT', 'COMPONENT']),
    vehicle_model: findIndex(['VEHICLE MODEL', 'VEHICLE']),
    report_number: findIndex(['REPORT NUMBER', 'REPORT NO', 'REPORT NO.']),
    test_name: findIndex(['TEST NAME', 'TESTNAME']),
    test_description: findIndex(['TEST DESCRIPTION', 'DESCRIPTION']),
    requested_by: findIndex(['REQUESTED BY', 'REQUESTEDBY']),
    test_location: findIndex(['TEST LOCATION', 'LOCATION']),
    start_date: findIndex(['START DATE', 'STARTDATE']),
    end_date: findIndex(['END DATE', 'ENDDATE']),
    report_date: findIndex(['REPORT DATE', 'REPORTDATE', 'DATE']),
    test_engineer: findIndex(['TEST ENGINEER', 'TEST ENGINNER', 'ENGINEER']),
    test_decision: findIndex(['TEST DECISION', 'DECISION']),
    remark: findIndex(['REMARK', 'REMARKS']),
    test_data: findIndex(['TEST DATA', 'DATA']),
    category_first: findIndex(['CATEGORY']),
    category_last: findLastIndex(['CATEGORY']),
    ord_report_number: findIndex(['ORD REPORT NUMBER', 'ORD REPORT NO']),
    engineers: findIndex(['ENGINEERS'])
  };

  // Defaults if not found
  if (colIndices.sl_no === -1) colIndices.sl_no = 0;
  if (colIndices.test_component === -1) colIndices.test_component = 2;
  if (colIndices.vehicle_model === -1) colIndices.vehicle_model = 3;
  if (colIndices.report_number === -1) colIndices.report_number = 4;
  if (colIndices.test_name === -1) colIndices.test_name = 5;

  let added = 0, skipped = 0, errors = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0 || !row[0]) continue;

    const getVal = (field) => {
      const idx = colIndices[field];
      return idx !== undefined && idx !== -1 ? (row[idx] || '') : '';
    };

    const categoryLastVal = String(getVal('category_last')).toString().trim().toUpperCase();
    const categoryFirstVal = String(getVal('category_first')).toString().trim().toUpperCase();
    const reportNumVal = String(getVal('report_number')).toString().trim().toUpperCase();

    // Determine category (sub-category)
    let categoryVal = categoryLastVal && categoryLastVal !== 'ALL' ? categoryLastVal : (categoryFirstVal || 'ALL');

    let testingGroupVal = colIndices.testing_group !== -1 ? String(getVal('testing_group')).toString().trim().toUpperCase() : '';

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

    const record = {
      sl_no: parseInt(getVal('sl_no')) || i,
      testing_group: testingGroupVal,
      test_component: String(getVal('test_component')).toString().trim(),
      vehicle_model: String(getVal('vehicle_model')).toString().trim().toUpperCase(),
      report_number: String(getVal('report_number')).toString().trim(),
      test_name: String(getVal('test_name')).toString().trim(),
      test_description: String(getVal('test_description')).toString().trim(),
      requested_by: String(getVal('requested_by')).toString().trim(),
      test_location: String(getVal('test_location')).toString().trim().toUpperCase(),
      start_date: excelDateToString(parseFloat(getVal('start_date'))) || null,
      end_date: excelDateToString(parseFloat(getVal('end_date'))) || null,
      report_date: excelDateToString(parseFloat(getVal('report_date'))) || null,
      test_engineer: String(getVal('test_engineer')).toString().trim().toUpperCase(),
      test_decision: String(getVal('test_decision')).toString().trim().toUpperCase(),
      test_data: String(getVal('test_data')).toString().trim(),
      remark: String(getVal('remark')).toString().trim(),
      category: categoryVal || 'ALL',
      ord_report_number: String(getVal('ord_report_number')).toString().trim(),
      engineers: String(getVal('engineers')).toString().trim().toUpperCase(),
      source: 'excel_import',
    };

    if (!record.report_number) {
      console.log(`  ⚠️  Row ${i + 1}: Missing report number, skipping`);
      skipped++;
      continue;
    }

    try {
      await TestReport.upsert(record, { conflictFields: ['report_number'] });
      added++;
      if (added % 20 === 0) console.log(`  ✅ Imported ${added} records...`);
    } catch (err) {
      skipped++;
      errors.push(`Row ${i + 1} (${record.report_number}): ${err.message}`);
    }
  }

  console.log('');
  console.log('═══════════════════════════════════');
  console.log(`✅ Seeding complete!`);
  console.log(`   Records imported: ${added}`);
  console.log(`   Records skipped:  ${skipped}`);
  if (errors.length > 0) {
    console.log(`   Errors (first 5):`);
    errors.slice(0, 5).forEach(e => console.log(`     - ${e}`));
  }
  console.log('═══════════════════════════════════');
  console.log('');
  console.log('👉 Now run: node server.js');

  process.exit(0);
}

seedDatabase().catch(err => {
  console.error('Seed failed:', err);
  process.exit(1);
});
