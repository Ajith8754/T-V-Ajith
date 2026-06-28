const { initDB, TestReport } = require('./db');
const { readFromGoogleSheet } = require('./services/googleSheets');

async function test() {
  await initDB();
  const rows = await readFromGoogleSheet();
  const row = rows[1]; // first record of Sheet6 (Lady Foot Rest)
  console.log('Upserting row:', row);
  const res = await TestReport.upsert(row, { returning: true, conflictFields: ['report_number'] });
  console.log('Result:', res);
  // Let's select it back
  const found = await TestReport.findOne({ where: { report_number: row.report_number }, raw: true });
  console.log('Found in DB:', found);
}
test().catch(console.error);
