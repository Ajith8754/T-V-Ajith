const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');
require('dotenv').config({ path: path.join(__dirname, '.env') });

function getAuthClient() {
  const credentialsPath = path.resolve(process.env.GOOGLE_CREDENTIALS_PATH || './google-credentials.json');
  if (!fs.existsSync(credentialsPath)) {
    throw new Error(`Google credentials file not found at: ${credentialsPath}`);
  }
  const credentials = JSON.parse(fs.readFileSync(credentialsPath, 'utf8'));
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function run() {
  const auth = getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = process.env.GOOGLE_SHEET_ID;

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `test!A:V`,
  });
  const rows = response.data.values || [];
  
  console.log(`Total rows: ${rows.length}`);
  const headers = rows[0];
  
  console.log("\nRows containing DUR in report number or DURABILITY in category:");
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const reportNum = String(row[2] || '');
    const cat1 = String(row[5] || '');
    const cat2 = String(row[15] || '');
    if (reportNum.includes('DUR') || cat1 === 'DURABILITY' || cat2 === 'DURABILITY') {
      console.log(`Row ${i + 1} (SL ${row[0]}):`);
      console.log(`  Report Num (Col 2): "${reportNum}"`);
      console.log(`  Category 1 (Col 5): "${cat1}"`);
      console.log(`  Category 2 (Col 15): "${cat2}"`);
      console.log(`  Report Num (Col 16): "${row[16] || ''}"`);
    }
  }
}

run().catch(console.error);
