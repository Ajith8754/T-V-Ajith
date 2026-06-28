const { TestReport } = require('./db');

async function runMigration() {
  console.log('🔄 Starting database category/testing_group migration...');
  try {
    const records = await TestReport.findAll();
    console.log(`📋 Found ${records.length} records in SQLite.`);
    let updated = 0;
    
    for (const record of records) {
      if (record.raw_data) {
        try {
          const raw = JSON.parse(record.raw_data);
          const categoryVal = String(raw.CATEGORY || '').trim().toUpperCase();
          let testingGroupVal = record.testing_group;
          
          if (!testingGroupVal || testingGroupVal === 'ALL' || testingGroupVal === '') {
            if (categoryVal) {
              if (categoryVal === 'RLDA') {
                testingGroupVal = 'RELIABILITY';
              } else if (categoryVal === 'PERFORMANCE') {
                testingGroupVal = 'MOTOR PERFORMANCE';
              } else {
                testingGroupVal = categoryVal;
              }
            } else {
              testingGroupVal = 'ALL';
            }
          }
          
          let reportDateVal = record.report_date;
          if (!reportDateVal && raw.DATE) {
            const rawDate = String(raw.DATE).trim();
            const monthMap = {
              jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
              jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12'
            };
            function parseDate(str) {
              if (!str) return null;
              if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(str)) {
                return str.replace(/\//g, '-');
              }
              const matchDMY = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
              if (matchDMY) {
                const [_, d, m, y] = matchDMY;
                return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
              }
              const matchDMMMYY = str.match(/^(\d{1,2})[-/]([A-Za-z]{3})[-/](\d{2})$/);
              if (matchDMMMYY) {
                const [_, d, mName, yShort] = matchDMMMYY;
                const m = monthMap[mName.toLowerCase()];
                if (m) {
                  const y = parseInt(yShort) < 50 ? `20${yShort}` : `19${yShort}`;
                  return `${y}-${m}-${d.padStart(2, '0')}`;
                }
              }
              return str;
            }
            reportDateVal = parseDate(rawDate);
          }

          await record.update({
            category: categoryVal || 'ALL',
            testing_group: testingGroupVal,
            report_date: reportDateVal
          });
          updated++;
        } catch (e) {
          console.error(`⚠️ Failed to parse/migrate record ID ${record.id}:`, e.message);
        }
      }
    }
    console.log(`✅ Migration complete: ${updated} records updated.`);
    process.exit(0);
  } catch (err) {
    console.error('❌ Migration failed:', err.message);
    process.exit(1);
  }
}

runMigration();
