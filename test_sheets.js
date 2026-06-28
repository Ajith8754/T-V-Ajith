const { initDB, TestReport } = require('./db');
const { Op } = require('sequelize');

async function check() {
  await initDB();
  const records = await TestReport.findAll({
    where: {
      source: 'google_sheets:Sheet3',
      report_number: {
        [Op.notLike]: 'SR-Sheet3-%'
      }
    },
    attributes: ['id', 'sl_no', 'report_number', 'test_name', 'source'],
    limit: 10,
    raw: true
  });
  console.log('Real reports under Sheet3:');
  console.log(records);
}

check().catch(console.error);
