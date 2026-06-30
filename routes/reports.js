// ============================================================
// reports.js — Test Reports API Routes
// All CRUD operations + filtering + PDF download
// ============================================================

const express = require('express');
const router = express.Router();
const { Op } = require('sequelize');
const { TestReport, SyncLog } = require('../db');
const { generatePDF } = require('../services/pdfGenerator');
const { writeToGoogleSheet, deleteFromGoogleSheet, updateInGoogleSheet, writeMultipleToGoogleSheet, syncSheetMetadata, clearUploadDataInGoogleSheet } = require('../services/googleSheets');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Month Name → Numeric string mapping
const monthMap = {
  jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
  jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12'
};

// Standardize Date Strings to YYYY-MM-DD
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

// Map raw row values to a structured record dynamically based on headers
function mapRowToRecord(row, headers, index, sourceName = 'upload') {
  const formatHeader = h => String(h || '').trim().toUpperCase();
  const formattedHeaders = headers.map(formatHeader);

  const findIndex = (names) => formattedHeaders.findIndex(h => names.some(n => h === n || h.replace(/[^A-Z0-9]/g, '') === n.replace(/[^A-Z0-9]/g, '')));
  const findLastIndex = (names) => {
    for (let i = formattedHeaders.length - 1; i >= 0; i--) {
      const h = formattedHeaders[i];
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
    vin_number: findIndex(['VIN', 'VIN NUMBER', 'VIN NO']),
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
    category_last: findLastIndex(['CATEGORY'])
  };



  const getVal = (field) => {
    const idx = colIndices[field];
    return idx !== undefined && idx !== -1 ? (row[idx] || '') : '';
  };

  const formatDate = (val) => {
    if (!val) return null;
    if (val instanceof Date) {
      const year = val.getFullYear();
      const month = String(val.getMonth() + 1).padStart(2, '0');
      const day = String(val.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    if (typeof val === 'number') {
      const d = new Date((val - 25569) * 86400 * 1000);
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    if (typeof val === 'string') {
      const s = val.trim();
      if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(s)) {
        return s.replace(/\//g, '-');
      }
      const matchDMY = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
      if (matchDMY) {
        const [_, d, m, y] = matchDMY;
        return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
      }
    }
    return parseDateStringToYYYYMMDD(val);
  };

  // Build raw key-value object using original headers
  const rawObj = {};
  headers.forEach((header, colIdx) => {
    if (header) {
      let val = row[colIdx];
      if (val === undefined || val === null) val = '';
      if (val instanceof Date) {
        val = val.toISOString().split('T')[0];
      }
      rawObj[header] = val;
    }
  });

  const sheetLabel = sourceName.includes(':') ? sourceName.split(':')[1] : 'Upload';

  let reportNum = String(getVal('report_number')).trim();
  if (!reportNum) {
    reportNum = `SR-${sheetLabel.replace(/[^A-Za-z0-9]/g, '-')}-${index + 1}`;
  }

  let slNo = parseInt(getVal('sl_no')) || (index + 1);

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
    start_date: formatDate(getVal('start_date')),
    end_date: formatDate(getVal('end_date')),
    report_date: formatDate(getVal('report_date')),
    test_engineer: String(getVal('test_engineer')).trim().toUpperCase(),
    test_decision: String(getVal('test_decision')).trim().toUpperCase() || 'NO DECISION',
    remark: String(getVal('remark')).trim(),
    test_data: String(getVal('test_data')).trim(),
    category: categoryVal || 'ALL',
    raw_data: JSON.stringify(rawObj),
    source: sourceName,
  };
}

// Multer config for file uploads (stores in memory for processing)
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

// Storage configuration for images
const imageStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.join(__dirname, '..', 'uploads'));
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    const ext = path.extname(file.originalname);
    cb(null, 'img-' + uniqueSuffix + ext);
  }
});

const imageUpload = multer({
  storage: imageStorage,
  limits: { fileSize: 5 * 1024 * 1024 }, // limit to 5MB
  fileFilter: (req, file, cb) => {
    const filetypes = /jpeg|jpg|png|gif|webp/;
    const mimetype = filetypes.test(file.mimetype);
    const extname = filetypes.test(path.extname(file.originalname).toLowerCase());
    if (mimetype && extname) {
      return cb(null, true);
    }
    cb(new Error('Only image files (jpg, jpeg, png, gif, webp) are allowed!'));
  }
});

// -------------------------------------------------------
// Helper: Build Sequelize WHERE clause from query params
// -------------------------------------------------------
function buildWhereClause(query) {
  const where = {};
  if (query.testing_group && query.testing_group !== 'ALL') {
    where.testing_group = query.testing_group;
  }
  if (query.vehicle_model && query.vehicle_model !== 'ALL') {
    where.vehicle_model = query.vehicle_model;
  }
  if (query.test_decision && query.test_decision !== 'ALL') {
    where.test_decision = query.test_decision;
  }
  if (query.test_location && query.test_location !== 'ALL') {
    where.test_location = query.test_location;
  }
  if (query.category && query.category !== 'ALL') {
    where.category = query.category;
  }
  if (query.test_engineer && query.test_engineer !== 'ALL') {
    where.test_engineer = query.test_engineer;
  }
  // Date range filter on report_date
  if (query.from_date || query.to_date) {
    where.report_date = {};
    if (query.from_date) where.report_date[Op.gte] = query.from_date;
    if (query.to_date) where.report_date[Op.lte] = query.to_date;
  }
  // Search across multiple text fields and the dynamic raw_data JSON string
  if (query.search) {
    const searchTerm = `%${query.search}%`;
    where[Op.or] = [
      { report_number: { [Op.like]: searchTerm } },
      { test_name: { [Op.like]: searchTerm } },
      { test_component: { [Op.like]: searchTerm } },
      { test_description: { [Op.like]: searchTerm } },
      { requested_by: { [Op.like]: searchTerm } },
      { testing_group: { [Op.like]: searchTerm } },
      { vehicle_model: { [Op.like]: searchTerm } },
      { vin_number: { [Op.like]: searchTerm } },
      { test_location: { [Op.like]: searchTerm } },
      { test_engineer: { [Op.like]: searchTerm } },
      { category: { [Op.like]: searchTerm } },
      { remark: { [Op.like]: searchTerm } },
      { test_data: { [Op.like]: searchTerm } },
      { raw_data: { [Op.like]: searchTerm } },
    ];
  }

  // Filter by source (dynamic sheets)
  if (query.sheet) {
    if (query.sheet === 'google_sheets:upload data') {
      const uploadTabSource = 'google_sheets:upload data';
      where[Op.or] = [
        { source: { [Op.like]: '%upload%' } },
        { source: { [Op.like]: '%manual%' } },
        { source: { [Op.like]: '%url_import%' } },
        { source: { [Op.like]: '%google_drive%' } },
        { source: { [Op.like]: `%${uploadTabSource}%` } }
      ];
    } else {
      where.source = { [Op.like]: `%${query.sheet}%` };
    }
  } else {
    const uploadTabSource = 'google_sheets:upload data';
    if (query.preview === 'true') {
      where[Op.or] = [
        { source: { [Op.like]: '%upload%' } },
        { source: { [Op.like]: '%manual%' } },
        { source: { [Op.like]: '%url_import%' } },
        { source: { [Op.like]: '%google_drive%' } },
        { source: { [Op.like]: `%${uploadTabSource}%` } }
      ];
    } else {
      where.source = {
        [Op.like]: 'google_sheets:%',
        [Op.ne]: uploadTabSource
      };
    }
  }

  return where;
}

// -------------------------------------------------------
// GET /api/reports — Get paginated list of reports
// Supports all filter and search query params
// -------------------------------------------------------
router.get('/', async (req, res) => {
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = parseInt(req.query.limit) || 50;
    const offset = (page - 1) * limit;
    const where = buildWhereClause(req.query);

    const { count, rows } = await TestReport.findAndCountAll({
      where,
      order: [['sl_no', 'ASC']],
      limit,
      offset,
    });

    res.json({
      data: rows,
      total: count,
      page,
      totalPages: Math.ceil(count / limit),
    });
  } catch (err) {
    console.error('Error fetching reports:', err);
    res.status(500).json({ error: 'Failed to fetch reports' });
  }
});

// -------------------------------------------------------
// GET /api/reports/stats — Summary statistics for dashboard
// -------------------------------------------------------
router.get('/stats', async (req, res) => {
  try {
    const where = buildWhereClause(req.query);

    const total = await TestReport.count({ where });
    const passed = await TestReport.count({ where: { ...where, test_decision: 'PASSED' } });
    const failed = await TestReport.count({ where: { ...where, test_decision: 'FAILED' } });
    const no_decision = await TestReport.count({ where: { ...where, test_decision: 'NO DECISION' } });
    const completed = await TestReport.count({ where: { ...where, test_decision: 'COMPLETED' } });

    // Vehicle model breakdown
    const vehicleStats = await TestReport.findAll({
      attributes: [
        'vehicle_model',
        'test_decision',
        [TestReport.sequelize.fn('COUNT', TestReport.sequelize.col('id')), 'count']
      ],
      where,
      group: ['vehicle_model', 'test_decision'],
      raw: true,
    });

    // Testing group breakdown
    const groupStats = await TestReport.findAll({
      attributes: [
        'testing_group',
        [TestReport.sequelize.fn('COUNT', TestReport.sequelize.col('id')), 'count'],
        [TestReport.sequelize.fn('SUM', TestReport.sequelize.literal("CASE WHEN test_decision='PASSED' THEN 1 ELSE 0 END")), 'passed'],
        [TestReport.sequelize.fn('SUM', TestReport.sequelize.literal("CASE WHEN test_decision='FAILED' THEN 1 ELSE 0 END")), 'failed'],
      ],
      where,
      group: ['testing_group'],
      raw: true,
    });

    // Last sync info
    const lastSync = await SyncLog.findOne({ order: [['synced_at', 'DESC']] });

    res.json({
      total,
      passed,
      failed,
      no_decision,
      completed,
      pass_rate: total > 0 ? ((passed / total) * 100).toFixed(1) : 0,
      vehicle_stats: vehicleStats,
      group_stats: groupStats,
      last_sync: lastSync,
    });
  } catch (err) {
    console.error('Stats error:', err);
    res.status(500).json({ error: 'Failed to fetch stats' });
  }
});

// -------------------------------------------------------
// GET /api/reports/filter-options — All unique filter values
// Used to populate filter dropdown menus
// -------------------------------------------------------
router.get('/filter-options', async (req, res) => {
  try {
    // Real-time synchronization of sheet tabs from Google Sheets metadata
    try {
      await syncSheetMetadata();
    } catch (e) {
      console.warn('⚠️ Dynamic sheet metadata sync failed:', e.message);
    }

    const where = {};
    if (req.query.sheet) {
      if (req.query.sheet === 'google_sheets:upload data') {
        const uploadTabSource = 'google_sheets:upload data';
        where[Op.or] = [
          { source: { [Op.like]: '%upload%' } },
          { source: { [Op.like]: '%manual%' } },
          { source: { [Op.like]: '%url_import%' } },
          { source: { [Op.like]: '%google_drive%' } },
          { source: { [Op.like]: `%${uploadTabSource}%` } }
        ];
      } else {
        where.source = { [Op.like]: `%${req.query.sheet}%` };
      }
    } else {
      where.source = {
        [Op.like]: 'google_sheets:%',
        [Op.ne]: 'google_sheets:upload data'
      };
    }

    const [groups, vehicles, decisions, locations, categories, engineers] = await Promise.all([
      TestReport.findAll({ where, attributes: [[TestReport.sequelize.fn('DISTINCT', TestReport.sequelize.col('testing_group')), 'testing_group']], raw: true }),
      TestReport.findAll({ where, attributes: [[TestReport.sequelize.fn('DISTINCT', TestReport.sequelize.col('vehicle_model')), 'vehicle_model']], raw: true }),
      TestReport.findAll({ where, attributes: [[TestReport.sequelize.fn('DISTINCT', TestReport.sequelize.col('test_decision')), 'test_decision']], raw: true }),
      TestReport.findAll({ where, attributes: [[TestReport.sequelize.fn('DISTINCT', TestReport.sequelize.col('test_location')), 'test_location']], raw: true }),
      TestReport.findAll({ where, attributes: [[TestReport.sequelize.fn('DISTINCT', TestReport.sequelize.col('category')), 'category']], raw: true }),
      TestReport.findAll({ where, attributes: [[TestReport.sequelize.fn('DISTINCT', TestReport.sequelize.col('test_engineer')), 'test_engineer']], raw: true }),
    ]);

    const sources = await TestReport.findAll({
      attributes: [[TestReport.sequelize.fn('DISTINCT', TestReport.sequelize.col('source')), 'source']],
      raw: true
    });

    // 1. Read Google Sheets tab list from mappings (includes empty ones)
    let googleSheetsOptions = [];
    const mappingPath = path.resolve(__dirname, '..', 'sheet_mappings.json');
    if (fs.existsSync(mappingPath)) {
      try {
        const mappings = JSON.parse(fs.readFileSync(mappingPath, 'utf8'));
        googleSheetsOptions = Object.values(mappings)
          .map(title => ({
            value: `google_sheets:${title}`,
            label: title
          }));
      } catch (e) {
        console.error('Failed to parse sheet_mappings.json in filter-options:', e);
      }
    }

    const existingGoogleValues = new Set(googleSheetsOptions.map(o => o.value));
    const sheets = [...googleSheetsOptions];

    // 2. Merge database sources (uploaded Excel tabs, manual, etc.)
    sources.forEach(r => {
      const src = r.source || '';
      if (
        src === 'upload' ||
        src === 'manual' ||
        src === 'url_import' ||
        src === 'google_drive' ||
        src.startsWith('upload:') ||
        src.startsWith('url_import:')
      ) {
        return; // Skip upload-related sources in standard dropdown
      }
      if (src.startsWith('google_sheets:')) {
        // Skip adding google_sheets sources that are not in sheet_mappings.json
        return;
      } else {
        if (src) {
          sheets.push({ value: src, label: src });
        }
      }
    });

    res.json({
      testing_groups: groups.map(r => r.testing_group).filter(Boolean),
      vehicle_models: vehicles.map(r => r.vehicle_model).filter(v => Boolean(v) && v.toUpperCase() !== 'FE FE FE'),
      test_decisions: decisions.map(r => r.test_decision).filter(d => Boolean(d) && d.toUpperCase() !== 'D D D D'),
      test_locations: locations.map(r => r.test_location).filter(Boolean),
      categories: categories.map(r => r.category).filter(Boolean),
      engineers: engineers.map(r => r.test_engineer).filter(Boolean),
      sheets: sheets,
    });
  } catch (err) {
    res.status(500).json({ error: 'Failed to fetch filter options' });
  }
});

// -------------------------------------------------------
// GET /api/reports/chart-data — Data for chart report builder
// -------------------------------------------------------
router.get('/chart-data', async (req, res) => {
  try {
    const where = buildWhereClause(req.query);
    const xAxis = req.query.x_axis || 'testing_group';
    const yAxis = req.query.y_axis || 'count';

    let data;
    if (yAxis === 'count') {
      data = await TestReport.findAll({
        attributes: [
          xAxis,
          [TestReport.sequelize.fn('COUNT', TestReport.sequelize.col('id')), 'value'],
          [TestReport.sequelize.fn('SUM', TestReport.sequelize.literal("CASE WHEN test_decision='PASSED' THEN 1 ELSE 0 END")), 'passed'],
          [TestReport.sequelize.fn('SUM', TestReport.sequelize.literal("CASE WHEN test_decision='FAILED' THEN 1 ELSE 0 END")), 'failed'],
        ],
        where,
        group: [xAxis],
        raw: true,
      });
    } else {
      data = await TestReport.findAll({
        attributes: [
          xAxis,
          yAxis,
          [TestReport.sequelize.fn('COUNT', TestReport.sequelize.col('id')), 'value'],
        ],
        where,
        group: [xAxis, yAxis],
        raw: true,
      });
    }

    res.json({ data: data || [], x_axis: xAxis, y_axis: yAxis });
  } catch (err) {
    res.status(500).json({ error: 'Failed to fetch chart data' });
  }
});

// -------------------------------------------------------
// GET /api/reports/export/pdf — Download filtered data as PDF
// -------------------------------------------------------
router.get('/export/pdf', async (req, res) => {
  try {
    const where = buildWhereClause(req.query);
    const records = await TestReport.findAll({ where, order: [['sl_no', 'ASC']], raw: true });

    const pdfBuffer = await generatePDF(records, req.query);

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="TV_Report_${new Date().toISOString().split('T')[0]}.pdf"`,
      'Content-Length': pdfBuffer.length,
    });
    res.send(pdfBuffer);
  } catch (err) {
    console.error('PDF export error:', err);
    res.status(500).json({ error: 'Failed to generate PDF' });
  }
});

// -------------------------------------------------------
// POST /api/reports — Add a single new report manually
// Also pushes the new row to Google Sheets
// -------------------------------------------------------
router.post('/', async (req, res) => {
  try {
    const data = req.body;
    if (!data.report_number) {
      return res.status(400).json({ error: 'Report number is required' });
    }

    // Save to SQLite database
    const record = await TestReport.create({ ...data, source: data.source || 'manual' });

    // Try to push to Google Sheets (don't fail if sheets not configured)
    try {
      await writeToGoogleSheet(data);
    } catch (sheetsErr) {
      console.warn('⚠️  Could not write to Google Sheets:', sheetsErr.message);
    }

    // Notify connected clients via WebSocket (attached to app)
    const io = req.app.get('io');
    if (io) {
      io.emit('data_updated', { type: 'new_record', record: record.toJSON() });
    }

    res.status(201).json({ success: true, record });
  } catch (err) {
    if (err.name === 'SequelizeUniqueConstraintError') {
      return res.status(409).json({ error: 'A report with this number already exists' });
    }
    console.error('Create report error:', err);
    res.status(500).json({ error: 'Failed to create report' });
  }
});

// -------------------------------------------------------
// POST /api/reports/upload — Bulk upload via Excel/CSV file
// -------------------------------------------------------
router.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const uploadTabSource = 'google_sheets:upload data';
    
    // Clear existing uploaded database records
    await TestReport.destroy({
      where: {
        [Op.or]: [
          { source: { [Op.in]: ['upload', 'manual', 'url_import', 'google_drive', uploadTabSource] } },
          { source: { [Op.like]: 'upload:%' } }
        ]
      }
    });

    let added = 0, skipped = 0, errors = [];
    const uploadedRecords = [];

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      if (rows.length < 2) continue; // Skip empty sheets

      const headers = rows[0];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0 || !row.some(cell => String(cell || '').trim() !== '')) continue;

        const record = mapRowToRecord(row, headers, i, `upload:${sheetName}`);

        try {
          const existing = await TestReport.findOne({ where: { report_number: record.report_number } });
          if (existing) {
            let newSource = record.source;
            if (existing.source && existing.source !== record.source) {
              const sourcesSet = new Set(existing.source.split(',').map(s => s.trim()));
              sourcesSet.add(record.source);
              newSource = Array.from(sourcesSet).join(',');
            }
            await existing.update({
              ...record,
              source: newSource
            });
          } else {
            await TestReport.create(record);
          }
          added++;
          uploadedRecords.push(record);
        } catch (e) {
          skipped++;
          errors.push(`Sheet "${sheetName}" Row ${i + 1}: ${e.message}`);
        }
      }
    }

    if (uploadedRecords.length > 0) {
      writeMultipleToGoogleSheet(uploadedRecords).catch(sheetsErr => {
        console.warn('⚠️  Could not bulk write to Google Sheets:', sheetsErr.message);
      });
    }

    // Notify connected clients
    const io = req.app.get('io');
    if (io) io.emit('data_updated', { type: 'bulk_upload', added });

    res.json({ success: true, added, skipped, errors: errors.slice(0, 10) });
  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ error: 'Failed to process uploaded file' });
  }
});

// -------------------------------------------------------
// DELETE /api/reports/clear-all — Clear all reports from database
// MUST be defined before /api/reports/:id route
// -------------------------------------------------------
router.delete('/clear-all', async (req, res) => {
  try {
    const { preview } = req.query;
    if (preview === 'true') {
      const uploadTabSource = 'google_sheets:upload data';

      // Clear Google Sheets 'upload data' tab content first
      try {
        await clearUploadDataInGoogleSheet();
      } catch (sheetsErr) {
        console.warn("⚠️ Could not clear 'upload data' tab in Google Sheets:", sheetsErr.message);
      }

      // Clear only uploaded records
      await TestReport.destroy({
        where: {
          [Op.or]: [
            { source: { [Op.like]: '%upload%' } },
            { source: { [Op.like]: '%manual%' } },
            { source: { [Op.like]: '%url_import%' } },
            { source: { [Op.like]: '%google_drive%' } },
            { source: { [Op.like]: `%${uploadTabSource}%` } }
          ]
        }
      });
    } else {
      // Delete all records from SQLite
      await TestReport.destroy({ where: {} });
      await SyncLog.destroy({ where: {} });
    }

    // Automatically trigger Google Sheets sync to reload the clean data stored in the Google Sheet
    const syncService = req.app.get('syncService');
    if (syncService) {
      await syncService();
    }

    // Notify connected clients via WebSocket
    const io = req.app.get('io');
    if (io) {
      io.emit('data_updated', { type: 'bulk_upload', added: 0 }); // Trigger lists to reload
    }

    res.json({ success: true, message: 'All reports cleared and synced successfully' });
  } catch (err) {
    console.error('Clear all reports error:', err);
    res.status(500).json({ error: 'Failed to clear database' });
  }
});

// Helper to remove upload sources from a comma-separated source string
function removeUploadSources(sourceStr) {
  if (!sourceStr) return '';
  const uploadTabSource = 'google_sheets:upload data';
  const sources = sourceStr.split(',').map(s => s.trim());
  const cleanSources = sources.filter(s => {
    const isUpload = s === 'upload' || 
                     s === 'manual' || 
                     s === 'google_drive' || 
                     s === uploadTabSource ||
                     s.startsWith('upload:') ||
                     s.startsWith('url_import:') ||
                     s.startsWith('google_sheets:upload data');
    return !isUpload;
  });
  return cleanSources.join(',');
}

// -------------------------------------------------------
// DELETE /api/reports/:id — Delete a record
// Also removes the row from Google Sheets
// -------------------------------------------------------
router.delete('/:id', async (req, res) => {
  try {
    const { preview } = req.query;

    // Fetch the record first so we have the report_number and source
    const record = await TestReport.findOne({ where: { id: req.params.id } });
    if (!record) return res.status(404).json({ error: 'Record not found' });

    if (preview === 'true') {
      // 1. Delete SPECIFICALLY from the 'upload data' tab in Google Sheets
      if (record.report_number) {
        deleteFromGoogleSheet(record.report_number, 'upload data').catch(err => {
          console.warn("⚠️  Could not delete from Google Sheets 'upload data' tab:", err.message);
        });
      }

      // 2. Remove upload-related sources in SQLite, or destroy if no other sources remain
      const cleanedSource = removeUploadSources(record.source);
      if (cleanedSource) {
        await record.update({ source: cleanedSource });
      } else {
        await record.destroy();
      }
    } else {
      // Default: Destroy SQLite record completely and delete from all Google Sheet tabs
      const reportNumber = record.report_number;
      await record.destroy();

      if (reportNumber) {
        deleteFromGoogleSheet(reportNumber).catch(err => {
          console.warn('⚠️  Could not delete from Google Sheets:', err.message);
        });
      }
    }

    res.json({ success: true });
  } catch (err) {
    console.error('Delete error:', err);
    res.status(500).json({ error: 'Failed to delete record' });
  }
});

// -------------------------------------------------------
// PUT /api/reports/:id — Update a record
// Also updates the row in Google Sheets
// -------------------------------------------------------
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const updatedData = req.body;

    // Fetch the original record so we know the old report_number
    const record = await TestReport.findOne({ where: { id } });
    if (!record) {
      return res.status(404).json({ error: 'Record not found' });
    }

    const oldReportNumber = record.report_number;

    // Update SQLite DB
    await record.update(updatedData);

    // Update in Google Sheets in background (don't fail if Sheets not set)
    if (oldReportNumber) {
      updateInGoogleSheet(oldReportNumber, record.toJSON()).catch(err => {
        console.warn('⚠️  Could not update in Google Sheets:', err.message);
      });
    }

    // Notify connected clients via WebSocket
    const io = req.app.get('io');
    if (io) {
      io.emit('data_updated', { type: 'edit_record', record: record.toJSON() });
    }

    res.json({ success: true, record });
  } catch (err) {
    if (err.name === 'SequelizeUniqueConstraintError') {
      return res.status(409).json({ error: 'A report with this number already exists' });
    }
    console.error('Update report error:', err);
    res.status(500).json({ error: 'Failed to update report' });
  }
});

// -------------------------------------------------------
// GET /api/reports/export/excel — Download filtered data as Excel
// -------------------------------------------------------
router.get('/export/excel', async (req, res) => {
  try {
    const where = buildWhereClause(req.query);
    const records = await TestReport.findAll({ where, order: [['sl_no', 'ASC']], raw: true });

    const headers = [
      'SL NO', 'Test Group', 'Test Component', 'Vehicle Model', 'VIN Number', 'Report Number',
      'Test Name', 'Test Description', 'Requested By', 'Test Location', 'Start Date',
      'End Date', 'Report Date', 'Test Engineer', 'Test Decision', 'Remark', 'Test Data'
    ];

    const rows = records.map(r => [
      r.sl_no || '',
      r.testing_group || '',
      r.test_component || '',
      r.vehicle_model || '',
      r.vin_number || '',
      r.report_number || '',
      r.test_name || '',
      r.test_description || '',
      r.requested_by || '',
      r.test_location || '',
      r.start_date || '',
      r.end_date || '',
      r.report_date || '',
      r.test_engineer || '',
      r.test_decision || '',
      r.remark || '',
      r.test_data || ''
    ]);

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

    // Optional: add some basic column sizing
    const wscols = [
      {wch: 8},  // SL NO
      {wch: 15}, // Test Group
      {wch: 20}, // Test Component
      {wch: 15}, // Vehicle Model
      {wch: 20}, // VIN Number
      {wch: 25}, // Report Number
      {wch: 25}, // Test Name
      {wch: 40}, // Test Description
      {wch: 15}, // Requested By
      {wch: 15}, // Test Location
      {wch: 12}, // Start Date
      {wch: 12}, // End Date
      {wch: 12}, // Report Date
      {wch: 15}, // Test Engineer
      {wch: 15}, // Test Decision
      {wch: 40}, // Remark
      {wch: 30}  // Test Data
    ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'T&V Export');

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="TV_Reports_Export_${new Date().toISOString().split('T')[0]}.xlsx"`,
      'Content-Length': buf.length,
    });

    res.send(buf);
  } catch (err) {
    console.error('Excel export error:', err);
    res.status(500).json({ error: 'Failed to generate Excel export' });
  }
});

// -------------------------------------------------------
// POST /api/reports/upload-url — Fetch & parse Excel/CSV from URL
// -------------------------------------------------------
router.post('/upload-url', async (req, res) => {
  try {
    const { url } = req.body;
    if (!url) return res.status(400).json({ error: 'URL is required' });

    // Fetch the file using native fetch
    const response = await fetch(url);
    if (!response.ok) {
      return res.status(400).json({ error: `Failed to fetch file from URL: status ${response.status}` });
    }

    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });

    const uploadTabSource = 'google_sheets:upload data';
    // Clear only existing uploaded database records before bulk upload
    await TestReport.destroy({
      where: {
        [Op.or]: [
          { source: { [Op.in]: ['upload', 'manual', 'url_import', 'google_drive', uploadTabSource] } },
          { source: { [Op.like]: 'upload:%' } },
          { source: { [Op.like]: 'url_import:%' } }
        ]
      }
    });

    let added = 0, skipped = 0, errors = [];
    const uploadedRecords = [];

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      if (rows.length < 2) continue; // Skip empty sheets

      const headers = rows[0];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0 || !row.some(cell => String(cell || '').trim() !== '')) continue;

        const record = mapRowToRecord(row, headers, i, `url_import:${sheetName}`);

        try {
          const existing = await TestReport.findOne({ where: { report_number: record.report_number } });
          if (existing) {
            let newSource = record.source;
            if (existing.source && existing.source !== record.source) {
              const sourcesSet = new Set(existing.source.split(',').map(s => s.trim()));
              sourcesSet.add(record.source);
              newSource = Array.from(sourcesSet).join(',');
            }
            await existing.update({
              ...record,
              source: newSource
            });
          } else {
            await TestReport.create(record);
          }
          added++;
          uploadedRecords.push(record);
        } catch (e) {
          skipped++;
          errors.push(`Sheet "${sheetName}" Row ${i + 1}: ${e.message}`);
        }
      }
    }

    if (uploadedRecords.length > 0) {
      writeMultipleToGoogleSheet(uploadedRecords).catch(sheetsErr => {
        console.warn('⚠️  Could not bulk write URL data to Google Sheets:', sheetsErr.message);
      });
    }

    // Notify connected clients
    const io = req.app.get('io');
    if (io) io.emit('data_updated', { type: 'bulk_upload', added });

    res.json({ success: true, added, skipped, errors: errors.slice(0, 10) });
  } catch (err) {
    console.error('URL upload error:', err);
    res.status(500).json({ error: 'Failed to process spreadsheet from URL' });
  }
});

// -------------------------------------------------------
// POST /api/reports/upload-image — Upload image file
// -------------------------------------------------------
router.post('/upload-image', imageUpload.single('image'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No image file uploaded' });
    }

    // Generate static URL path
    const imagePath = `/uploads/${req.file.filename}`;

    res.json({
      success: true,
      filename: req.file.filename,
      path: imagePath,
      mimetype: req.file.mimetype,
      size: req.file.size
    });
  } catch (err) {
    console.error('Image upload error:', err);
    res.status(500).json({ error: err.message || 'Failed to upload image' });
  }
});

module.exports = router;
