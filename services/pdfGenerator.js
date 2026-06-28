// ============================================================
// pdfGenerator.js — PDF Report Generation Service
// Generates beautiful PDF reports from filtered test data
// Uses pdfkit library to create professional PDF documents
// ============================================================

const PDFDocument = require('pdfkit');

// -------------------------------------------------------
// Colors used in the PDF
// -------------------------------------------------------
const COLORS = {
  primary: '#1a1a2e',
  accent: '#6c63ff',
  success: '#00d4aa',
  danger: '#ff4757',
  warning: '#ffa502',
  light: '#f8f9fa',
  dark: '#2d2d44',
  white: '#ffffff',
  gray: '#666666',
  border: '#e0e0e0',
};

// -------------------------------------------------------
// Main function: Generate PDF buffer from data
// @param {Array} data - Array of test report objects
// @param {Object} filters - Applied filters (for display)
// @param {Object} stats - Summary statistics
// @returns {Buffer} - PDF file buffer
// -------------------------------------------------------
async function generatePDF(data, filters = {}, stats = {}) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: 'A4',
      layout: 'landscape', // Wide layout for the table
      margins: { top: 40, bottom: 40, left: 40, right: 40 },
    });

    const chunks = [];
    doc.on('data', chunk => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const pageWidth = doc.page.width - 80; // Account for margins

    // -------------------------------------------------------
    // HEADER SECTION
    // -------------------------------------------------------
    // Dark header background
    doc.rect(0, 0, doc.page.width, 90).fill(COLORS.primary);

    // Company name
    doc.fillColor(COLORS.white)
       .font('Helvetica-Bold')
       .fontSize(20)
       .text('SIMPLE ENERGY', 40, 20);

    // Subtitle
    doc.fillColor(COLORS.accent)
       .font('Helvetica')
       .fontSize(11)
       .text('Testing & Validation Department', 40, 45);

    // Report title on right
    doc.fillColor(COLORS.white)
       .font('Helvetica-Bold')
       .fontSize(14)
       .text('TEST REPORTS SUMMARY', 40, 65, { align: 'right', width: pageWidth });

    // -------------------------------------------------------
    // REPORT META INFO
    // -------------------------------------------------------
    doc.fillColor(COLORS.dark).rect(0, 90, doc.page.width, 40).fill('#f0f0f8');

    doc.fillColor(COLORS.gray)
       .font('Helvetica')
       .fontSize(9)
       .text(
         `Generated: ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })}  |  Total Records: ${data.length}  |  Filters: ${formatFilters(filters)}`,
         40, 100, { width: pageWidth }
       );

    // -------------------------------------------------------
    // STATS SUMMARY BOX
    // -------------------------------------------------------
    const statsY = 145;

    // Calculate stats from data if not provided
    const totalTests = stats.total || data.length;
    const passed = stats.passed || data.filter(r => r.test_decision === 'PASSED').length;
    const failed = stats.failed || data.filter(r => r.test_decision === 'FAILED').length;
    const noDecision = stats.no_decision || data.filter(r => !['PASSED', 'FAILED'].includes(r.test_decision)).length;
    const passRate = totalTests > 0 ? ((passed / totalTests) * 100).toFixed(1) : 0;

    const boxWidth = pageWidth / 4 - 5;

    drawStatBox(doc, 40, statsY, boxWidth, 55, 'TOTAL TESTS', totalTests, COLORS.accent);
    drawStatBox(doc, 40 + boxWidth + 5, statsY, boxWidth, 55, 'PASSED', passed, COLORS.success);
    drawStatBox(doc, 40 + (boxWidth + 5) * 2, statsY, boxWidth, 55, 'FAILED', failed, COLORS.danger);
    drawStatBox(doc, 40 + (boxWidth + 5) * 3, statsY, boxWidth, 55, 'PASS RATE', `${passRate}%`, COLORS.warning);

    // -------------------------------------------------------
    // DATA TABLE
    // -------------------------------------------------------
    const tableY = statsY + 70;
    const columns = [
      { header: 'SL', key: 'sl_no', width: 30 },
      { header: 'Group', key: 'testing_group', width: 80 },
      { header: 'Component', key: 'test_component', width: 120 },
      { header: 'Vehicle', key: 'vehicle_model', width: 70 },
      { header: 'Report No.', key: 'report_number', width: 110 },
      { header: 'Engineer', key: 'test_engineer', width: 80 },
      { header: 'Date', key: 'report_date', width: 65 },
      { header: 'Decision', key: 'test_decision', width: 70 },
      { header: 'Location', key: 'test_location', width: 65 },
    ];

    // Table header row
    let currentX = 40;
    doc.rect(40, tableY, pageWidth, 22).fill(COLORS.primary);

    columns.forEach(col => {
      doc.fillColor(COLORS.white)
         .font('Helvetica-Bold')
         .fontSize(8)
         .text(col.header, currentX + 3, tableY + 7, { width: col.width - 6, ellipsis: true });
      currentX += col.width;
    });

    // Table data rows
    let rowY = tableY + 22;
    let rowCount = 0;

    for (const record of data) {
      // Check if we need a new page
      if (rowY > doc.page.height - 60) {
        doc.addPage({ layout: 'landscape' });
        rowY = 40;

        // Repeat header on new page
        currentX = 40;
        doc.rect(40, rowY, pageWidth, 22).fill(COLORS.primary);
        columns.forEach(col => {
          doc.fillColor(COLORS.white).font('Helvetica-Bold').fontSize(8)
             .text(col.header, currentX + 3, rowY + 7, { width: col.width - 6, ellipsis: true });
          currentX += col.width;
        });
        rowY += 22;
        rowCount = 0;
      }

      // Alternating row background
      const bgColor = rowCount % 2 === 0 ? COLORS.white : '#f8f8ff';
      doc.rect(40, rowY, pageWidth, 18).fill(bgColor);

      // Decision color indicator
      const decisionColor = getDecisionColor(record.test_decision);
      doc.rect(40, rowY, 3, 18).fill(decisionColor);

      // Row data
      currentX = 40;
      columns.forEach(col => {
        let value = record[col.key] || '';
        // Special formatting for decision column
        if (col.key === 'test_decision') {
          doc.fillColor(decisionColor).font('Helvetica-Bold');
        } else {
          doc.fillColor(COLORS.dark).font('Helvetica');
        }
        doc.fontSize(7.5).text(String(value), currentX + 4, rowY + 5, {
          width: col.width - 8,
          ellipsis: true,
          lineBreak: false,
        });
        currentX += col.width;
      });

      // Row border line
      doc.moveTo(40, rowY + 18).lineTo(40 + pageWidth, rowY + 18)
         .strokeColor('#e8e8f0').lineWidth(0.5).stroke();

      rowY += 18;
      rowCount++;
    }

    // -------------------------------------------------------
    // FOOTER
    // -------------------------------------------------------
    const footerY = doc.page.height - 35;
    doc.rect(0, footerY, doc.page.width, 35).fill(COLORS.primary);
    doc.fillColor(COLORS.gray)
       .font('Helvetica')
       .fontSize(8)
       .text(
         'CONFIDENTIAL — Simple Energy Pvt. Ltd. | Testing & Validation Department',
         40, footerY + 12, { align: 'center', width: pageWidth }
       );

    doc.end();
  });
}

// Helper: Draw a statistics box
function drawStatBox(doc, x, y, w, h, label, value, color) {
  doc.rect(x, y, w, h).fill(color + '20'); // Light background
  doc.rect(x, y, 3, h).fill(color); // Left accent bar
  doc.fillColor(color).font('Helvetica-Bold').fontSize(22)
     .text(String(value), x + 10, y + 6, { width: w - 15 });
  doc.fillColor('#888').font('Helvetica').fontSize(8)
     .text(label, x + 10, y + 36, { width: w - 15 });
}

// Helper: Get color for test decision
function getDecisionColor(decision) {
  switch ((decision || '').toUpperCase()) {
    case 'PASSED': return '#00c851';
    case 'FAILED': return '#ff4444';
    case 'NO DECISION': return '#ffbb33';
    case 'NA': return '#aaaaaa';
    default: return '#33b5e5';
  }
}

// Helper: Format applied filters for display
function formatFilters(filters) {
  const parts = [];
  if (filters.testing_group) parts.push(`Group: ${filters.testing_group}`);
  if (filters.vehicle_model) parts.push(`Vehicle: ${filters.vehicle_model}`);
  if (filters.test_decision) parts.push(`Decision: ${filters.test_decision}`);
  if (filters.from_date) parts.push(`From: ${filters.from_date}`);
  if (filters.to_date) parts.push(`To: ${filters.to_date}`);
  return parts.length > 0 ? parts.join(' | ') : 'None';
}

module.exports = { generatePDF };
