'use strict';

const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const fs = require('fs');

const MARGIN = 30;
const ROW_HEIGHT = 20;
const FONT_SIZE = 9;
const TITLE_FONT_SIZE = 14;
const HEADER_COLOR = '#4472C4';
const ALT_ROW_COLOR = '#EEF2F8';
const BORDER_COLOR = '#CCCCCC';

function getCellText(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toLocaleDateString();
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(r => r.text).join('');
    if (v.formula !== undefined) return v.result != null ? String(v.result) : '';
    if (v.error) return v.error;
    if (v.text) return String(v.text);
  }
  return String(v);
}

function drawSheet(doc, worksheet) {
  const rows = [];
  let maxCols = 0;

  worksheet.eachRow({ includeEmpty: false }, row => {
    const cells = [];
    row.eachCell({ includeEmpty: true }, cell => cells.push(getCellText(cell)));
    maxCols = Math.max(maxCols, cells.length);
    rows.push(cells);
  });

  if (rows.length === 0) return;

  // Normalize row lengths
  rows.forEach(r => { while (r.length < maxCols) r.push(''); });

  const pageWidth = doc.page.width - MARGIN * 2;
  const colWidth = pageWidth / maxCols;
  const startX = MARGIN;
  let y = doc.y;

  rows.forEach((row, rowIdx) => {
    if (y + ROW_HEIGHT > doc.page.height - MARGIN) {
      doc.addPage();
      y = MARGIN;
    }

    const isHeader = rowIdx === 0;
    const bgColor = isHeader ? HEADER_COLOR : (rowIdx % 2 === 0 ? '#FFFFFF' : ALT_ROW_COLOR);

    doc.rect(startX, y, pageWidth, ROW_HEIGHT).fill(bgColor);

    row.forEach((text, colIdx) => {
      doc
        .fillColor(isHeader ? '#FFFFFF' : '#000000')
        .font(isHeader ? 'Helvetica-Bold' : 'Helvetica')
        .fontSize(FONT_SIZE)
        .text(text, startX + colIdx * colWidth + 3, y + 5, {
          width: colWidth - 6,
          height: ROW_HEIGHT - 5,
          lineBreak: false,
          ellipsis: true,
        });
    });

    // Horizontal border
    doc.strokeColor(BORDER_COLOR).lineWidth(0.5)
      .moveTo(startX, y).lineTo(startX + pageWidth, y).stroke();

    // Vertical borders
    for (let c = 0; c <= maxCols; c++) {
      doc.moveTo(startX + c * colWidth, y)
        .lineTo(startX + c * colWidth, y + ROW_HEIGHT).stroke();
    }

    y += ROW_HEIGHT;
  });

  // Bottom border
  doc.strokeColor(BORDER_COLOR).lineWidth(0.5)
    .moveTo(startX, y).lineTo(startX + pageWidth, y).stroke();

  doc.fillColor('#000000');
  doc.y = y + 10;
}

async function convertXlsxToPdf(inputPath, outputPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);

  const doc = new PDFDocument({ margin: MARGIN, size: 'A4', layout: 'landscape' });
  const out = fs.createWriteStream(outputPath);
  doc.pipe(out);

  let firstSheet = true;
  workbook.eachSheet(worksheet => {
    if (!firstSheet) doc.addPage();
    firstSheet = false;

    doc
      .font('Helvetica-Bold')
      .fontSize(TITLE_FONT_SIZE)
      .fillColor('#000000')
      .text(worksheet.name, { align: 'center' });
    doc.moveDown(0.5);

    drawSheet(doc, worksheet);
  });

  doc.end();

  return new Promise((resolve, reject) => {
    out.on('finish', resolve);
    out.on('error', reject);
  });
}

// CLI
const [,, inputPath, outputPath] = process.argv;

if (!inputPath || !outputPath) {
  console.error('Usage: node index.js <input.xlsx> <output.pdf>');
  process.exit(1);
}

if (!fs.existsSync(inputPath)) {
  console.error(`File not found: ${inputPath}`);
  process.exit(1);
}

convertXlsxToPdf(inputPath, outputPath)
  .then(() => console.log(`PDF saved to: ${outputPath}`))
  .catch(err => {
    console.error('Conversion failed:', err.message);
    process.exit(1);
  });
