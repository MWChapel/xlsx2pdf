'use strict';

const AdmZip = require('adm-zip');
const { XMLParser } = require('fast-xml-parser');
const PDFDocument = require('pdfkit');
const fs = require('fs');

const MARGIN = 30;
const ROW_HEIGHT = 20;
const FONT_SIZE = 9;
const TITLE_FONT_SIZE = 14;
const HEADER_COLOR = '#4472C4';
const ALT_ROW_COLOR = '#EEF2F8';
const BORDER_COLOR = '#CCCCCC';

// Date serial numbers in XLSX are days since 1899-12-30
const XLSX_EPOCH = new Date(1899, 11, 30).getTime();
// Built-in date format IDs per OOXML spec
const DATE_FMT_IDS = new Set([14,15,16,17,18,19,20,21,22,45,46,47]);

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  isArray: (name) => ['sheet', 'Relationship', 'row', 'c', 'si', 'r', 'xf', 'numFmt'].includes(name),
});

function colLetterToIndex(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n - 1;
}

function parseRef(ref) {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  return m ? { col: colLetterToIndex(m[1]), row: parseInt(m[2]) - 1 } : null;
}

function readXlsx(filePath) {
  const zip = new AdmZip(filePath);

  const workbookXml = zip.readAsText('xl/workbook.xml');
  const wb = xmlParser.parse(workbookXml);
  const sheets = wb.workbook.sheets.sheet;

  const relsXml = zip.readAsText('xl/_rels/workbook.xml.rels');
  const rels = xmlParser.parse(relsXml);
  const relMap = {};
  rels.Relationships.Relationship.forEach(r => { relMap[r['@_Id']] = r['@_Target']; });

  // Shared strings
  const sharedStrings = [];
  if (zip.getEntry('xl/sharedStrings.xml')) {
    const ss = xmlParser.parse(zip.readAsText('xl/sharedStrings.xml'));
    (ss.sst.si || []).forEach(si => {
      if (si.t !== undefined) sharedStrings.push(String(si.t));
      else if (si.r) sharedStrings.push((Array.isArray(si.r) ? si.r : [si.r]).map(r => String(r.t ?? '')).join(''));
      else sharedStrings.push('');
    });
  }

  // Date format detection via styles.xml
  const dateFmtIds = new Set(DATE_FMT_IDS);
  if (zip.getEntry('xl/styles.xml')) {
    const styles = xmlParser.parse(zip.readAsText('xl/styles.xml'));
    const numFmts = styles.styleSheet?.numFmts?.numFmt || [];
    numFmts.forEach(fmt => {
      const code = (fmt['@_formatCode'] || '').toLowerCase();
      if (/[ymd]/.test(code) && !/[#0]/.test(code)) dateFmtIds.add(Number(fmt['@_numFmtId']));
    });
    // xfIdx -> numFmtId lookup
    const xfs = styles.styleSheet?.cellXfs?.xf || [];
    const xfFmtIds = xfs.map(xf => Number(xf['@_numFmtId'] || 0));
    // Store on relMap so sheet parser can reach it
    relMap.__xfFmtIds__ = xfFmtIds;
  }

  const xfFmtIds = relMap.__xfFmtIds__ || [];

  return sheets.map(sheet => {
    const name = sheet['@_name'];
    const rId = sheet['@_r:id'];
    const target = relMap[rId];
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;

    const sd = xmlParser.parse(zip.readAsText(path));
    const rawRows = sd.worksheet?.sheetData?.row || [];

    const rowMap = {};
    let maxCol = 0;

    rawRows.forEach(row => {
      const cells = Array.isArray(row.c) ? row.c : row.c ? [row.c] : [];
      cells.forEach(cell => {
        const ref = parseRef(cell['@_r'] || '');
        if (!ref) return;
        const { col, row: r } = ref;
        maxCol = Math.max(maxCol, col);
        const t = cell['@_t'];
        const s = Number(cell['@_s'] ?? -1);
        const v = cell.v;

        let text = '';
        if (v !== undefined && v !== null) {
          if (t === 's') {
            text = sharedStrings[parseInt(v)] ?? '';
          } else if (t === 'b') {
            text = v === '1' || v === 1 ? 'TRUE' : 'FALSE';
          } else if (t === 'str' || t === 'inlineStr') {
            text = String(v);
          } else if (t === 'e') {
            text = String(v);
          } else {
            // Numeric — check if it's a date style
            const fmtId = s >= 0 ? xfFmtIds[s] : -1;
            if (fmtId >= 0 && dateFmtIds.has(fmtId)) {
              const ms = XLSX_EPOCH + parseFloat(v) * 86400000;
              text = new Date(ms).toLocaleDateString();
            } else {
              text = String(v);
            }
          }
        }

        if (!rowMap[r]) rowMap[r] = {};
        rowMap[r][col] = text;
      });
    });

    const rowIndices = Object.keys(rowMap).map(Number).sort((a, b) => a - b);
    const rows = rowIndices.map(r => {
      const arr = [];
      for (let c = 0; c <= maxCol; c++) arr.push(rowMap[r][c] ?? '');
      return arr;
    });

    return { name, rows };
  });
}

function drawSheet(doc, { name, rows }) {
  doc
    .font('Helvetica-Bold')
    .fontSize(TITLE_FONT_SIZE)
    .fillColor('#000000')
    .text(name, { align: 'center' });
  doc.moveDown(0.5);

  if (rows.length === 0) return;

  const maxCols = Math.max(...rows.map(r => r.length));
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

    doc.strokeColor(BORDER_COLOR).lineWidth(0.5)
      .moveTo(startX, y).lineTo(startX + pageWidth, y).stroke();

    for (let c = 0; c <= maxCols; c++) {
      doc.moveTo(startX + c * colWidth, y)
        .lineTo(startX + c * colWidth, y + ROW_HEIGHT).stroke();
    }

    y += ROW_HEIGHT;
  });

  doc.strokeColor(BORDER_COLOR).lineWidth(0.5)
    .moveTo(startX, y).lineTo(startX + pageWidth, y).stroke();

  doc.fillColor('#000000');
  doc.y = y + 10;
}

function convertXlsxToPdf(inputPath, outputPath) {
  const sheets = readXlsx(inputPath);

  const doc = new PDFDocument({ margin: MARGIN, size: 'A4', layout: 'landscape' });
  const out = fs.createWriteStream(outputPath);
  doc.pipe(out);

  sheets.forEach((sheet, i) => {
    if (i > 0) doc.addPage();
    drawSheet(doc, sheet);
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
