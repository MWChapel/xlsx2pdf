# XLSX2PDF

A Node.js CLI tool that converts Excel (`.xlsx`) files to PDF.

## Dependencies

- [adm-zip](https://www.npmjs.com/package/adm-zip) — extracts the ZIP archive that makes up an XLSX file
- [fast-xml-parser](https://www.npmjs.com/package/fast-xml-parser) — parses the XML files inside the archive
- [pdfkit](https://www.npmjs.com/package/pdfkit) — generates PDF output

## Installation

```bash
npm install
```

## Usage

```bash
node index.js <input.xlsx> <output.pdf>
```

**Example:**

```bash
node index.js report.xlsx report.pdf
```

## Features

- Converts all worksheets — each sheet gets its own page with its name as a title
- Styled header row (blue background, white bold text)
- Alternating row colors for readability
- Automatic column width distribution
- Page breaks when content exceeds a single page
- Handles cell types: plain values, dates, formula results, rich text, and errors

## Output Format

- Page size: A4 landscape
- Font: Helvetica
- Font size: 9pt (data), 14pt (sheet title)
