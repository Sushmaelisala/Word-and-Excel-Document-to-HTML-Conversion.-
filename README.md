# DOCX & XLSX to HTML Converter

## Overview
Simple Node.js + Express app that converts Microsoft Word (.docx) and Excel (.xlsx) files to standalone HTML files.
- `.docx` conversion uses `mammoth` and embeds images as base64 data URIs.
- `.xlsx` conversion uses `exceljs` and outputs basic HTML tables with some inline styles (bold/italic/background color).

## How to run
1. Install Node.js (v14+ recommended).
2. Extract the project and open a terminal in the project folder.
3. Run `npm install` to install dependencies.
4. Run `npm start`
5. Open `http://localhost:3000` in your browser, upload a `.docx` or `.xlsx`, and download the converted HTML.

## Limitations & Notes
- Mammoth aims to produce clean HTML but does not preserve every Word feature (e.g., complex floating objects, exact page breaks, advanced WordArt).
- ExcelJS reads cell values and some styles; it does not perfectly replicate Excel rendering (column widths, row heights, text wrapping, charts, pivot tables).
- For production-grade conversions you may require more advanced paid libraries or headless office suites (LibreOffice headless, Aspose, etc.).

## Files
- `server.js` — Express server handling uploads and conversions.
- `public/index.html` — Frontend upload form.
- `package.json` — Dependencies and start script.

## License
MIT
