# DOCX & XLSX to HTML Converter

## Overview
Simple Node.js + Express app that converts Microsoft Word (.docx) and Excel (.xlsx) files to standalone HTML files.
- `.docx` conversion uses `mammoth` and embeds images as base64 data URIs.
- `.xlsx` conversion uses `exceljs` and outputs basic HTML tables with some inline styles (bold/italic/background color).

 # 📄 Word & Excel to HTML Converter

A simple, lightweight **Node.js + Express** application to convert Microsoft Word (`.docx`) and Excel (`.xlsx`) documents into **clean, standalone HTML** files.  
Designed to preserve essential formatting, table structures, and embedded images while producing HTML that works across browsers and devices.

---

## ✨ Features
- 📂 **Upload & Convert** — Easily upload `.docx` or `.xlsx` files for conversion.
- 🖼 **Embedded Images** — Word images are converted to inline Base64 for portability.
- 📊 **Excel Table Output** — Preserves basic formatting like **bold**, *italic*, and background colors.
- 💻 **Cross-platform HTML** — Clean HTML that works across browsers and devices.
- ⚡ **Fast & Local** — All processing happens locally on your machine, no cloud dependency.

---

## 🛠 Tech Stack
- **Node.js** — Backend runtime
- **Express.js** — Web framework
- **Multer** — File upload handling
- **Mammoth.js** — Word to HTML conversion
- **ExcelJS** — Excel to HTML table conversion
- **HTML/CSS** — Simple upload UI

---

## 📦 Installation & Setup
1. **Clone this repository**
   ```bash
   git clone https://github.com/YOUR-USERNAME/docx-xlsx-to-html.git
   cd docx-xlsx-to-html


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
