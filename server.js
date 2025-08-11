const express = require('express');
const multer = require('multer');
const mammoth = require('mammoth');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const mime = require('mime-types');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static(path.join(__dirname, 'public')));

// Upload form
app.post('/convert', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');

    const ext = path.extname(req.file.originalname).toLowerCase();

    try {
        if (ext === '.docx' || ext === '.doc') {
            // Convert docx using mammoth
            const options = {
                convertImage: mammoth.images.inline(function() {
                    return function(element) {
                        return element.read("base64").then(function(imageBuffer) {
                            return { src: "data:" + element.contentType + ";base64," + imageBuffer };
                        });
                    };
                })
            };
            const result = await mammoth.convertToHtml({path: req.file.path}, options);
            const html = wrapHtml(req.file.originalname, result.value);
            sendHtmlAsFile(res, html, req.file.originalname + '.html');
        } else if (ext === '.xlsx' || ext === '.xls') {
            // Convert Excel using ExcelJS
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(req.file.path);
            let combinedHtml = '';
            workbook.eachSheet((worksheet, sheetId) => {
                combinedHtml += `<h2>Sheet: ${escapeHtml(worksheet.name)}</h2>`;
                combinedHtml += worksheetToTable(worksheet);
            });
            const html = wrapHtml(req.file.originalname, combinedHtml);
            sendHtmlAsFile(res, html, req.file.originalname + '.html');
        } else {
            res.status(400).send('Unsupported file type. Upload a .docx or .xlsx file.');
        }
    } catch (err) {
        console.error(err);
        res.status(500).send('Conversion error: ' + err.message);
    } finally {
        // cleanup uploaded file
        fs.unlink(req.file.path, () => {});
    }
});

function sendHtmlAsFile(res, htmlContent, filename) {
    res.setHeader('Content-disposition', 'attachment; filename=' + filename);
    res.setHeader('Content-Type', 'text/html');
    res.send(htmlContent);
}

function wrapHtml(title, bodyContent) {
    return `<!doctype html>
<html>
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>${escapeHtml(title)}</title>
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; line-height:1.4; }
    table { border-collapse: collapse; margin-bottom: 1rem; }
    table, th, td { border: 1px solid #bbb; padding: 6px 8px; }
    th { background: #f0f0f0; }
  </style>
</head>
<body>
  ${bodyContent}
</body>
</html>`;
}

function escapeHtml(s) {
    if (!s) return '';
    return String(s).replace(/[&<>"']/g, function(m) {
        return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]);
    });
}

function worksheetToTable(worksheet) {
    // Determine the range
    const range = worksheet.dimensions; // { top, left, bottom, right }
    let html = '<table>';
    for (let r = range.top; r <= range.bottom; r++) {
        html += '<tr>';
        for (let c = range.left; c <= range.right; c++) {
            const cell = worksheet.getCell(r, c);
            const value = cell.value === null ? '' : cell.value;
            // Basic inline styles for bold/italic and fill color
            let styles = '';
            try {
                if (cell.font) {
                    if (cell.font.bold) styles += 'font-weight:bold;';
                    if (cell.font.italic) styles += 'font-style:italic;';
                }
                if (cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb) {
                    const argb = cell.fill.fgColor.argb;
                    // convert ARGB to hex (#RRGGBB)
                    const rgb = argb.length === 8 ? '#' + argb.slice(2,8) : '#'+argb;
                    styles += 'background:' + rgb + ';';
                }
            } catch(e) {}
            const cellTag = (worksheet.getRow(r).getCell(c).isMerged) ? 'td' : 'td';
            html += `<${cellTag} style="${styles}">${escapeHtml(String(value))}</${cellTag}>`;
        }
        html += '</tr>';
    }
    html += '</table>';
    return html;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server started on http://localhost:${PORT}`);
});
