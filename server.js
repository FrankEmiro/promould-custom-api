const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { Readable } = require('stream');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Map ARGB hex to human-readable color names
// Excel stores colors as ARGB (e.g. "FF00B050" = fully opaque green)
function argbToColorName(argb) {
  if (!argb) return null;

  // Strip alpha channel → get RGB
  const rgb = argb.replace(/^FF/i, '').toUpperCase();

  const colorMap = {
    // Greens
    '00B050': 'green',
    '92D050': 'green',
    '00FF00': 'green',
    '70AD47': 'green',
    'E2EFDA': 'green_light',
    // Yellows
    'FFFF00': 'yellow',
    'FFEB9C': 'yellow',
    'FFC000': 'orange',
    'FFD966': 'yellow',
    'FFF2CC': 'yellow_light',
    // Reds
    'FF0000': 'red',
    'FF0000': 'red',
    'C00000': 'red',
    'FF4444': 'red',
    'FFE7E7': 'red_light',
    'FCE4D6': 'red_light',
    'FA8072': 'red',
    // Blues
    '0070C0': 'blue',
    '4472C4': 'blue',
    '2E75B6': 'blue',
    'BDD7EE': 'blue_light',
    'DEEAF1': 'blue_light',
    // Orange
    'ED7D31': 'orange',
    'F4B942': 'orange',
    // White / no fill
    'FFFFFF': null,
    // Grey
    'A6A6A6': 'grey',
    'D9D9D9': 'grey_light',
    'BFBFBF': 'grey',
  };

  if (colorMap.hasOwnProperty(rgb)) return colorMap[rgb];

  // Fallback: return hex string so the AI can still reason about it
  return `#${rgb}`;
}

function getCellBgColor(cell) {
  const fill = cell.fill;
  if (!fill) return null;
  if (fill.type === 'pattern' && fill.fgColor) {
    const argb = fill.fgColor.argb || fill.fgColor.theme;
    if (argb && typeof argb === 'string') {
      return argbToColorName(argb);
    }
  }
  return null;
}

// POST /parse-excel
// Accepts multipart/form-data with field "file" (binary xlsx)
// Returns array of row objects, each with _cell_styles
app.post('/parse-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded. Use field name "file".' });
    }

    const workbook = new ExcelJS.Workbook();
    const stream = Readable.from(req.file.buffer);
    await workbook.xlsx.read(stream);

    // Use the first sheet by default, or ?sheet=SheetName
    const sheetName = req.query.sheet;
    const worksheet = sheetName
      ? workbook.getWorksheet(sheetName)
      : workbook.worksheets[0];

    if (!worksheet) {
      return res.status(404).json({ error: `Sheet "${sheetName}" not found.` });
    }

    const rows = [];
    let headers = [];

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) {
        // First row = headers
        headers = row.values.slice(1).map((v, i) => {
          if (v === null || v === undefined || v === '') return `Col${i + 1}`;
          return String(v).trim();
        });
        return;
      }

      const rowObj = {};
      const cellStyles = {};

      headers.forEach((header, i) => {
        const cell = row.getCell(i + 1);
        const value = cell.value;

        // Resolve formula result if present
        if (value && typeof value === 'object' && value.result !== undefined) {
          rowObj[header] = value.result;
        } else if (value && typeof value === 'object' && value instanceof Date) {
          rowObj[header] = value.toISOString();
        } else {
          rowObj[header] = value ?? null;
        }

        const bg = getCellBgColor(cell);
        if (bg !== null) {
          cellStyles[header] = { bg };
        } else {
          cellStyles[header] = { bg: null };
        }
      });

      rowObj._cell_styles = cellStyles;
      rows.push(rowObj);
    });

    res.json(rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Health check
app.get('/health', (req, res) => res.json({ status: 'ok' }));

// List sheets
app.post('/list-sheets', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });
    const workbook = new ExcelJS.Workbook();
    const stream = Readable.from(req.file.buffer);
    await workbook.xlsx.read(stream);
    res.json(workbook.worksheets.map(ws => ws.name));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3200;
app.listen(PORT, () => console.log(`Excel Style API running on port ${PORT}`));
