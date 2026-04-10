const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { Readable } = require('stream');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

function log(label, data) {
  const ts = new Date().toISOString();
  console.log(`[${ts}] ${label}`, data !== undefined ? data : '');
}

// Request logger
app.use((req, res, next) => {
  log(`→ ${req.method} ${req.path}`, { query: req.query, 'content-type': req.headers['content-type'] });
  next();
});

// Raw binary body parser (for n8n "Binary" body type)
app.use((req, res, next) => {
  const ct = req.headers['content-type'] || '';
  if (ct.includes('multipart/form-data')) {
    log('body-parser', 'multipart/form-data detected → delegating to multer');
    return next();
  }
  if (req.method === 'POST') {
    log('body-parser', 'raw binary body — reading stream');
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      req.rawBody = Buffer.concat(chunks);
      log('body-parser', `raw body received: ${req.rawBody.length} bytes`);
      next();
    });
  } else {
    next();
  }
});

// Returns the background color of a cell as #RRGGBB hex, or null if no fill
function getCellBgColor(cell) {
  const fill = cell.fill;
  if (!fill || fill.type !== 'pattern') return null;
  const argb = fill.fgColor?.argb;
  if (!argb || typeof argb !== 'string' || argb.toUpperCase() === 'FFFFFFFF') return null;
  const rgb = argb.replace(/^FF/i, '').toUpperCase();
  return `#${rgb}`;
}

// POST /parse-excel
// Accepts:
//   - multipart/form-data with field "file"  (curl / form)
//   - raw binary body                         (n8n "Binary" body type)
// Returns array of row objects, each with _cell_styles
app.post('/parse-excel', upload.single('file'), async (req, res) => {
  try {
    const source = req.file ? 'form-data' : 'raw-binary';
    const buffer = req.file?.buffer ?? req.rawBody;
    log('parse-excel', `buffer source: ${source}, size: ${buffer?.length ?? 0} bytes`);

    if (!buffer || buffer.length === 0) {
      log('parse-excel', 'ERROR: no file received');
      return res.status(400).json({ error: 'No file received. Send as form-data field "file" or raw binary body.' });
    }

    const workbook = new ExcelJS.Workbook();
    const stream = Readable.from(buffer);
    await workbook.xlsx.read(stream);

    const sheets = workbook.worksheets.map(ws => ws.name);
    log('parse-excel', `workbook loaded — sheets: ${JSON.stringify(sheets)}`);

    // Use the first sheet by default, or ?sheet=SheetName
    const sheetName = req.query.sheet;
    const worksheet = sheetName
      ? workbook.getWorksheet(sheetName)
      : workbook.worksheets[0];

    if (!worksheet) {
      log('parse-excel', `ERROR: sheet "${sheetName}" not found`);
      return res.status(404).json({ error: `Sheet "${sheetName}" not found.` });
    }

    log('parse-excel', `using sheet: "${worksheet.name}", rowCount: ${worksheet.rowCount}`);

    const rows = [];
    let headers = [];

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) {
        headers = row.values.slice(1).map((v, i) => {
          if (v === null || v === undefined || v === '') return `Col${i + 1}`;
          return String(v).trim();
        });
        log('parse-excel', `headers (${headers.length}): ${JSON.stringify(headers)}`);
        return;
      }

      const rowObj = {};
      const cellStyles = {};

      headers.forEach((header, i) => {
        const cell = row.getCell(i + 1);
        const value = cell.value;

        if (value && typeof value === 'object' && value.result !== undefined) {
          rowObj[header] = value.result;
        } else if (value && typeof value === 'object' && value instanceof Date) {
          rowObj[header] = value.toISOString();
        } else {
          rowObj[header] = value ?? null;
        }

        const bg = getCellBgColor(cell);
        cellStyles[header] = { bg: bg ?? null };
      });

      rowObj._cell_styles = cellStyles;
      rows.push(rowObj);
    });

    log('parse-excel', `done — ${rows.length} data rows returned`);
    res.json(rows);
  } catch (err) {
    log('parse-excel', `EXCEPTION: ${err.message}`);
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Health check
app.get('/health', (req, res) => {
  log('health', 'ok');
  res.json({ status: 'ok' });
});

// List sheets
app.post('/list-sheets', upload.single('file'), async (req, res) => {
  try {
    const buffer = req.file?.buffer ?? req.rawBody;
    log('list-sheets', `buffer size: ${buffer?.length ?? 0} bytes`);
    if (!buffer || buffer.length === 0) {
      log('list-sheets', 'ERROR: no file received');
      return res.status(400).json({ error: 'No file received.' });
    }
    const workbook = new ExcelJS.Workbook();
    const stream = Readable.from(buffer);
    await workbook.xlsx.read(stream);
    const sheets = workbook.worksheets.map(ws => ws.name);
    log('list-sheets', `sheets: ${JSON.stringify(sheets)}`);
    res.json(sheets);
  } catch (err) {
    log('list-sheets', `EXCEPTION: ${err.message}`);
    res.status(500).json({ error: err.message });
  }
});

// Debug: inspect raw cell fills from first row of data
app.post('/debug-styles', upload.single('file'), async (req, res) => {
  try {
    const buffer = req.file?.buffer ?? req.rawBody;
    if (!buffer || buffer.length === 0) return res.status(400).json({ error: 'No file received.' });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.read(Readable.from(buffer));
    const ws = workbook.worksheets[0];
    const result = [];
    // Inspect first 5 data rows
    let count = 0;
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1 || count >= 2) return;
      count++;
      row.eachCell({ includeEmpty: false }, (cell) => {
        result.push({
          address: cell.address,
          fill: cell.fill,
          value: typeof cell.value === 'object' ? JSON.stringify(cell.value) : cell.value,
        });
      });
    });
    res.json(result);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3200;
app.listen(PORT, () => log('startup', `Promould Custom API running on port ${PORT}`));
