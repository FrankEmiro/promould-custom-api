const express = require('express');
const multer = require('multer');
const unzipper = require('unzipper');
const { XMLParser } = require('fast-xml-parser');
const { Readable } = require('stream');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 200 * 1024 * 1024 } });

function log(label, data) {
  const ts = new Date().toISOString();
  console.log(`[${ts}] ${label}`, data !== undefined ? data : '');
}

app.use((req, res, next) => {
  log(`→ ${req.method} ${req.path}`, { query: req.query, 'content-type': req.headers['content-type'] });
  next();
});

// Raw binary body parser (for n8n "Binary" body type)
app.use((req, res, next) => {
  const ct = req.headers['content-type'] || '';
  if (ct.includes('multipart/form-data')) {
    log('body-parser', 'multipart/form-data → multer');
    return next();
  }
  if (req.method === 'POST') {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      req.rawBody = Buffer.concat(chunks);
      log('body-parser', `raw binary: ${req.rawBody.length} bytes`);
      next();
    });
  } else {
    next();
  }
});

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', isArray: (name) => ['row', 'c', 'xf', 'fill', 'patternFill', 'fgColor', 'mrgCell'].includes(name) });

// Read only specific entries from the xlsx zip (skip xl/media/*)
async function readXlsxEntries(buffer, targets) {
  const result = {};
  const zip = unzipper.Open.buffer(buffer);
  const directory = await zip;
  for (const file of directory.files) {
    if (targets.includes(file.path)) {
      const content = await file.buffer();
      result[file.path] = content.toString('utf8');
    }
  }
  return result;
}

// Parse shared strings table
function parseSharedStrings(xml) {
  if (!xml) return [];
  const parsed = xmlParser.parse(xml);
  const sst = parsed.sst;
  if (!sst || !sst.si) return [];
  const si = Array.isArray(sst.si) ? sst.si : [sst.si];
  return si.map(item => {
    if (item.t !== undefined) return String(item.t);
    if (item.r) {
      const runs = Array.isArray(item.r) ? item.r : [item.r];
      return runs.map(r => (r.t !== undefined ? String(r.t) : '')).join('');
    }
    return '';
  });
}

// Theme color index → position in clrScheme
const THEME_COLOR_KEYS = ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];

function parseThemeColors(xml) {
  if (!xml) return {};
  const parsed = xmlParser.parse(xml);
  const scheme = parsed?.['a:theme']?.['a:themeElements']?.['a:clrScheme']
               ?? parsed?.theme?.themeElements?.clrScheme
               ?? {};
  const colors = {};
  THEME_COLOR_KEYS.forEach((key, idx) => {
    const node = scheme[`a:${key}`] ?? scheme[key];
    if (!node) return;
    const srgb = node['a:srgbClr']?.['@_val'] ?? node['a:srgbClr'];
    const sys  = node['a:sysClr']?.['@_lastClr'];
    colors[idx] = (srgb || sys || '').toUpperCase();
  });
  return colors;
}

// Apply Excel tint to an RGB hex string (e.g. "4EA72E", tint 0.4 → lighter green)
function applyTint(hex, tint) {
  if (!hex || tint === undefined || tint === null) return hex;
  const r = parseInt(hex.slice(0,2), 16);
  const g = parseInt(hex.slice(2,4), 16);
  const b = parseInt(hex.slice(4,6), 16);
  let nr, ng, nb;
  if (tint >= 0) {
    nr = Math.round(r + (255 - r) * tint);
    ng = Math.round(g + (255 - g) * tint);
    nb = Math.round(b + (255 - b) * tint);
  } else {
    nr = Math.round(r * (1 + tint));
    ng = Math.round(g * (1 + tint));
    nb = Math.round(b * (1 + tint));
  }
  return [nr, ng, nb].map(c => Math.min(255, Math.max(0, c)).toString(16).padStart(2,'0')).join('').toUpperCase();
}

// Parse styles — returns array of fill colors indexed by xfId
function parseStyles(xml, themeColors = {}) {
  if (!xml) return [];
  const parsed = xmlParser.parse(xml);
  const ss = parsed.styleSheet;

  // Build fills array: index → #RRGGBB or null
  const fills = [];
  const fillList = ss?.fills?.fill;
  if (fillList) {
    const arr = Array.isArray(fillList) ? fillList : [fillList];
    arr.forEach(f => {
      const pf = f.patternFill;
      const patternArr = Array.isArray(pf) ? pf : (pf ? [pf] : []);
      if (patternArr.length === 0) { fills.push(null); return; }
      const pattern = patternArr[0];
      if (!pattern || pattern['@_patternType'] === 'none' || pattern['@_patternType'] === 'gray125') {
        fills.push(null); return;
      }
      const fgColor = Array.isArray(pattern.fgColor) ? pattern.fgColor[0] : pattern.fgColor;
      if (!fgColor) { fills.push(null); return; }

      // Direct RGB color
      const argb = fgColor['@_rgb'] || fgColor['@_argb'];
      if (argb) {
        const upper = argb.toUpperCase();
        if (upper === 'FFFFFFFF' || upper === 'FF000000' || upper === 'FFFFFF' || upper === '000000') {
          fills.push(null);
        } else {
          const rgb = argb.replace(/^FF/i, '').toUpperCase();
          fills.push(`#${rgb}`);
        }
        return;
      }

      // Theme color
      const themeIdx = fgColor['@_theme'];
      if (themeIdx !== undefined) {
        const baseHex = themeColors[parseInt(themeIdx, 10)];
        if (!baseHex) { fills.push(null); return; }
        const tint = parseFloat(fgColor['@_tint'] ?? '0');
        const finalHex = applyTint(baseHex, tint);
        // Skip near-white and near-black
        if (finalHex === 'FFFFFF' || finalHex === '000000') { fills.push(null); return; }
        fills.push(`#${finalHex}`);
        return;
      }

      fills.push(null);
    });
  }

  // Build cellXfs: xfId → fill color
  const xfs = [];
  const xfList = ss?.cellXfs?.xf;
  if (xfList) {
    const arr = Array.isArray(xfList) ? xfList : [xfList];
    arr.forEach(xf => {
      const fillId = parseInt(xf['@_fillId'] ?? '0', 10);
      xfs.push(fills[fillId] ?? null);
    });
  }
  return xfs;
}

// Convert Excel serial date to ISO string
function excelDateToISO(serial) {
  if (typeof serial !== 'number' || serial < 1) return serial;
  const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
  return date.toISOString();
}

// Parse sheet XML into rows with values and cell styles
function parseSheet(xml, sharedStrings, xfColors) {
  const parsed = xmlParser.parse(xml);
  const sheetData = parsed.worksheet?.sheetData;
  if (!sheetData) return [];

  const rawRows = Array.isArray(sheetData.row) ? sheetData.row : (sheetData.row ? [sheetData.row] : []);

  // Helper: column letter(s) to 0-based index
  function colToIndex(col) {
    let n = 0;
    for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n - 1;
  }

  // Helper: parse address like "A1" → { col: 0, row: 1 }
  function parseAddr(addr) {
    const m = addr.match(/^([A-Z]+)(\d+)$/);
    if (!m) return null;
    return { col: colToIndex(m[1]), row: parseInt(m[2], 10) };
  }

  const rows = [];

  for (const row of rawRows) {
    const rowNum = parseInt(row['@_r'], 10);
    const cells = Array.isArray(row.c) ? row.c : (row.c ? [row.c] : []);

    const rowData = { rowNum, values: {}, styles: {} };

    for (const cell of cells) {
      const addr = parseAddr(cell['@_r'] || '');
      if (!addr) continue;
      const colIdx = addr.col;
      const t = cell['@_t']; // type: s=shared string, str=formula string, b=boolean, n=number
      const s = parseInt(cell['@_s'] ?? '-1', 10); // style index

      let value;
      const v = cell.v;
      const f = cell.f;

      if (t === 's') {
        value = sharedStrings[parseInt(v, 10)] ?? null;
      } else if (t === 'str') {
        value = v !== undefined ? String(v) : (f ? { formula: typeof f === 'object' ? f['#text'] || f : f } : null);
      } else if (t === 'b') {
        value = v === '1' || v === 1;
      } else if (v !== undefined && v !== '') {
        const num = parseFloat(v);
        value = isNaN(num) ? String(v) : num;
      } else if (f) {
        value = { formula: typeof f === 'object' ? f['#text'] || String(f) : String(f) };
      } else {
        value = null;
      }

      rowData.values[colIdx] = value;
      rowData.styles[colIdx] = s >= 0 ? (xfColors[s] ?? null) : null;
    }

    rows.push(rowData);
  }

  return rows;
}

async function parseExcel(buffer, sheetName) {
  log('parse', `file size: ${buffer.length} bytes`);

  // List entries to find the right sheet
  const zip = unzipper.Open.buffer(buffer);
  const directory = await zip;

  const entryPaths = directory.files.map(f => f.path);
  log('parse', `zip entries: ${entryPaths.filter(p => !p.startsWith('xl/media/')).join(', ')}`);

  // Find workbook.xml to map sheet names → rId → sheet files
  const themeEntry = entryPaths.find(p => p.match(/xl\/theme\/theme\d+\.xml/)) || 'xl/theme/theme1.xml';
  const targets = ['xl/workbook.xml', 'xl/sharedStrings.xml', 'xl/styles.xml', 'xl/_rels/workbook.xml.rels', themeEntry];
  const sheets = entryPaths.filter(p => p.match(/^xl\/worksheets\/sheet\d+\.xml$/));
  targets.push(...sheets);

  const entries = await readXlsxEntries(buffer, targets);

  // Parse workbook to get sheet name → file mapping
  let sheetFileIndex = 0;
  if (entries['xl/workbook.xml'] && entries['xl/_rels/workbook.xml.rels']) {
    const wb = xmlParser.parse(entries['xl/workbook.xml']);
    const rels = xmlParser.parse(entries['xl/_rels/workbook.xml.rels']);

    const wbSheets = wb.workbook?.sheets?.sheet;
    const sheetArr = Array.isArray(wbSheets) ? wbSheets : (wbSheets ? [wbSheets] : []);
    log('parse', `sheets in workbook: ${sheetArr.map(s => s['@_name']).join(', ')}`);

    if (sheetName) {
      const idx = sheetArr.findIndex(s => s['@_name'] === sheetName);
      if (idx >= 0) sheetFileIndex = idx;
    }
  }

  const themeColors = parseThemeColors(entries[themeEntry]);
  log('parse', `theme colors: ${JSON.stringify(themeColors)}`);
  const sharedStrings = parseSharedStrings(entries['xl/sharedStrings.xml']);
  const xfColors = parseStyles(entries['xl/styles.xml'], themeColors);
  log('parse', `shared strings: ${sharedStrings.length}, xf styles: ${xfColors.length}`);

  const sheetFile = sheets[sheetFileIndex];
  if (!sheetFile || !entries[sheetFile]) throw new Error('Sheet not found in zip');
  log('parse', `parsing sheet file: ${sheetFile}`);

  const rawRows = parseSheet(entries[sheetFile], sharedStrings, xfColors);
  if (rawRows.length === 0) return [];

  // First row = headers
  const headerRow = rawRows[0];
  const maxCol = Math.max(...rawRows.map(r => Math.max(...Object.keys(r.values).map(Number))));
  const headers = [];
  for (let i = 0; i <= maxCol; i++) {
    const v = headerRow.values[i];
    headers[i] = (v !== null && v !== undefined && v !== '') ? String(v).trim() : `Col${i + 1}`;
  }
  log('parse', `headers (${headers.length}): ${JSON.stringify(headers.slice(0, 10))}...`);

  const result = [];
  for (let ri = 1; ri < rawRows.length; ri++) {
    const row = rawRows[ri];
    const rowObj = {};
    const cellStyles = {};

    for (let i = 0; i <= maxCol; i++) {
      const header = headers[i];
      let value = row.values[i] ?? null;

      // Detect Excel date serials (numbers in date columns — heuristic: > 40000 likely a date)
      if (typeof value === 'number' && value > 40000 && value < 100000) {
        value = excelDateToISO(value);
      }

      rowObj[header] = value;
      cellStyles[header] = { bg: row.styles[i] ?? null };
    }

    rowObj._cell_styles = cellStyles;
    result.push(rowObj);
  }

  log('parse', `done — ${result.length} rows`);
  return result;
}

// POST /parse-excel
app.post('/parse-excel', upload.single('file'), async (req, res) => {
  try {
    const buffer = req.file?.buffer ?? req.rawBody;
    log('parse-excel', `source: ${req.file ? 'form-data' : 'raw-binary'}, size: ${buffer?.length ?? 0}`);
    if (!buffer || buffer.length === 0) return res.status(400).json({ error: 'No file received.' });

    const rows = await parseExcel(buffer, req.query.sheet);
    res.json(rows);
  } catch (err) {
    log('parse-excel', `EXCEPTION: ${err.message}`);
    console.error(err.stack);
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
    if (!buffer || buffer.length === 0) return res.status(400).json({ error: 'No file received.' });

    const directory = await unzipper.Open.buffer(buffer);
    const entries = {};
    for (const file of directory.files) {
      if (file.path === 'xl/workbook.xml') {
        entries['xl/workbook.xml'] = (await file.buffer()).toString('utf8');
      }
    }
    const wb = xmlParser.parse(entries['xl/workbook.xml'] || '<workbook/>');
    const wbSheets = wb.workbook?.sheets?.sheet;
    const sheetArr = Array.isArray(wbSheets) ? wbSheets : (wbSheets ? [wbSheets] : []);
    res.json(sheetArr.map(s => s['@_name']));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3200;
app.listen(PORT, () => log('startup', `Promould Custom API running on port ${PORT}`));
