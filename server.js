'use strict';

const express = require('express');
const helmet = require('helmet');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { runConsolidation } = require('./src/engine');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const INVENTORY_PATH = path.join(DATA_DIR, 'inventory.json');
const MAX_FILE_SIZE = 25 * 1024 * 1024;
fs.mkdirSync(DATA_DIR, { recursive: true });

if (!process.env.APP_USER || !process.env.APP_PASSWORD) {
  console.error('APP_USER and APP_PASSWORD must be configured before starting the server.');
  process.exit(1);
}

app.disable('x-powered-by');
app.use(helmet({ contentSecurityPolicy: false }));
app.use(express.json({ limit: '1mb' }));

function safeEqual(a, b) {
  const aa = Buffer.from(String(a));
  const bb = Buffer.from(String(b));
  return aa.length === bb.length && crypto.timingSafeEqual(aa, bb);
}

function requireAuth(req, res, next) {
  const header = req.headers.authorization || '';
  const match = header.match(/^Basic\s+(.+)$/i);
  if (!match) {
    res.set('WWW-Authenticate', 'Basic realm="Shopee Order Processing", charset="UTF-8"');
    return res.status(401).send('Authentication required');
  }
  let decoded = '';
  try { decoded = Buffer.from(match[1], 'base64').toString('utf8'); } catch (_) {}
  const split = decoded.indexOf(':');
  const user = split >= 0 ? decoded.slice(0, split) : '';
  const password = split >= 0 ? decoded.slice(split + 1) : '';
  if (!safeEqual(user, process.env.APP_USER) || !safeEqual(password, process.env.APP_PASSWORD)) {
    res.set('WWW-Authenticate', 'Basic realm="Shopee Order Processing", charset="UTF-8"');
    return res.status(401).send('Invalid credentials');
  }
  next();
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_FILE_SIZE, files: 1 },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname || '').toLowerCase();
    if (!['.xlsx', '.xls', '.csv'].includes(ext)) return cb(new Error('Only .xlsx, .xls and .csv files are supported.'));
    cb(null, true);
  }
});

function workbookRows(buffer, filename, preferRaw = true) {
  const ext = path.extname(filename || '').toLowerCase();
  const workbook = XLSX.read(buffer, { type: 'buffer', raw: false, cellDates: false });
  if (!workbook.SheetNames.length) throw new Error('The uploaded file contains no worksheets.');
  const preferred = ext === '.csv' || !preferRaw ? workbook.SheetNames[0] :
    (workbook.SheetNames.find(n => n.toLowerCase() === 'raw') || workbook.SheetNames[0]);
  const sheet = workbook.Sheets[preferred];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: false });
}

function normalizeInventoryRows(rows) {
  if (!rows || rows.length < 2) throw new Error('Inventory file must contain a header row and at least one product.');
  const header = rows[0].map(v => String(v || '').replace(/^\uFEFF/, '').trim().toLowerCase());
  const skuIndex = header.findIndex(v => /^(sku\s*\/\s*item\s*code|sku|item\s*code)$/.test(v));
  const nameIndex = header.findIndex(v => /^(product\s*name|item\s*name|name)$/.test(v));
  const categoryIndex = header.findIndex(v => /^category(?:\s*\(.*\))?$/.test(v));
  const tagsIndex = header.findIndex(v => /^tags?$/.test(v));
  // Existing inventory export format is A=SKU, B=Product Name, C=Category, D=Tags.
  const indices = {
    sku: skuIndex >= 0 ? skuIndex : 0,
    name: nameIndex >= 0 ? nameIndex : 1,
    category: categoryIndex >= 0 ? categoryIndex : 2,
    tags: tagsIndex >= 0 ? tagsIndex : 3
  };
  const clean = [['SKU / Item Code', 'Product Name', 'Category (ID or Name)', 'Tags']];
  for (const row of rows.slice(1)) {
    const values = [indices.sku, indices.name, indices.category, indices.tags].map(i => String(row[i] ?? '').trim());
    if (!values[0] && !values[1]) continue;
    clean.push(values);
  }
  if (clean.length < 2) throw new Error('No usable Inventory rows were found.');
  return clean;
}

function loadInventory() {
  if (!fs.existsSync(INVENTORY_PATH)) return null;
  try { return JSON.parse(fs.readFileSync(INVENTORY_PATH, 'utf8')); }
  catch (error) { throw new Error('Saved Inventory Database could not be read: ' + error.message); }
}

app.get('/health', (_req, res) => res.json({ ok: true }));
app.use(requireAuth);
app.use(express.static(path.join(__dirname, 'public'), { index: 'index.html', etag: true, maxAge: '1h' }));

app.get('/api/status', (_req, res) => {
  const inventory = loadInventory();
  res.json({ inventoryReady: !!inventory, inventoryRows: inventory ? Math.max(0, inventory.length - 1) : 0 });
});

app.post('/api/inventory', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Choose an Inventory Excel or CSV file first.' });
    const rows = normalizeInventoryRows(workbookRows(req.file.buffer, req.file.originalname, false));
    const tempPath = INVENTORY_PATH + '.tmp';
    fs.writeFileSync(tempPath, JSON.stringify(rows));
    fs.renameSync(tempPath, INVENTORY_PATH);
    res.json({ ok: true, inventoryRows: rows.length - 1, filename: path.basename(req.file.originalname) });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.post('/api/process', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Choose an order Excel file first.' });
    const inventory = loadInventory();
    if (!inventory) return res.status(400).json({ error: 'Upload the Inventory Database before processing orders.' });
    const rows = workbookRows(req.file.buffer, req.file.originalname);
    const result = runConsolidation({ orderRows: rows, inventoryRows: inventory });
    if (!result.rows.length) return res.status(422).json({ error: 'No order line items were parsed. Check that the order file has Tracking Number, Order ID and Item Data in columns A-C.' });
    res.json({
      ok: true,
      filename: path.basename(req.file.originalname),
      columns: ['Order Code', 'Product', 'Qty', 'Tracking', 'Variation', 'SKU Reference No.', 'Matched SKU', '', 'Notes'],
      rows: result.rows,
      summary: result.stats
    });
  } catch (error) {
    console.error('Order processing failed:', error);
    res.status(400).json({ error: error.message || 'Processing failed.' });
  }
});

app.use((error, _req, res, _next) => {
  const message = error && error.message ? error.message : 'Unexpected server error.';
  res.status(400).json({ error: message });
});

app.listen(PORT, () => console.log(`Shopee Order Processing listening on ${PORT}`));
