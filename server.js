const express = require('express');
const multer = require('multer');
const mysql = require('mysql2');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');


// Config
const app = express();
const port = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());


// Multer: accept typical Excel types
const upload = multer({
  dest: 'uploads/',
  fileFilter: (req, file, cb) => {
    const allowedMimes = new Set([
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
      'application/vnd.ms-excel' // legacy .xls (many systems still use this)
    ]);
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowedMimes.has(file.mimetype) || ['.xlsx', '.xls', '.xlsm'].includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files (.xlsx/.xls) are allowed'), false);
    }
  }
});

// MySQL connection
const db = mysql.createConnection({
  host: process.env.DB_HOST || 'localhost',
  user: process.env.DB_USER || 'root',
  password: process.env.DB_PASS || 'Jackboy@12',
  database: process.env.DB_NAME || 'financial_db',
  multipleStatements: false
});

db.connect(err => {
  if (err) {
    console.error('Database connection failed: ' + err.stack);
    process.exit(1);
  }
  console.log('Connected to database');
});


// Helpers
const MONTHS = [
  'January','February','March','April','May','June',
  'July','August','September','October','November','December'
];

const indexToMonthName = (idx) => (idx >= 0 && idx <= 11 ? MONTHS[idx] : null);

/**
 * Convert various month cell forms to month index (0–11).
 * Accepts: "January", "Jan", "1", "01", Date cells, "2023-01-01"
 */
const monthStringToIndex = (val) => {
  if (val === null || val === undefined) return null;

  // Excel date -> JS Date if cellDates:true & raw:true
  if (val instanceof Date && !isNaN(val)) {
    return val.getMonth();
  }

  if (typeof val === 'number') {
    if (val >= 1 && val <= 12) return val - 1;
    if (val >= 0 && val <= 11) return val; // some sheets store 0-based
  }

  const s = String(val).trim().toLowerCase();

  // Try a parseable date string
  const d = new Date(s);
  if (!isNaN(d)) return d.getMonth();

  const cleaned = s.replace(/[.,]/g, '');
  // numeric string
  if (/^\d{1,2}$/.test(cleaned)) {
    const n = parseInt(cleaned, 10);
    if (n >= 1 && n <= 12) return n - 1;
  }

  const abbr = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,sept:8,oct:9,nov:10,dec:11 };
  const full = MONTHS.map(m => m.toLowerCase());

  if (full.includes(cleaned)) return full.indexOf(cleaned);
  if (abbr.hasOwnProperty(cleaned)) return abbr[cleaned];

  return null;
};

/** Strip currency symbols/commas; keep digits, minus, decimal */
const parseAmount = (val) => {
  if (val === null || val === undefined) return null;
  if (typeof val === 'number') return val;

  const s = String(val).trim().replace(/[^0-9.\-]/g, '');
  if (!s || s === '-' || s === '.' || s === '-.') return null;

  const num = parseFloat(s);
  return isNaN(num) ? null : num;
};

/**
 * Find header row and Month/Amount column indexes regardless of case/spacing/position.
 * rows: array-of-arrays (from sheet_to_json header:1)
 */
const findHeaderIndexes = (rows) => {
  for (let r = 0; r < rows.length && r < 30; r++) {
    const row = rows[r] || [];
    const norm = row.map(c => String(c ?? '').trim().toLowerCase().replace(/\s+/g, ''));

    // Accept common synonyms for robustness
    const monthIdx = norm.findIndex(x => ['month','months'].includes(x));
    const amountIdx = norm.findIndex(x => ['amount','value','amt'].includes(x));

    if (monthIdx !== -1 && amountIdx !== -1) {
      return { headerRow: r, monthIdx, amountIdx };
    }
  }
  return null;
};

// Utility to ensure uploaded file is removed (best-effort)
const cleanupFile = (filePath) => {
  try {
    if (filePath && fs.existsSync(filePath)) fs.unlinkSync(filePath);
  } catch {
    // ignore
  }
};


// Routes
app.get('/health', (_req, res) => {
  res.json({ ok: true, db: !!db.threadId, ts: new Date().toISOString() });
});

/**
 * Upload and ingest Excel
 * Path params: :userid, :year
 * Body: multipart/form-data with field 'file'
 */
app.post('/api/finances/upload/:userid/:year', upload.single('file'), (req, res) => {
  const done = (status, payload) => {
    cleanupFile(req.file?.path);
    return res.status(status).json(payload);
  };

  try {
    const userId = parseInt(req.params.userid, 10);
    const year = parseInt(req.params.year, 10);

    if (!req.file) return done(400, { error: 'No file uploaded' });
    if (!Number.isInteger(userId) || !Number.isInteger(year)) {
      return done(400, { error: 'Invalid user or year parameter' });
    }

    // Validate user exists
    db.query('SELECT * FROM users WHERE user_id = ?', [userId], (uErr, users) => {
      if (uErr) return done(500, { error: 'Database error (user lookup)' });
      if (users.length === 0) return done(404, { error: 'User not found' });

      // Read the Excel
      let workbook;
      try {
        workbook = XLSX.readFile(req.file.path, { cellDates: true, raw: true });
      } catch (e) {
        return done(400, { error: 'Failed to read Excel: ' + e.message });
      }

      const sheetName = workbook.SheetNames[0];
      const ws = workbook.Sheets[sheetName];
      if (!ws) return done(400, { error: 'No worksheet found' });

      // Rows as arrays so we can find the header manually
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
      if (!rows || rows.length === 0) return done(400, { error: 'Empty worksheet' });

      const headerInfo = findHeaderIndexes(rows);
      if (!headerInfo) {
        return done(400, { error: "Invalid Excel format. Could not find 'Month' and 'Amount' columns." });
      }

      const { headerRow, monthIdx, amountIdx } = headerInfo;
      const dataRows = rows.slice(headerRow + 1).filter(r => Array.isArray(r) && r.length > 0);

      // Normalize the data
      const normalized = [];
      for (const r of dataRows) {
        const rawMonth = r[monthIdx];
        const rawAmount = r[amountIdx];

        const mi = monthStringToIndex(rawMonth); // 0–11
        const amt = parseAmount(rawAmount);

        // Skip invalid/blank lines; optional: collect to report back
        if (mi === null || amt === null) continue;

        const monthNum = mi + 1; // store 1–12 in DB
        normalized.push({ monthNum, amount: amt });
      }

      if (normalized.length === 0) {
        return done(400, { error: 'No valid rows found. Ensure Month and Amount contain valid values.' });
      }

      // Insert with upsert; requires UNIQUE(user_id, year, month_num)
      const insertQuery = `
        INSERT INTO financial_records (user_id, year, month_num, amount)
        VALUES (?, ?, ?, ?)
        ON DUPLICATE KEY UPDATE amount = VALUES(amount), updated_at = CURRENT_TIMESTAMP
      `;

      const promises = normalized.map(({ monthNum, amount }) =>
        new Promise((resolve, reject) => {
          db.query(insertQuery, [userId, year, monthNum, amount], (iErr, result) => {
            if (iErr) return reject(iErr);
            resolve(result);
          });
        })
      );

      Promise.all(promises)
        .then(() => done(200, { message: 'Data uploaded successfully' }))
        .catch(e => done(500, { error: 'Error inserting data: ' + e.message }));
    });
  } catch (e) {
    return done(500, { error: e.message });
  }
});

/**
 * Fetch data for a user/year
 * Returns pretty month names; ordered by month_num
 */
app.get('/api/finances/:userid/:year', (req, res) => {
  const userId = req.params.userid;
  const year = req.params.year;

  const query = `
    SELECT fr.month_num, fr.amount, u.name
    FROM financial_records fr
    JOIN users u ON fr.user_id = u.user_id
    WHERE fr.user_id = ? AND fr.year = ?
    ORDER BY fr.month_num ASC
  `;

  db.query(query, [userId, year], (err, results) => {
    if (err) return res.status(500).json({ error: 'Database error' });
    if (results.length === 0) return res.status(404).json({ error: 'No data found for this user and year' });

    res.json({
      name: results[0].name,
      year,
      records: results.map(r => ({
        month: indexToMonthName((r.month_num || 1) - 1),
        amount: Number(r.amount)
      }))
    });
  });
});


// Start server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
