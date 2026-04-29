'use strict';
require('dotenv').config();

const express = require('express');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');
const sql     = require('mssql');
const zlib    = require('zlib');

const app  = express();
const PORT = process.env.PORT || 3000;

// ─── DB Config ──────────────────────────────────────────────────────────────
const dbConfig = {
  server  : process.env.DB_SERVER,
  port    : parseInt(process.env.DB_PORT, 10),
  database: process.env.DB_NAME,
  user    : process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  options : {
    encrypt               : true,    // server requires TLS encryption
    trustServerCertificate: true,    // accept self-signed / internal certs
    enableArithAbort      : true,
  },
  requestTimeout: 60000,
  connectionTimeout: 30000,
  pool: {
    max    : 10,
    min    : 0,
    idleTimeoutMillis: 30000,
  },
};

// Default query timeout (ms) – views without indexes can be slow
const QUERY_TIMEOUT = 60000;
dbConfig.requestTimeout = QUERY_TIMEOUT;

const DB_VIEW = process.env.DB_VIEW || 'ops.vw_SUM_PRODUCT_ALLOCATION';

let pool = null;

async function getPool() {
  if (!pool) {
    pool = await sql.connect(dbConfig);
    console.log(`[DB] Connected to ${dbConfig.server}:${dbConfig.port} / ${dbConfig.database}`);
  }
  return pool;
}

// ─── In-memory query cache ────────────────────────────────────────────────────
// Stores { data, ts } per cache key to avoid repeated DB queries for the
// same WW/day within the TTL window.
const queryCache  = new Map();
const CACHE_TTL   = 5 * 60 * 1000; // 5 minutes

function getCached(key) {
  const entry = queryCache.get(key);
  if (!entry) return null;
  if (Date.now() - entry.ts > CACHE_TTL) { queryCache.delete(key); return null; }
  return entry.data;
}
function setCached(key, data) {
  queryCache.set(key, { data, ts: Date.now() });
}

// ─── Middleware ──────────────────────────────────────────────────────────────
// Gzip middleware using built-in zlib (no external dependency)
app.use((req, res, next) => {
  const ae = req.headers['accept-encoding'] || '';
  if (!ae.includes('gzip')) return next();
  const _json = res.json.bind(res);
  res.json = (body) => {
    const buf = Buffer.from(JSON.stringify(body));
    zlib.gzip(buf, (err, compressed) => {
      if (err) return _json(body);
      res.setHeader('Content-Encoding', 'gzip');
      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Content-Length', compressed.length);
      res.end(compressed);
    });
  };
  next();
});
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Helper: validate column name (prevent SQL injection via dynamic names) ──
const VALID_IDENT = /^[A-Za-z_][A-Za-z0-9_]*$/;
function safeCol(name) {
  if (!name || !VALID_IDENT.test(name)) {
    throw new Error(`Invalid column name: "${name}"`);
  }
  return `[${name}]`;
}

const DEFAULT_COLS = {
  toolCol: 'MACHINE_NAME',
  cellCol: 'CELL',
  productCol: 'PRODUCT',
  wwCol: 'ww',
  dayCol: 'day_',
  tosCol: 'TESTER_NAME',
};

function colFromQuery(query, key, optional = false) {
  const col = query[key] || DEFAULT_COLS[key];
  if (optional && !col) return null;
  return safeCol(col);
}

function parseCalendarWeekForDate(date, csvText) {
  const dayOfMonth = date.getDate();
  const dow = date.getDay(); // Su=0, Mo=1, Tu=2, We=3, Th=4, Fr=5, Sa=6

  const lines = String(csvText || '').split(/\r?\n/);

  // Scan all lines to find the day in any WW row
  for (const line of lines) {
    const cells = line.split(',').map(c => String(c || '').replace(/\u00a0/g, ' ').trim());
    const first = cells[0] || '';

    // Extract WW number from first cell (may contain just digits, e.g., "18" or "WW,18")
    const wwText = first.replace(/\D/g, '');
    if (!wwText) continue;

    // Get the day value at the correct column position
    // Column indices: 0=WW, 1=Sunday, 2=Monday, 3=Tuesday, etc.
    // dow=0 (Sunday) maps to cells[1], dow=1 (Monday) maps to cells[2], etc.
    const dayCell = String(cells[dow + 1] || '').replace(/\D/g, '');
    if (!dayCell) continue;

    // Check if this day matches the date we're looking for
    if (Number(dayCell) === dayOfMonth) {
      const ww = Number(wwText);
      if (!Number.isFinite(ww) || ww <= 0) continue;
      return ww;
    }
  }

  return null;
}

function getDefaultPeriodFromCalendar() {
  const now = new Date();
  const calendarPath = path.join(__dirname, 'Intel Calendar.csv');
  const csv = fs.readFileSync(calendarPath, 'utf8');
  const ww = parseCalendarWeekForDate(now, csv);

  if (!ww) {
    throw new Error('Could not resolve current WW from Intel Calendar.csv');
  }

  const year = now.getFullYear();
  return {
    ww: `${year}${String(ww).padStart(2, '0')}`,
    day: String(now.getDay()),
    source: 'intel-calendar',
  };
}

// ─── Routes ─────────────────────────────────────────────────────────────────

// GET /api/schema
// Returns all column names from the allocation view so the UI can map them.
app.get('/api/schema', async (req2, res) => {
  try {
    const p = await getPool();
    const r = p.request();
    r.timeout = QUERY_TIMEOUT;
    const result = await r.query(`SELECT TOP 0 * FROM ${DB_VIEW}`);
    const columns = result.recordset.columns
      ? Object.keys(result.recordset.columns)
      : Object.keys(result.recordset[0] || {});
    res.json({ columns });
  } catch (err) {
    console.error('[/api/schema]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/filters?wwCol=ww&dayCol=day_
// Uses TOP 1 ORDER BY ww DESC to get latest WW (fast even without index).
// Days 1-7 are returned without a DB query.
app.get('/api/filters', async (req2, res) => {
  try {
    const wCol = colFromQuery(req2.query, 'wwCol');
    const p    = await getPool();

    const r = p.request();
    r.timeout = QUERY_TIMEOUT;
    const latestRes = await r.query(
      `SELECT TOP 1 ${wCol} AS ww FROM ${DB_VIEW}
       WHERE ${wCol} IS NOT NULL ORDER BY ${wCol} DESC`);

    let weeks = [];
    // Intel calendar mapping: Su=0, Mo=1, Tu=2, We=3, Th=4, Fr=5, Sa=6
    const days = ['0','1','2','3','4','5','6'];

    if (latestRes.recordset.length) {
      const latestWW = Number(latestRes.recordset[0].ww);
      for (let i = 7; i >= 0; i--) { const w = latestWW - i; if (w > 0) weeks.push(String(w)); }
    }

    res.json({ weeks, days });
  } catch (err) {
    console.error('[/api/filters]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/default-period
// Uses Intel Calendar.csv + today's date to return default WW/day.
app.get('/api/default-period', async (req2, res) => {
  try {
    const period = getDefaultPeriodFromCalendar();
    res.json(period);
  } catch (err) {
    console.error('[/api/default-period]', err.message);
    res.status(500).json({ error: err.message });
  }
});


// GET /api/machines?toolCol=MACHINE_NAME&wwCol=ww&dayCol=day_&ww=17&day=2
app.get('/api/machines', async (req2, res) => {
  const { ww, day } = req2.query;
  try {
    const tCol    = colFromQuery(req2.query, 'toolCol');
    const wCol    = colFromQuery(req2.query, 'wwCol');
    const dCol    = colFromQuery(req2.query, 'dayCol');
    const p       = await getPool();
    const r       = p.request();
    r.timeout     = QUERY_TIMEOUT;
    const where   = [`${tCol} IS NOT NULL`];
    if (ww)  { r.input('ww',  sql.NVarChar, String(ww));  where.push(`${wCol} = @ww`);  }
    if (day) { r.input('day', sql.NVarChar, String(day)); where.push(`${dCol} = @day`); }
    const result = await r.query(
      `SELECT DISTINCT ${tCol} AS machine FROM ${DB_VIEW}
       WHERE ${where.join(' AND ')} ORDER BY machine`
    );
    res.json({ machines: result.recordset.map(r => r.machine) });
  } catch (err) {
    console.error('[/api/machines]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/allocations
app.get('/api/allocations', async (req2, res) => {
  const { ww, day } = req2.query;
  let tools = req2.query.tools;

  try {
    const tCol = colFromQuery(req2.query, 'toolCol');
    const cCol = colFromQuery(req2.query, 'cellCol');
    const pCol = colFromQuery(req2.query, 'productCol');
    const wCol = colFromQuery(req2.query, 'wwCol');
    const dCol = colFromQuery(req2.query, 'dayCol');
    const oCol = colFromQuery(req2.query, 'tosCol', true);

    const selectList = [tCol, cCol, pCol, wCol, dCol];
    if (oCol) selectList.push(oCol);

    const p = await getPool();
    const r = p.request();
    r.timeout = QUERY_TIMEOUT;
    const where = [];

    if (ww !== undefined && ww !== '') {
      r.input('ww', sql.NVarChar, String(ww));
      where.push(`${wCol} = @ww`);
    }
    if (day !== undefined && day !== '') {
      r.input('day', sql.NVarChar, String(day));
      where.push(`${dCol} = @day`);
    }

    if (tools) {
      const toolList = (Array.isArray(tools) ? tools : tools.split(',')).map(t => t.trim()).filter(Boolean);
      if (toolList.length > 0) {
        const params = toolList.map((t, i) => {
          r.input(`tool${i}`, sql.NVarChar, t);
          return `@tool${i}`;
        });
        where.push(`${tCol} IN (${params.join(', ')})`);
      }
    }

    const whereClause = where.length ? `WHERE ${where.join(' AND ')}` : '';

    // Serve from cache when no tool filter is active (tool filter data is
    // a subset of the full WW/day result, so we skip caching those slices).
    const cacheKey = `alloc|${ww || ''}|${day || ''}`;
    const useCache = !tools; // only cache unfiltered (full) loads
    if (useCache) {
      const cached = getCached(cacheKey);
      if (cached) {
        console.log(`[/api/allocations] cache hit for ${cacheKey}`);
        return res.json({ data: cached });
      }
    }

    const result = await r.query(
      `SELECT ${selectList.join(', ')} FROM ${DB_VIEW} ${whereClause} ORDER BY ${tCol}, ${cCol}`
    );

    if (useCache) setCached(cacheKey, result.recordset);
    console.log(`[/api/allocations] ${result.recordset.length} rows for ${cacheKey}`);
    res.json({ data: result.recordset });
  } catch (err) {
    console.error('[/api/allocations]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/product
app.get('/api/product', async (req2, res) => {
  const { q, ww, day } = req2.query;
  if (!q) return res.json({ data: [] });

  try {
    const tCol = colFromQuery(req2.query, 'toolCol');
    const cCol = colFromQuery(req2.query, 'cellCol');
    const pCol = colFromQuery(req2.query, 'productCol');
    const wCol = colFromQuery(req2.query, 'wwCol');
    const dCol = colFromQuery(req2.query, 'dayCol');
    const oCol = colFromQuery(req2.query, 'tosCol', true);

    const selectList = [tCol, cCol, pCol, wCol, dCol];
    if (oCol) selectList.push(oCol);

    const p = await getPool();
    const r = p.request();
    r.timeout = QUERY_TIMEOUT;
    r.input('q', sql.NVarChar, `%${q}%`);

    const where = [`${pCol} LIKE @q`];
    if (ww  && ww  !== '') { r.input('ww',  sql.NVarChar, String(ww));  where.push(`${wCol} = @ww`);  }
    if (day && day !== '') { r.input('day', sql.NVarChar, String(day)); where.push(`${dCol} = @day`); }

    const result = await r.query(
      `SELECT ${selectList.join(', ')} FROM ${DB_VIEW}
       WHERE ${where.join(' AND ')} ORDER BY ${tCol}, ${cCol}`
    );
    res.json({ data: result.recordset });
  } catch (err) {
    console.error('[/api/product]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/export
app.get('/api/export', async (req2, res) => {
  const { ww, day } = req2.query;
  let tools = req2.query.tools;

  try {
    const tCol = colFromQuery(req2.query, 'toolCol');
    const cCol = colFromQuery(req2.query, 'cellCol');
    const pCol = colFromQuery(req2.query, 'productCol');
    const wCol = colFromQuery(req2.query, 'wwCol');
    const dCol = colFromQuery(req2.query, 'dayCol');
    const oCol = colFromQuery(req2.query, 'tosCol', true);

    const selectList = [tCol, cCol, pCol, wCol, dCol];
    if (oCol) selectList.push(oCol);

    const p = await getPool();
    const r = p.request();
    r.timeout = QUERY_TIMEOUT;
    const where = [];

    if (ww  && ww  !== '') { r.input('ww',  sql.NVarChar, String(ww));  where.push(`${wCol} = @ww`);  }
    if (day && day !== '') { r.input('day', sql.NVarChar, String(day)); where.push(`${dCol} = @day`); }
    if (tools) {
      const toolList = (Array.isArray(tools) ? tools : tools.split(',')).map(t => t.trim()).filter(Boolean);
      if (toolList.length > 0) {
        const params = toolList.map((t, i) => { r.input(`tool${i}`, sql.NVarChar, t); return `@tool${i}`; });
        where.push(`${tCol} IN (${params.join(', ')})`);
      }
    }

    const whereClause = where.length ? `WHERE ${where.join(' AND ')}` : '';
    const result = await r.query(
      `SELECT ${selectList.join(', ')} FROM ${DB_VIEW} ${whereClause} ORDER BY ${tCol}, ${cCol}`
    );
    const rows = result.recordset;
    if (!rows.length) return res.status(200).send('No data');

    const headers = Object.keys(rows[0]);
    const csvRows = [headers.join(',')];
    for (const row of rows) {
      csvRows.push(headers.map(h => `"${(row[h] == null ? '' : String(row[h])).replace(/"/g, '""')}"`).join(','));
    }

    const filename = `tool_allocation_ww${ww || 'all'}_day${day || 'all'}.csv`;
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(csvRows.join('\r\n'));
  } catch (err) {
    console.error('[/api/export]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── Fallback: serve SPA ─────────────────────────────────────────────────────
app.get('*', (req2, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ─── Start ───────────────────────────────────────────────────────────────────
app.listen(PORT, async () => {
  console.log(`[Server] Listening on http://localhost:${PORT}`);
  try {
    await getPool();
  } catch (err) {
    console.error('[DB] Initial connection failed:', err.message);
  }
});
