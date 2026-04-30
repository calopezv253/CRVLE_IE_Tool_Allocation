'use strict';
require('dotenv').config();

// Keep process alive and surface hidden crashes
process.on('uncaughtException',  err => console.error('[UNCAUGHT]', err));
process.on('unhandledRejection', err => console.error('[UNHANDLED]', err));

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
    idleTimeoutMillis: 60 * 60 * 1000, // 1 hour
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
const CACHE_TTL   = 13 * 60 * 60 * 1000; // 13 hours (longer than the 12h warm-up refresh interval so the cache never expires between refreshes)

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
// Derives weeks entirely from Intel Calendar.csv (no DB query).
// The DB query for the latest WW was removed because the view has no index
// on (ww), causing a 60-second full-table scan on every page load.
app.get('/api/filters', (req2, res) => {
  try {
    const period = getDefaultPeriodFromCalendar();
    const currentWW = Number(String(period.ww).slice(-2));
    const year      = String(period.ww).slice(0, 4);

    // Build last 8 weeks from the calendar-resolved current WW
    const weeks = [];
    for (let i = 7; i >= 0; i--) {
      const w = currentWW - i;
      if (w > 0) weeks.push(`${year}${String(w).padStart(2, '0')}`);
    }

    // Intel calendar: Su=0, Mo=1, Tu=2, We=3, Th=4, Fr=5, Sa=6
    const days = ['0', '1', '2', '3', '4', '5', '6'];

    console.log(`[/api/filters] served from calendar (no DB) — ${weeks.length} weeks, current=${period.ww}`);
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

// ─── Changes result cache ─────────────────────────────────────────────────────
// Keyed by column mapping so different mappings don't share results.
const changesResultCache = new Map();
const CHANGES_CACHE_TTL  = 30 * 60 * 1000; // 30 minutes

// Compute up to `n` most-recent business-day periods going back from today.
// Uses Intel Calendar.csv so no DB round-trip is needed to find WW numbers.
function recentBusinessPeriods(n) {
  const calendarPath = path.join(__dirname, 'Intel Calendar.csv');
  const csv = fs.readFileSync(calendarPath, 'utf8');
  const periods = [];
  const cursor = new Date();
  // Safety cap: scan at most 30 calendar days back to find n business days
  for (let i = 0; i < 30 && periods.length < n; i++) {
    const dow = cursor.getDay();
    if (dow !== 0 && dow !== 6) { // skip Sun(0) / Sat(6)
      const ww = parseCalendarWeekForDate(cursor, csv);
      if (ww) {
        const year = cursor.getFullYear();
        periods.push({
          ww : `${year}${String(ww).padStart(2, '0')}`,
          day: String(dow),
        });
      }
    }
    cursor.setDate(cursor.getDate() - 1);
  }
  return periods; // newest first
}

// Fetch allocation data for one (ww, day) period.
// Re-uses queryCache so warm-up data is never fetched twice.
async function fetchPeriodMap(p, tCol, cCol, pCol, wCol, dCol, ww, day) {
  const cacheKey = `alloc|${ww}|${day}`;
  let rows = getCached(cacheKey);
  if (!rows) {
    const r = p.request();
    r.timeout = QUERY_TIMEOUT;
    r.input('ww',  sql.NVarChar, String(ww));
    r.input('day', sql.NVarChar, String(day));
    const result = await r.query(
      `SELECT ${tCol} AS MACHINE_NAME, ${cCol} AS CELL, ${pCol} AS PRODUCT
       FROM ${DB_VIEW} WHERE ${wCol} = @ww AND ${dCol} = @day`
    );
    rows = result.recordset;
    setCached(cacheKey, rows);
    console.log(`[/api/changes] fetched ${rows.length} rows for ${cacheKey}`);
  } else {
    console.log(`[/api/changes] cache hit for ${cacheKey}`);
  }

  const map = {};
  for (const row of rows) {
    const key = `${row.MACHINE_NAME}||${row.CELL}`;
    map[key] = row.PRODUCT == null ? null : String(row.PRODUCT).trim();
  }
  return map;
}

// GET /api/changes
// Returns allocation differences across the last 3 business days.
// Uses queryCache and changesResultCache for near-instant repeat calls.
app.get('/api/changes', async (req2, res) => {
  try {
    const tCol = colFromQuery(req2.query, 'toolCol');
    const cCol = colFromQuery(req2.query, 'cellCol');
    const pCol = colFromQuery(req2.query, 'productCol');
    const wCol = colFromQuery(req2.query, 'wwCol');
    const dCol = colFromQuery(req2.query, 'dayCol');

    // Return cached changes if still fresh
    const changeCacheKey = `${tCol}|${cCol}|${pCol}|${wCol}|${dCol}`;
    const cached = changesResultCache.get(changeCacheKey);
    if (cached && Date.now() - cached.ts < CHANGES_CACHE_TTL) {
      console.log('[/api/changes] changes cache hit');
      return res.json(cached.data);
    }

    // Compute the 3 most-recent business-day periods from the Intel Calendar
    const periods = recentBusinessPeriods(3);
    if (periods.length < 2) {
      return res.json({ changes: [], periods: [] });
    }

    const p = await getPool();

    // Fetch each period's data, re-using queryCache where possible
    const periodData = [];
    for (const period of periods) {
      const map = await fetchPeriodMap(p, tCol, cCol, pCol, wCol, dCol, period.ww, period.day);
      periodData.push({ ww: period.ww, day: period.day, map });
    }

    // Diff consecutive periods (newest first)
    const changes = [];
    for (let i = 0; i < periodData.length - 1; i++) {
      const newer = periodData[i];
      const older = periodData[i + 1];
      const allKeys = new Set([...Object.keys(newer.map), ...Object.keys(older.map)]);
      for (const key of allKeys) {
        const [machine, cell] = key.split('||');
        const oldProd = Object.prototype.hasOwnProperty.call(older.map, key) ? older.map[key] : null;
        const newProd = Object.prototype.hasOwnProperty.call(newer.map, key) ? newer.map[key] : null;
        if (oldProd !== newProd) {
          changes.push({
            machine, cell,
            oldProduct: oldProd, newProduct: newProd,
            fromWW: older.ww, fromDay: older.day,
            toWW  : newer.ww, toDay  : newer.day,
          });
        }
      }
    }

    const result = { changes, periods: periodData.map(pd => ({ ww: pd.ww, day: pd.day })) };
    changesResultCache.set(changeCacheKey, { data: result, ts: Date.now() });
    res.json(result);
  } catch (err) {
    console.error('[/api/changes]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── Fallback: serve SPA ─────────────────────────────────────────────────────
app.get('*', (req2, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ─── Cache Warm-up ───────────────────────────────────────────────────────────
// Runs an /api/allocations-equivalent query for the current WW/day and stores
// the result in `queryCache`. This way the first user request is instant
// instead of waiting up to 60s for the SQL view to scan.
async function warmCache(reason = 'scheduled') {
  try {
    const period = getDefaultPeriodFromCalendar();
    if (!period || !period.ww) {
      console.warn('[CacheWarm] No default period available, skipping');
      return;
    }
    const ww  = String(period.ww);
    const day = String(period.day);

    const tCol = safeCol(DEFAULT_COLS.toolCol);
    const cCol = safeCol(DEFAULT_COLS.cellCol);
    const pCol = safeCol(DEFAULT_COLS.productCol);
    const wCol = safeCol(DEFAULT_COLS.wwCol);
    const dCol = safeCol(DEFAULT_COLS.dayCol);
    const oCol = safeCol(DEFAULT_COLS.tosCol);

    const selectList = [tCol, cCol, pCol, wCol, dCol, oCol];

    const t0 = Date.now();
    const p  = await getPool();
    const r  = p.request();
    r.timeout = 120000; // 2 min — DB is slow without index; give it more time than QUERY_TIMEOUT
    r.input('ww',  sql.NVarChar, ww);
    r.input('day', sql.NVarChar, day);

    const result = await r.query(
      `SELECT ${selectList.join(', ')} FROM ${DB_VIEW} WHERE ${wCol} = @ww AND ${dCol} = @day ORDER BY ${tCol}, ${cCol}`
    );

    const cacheKey = `alloc|${ww}|${day}`;
    setCached(cacheKey, result.recordset);
    console.log(`[CacheWarm] (${reason}) ${result.recordset.length} rows for ${cacheKey} in ${Date.now() - t0}ms`);
  } catch (err) {
    console.error('[CacheWarm] Failed:', err.message);
    // Retry once after 2 minutes if this was startup warm-up
    if (reason === 'startup') {
      console.log('[CacheWarm] Will retry in 2 minutes...');
      setTimeout(() => warmCache('startup-retry'), 2 * 60 * 1000);
    }
  }
}

// Auto-refresh cache every 12 hours.
// NOTE: do NOT call .unref() — the interval must keep the Node.js process
// alive so the server doesn't exit after the warm-up completes.
const WARM_INTERVAL_MS = 12 * 60 * 60 * 1000; // 12 hours
setInterval(() => warmCache('12h-refresh'), WARM_INTERVAL_MS);

// Extra keep-alive: a lightweight heartbeat that prevents Node from exiting
// when there are no active DB connections in the pool.
setInterval(() => {}, 60 * 1000); // ping every 60s — no-op, just holds the event loop

// ─── Start ───────────────────────────────────────────────────────────────────
app.listen(PORT, async () => {
  console.log(`[Server] Listening on http://localhost:${PORT}`);
  try {
    await getPool();
    // Pre-warm cache in background — does not block server startup
    warmCache('startup');
  } catch (err) {
    console.error('[DB] Initial connection failed:', err.message);
  }
});
