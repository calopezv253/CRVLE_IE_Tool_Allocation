'use strict';
require('dotenv').config();

// Keep process alive and surface hidden crashes
process.on('uncaughtException',  err => console.error('[UNCAUGHT]', err));
process.on('unhandledRejection', err => console.error('[UNHANDLED]', err));

const express    = require('express');
const cors       = require('cors');
const path       = require('path');
const fs         = require('fs');
const sql        = require('mssql');
const zlib       = require('zlib');
const nodemailer = require('nodemailer');
const cron       = require('node-cron');
const os         = require('os');
const { execSync } = require('child_process');

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
    // CRITICAL: mssql pool emits 'error' events when connections drop.
    // Without a listener, Node.js treats it as an unhandled error and exits.
    pool.on('error', err => {
      console.error('[Pool] Connection error (pool will auto-reconnect):', err.message);
      pool = null; // force reconnect on next request
    });
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
    // Include only the immediately next WW so forecast slots are selectable.
    weeks.push(`${year}${String(currentWW + 1).padStart(2, '0')}`);

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

// GET /api/forecast-slots
// Returns today's WW/day/date plus all future slots:
//   • remaining days of the current WW after today (e.g. WW19.6 when today is WW19.5)
//   • all 7 days (Su–Sa) of the next WW
app.get('/api/forecast-slots', (req2, res) => {
  try {
    const calendarPath = path.join(__dirname, 'Intel Calendar.csv');
    const csv = fs.readFileSync(calendarPath, 'utf8');
    const todayPeriod  = getDefaultPeriodFromCalendar();
    const todayDateStr = new Date().toISOString().slice(0, 10);
    const currentWW    = Number(String(todayPeriod.ww).slice(-2));
    const nextWW       = currentWW + 1;
    const currentWWStr = String(todayPeriod.ww);
    const nextWWStr    = `${String(todayPeriod.ww).slice(0, 4)}${String(nextWW).padStart(2, '0')}`;

    const slots = [];
    const cursor = new Date();
    cursor.setDate(cursor.getDate() + 1); // start from tomorrow

    // Collect remaining current-WW days + all next-WW days (up to 28 days look-ahead)
    let nextWWCount = 0;
    for (let safety = 0; safety < 28; safety++) {
      const ww = parseCalendarWeekForDate(cursor, csv);
      if (ww === currentWW) {
        // Remaining day(s) in the current WW after today
        slots.push({
          ww  : currentWWStr,
          day : String(cursor.getDay()),
          date: cursor.toISOString().slice(0, 10),
        });
      } else if (ww === nextWW) {
        slots.push({
          ww  : nextWWStr,
          day : String(cursor.getDay()),
          date: cursor.toISOString().slice(0, 10),
        });
        nextWWCount++;
        if (nextWWCount >= 7) break; // collected the full next WW
      }
      cursor.setDate(cursor.getDate() + 1);
    }

    res.json({
      today: { ww: todayPeriod.ww, day: todayPeriod.day, date: todayDateStr },
      slots,
    });
  } catch (err) {
    console.error('[/api/forecast-slots]', err.message);
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

// In-flight deduplication maps:
//   pendingPeriodFetches — one Promise per (cacheKey) period fetch
//   pendingChangesCalc  — one Promise per (changeCacheKey) /api/changes computation
// This prevents N simultaneous panel requests from each firing their own DB queries.
const pendingPeriodFetches = new Map();
const pendingChangesCalc   = new Map();

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
// Uses pendingPeriodFetches to deduplicate concurrent requests for the same period,
// so multiple simultaneous callers share a single DB query instead of each launching one.
function fetchPeriodMap(p, tCol, cCol, pCol, wCol, dCol, ww, day) {
  const cacheKey = `alloc|${ww}|${day}`;

  // Return immediately if already in queryCache
  const cached = getCached(cacheKey);
  if (cached) {
    console.log(`[/api/changes] cache hit for ${cacheKey}`);
    const map = {};
    for (const row of cached) {
      const key = `${row.MACHINE_NAME}||${row.CELL}`;
      map[key] = row.PRODUCT == null ? null : String(row.PRODUCT).trim();
    }
    return Promise.resolve(map);
  }

  // Return the existing in-flight promise if one is already running
  if (pendingPeriodFetches.has(cacheKey)) {
    console.log(`[/api/changes] dedup in-flight fetch for ${cacheKey}`);
    return pendingPeriodFetches.get(cacheKey);
  }

  // Start a new fetch and register it so concurrent callers can share it
  const promise = (async () => {
    try {
      const r = p.request();
      r.timeout = 120000; // 120 s — each query gets its own slot
      r.input('ww',  sql.NVarChar, String(ww));
      r.input('day', sql.NVarChar, String(day));
      const result = await r.query(
        `SELECT ${tCol} AS MACHINE_NAME, ${cCol} AS CELL, ${pCol} AS PRODUCT
         FROM ${DB_VIEW} WHERE ${wCol} = @ww AND ${dCol} = @day`
      );
      const rows = result.recordset;
      setCached(cacheKey, rows);
      console.log(`[/api/changes] fetched ${rows.length} rows for ${cacheKey}`);
      const map = {};
      for (const row of rows) {
        const key = `${row.MACHINE_NAME}||${row.CELL}`;
        map[key] = row.PRODUCT == null ? null : String(row.PRODUCT).trim();
      }
      return map;
    } finally {
      pendingPeriodFetches.delete(cacheKey);
    }
  })();

  pendingPeriodFetches.set(cacheKey, promise);
  return promise;
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

    // If another request is already computing changes for this mapping, share its result
    if (pendingChangesCalc.has(changeCacheKey)) {
      console.log('[/api/changes] dedup in-flight changes computation');
      const result = await pendingChangesCalc.get(changeCacheKey);
      return res.json(result);
    }

    // Compute the 3 most-recent business-day periods from the Intel Calendar
    const periods = recentBusinessPeriods(3);
    if (periods.length < 2) {
      return res.json({ changes: [], periods: [] });
    }

    const p = await getPool();

    const computePromise = (async () => {
      try {
        // Fetch all periods in parallel — cuts wall-time from sum(queries) to max(query)
        // pendingPeriodFetches ensures each period is fetched only once even under parallel load
        const maps = await Promise.all(
          periods.map(period => fetchPeriodMap(p, tCol, cCol, pCol, wCol, dCol, period.ww, period.day))
        );
        const periodData = periods.map((period, i) => ({ ww: period.ww, day: period.day, map: maps[i] }));

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
        return result;
      } finally {
        pendingChangesCalc.delete(changeCacheKey);
      }
    })();

    pendingChangesCalc.set(changeCacheKey, computePromise);

    const result = await computePromise;
    res.json(result);
  } catch (err) {
    console.error('[/api/changes]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── Planned Conversions ──────────────────────────────────────────────────────
// Stored in planned-conversions.json in the project root (server-side, shared).
const PLANNED_FILE = path.join(__dirname, 'planned-conversions.json');

function readPlanned() {
  try {
    if (!fs.existsSync(PLANNED_FILE)) return [];
    return JSON.parse(fs.readFileSync(PLANNED_FILE, 'utf8'));
  } catch { return []; }
}
function writePlanned(data) {
  fs.writeFileSync(PLANNED_FILE, JSON.stringify(data, null, 2), 'utf8');
}

// Validate that an ID is a safe numeric-only string (timestamp-based)
const VALID_ID = /^\d{1,20}$/;

// GET /api/planned-conversions
app.get('/api/planned-conversions', (req2, res) => {
  try {
    res.json(readPlanned());
  } catch (err) {
    console.error('[/api/planned-conversions GET]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/planned-conversions
app.post('/api/planned-conversions', (req2, res) => {
  try {
    const { area, tool, cell, newProduct, date, notes } = req2.body;
    if (!area || !tool || !cell || !newProduct || !date) {
      return res.status(400).json({ error: 'area, tool, cell, newProduct y date son obligatorios.' });
    }
    // Sanitise inputs — no SQL here (flat-file storage), but trim/truncate for safety.
    const entry = {
      id        : Date.now().toString(),
      area      : String(area).trim().slice(0, 64),
      tool      : String(tool).trim().slice(0, 128),
      cell      : String(cell).trim().slice(0, 32),
      newProduct: String(newProduct).trim().slice(0, 256),
      date      : String(date).trim().slice(0, 10),
      notes     : String(notes || '').trim().slice(0, 512),
      createdAt : new Date().toISOString(),
    };
    const list = readPlanned();
    list.push(entry);
    writePlanned(list);
    console.log(`[/api/planned-conversions] added id=${entry.id} tool=${entry.tool} cell=${entry.cell}`);
    res.status(201).json(entry);
  } catch (err) {
    console.error('[/api/planned-conversions POST]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// DELETE /api/planned-conversions/:id
app.delete('/api/planned-conversions/:id', (req2, res) => {
  try {
    const { id } = req2.params;
    if (!VALID_ID.test(id)) return res.status(400).json({ error: 'Invalid ID.' });
    const list = readPlanned();
    const idx = list.findIndex(e => e.id === id);
    if (idx === -1) return res.status(404).json({ error: 'Entry not found.' });
    list.splice(idx, 1);
    writePlanned(list);
    console.log(`[/api/planned-conversions] deleted id=${id}`);
    res.json({ ok: true });
  } catch (err) {
    console.error('[/api/planned-conversions DELETE]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── Disabled Tools ───────────────────────────────────────────────────────────
// Stored in disabled-tools.json in the project root (server-side, shared).
const DISABLED_TOOLS_FILE = path.join(__dirname, 'disabled-tools.json');

function readDisabledTools() {
  try {
    if (!fs.existsSync(DISABLED_TOOLS_FILE)) return [];
    const parsed = JSON.parse(fs.readFileSync(DISABLED_TOOLS_FILE, 'utf8'));
    if (!Array.isArray(parsed)) return [];
    return parsed;
  } catch { return []; }
}
function writeDisabledTools(list) {
  fs.writeFileSync(DISABLED_TOOLS_FILE, JSON.stringify(list, null, 2), 'utf8');
}

// GET /api/disabled-tools
app.get('/api/disabled-tools', (req2, res) => {
  try {
    res.json(readDisabledTools());
  } catch (err) {
    console.error('[/api/disabled-tools GET]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// PUT /api/disabled-tools  — body: { tools: ["id1","id2",...] }
app.put('/api/disabled-tools', (req2, res) => {
  try {
    const { tools } = req2.body;
    if (!Array.isArray(tools)) return res.status(400).json({ error: 'tools must be an array.' });
    // Sanitise: only keep non-empty strings, max 256 chars each
    const clean = tools
      .filter(t => t && typeof t === 'string')
      .map(t => t.trim().slice(0, 256))
      .filter(Boolean);
    writeDisabledTools(clean);
    console.log(`[/api/disabled-tools] saved ${clean.length} disabled tools`);
    res.json({ ok: true, count: clean.length });
  } catch (err) {
    console.error('[/api/disabled-tools PUT]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── Manual email trigger (test endpoint) ───────────────────────────────────
app.get('/api/send-report-now', (req2, res) => {
  res.json({ ok: true, message: 'Report triggered — check server logs for details.' });
  // Run async; errors are logged inside sendDailyReport
  sendDailyReport().catch(err => console.error('[/api/send-report-now]', err.message));
});

// ─── Fallback: serve SPA ─────────────────────────────────────────────────────
app.get('*', (req2, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ─── Cache Warm-up ───────────────────────────────────────────────────────────
// Fetches allocation data for one (ww, day) period and stores in queryCache.
async function warmOnePeriod(p, ww, day, reason) {
  const cacheKey = `alloc|${ww}|${day}`;
  if (getCached(cacheKey)) {
    console.log(`[CacheWarm] (${reason}) already cached for ${cacheKey}`);
    return;
  }
  const tCol = safeCol(DEFAULT_COLS.toolCol);
  const cCol = safeCol(DEFAULT_COLS.cellCol);
  const pCol = safeCol(DEFAULT_COLS.productCol);
  const wCol = safeCol(DEFAULT_COLS.wwCol);
  const dCol = safeCol(DEFAULT_COLS.dayCol);
  const oCol = safeCol(DEFAULT_COLS.tosCol);
  const selectList = [tCol, cCol, pCol, wCol, dCol, oCol];

  const t0 = Date.now();
  const r  = p.request();
  r.timeout = 120000;
  r.input('ww',  sql.NVarChar, String(ww));
  r.input('day', sql.NVarChar, String(day));
  const result = await r.query(
    `SELECT ${selectList.join(', ')} FROM ${DB_VIEW} WHERE ${wCol} = @ww AND ${dCol} = @day ORDER BY ${tCol}, ${cCol}`
  );
  setCached(cacheKey, result.recordset);
  console.log(`[CacheWarm] (${reason}) ${result.recordset.length} rows for ${cacheKey} in ${Date.now() - t0}ms`);
}

// Warm the current period + up to 2 previous business days so /api/changes
// can diff without any cold queries.
async function warmCache(reason = 'scheduled') {
  try {
    const periods = recentBusinessPeriods(3);
    if (!periods.length) {
      console.warn('[CacheWarm] No periods resolved from calendar, skipping');
      return;
    }
    const p = await getPool();
    // Warm all periods in parallel so startup completes faster
    await Promise.all(periods.map(period => warmOnePeriod(p, period.ww, period.day, reason)));
  } catch (err) {
    console.error('[CacheWarm] Failed:', err.message);
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

// ─── Daily 6 PM Email Report ─────────────────────────────────────────────────
// Computes product-change diffs for the last 24 h and sends an HTML table email.

const DAY_NAMES = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

function buildChangesEmailHtml(changes, periods, reportDate) {
  const dateStr = reportDate.toLocaleDateString('es-MX', {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric',
  });

  const periodLabels = periods.map(p => `WW${String(p.ww).slice(-2)} / ${DAY_NAMES[Number(p.day)] || p.day}`).join(' → ');

  let rows = '';
  if (!changes.length) {
    rows = `<tr><td colspan="6" style="text-align:center;padding:16px;color:#6b7280;font-style:italic;">
      Sin cambios de producto registrados en las últimas 24 horas.
    </td></tr>`;
  } else {
    for (const c of changes) {
      const fromLabel = `WW${String(c.fromWW).slice(-2)}/Día ${c.fromDay}`;
      const toLabel   = `WW${String(c.toWW).slice(-2)}/Día ${c.toDay}`;
      const oldColor  = c.oldProduct ? '#374151' : '#9ca3af';
      const newColor  = c.newProduct ? '#0052cc' : '#9ca3af';
      rows += `
        <tr style="border-bottom:1px solid #e5e7eb;">
          <td style="padding:8px 12px;font-weight:600;color:#111827;">${c.machine || '—'}</td>
          <td style="padding:8px 12px;color:#374151;">${c.cell || '—'}</td>
          <td style="padding:8px 12px;color:${oldColor};font-style:${c.oldProduct ? 'normal' : 'italic'};">${c.oldProduct || 'Empty'}</td>
          <td style="padding:8px 12px;color:#6b7280;text-align:center;">→</td>
          <td style="padding:8px 12px;color:${newColor};font-weight:600;font-style:${c.newProduct ? 'normal' : 'italic'};">${c.newProduct || 'Empty'}</td>
          <td style="padding:8px 12px;color:#6b7280;font-size:12px;">${fromLabel} → ${toLabel}</td>
        </tr>`;
    }
  }

  return `<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,Helvetica,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f3f4f6;padding:32px 0;">
    <tr><td align="center">
      <table width="700" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,.12);">

        <!-- Header -->
        <tr>
          <td style="background:linear-gradient(135deg,#0052cc 0%,#0071C5 100%);padding:28px 32px;">
            <p style="margin:0;font-size:13px;color:#bfdbfe;letter-spacing:1px;text-transform:uppercase;">Intel — CRVLE Operations Planning</p>
            <h1 style="margin:8px 0 0;font-size:22px;color:#ffffff;font-weight:700;">
              Reporte Diario de Cambios de Producto
            </h1>
            <p style="margin:6px 0 0;font-size:13px;color:#93c5fd;">${dateStr}</p>
          </td>
        </tr>

        <!-- Greeting -->
        <tr>
          <td style="padding:28px 32px 12px;">
            <p style="margin:0;font-size:15px;color:#1f2937;line-height:1.6;">
              Estimado equipo de Ingeniería Industrial,
            </p>
            <p style="margin:12px 0 0;font-size:14px;color:#374151;line-height:1.6;">
              A continuación se presenta el resumen de los cambios de producto registrados en las celdas del sistema
              de asignación de herramientas durante las últimas 24 horas.
              Por favor revisen los cambios y tomen las acciones correspondientes de ser necesario.
            </p>
            ${periods.length ? `<p style="margin:10px 0 0;font-size:12px;color:#6b7280;">Períodos comparados: <strong>${periodLabels}</strong></p>` : ''}
          </td>
        </tr>

        <!-- Summary badge -->
        <tr>
          <td style="padding:8px 32px 20px;">
            <span style="display:inline-block;background:${changes.length ? '#eff6ff' : '#f0fdf4'};
              color:${changes.length ? '#1d4ed8' : '#15803d'};
              border:1px solid ${changes.length ? '#bfdbfe' : '#bbf7d0'};
              border-radius:20px;padding:4px 14px;font-size:13px;font-weight:700;">
              ${changes.length} cambio${changes.length !== 1 ? 's' : ''} detectado${changes.length !== 1 ? 's' : ''}
            </span>
          </td>
        </tr>

        <!-- Table -->
        <tr>
          <td style="padding:0 32px 32px;">
            <table width="100%" cellpadding="0" cellspacing="0"
              style="border-collapse:collapse;border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;font-size:13px;">
              <thead>
                <tr style="background:#0052cc;">
                  <th style="padding:10px 12px;text-align:left;color:#fff;font-weight:700;white-space:nowrap;">Tool / Máquina</th>
                  <th style="padding:10px 12px;text-align:left;color:#fff;font-weight:700;">Celda</th>
                  <th style="padding:10px 12px;text-align:left;color:#fff;font-weight:700;">Producto Anterior</th>
                  <th style="padding:10px 12px;text-align:center;color:#fff;font-weight:700;"></th>
                  <th style="padding:10px 12px;text-align:left;color:#fff;font-weight:700;">Producto Nuevo</th>
                  <th style="padding:10px 12px;text-align:left;color:#fff;font-weight:700;white-space:nowrap;">Período</th>
                </tr>
              </thead>
              <tbody>
                ${rows}
              </tbody>
            </table>
          </td>
        </tr>

        <!-- Footer -->
        <tr>
          <td style="background:#f9fafb;border-top:1px solid #e5e7eb;padding:18px 32px;">
            <p style="margin:0;font-size:12px;color:#9ca3af;line-height:1.5;">
              Este correo fue generado automáticamente por el <strong>Tool Allocation Dashboard</strong> — CRVLE Ops Planning.<br>
              Por favor no responder directamente a este mensaje.
            </p>
          </td>
        </tr>

      </table>
    </td></tr>
  </table>
</body>
</html>`;
}

async function sendDailyReport() {
  const emailFrom     = process.env.EMAIL_FROM;
  const emailTo       = process.env.EMAIL_TO;
  const emailPassword = process.env.EMAIL_PASSWORD;
  const smtpHost      = process.env.EMAIL_SMTP_HOST || 'smtp.office365.com';
  const smtpPort      = parseInt(process.env.EMAIL_SMTP_PORT || '587', 10);

  if (!emailFrom || !emailTo || !emailPassword) {
    console.warn('[Email] Missing EMAIL_FROM, EMAIL_TO or EMAIL_PASSWORD in .env — skipping report.');
    return;
  }

  try {
    console.log('[Email] Building daily report...');

    const tCol = safeCol(DEFAULT_COLS.toolCol);
    const cCol = safeCol(DEFAULT_COLS.cellCol);
    const pCol = safeCol(DEFAULT_COLS.productCol);
    const wCol = safeCol(DEFAULT_COLS.wwCol);
    const dCol = safeCol(DEFAULT_COLS.dayCol);

    // Use the last 2 business days so we cover "last 24 h"
    const periods = recentBusinessPeriods(2);
    if (periods.length < 2) {
      console.warn('[Email] Not enough periods to compute changes — skipping report.');
      return;
    }

    // Use a dedicated pool with a 5-minute request timeout so slow DB queries don't fail
    const emailPool = new sql.ConnectionPool({ ...dbConfig, requestTimeout: 300000 });
    await emailPool.connect();
    let maps;
    try {
      maps = await Promise.all(
        periods.map(period => fetchPeriodMap(emailPool, tCol, cCol, pCol, wCol, dCol, period.ww, period.day))
      );
    } finally {
      emailPool.close().catch(() => {});
    }

    const newer = maps[0];
    const older  = maps[1];
    const allKeys = new Set([...Object.keys(newer), ...Object.keys(older)]);
    const changes = [];
    for (const key of allKeys) {
      const [machine, cell] = key.split('||');
      const oldProd = Object.prototype.hasOwnProperty.call(older, key) ? older[key] : null;
      const newProd = Object.prototype.hasOwnProperty.call(newer, key) ? newer[key] : null;
      if (oldProd !== newProd) {
        changes.push({
          machine, cell,
          oldProduct: oldProd,
          newProduct: newProd,
          fromWW : periods[1].ww, fromDay: periods[1].day,
          toWW   : periods[0].ww, toDay  : periods[0].day,
        });
      }
    }

    // Sort: machines with actual product changes first, then alphabetically
    changes.sort((a, b) => {
      const aHas = Boolean(a.newProduct);
      const bHas = Boolean(b.newProduct);
      if (aHas !== bHas) return aHas ? -1 : 1;
      return (a.machine || '').localeCompare(b.machine || '');
    });

    const reportDate = new Date();
    const html = buildChangesEmailHtml(changes, periods, reportDate);

    const subject = `[Tool Allocation] Reporte de Cambios — ${reportDate.toLocaleDateString('es-MX', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}`;

    // Send via Outlook COM (bypasses SMTP — uses existing Outlook session)
    const tmpDir   = os.tmpdir();
    const htmlFile = path.join(tmpDir, 'tool_alloc_report.html').replace(/\\/g, '/');
    const ps1File  = path.join(tmpDir, 'send_tool_alloc.ps1');

    fs.writeFileSync(htmlFile, html, 'utf8');

    // Build PowerShell script: try sending from the service mailbox first,
    // fall back to the default Outlook account if Send-As is not granted.
    const ps1 = [
      '$ErrorActionPreference = "Stop"',
      '$ol   = New-Object -ComObject Outlook.Application',
      '$mail = $ol.CreateItem(0)',
      `$mail.To      = '${emailTo}'`,
      `$mail.Subject = '${subject.replace(/'/g, "''")}'`,
      `$mail.HTMLBody = [System.IO.File]::ReadAllText('${htmlFile}')`,
      // Try setting the From account to the service mailbox
      'try {',
      `  $fromAddr = '${emailFrom}'`,
      '  $accounts = $ol.Session.Accounts',
      '  for ($i = 1; $i -le $accounts.Count; $i++) {',
      '    if ($accounts.Item($i).SmtpAddress -eq $fromAddr) {',
      '      $mail.SendUsingAccount = $accounts.Item($i)',
      '      break',
      '    }',
      '  }',
      '} catch {}',
      '$mail.Send()',
      'Write-Output "SENT"',
    ].join('\n');

    fs.writeFileSync(ps1File, ps1, 'utf8');

    const psOut = execSync(
      `powershell -ExecutionPolicy Bypass -NoProfile -File "${ps1File}"`,
      { timeout: 30000, encoding: 'utf8' }
    ).trim();

    // Clean up temp files
    try { fs.unlinkSync(htmlFile); } catch (_) {}
    try { fs.unlinkSync(ps1File);  } catch (_) {}

    if (!psOut.includes('SENT')) throw new Error('PowerShell did not confirm send: ' + psOut);

    console.log(`[Email] Daily report sent to ${emailTo} via Outlook — ${changes.length} change(s) reported.`);
  } catch (err) {
    console.error('[Email] Failed to send daily report:', err.message);
  }
}

// Schedule: every weekday at 18:10 (6:10 PM) Costa Rica time
// Cron format: minute hour day-of-month month day-of-week
cron.schedule('10 18 * * 1-5', () => {
  console.log('[Email] Cron triggered — 6:10 PM weekday report');
  sendDailyReport();
}, { timezone: 'America/Costa_Rica' });

console.log('[Email] Daily report scheduled for 18:10 Mon–Fri (America/Costa_Rica).');

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
