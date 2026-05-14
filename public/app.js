/* ═══════════════════════════════════════════════════════════════════════════
   Tool Allocation Dashboard — app.js
   All DB interaction goes through the Node.js backend REST API.
   ═══════════════════════════════════════════════════════════════════════════ */

'use strict';

// ─── HDBI/BDC Layout definitions (2D physical layout) ────────────────────────
const HDBI_LAYOUTS = {
  'HDBI 4786': { columns: ['LA', 'RA'], rows: ['601', '501', '401', '301', '201', '101'] },
  'HDBI 5082': { columns: ['LA', 'LB', 'LC', 'RA', 'RB', 'RC'], rows: ['601', '501', '401', '301', '201', '101'] },
  'HDBI 4344': { columns: ['LA', 'LB', 'LC', 'RA', 'RB', 'RC'], rows: ['601', '501', '401', '301', '201', '101'] },
  'BDC 5257':  { columns: [''], rows: ['201', '101'] },
  'BDC 4944':  { columns: [''], rows: ['201', '101'] },
  'PPV HST':   { columns: ['A', 'B', 'C', 'D'], rows: ['601', '501', '401', '301', '201', '101'] },
  'PPV SST':   { columns: ['A', 'B', 'C', 'D', 'E'], rows: ['401', '301', '201', '101'] },
};

// ─── Area → Sub → Tools mapping (from HDBI-HDMX-PPV toolsxlsx.csv) ─────────
const AREA_MAP = {
  HDBI: {
    subs: ['HDBI', 'BDC'],
    tools: {
      HDBI: ['HDBI 5082', 'HDBI 4344', 'HDBI 4786'],
      BDC : ['BDC 5257', 'BDC 4944'],
    },
  },
  HDMX: {
    subs: [],
    tools: {
      '': [
        '1439','1455','1701','1785','1823','1844',
        '1848','1852','1866','1878','2117','2169',
        '2207','2930','2989','3415','3707','5275',
      ],
    },
  },
  CLASS: {
    subs: [],
    tools: { '': [] },
  },
  PPV: {
    subs: ['HST', 'SST', 'PTC'],
    tools: {
      HST: ['CR4326','CR4467','CR4938','CR4739','CR4711','CR4906','CR5351','CR5209'],
      SST: ['SST0001','SST0002','SST0003','SST0004','SST0005','SST0008','SST0009','SST0010'],
      PTC: [
        'CR4678','CR4681','CR4682','CR 4683','CR4686',
        'CR4954','CR4956','CR5155','CR5158','CR5256',
        'CR4677','CR4685','CR4955','CR7157','CR7158','CR4953',
      ],
    },
  },
};

// ─── Product family palette ──────────────────────────────────────────────────
const PRODUCT_CLASSES = [
  'cell-p0','cell-p1','cell-p2','cell-p3','cell-p4',
  'cell-p5','cell-p6','cell-p7','cell-p8','cell-p9',
  'cell-p10','cell-p11','cell-p12','cell-p13','cell-p14',
  'cell-p15','cell-p16','cell-p17','cell-p18','cell-p19',
  'cell-p20','cell-p21','cell-p22','cell-p23',
];
const productColorMap = {};

// Stable hash of a product name → deterministic color index (djb2 variant).
function hashProductName(str) {
  let h = 5381;
  for (let i = 0; i < str.length; i++) {
    h = (((h << 5) + h) ^ str.charCodeAt(i)) >>> 0;
  }
  return h;
}
const CLASS_CELLS = ['A101','A102','A201','A202','A301','A302','A401','A402','A501','A502'];

// HDMX level groupings: each level = two cells that must share the same product.
const HDMX_LEVELS = [
  { label: 'L1', cells: ['A101', 'A102'] },
  { label: 'L2', cells: ['A201', 'A202'] },
  { label: 'L3', cells: ['A301', 'A302'] },
  { label: 'L4', cells: ['A401', 'A402'] },
  { label: 'L5', cells: ['A501', 'A502'] },
];

const PTC_CELLS = ['A101','A201','A301','A401','A501'];
const SST_CELLS = [
  'A101','A201','A301','A401',
  'B101','B201','B301','B401',
  'C101','C201','C301','C401',
  'D101','D201','D301','D401',
  'E101','E201','E301','E401',
];
const HST_CELLS = [
  'A101','A201','A301','A401','A501','A601',
  'B101','B201','B301','B401','B501','B601',
  'C101','C201','C301','C401','C501','C601',
  'D101','D201','D301','D401','D501','D601',
];

const HDMX_GMM_PRODUCTS = new Set([
  'GMMLCC10T0082',
  'GMMLCC10T0095',
  'GMMLCC10T0210',
  'GMMLCC10T0950',
  'GMMLCC10T0287',
  'GMMLCC10T1336',
  'GMMLCC10T1567',
  'GNRXEL10T0095',
]);

const HDBI_GMM_PRODUCTS = new Set([
  'GMM00014-0-A',
  'GMM00011-0-A',
  'GMM00003-0-A',
]);

// From "HDMX Tools.csv": tool -> coolant family
const HDMX_COOLANT_BY_TOOL = {
  '1439': 'EGDI',
  '1455': 'HFE',
  '1701': 'EGDI',
  '1785': 'EGDI',
  '1823': 'HFE',
  '1844': 'HFE',
  '1848': 'EGDI',
  '1852': 'EGDI',
  '1866': 'EGDI',
  '1878': 'EGDI',
  '2117': 'EGDI',
  '2169': 'EGDI',
  '2207': 'EGDI',
  '2930': 'HFE',
  '2989': 'EGDI',
  '3415': 'HFE',
  '3707': 'EGDI',
  '5275': 'HFE',
};

const ALL_VIEW_HIDDEN_MACHINES = new Set([
  'PDC14368',
  'PDC14542',
  'CR03HPDC14368',
  'CR03HPDC14542',
  'CR03ICDC3677',
  'CR03ICDC3679',
  'CR03WCMV0001',
  'CR03WCMV0001AC',
  'CR03WCMV0001CWF',
  'CR03WCMV0001FHF',
  'CR03WCMV0001JC',
  'CR03WCMV0001SSS',
  'CR03WCMV0001WC',
  'CR03WCMV0001WS',
  'CR03WCMV0002',
  'CR03WCMV0002AC',
  'CR03WCMV0002FHF',
  'CR03WCMV0002SSS',
  'CR03WCMV0002WS',
  'CR03WCMV0003',
  'CR03WCMV0003BNC',
  'CR03WCMV0003AC',
  'CR03WCMV0003AVC',
  'CR03WCMV0003WS',
  'CR03WCMV0004WS',
  'CR03WCMV0005AC',
  'CR03WCMV0005AVC',
  'CR03WCMV0006AC',
  'CR03WCMV0006BNC',
  'CR03WCMV0007AC',
  'CR03WCMV0008AC',
  'SC09WCMV0002AC',
  'SC09WCMV0002JC',
  'SC09WCMV0008JC',
  'SC09WCMV0024AVC',
  'SC09WCMV0025JC',
  'SC09WCMV0036JC',
]);

function normalizeMachineIdForBlock(value) {
  return String(value || '').replace(/\s+/g, '').toUpperCase();
}

function isMachineHiddenInAll(machineName) {
  const raw = normalizeMachineIdForBlock(machineName);
  const shown = normalizeMachineIdForBlock(formatMachineDisplayName(machineName));
  return ALL_VIEW_HIDDEN_MACHINES.has(raw) || ALL_VIEW_HIDDEN_MACHINES.has(shown);
}

function stripMpePrefix(p) {
  return String(p || '').replace(/^MPE_/i, '');
}

function normalizeHdmxProductForStats(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return 'Empty';

  const upper = raw.toUpperCase();
  if (['MPE_GNR-SP-LCC', 'MPE_GNR-SP-XCC', 'MPE_GNR-SP-HCC'].includes(upper)) {
    return 'GNR-SP';
  }
  if (HDMX_GMM_PRODUCTS.has(upper)) {
    return 'GMM';
  }

  return stripMpePrefix(raw);
}

function isHdbiGmmTool(toolId) {
  const tag = String(formatMachineDisplayName(toolId) || '').replace(/\s+/g, '').toUpperCase();
  return tag === 'HDBI4786' || tag === 'HDBI5082' || tag === 'HDBI4344';
}

function normalizeProductForCard(toolId, value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return 'Empty';

  const loc = resolveArea(toolId);
  const upper = raw.toUpperCase();

  if (loc.area === 'HDBI') {
    if (isHdbiGmmTool(toolId) && upper.startsWith('GMM')) return 'GMM';
    if (HDBI_GMM_PRODUCTS.has(upper)) return 'GMM';
  }

  if (loc.area === 'HDMX') {
    if (HDMX_GMM_PRODUCTS.has(upper)) return 'GMM';
    if (upper === 'MPE_NVL-S-28C') return 'NVL';
    // Alias DMR-AP → DMR_AP_SDE so it shares the same color as HDBI
    if (upper === 'DMR-AP' || upper === 'MPE_DMR-AP') return 'DMR_AP_SDE';
  }

  return stripMpePrefix(raw);
}

function formatToolTokenForFilter(toolId, areaHint = '') {
  const display = String(formatMachineDisplayName(toolId) || '').replace(/\s+/g, '').trim();
  if (!display) return '';

  const area = String(areaHint || resolveArea(toolId).area || '').toUpperCase();

  // formatMachineDisplayName already adds prefixes for HST/SST/PTC/HDBI/BDC
  // Only HDMX needs explicit prefix since formatMachineDisplayName returns only digits
  if (area === 'HDMX' && !display.startsWith('HDMX')) {
    return `HDMX${display}`;
  }

  return display;
}

const ALL_TOOL_GROUP_ORDER = ['HDBI', 'BDC', 'HDMX', 'HST', 'SST', 'PTC'];

function getAllFilterToolGroup(areaName, subName) {
  const area = String(areaName || '').toUpperCase();
  const sub = String(subName || '').toUpperCase();

  if (area === 'HDBI' && sub === 'HDBI') return 'HDBI';
  if (area === 'HDBI' && sub === 'BDC') return 'BDC';
  if (area === 'HDMX') return 'HDMX';
  if (area === 'PPV' && sub === 'HST') return 'HST';
  if (area === 'PPV' && sub === 'SST') return 'SST';
  if (area === 'PPV' && sub === 'PTC') return 'PTC';

  return 'OTHER';
}

function getHdbiLayoutKey(toolId) {
  if (!toolId) return '';

  if (HDBI_LAYOUTS[toolId]) return toolId;

  const norm = normalise(toolId);
  if ((norm.includes('HDBI') || norm.includes('HBI') || norm.includes('MBI')) && norm.includes('5082')) return 'HDBI 5082';
  if ((norm.includes('HDBI') || norm.includes('HBI') || norm.includes('MBI')) && norm.includes('4344')) return 'HDBI 4344';
  if ((norm.includes('HDBI') || norm.includes('HBI') || norm.includes('MBI')) && norm.includes('4786')) return 'HDBI 4786';
  if (norm.includes('BDC') && norm.includes('5257')) return 'BDC 5257';
  if (norm.includes('BDC') && norm.includes('4944')) return 'BDC 4944';

  return '';
}

function getLayoutCells(layoutKey) {
  const layout = HDBI_LAYOUTS[layoutKey];
  if (!layout) return null;

  const cells = [];
  for (const col of layout.columns) {
    for (const row of layout.rows) {
      cells.push(col ? `${col}${row}` : row);
    }
  }
  return cells;
}

function productClass(product) {
  if (!product || product === '' || /idle/i.test(product)) return 'cell-idle';
  if (/^tlo/i.test(product)) return 'cell-tlo';
  if (!productColorMap[product]) {
    productColorMap[product] = PRODUCT_CLASSES[hashProductName(product) % PRODUCT_CLASSES.length];
  }
  return productColorMap[product];
}

// Pre-populate the hash-based color cache for all products in the current dataset.
// Colors are assigned via stable hash so the same product always gets the same
// color across all areas and app restarts.
function buildProductColorMap(data) {
  const pc = colMapping.productCol;
  if (!pc) return;

  Object.keys(productColorMap).forEach(k => { delete productColorMap[k]; });
  for (const row of data) {
    const p = String(row[pc] == null ? '' : row[pc]).trim();
    if (!p || /^(idle|empty)$/i.test(p) || /^tlo/i.test(p)) continue;
    productClass(p); // triggers hash-based assignment into productColorMap
  }
}

function fixedCellsForTool(area, toolId) {
  const resolved = resolveArea(toolId);
  const resolvedArea = String((area || resolved.area || '')).toUpperCase();
  const resolvedSub = String(resolved.sub || '').toUpperCase();
  const layoutKey = getHdbiLayoutKey(toolId);

  if (layoutKey) {
    return getLayoutCells(layoutKey);
  }

  if (resolvedSub === 'PTC') return PTC_CELLS;
  if (resolvedSub === 'SST') return SST_CELLS;
  if (resolvedSub === 'HST') return HST_CELLS;
  if (['CLASS', 'HDMX'].includes(resolvedArea)) return CLASS_CELLS;
  return null;
}

// ─── Column mapping (persisted in localStorage) ──────────────────────────────
const MAPPING_KEYS = ['toolCol','cellCol','productCol','wwCol','dayCol','tosCol'];
const MAPPING_LABELS = {
  toolCol    : 'Tool / Machine ID',
  cellCol    : 'Cell / Slot',
  productCol : 'Product',
  wwCol      : 'Work Week (WW)',
  dayCol     : 'Week Day',
  tosCol     : 'TOS Version (optional)',
};

const MAPPING_CANDIDATES = {
  toolCol: ['MACHINE_NAME', 'TOOL', 'TOOL_ID'],
  cellCol: ['CELL', 'CELL_NAME', 'SLOT'],
  productCol: ['PRODUCT', 'PART', 'PART_NUMBER'],
  wwCol: ['ww', 'WW', 'WORK_WEEK'],
  dayCol: ['day_', 'DAY', 'DAY_'],
  tosCol: ['TESTER_NAME', 'TOS', 'TOS_VERSION'],
};

function loadMapping() {
  try {
    const raw = localStorage.getItem('colMapping');
    return raw ? JSON.parse(raw) : {};
  } catch { return {}; }
}
function saveMapping(map) {
  localStorage.setItem('colMapping', JSON.stringify(map));
}
function isMappingComplete(map) {
  // tosCol is optional; others are required
  return ['toolCol','cellCol','productCol','wwCol','dayCol'].every(k => map[k]);
}

function pickColumn(candidates, columns) {
  const canon = new Map(columns.map(c => [c.toUpperCase(), c]));
  for (const name of candidates) {
    const found = canon.get(String(name).toUpperCase());
    if (found) return found;
  }
  return '';
}

async function tryAutoDetectMapping() {
  if (isMappingComplete(colMapping)) return true;

  const { columns } = await apiFetch('/api/schema');
  discoveredCols = columns || [];

  const merged = { ...colMapping };
  for (const key of MAPPING_KEYS) {
    if (merged[key]) continue;
    merged[key] = pickColumn(MAPPING_CANDIDATES[key] || [], discoveredCols);
  }

  colMapping = merged;
  saveMapping(colMapping);
  return isMappingComplete(colMapping);
}

// ─── DOM refs ────────────────────────────────────────────────────────────────
const filterArea      = document.getElementById('filterArea');
const subAreaSection  = document.getElementById('subAreaSection');
const filterSub       = document.getElementById('filterSub');
const coolantSection  = document.getElementById('coolantSection');
const filterCoolant   = document.getElementById('filterCoolant');
const hdmxLayoutSection = document.getElementById('hdmxLayoutSection');
const hdmxBaggedSection = document.getElementById('hdmxBaggedSection');
const filterBagged      = document.getElementById('filterBagged');
const ppvLayoutSection  = document.getElementById('ppvLayoutSection');
const filterWW        = document.getElementById('filterWW');
const filterDay       = document.getElementById('filterDay');
const toolCheckList   = document.getElementById('toolCheckList');
const productSearch   = document.getElementById('productSearch');
const productSearchBtn= document.getElementById('productSearchBtn');
const applyBtn        = document.getElementById('applyBtn');
const exportBtn       = document.getElementById('exportBtn');
const clearBtn        = document.getElementById('clearBtn');
const machineGrid     = document.getElementById('machineGrid');
const emptyState      = document.getElementById('emptyState');
const statusMsg       = document.getElementById('statusMsg');
const recordCount     = document.getElementById('recordCount');
const areaBadge       = document.getElementById('areaBadge');
const tooltip         = document.getElementById('tooltip');
const settingsBtn     = document.getElementById('settingsBtn');
const settingsModal   = document.getElementById('settingsModal');
const discoverBtn     = document.getElementById('discoverBtn');
const discoverStatus  = document.getElementById('discoverStatus');
const mappingGrid     = document.getElementById('mappingGrid');
const saveMappingBtn  = document.getElementById('saveMappingBtn');
const cancelMappingBtn= document.getElementById('cancelMappingBtn');
const detailModal     = document.getElementById('detailModal');
const detailTitle     = document.getElementById('detailTitle');
const detailBody      = document.getElementById('detailBody');
const closeDetailBtn  = document.getElementById('closeDetailBtn');
const areaTabs        = [...document.querySelectorAll('.area-tab')];
const loadingOverlay  = document.getElementById('loadingOverlay');
const loadingMessage  = document.getElementById('loadingMessage');
const statsSection    = document.getElementById('statsSection');
const statsPanelTabs  = document.getElementById('statsPanelTabs');
const statsPanelBody  = document.getElementById('statsPanelBody');
const changesBtn      = document.getElementById('changesBtn');
const changesModal    = document.getElementById('changesModal');
const changesBody     = document.getElementById('changesBody');
const closeChangesBtn = document.getElementById('closeChangesBtn');
const changesSearch   = document.getElementById('changesSearch');
const changesExportBtn= document.getElementById('changesExportBtn');
const disableToolsBtn  = document.getElementById('disableToolsBtn');
const disableToolsModal= document.getElementById('disableToolsModal');
const disableToolsList = document.getElementById('disableToolsList');
const disableToolsSearch = document.getElementById('disableToolsSearch');
const saveDisableToolsBtn  = document.getElementById('saveDisableToolsBtn');
const closeDisableToolsBtn = document.getElementById('closeDisableToolsBtn');

// ─── State ───────────────────────────────────────────────────────────────────
let colMapping    = loadMapping();
let discoveredCols= [];
let currentData   = [];   // raw rows from last /api/allocations call
let baseDataCache = [];   // raw rows for active ww/day (single DB load)
let baseDataKey   = '';   // cache key = "ww|day"
let highlightSet  = new Set();  // tool+cell keys to highlight
let disabledTools = new Set();  // tool IDs that are disabled (cells shown as Empty)

// ─── Forecast state ───────────────────────────────────────────────────────────
// forecastSlots: next 7 days returned by /api/forecast-slots [{ww, day, date}]
// todayPeriod  : today's resolved period from the server {ww, day, date}
let forecastSlots = [];
let todayPeriod   = null;

// Returns true when the given WW/day combination is a future (forecast) slot.
function isForecastPeriod(ww, day) {
  if (!todayPeriod || !ww || day === '' || day === undefined || day === null) return false;
  const todayNum = Number(String(todayPeriod.ww).replace(/\D/g, ''));
  const wwNum    = Number(String(ww).replace(/\D/g, ''));
  if (!wwNum || !todayNum) return false;
  if (wwNum > todayNum) return true;
  if (wwNum === todayNum && Number(day) > Number(todayPeriod.day)) return true;
  return false;
}

// ─── API helpers ─────────────────────────────────────────────────────────────
async function apiFetch(path) {
  const res = await fetch(path);
  const contentType = res.headers.get('content-type') || '';
  if (!res.ok) {
    if (contentType.includes('application/json')) {
      const err = await res.json().catch(() => ({ error: res.statusText }));
      throw new Error(err.error || res.statusText);
    }
    throw new Error(`Server returned ${res.status} ${res.statusText}`);
  }
  if (!contentType.includes('application/json')) {
    throw new Error(`Unexpected server response (expected JSON, got ${contentType || 'unknown'}). Try reloading the page.`);
  }
  return res.json();
}

function mappingParams() {
  const p = new URLSearchParams();
  for (const k of MAPPING_KEYS) {
    if (colMapping[k]) p.set(k, colMapping[k]);
  }
  return p;
}

function showLoading(message) {
  if (!loadingOverlay) return;
  if (loadingMessage) loadingMessage.textContent = message || 'Loading data...';
  loadingOverlay.style.display = 'flex';
}

function hideLoading() {
  if (!loadingOverlay) return;
  loadingOverlay.style.display = 'none';
}

function activePeriodKey() {
  const ww = filterWW.value || 'ALL_WW';
  const day = filterDay.value || 'ALL_DAY';
  return `${ww}|${day}`;
}

function hdmxToolDigits(machineOrTool) {
  return String(machineOrTool || '').replace(/\D/g, '').slice(-4);
}

function getHdmxCoolant(machineOrTool) {
  const digits = hdmxToolDigits(machineOrTool);
  return String(HDMX_COOLANT_BY_TOOL[digits] || '').toUpperCase();
}

// ─── Forecast helpers ─────────────────────────────────────────────────────────

// Fetch forecast slots from the server and store in state.
async function loadForecastSlots() {
  try {
    const { today, slots } = await apiFetch('/api/forecast-slots');
    forecastSlots = slots || [];
    todayPeriod   = today || null;

    if (!forecastSlots.length) return;

    // Ensure any forecast WW values that weren't returned by /api/filters
    // are present in the selector (edge case: filters only goes +2 WW).
    const existingVals = new Set([...filterWW.options].map(o => o.value));
    const uniqueWWs    = [...new Set(forecastSlots.map(s => s.ww))];
    for (const wwVal of uniqueWWs) {
      if (!existingVals.has(wwVal)) {
        const opt = document.createElement('option');
        opt.value = wwVal;
        opt.textContent = formatWWLabel(wwVal);
        opt.dataset.forecast = 'true';
        filterWW.appendChild(opt);
      }
    }

    // Mark future WW options in the dataset (no visual indicator added)
    const todayWwNum = Number(String(today.ww).replace(/\D/g, ''));
    for (const opt of filterWW.options) {
      if (!opt.value) continue;
      const optWwNum = Number(String(opt.value).replace(/\D/g, ''));
      if (optWwNum > todayWwNum && !opt.dataset.forecast) {
        opt.dataset.forecast = 'true';
      }
    }

    // Build a set of forecast day values for the current WW (remaining days after today)
    const currentWwFutureDays = new Set(
      forecastSlots
        .filter(s => s.ww === today.ww)
        .map(s => s.day)
    );

    // Mark day selector options that are future slots within the current WW (no visual indicator)
    if (currentWwFutureDays.size > 0) {
      for (const opt of filterDay.options) {
        if (currentWwFutureDays.has(opt.value) && !opt.dataset.forecast) {
          opt.dataset.forecast = 'true';
        }
      }
    }
  } catch (err) {
    console.warn('[loadForecastSlots]', err.message);
  }
}

// Build a synthetic (forecast) dataset for the given future WW/day.
// Logic:
//   1. Load today's actual allocation from the DB (cached after first load).
//   2. Deep-copy those rows.
//   3. For each day in the chain from tomorrow → targetDay, apply any planned
//      conversions whose date matches that day.
//   4. Stamp the resulting rows with the target WW/day and store in baseDataCache.
async function buildForecastData(targetWw, targetDay) {
  if (!todayPeriod) {
    setStatus('⚠ Forecast: could not determine the current period.');
    baseDataCache = []; machineGrid.innerHTML = ''; emptyState.style.display = '';
    return false;
  }

  const targetSlotIdx = forecastSlots.findIndex(s => s.ww === targetWw && s.day === targetDay);
  if (targetSlotIdx === -1) {
    setStatus(`⚠ Forecast: slot WW${String(targetWw).slice(-2)}.${targetDay} not found.`);
    baseDataCache = []; machineGrid.innerHTML = ''; emptyState.style.display = '';
    return false;
  }

  const wwShort = String(targetWw).slice(-2);
  showLoading(`Building forecast WW${wwShort}.${targetDay}…`);
  setStatus(`<span class="spinner"></span> Building forecast WW${wwShort}.${targetDay}…`);
  recordCount.textContent = '';

  try {
    // ── 1. Obtain today's base data ──────────────────────────────────────────
    let todayData;
    const todayKey = `${todayPeriod.ww}|${todayPeriod.day}`;
    if (baseDataKey === todayKey && baseDataCache.length > 0) {
      todayData = baseDataCache; // already in memory — reuse
    } else {
      const p = mappingParams();
      p.set('ww',  todayPeriod.ww);
      p.set('day', todayPeriod.day);
      const { data } = await apiFetch(`/api/allocations?${p}`);
      todayData = data;
    }

    // ── 2. Deep-copy rows so we never mutate the cached today's snapshot ─────
    let forecastData = todayData.map(row => ({ ...row }));

    // ── 3. Fetch planned conversions (silently skip on error) ────────────────
    let planned = [];
    try { planned = await apiFetch('/api/planned-conversions'); } catch (_) {}

    // ── 4. Chain: apply each day's conversions in order ──────────────────────
    const chain = forecastSlots.slice(0, targetSlotIdx + 1);
    const tc = colMapping.toolCol;
    const cc = colMapping.cellCol;
    const pc = colMapping.productCol;
    const wc = colMapping.wwCol;
    const dc = colMapping.dayCol;

    for (const slot of chain) {
      const slotConversions = planned.filter(c => c.date === slot.date);
      for (const conv of slotConversions) {
        for (const row of forecastData) {
          // Use machineMatchesToken so that DB names like "CR03DHHX1439" match
          // form values like "1439", "CR4326", "HDBI 5082", etc.
          if (machineMatchesToken(String(row[tc] || ''), conv.tool) &&
              String(row[cc] || '').trim().toUpperCase() === String(conv.cell || '').trim().toUpperCase()) {
            row[pc] = conv.newProduct;
          }
        }
      }
    }

    // ── 5. Stamp all rows with the target WW/day ──────────────────────────────
    for (const row of forecastData) {
      if (wc) row[wc] = targetWw;
      if (dc) row[dc] = targetDay;
    }

    // Persist conversions applied in this forecast for status display
    const totalConversionsApplied = chain.reduce((acc, slot) =>
      acc + planned.filter(c => c.date === slot.date).length, 0);

    baseDataCache = forecastData;
    baseDataKey   = `${targetWw}|${targetDay}`;
    setStatus('');
    return true;
  } catch (err) {
    setStatus(`⚠ Forecast error: ${err.message}`);
    baseDataCache = []; baseDataKey = '';
    currentData = []; machineGrid.innerHTML = ''; emptyState.style.display = '';
    return false;
  } finally {
    hideLoading();
  }
}

async function ensureBaseDataLoaded(forceReload = false) {
  const ww = filterWW.value;
  const day = filterDay.value;

  // Keep same guard to avoid huge unbounded queries.
  if (!ww && !day) {
    setStatus('⚠ Select at least one week (WW) or day to load data.');
    machineGrid.innerHTML = '';
    emptyState.style.display = '';
    return false;
  }

  // ─── Forecast mode: future WW/day — synthesize from today's data ─────────
  if (ww && day !== '' && day !== undefined && day !== null && isForecastPeriod(ww, day)) {
    const key = activePeriodKey();
    if (!forceReload && baseDataKey === key && baseDataCache.length > 0) return true;
    return await buildForecastData(ww, day);
  }

  const key = activePeriodKey();
  if (!forceReload && baseDataCache.length && baseDataKey === key) {
    return true;
  }

  const p = mappingParams();
  if (ww) p.set('ww', ww);
  if (day) p.set('day', day);

  const wwDigits = String(ww || '').replace(/\D/g, '');
  const wwShort = wwDigits ? wwDigits.slice(-2) : 'ALL';
  const dayLabel = (day !== undefined && day !== null && String(day) !== '') ? String(day) : 'ALL';
  showLoading(`Loading data for WW${wwShort}.${dayLabel}`);
  setStatus('<span class="spinner"></span> Loading data from database...');
  recordCount.textContent = '';

  try {
    const { data } = await apiFetch(`/api/allocations?${p}`);
    baseDataCache = data;
    baseDataKey = key;
    return true;
  } catch (err) {
    setStatus(`⚠ ${err.message}`);
    baseDataCache = [];
    baseDataKey = '';
    currentData = [];
    machineGrid.innerHTML = '';
    emptyState.style.display = '';
    return false;
  } finally {
    hideLoading();
  }
}

// ─── Init ─────────────────────────────────────────────────────────────────────
// ─── Disabled Tools helpers ──────────────────────────────────────────────────
async function loadDisabledTools() {
  try {
    const list = await apiFetch('/api/disabled-tools');
    disabledTools = new Set(Array.isArray(list) ? list : []);
  } catch (e) {
    console.warn('[disabledTools] could not load:', e.message);
    disabledTools = new Set();
  }
}

async function saveDisabledTools(toolsSet) {
  await fetch('/api/disabled-tools', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ tools: [...toolsSet] }),
  });
}

/**
 * Returns true if the given toolId (or a DB machine name matching it) is disabled.
 * Uses machineMatchesToken for consistent name-matching (handles BDC 4944 ↔ CR03HBDC24944, etc.)
 */
function isToolDisabled(toolId) {
  if (!toolId || !disabledTools.size) return false;
  for (const d of disabledTools) {
    if (machineMatchesToken(toolId, d)) return true;
  }
  return false;
}

function openDisableToolsModal() {
  // Collect all known tools from AREA_MAP in canonical order
  const groups = [
    { label: 'HDBI', tools: [...(AREA_MAP.HDBI?.tools?.HDBI || []), ...(AREA_MAP.HDBI?.tools?.BDC || [])] },
    { label: 'HDMX', tools: AREA_MAP.HDMX?.tools?.[''] || [] },
    { label: 'PPV — HST', tools: AREA_MAP.PPV?.tools?.HST || [] },
    { label: 'PPV — SST', tools: AREA_MAP.PPV?.tools?.SST || [] },
    { label: 'PPV — PTC', tools: AREA_MAP.PPV?.tools?.PTC || [] },
  ];

  // Snapshot of currently disabled tools (mutable draft inside modal)
  const draft = new Set(disabledTools);

  disableToolsList.innerHTML = '';

  for (const grp of groups) {
    if (!grp.tools.length) continue;

    const section = document.createElement('div');
    section.className = 'dt-group';

    const hdr = document.createElement('div');
    hdr.className = 'dt-group-header';
    hdr.innerHTML = `<span>${grp.label}</span>`;
    section.appendChild(hdr);

    for (const toolId of grp.tools) {
      const displayName = formatMachineDisplayName(toolId);
      const row = document.createElement('label');
      row.className = 'dt-row';
      row.dataset.tool = toolId;
      row.dataset.search = normalise(displayName + ' ' + toolId);

      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.className = 'dt-checkbox';
      cb.checked = draft.has(toolId);
      cb.addEventListener('change', () => {
        if (cb.checked) draft.add(toolId);
        else draft.delete(toolId);
      });

      const nameSpan = document.createElement('span');
      nameSpan.textContent = displayName;

      const badge = document.createElement('span');
      badge.className = 'dt-badge-disabled';
      badge.textContent = 'BAGGED';

      row.appendChild(cb);
      row.appendChild(nameSpan);
      row.appendChild(badge);
      section.appendChild(row);
    }

    disableToolsList.appendChild(section);
  }

  // Search filter
  disableToolsSearch.value = '';
  disableToolsSearch.oninput = () => {
    const q = normalise(disableToolsSearch.value);
    disableToolsList.querySelectorAll('.dt-row').forEach(row => {
      row.style.display = (!q || row.dataset.search.includes(q)) ? '' : 'none';
    });
    // Hide group headers when all their rows are hidden
    disableToolsList.querySelectorAll('.dt-group').forEach(grp => {
      const anyVisible = [...grp.querySelectorAll('.dt-row')].some(r => r.style.display !== 'none');
      grp.style.display = anyVisible ? '' : 'none';
    });
  };

  // Save button
  saveDisableToolsBtn.onclick = async () => {
    disabledTools = new Set(draft);
    await saveDisabledTools(disabledTools);
    disableToolsModal.style.display = 'none';
    // Re-render grid with updated disabled state
    await fetchAndRender();
  };

  closeDisableToolsBtn.onclick = () => { disableToolsModal.style.display = 'none'; };

  disableToolsModal.style.display = '';
}

document.addEventListener('DOMContentLoaded', async () => {
  showLoading('Loading data');
  buildMappingGrid();

  if (!isMappingComplete(colMapping)) {
    try {
      const mapped = await tryAutoDetectMapping();
      buildMappingGrid();
      if (!mapped) {
        openSettings();
        setStatus('Configure column mapping first, then apply filters.');
        hideLoading();
        return;
      }
      setStatus('Column mapping auto-detected. Loading latest data...');
    } catch (err) {
      openSettings();
      setStatus(`⚠ Could not auto-detect mapping: ${err.message}`);
      hideLoading();
      return;
    }
  }

  await loadFilters();
  // Load forecast slots so future WW/day combos are recognised and marked
  // in the WW selector with a "→" indicator.
  await loadForecastSlots();
  // Load disabled tools list from server
  await loadDisabledTools();
  // On every page load/refresh, fetch a fresh snapshot from SQL once.
  // All subsequent filtering uses the in-memory cache for this session.
  await ensureBaseDataLoaded(true);
  // Build tool selector for ALL view on startup.
  populateToolList(filterArea.value, filterSub.value);
  // Auto-load all machines on startup with no filters applied
  await fetchAndRender();
  hideLoading();
});

// ─── Filter population ───────────────────────────────────────────────────────
// ─── Filter population ───────────────────────────────────────────────────────
async function loadFilters() {
  const p = mappingParams();
  let defaultWw = '';
  let defaultDay = '';

  // Fire both requests in parallel — eliminates sequential wait
  const [periodResult, filtersResult] = await Promise.allSettled([
    apiFetch('/api/default-period'),
    apiFetch(`/api/filters?${p}`),
  ]);

  if (periodResult.status === 'fulfilled') {
    defaultWw  = String(periodResult.value.ww  || '');
    defaultDay = String(periodResult.value.day || '');
  }

  try {
    if (filtersResult.status !== 'fulfilled') throw new Error(filtersResult.reason?.message || 'filters failed');
    const { weeks, days } = filtersResult.value;
    populateWWSelect(filterWW, weeks, 'All Weeks');
    populateSelect(filterDay, days,  'All Days');

    if (defaultWw && weeks.includes(defaultWw)) {
      filterWW.value = defaultWw;
    } else if (weeks.length && !filterWW.value) {
      filterWW.value = weeks[weeks.length - 1];
    }

    if (defaultDay && days.includes(defaultDay)) {
      filterDay.value = defaultDay;
    } else if (days.length && !filterDay.value) {
      filterDay.value = days[0];
    }
  } catch (err) {
    // Fallback when /api/filters is slow: still start with today's period.
    const wwTail = defaultWw.replace(/\D/g, '').slice(-2);
    const year = defaultWw.replace(/\D/g, '').slice(0, 4) || String(new Date().getFullYear());
    const currentWW = Number(wwTail || 0);

    let fallbackWeeks = [];
    if (currentWW > 0) {
      for (let i = 7; i >= 0; i--) {
        const w = currentWW - i;
        if (w > 0) fallbackWeeks.push(`${year}${String(w).padStart(2, '0')}`);
      }
    }
    if (!fallbackWeeks.length && defaultWw) fallbackWeeks = [defaultWw];

    populateWWSelect(filterWW, fallbackWeeks, 'All Weeks');
    populateSelect(filterDay, ['0','1','2','3','4','5','6'], 'All Days');

    if (defaultWw) filterWW.value = defaultWw;
    if (defaultDay !== '') filterDay.value = defaultDay;

    setStatus(`⚠ Could not load filters from DB: ${err.message}. Using calendar defaults.`);
  }
}

function populateSelect(sel, values, placeholder) {
  sel.innerHTML = `<option value="">${placeholder}</option>`;
  for (const v of values) {
    const opt = document.createElement('option');
    opt.value = opt.textContent = v;
    sel.appendChild(opt);
  }
}

function formatWWLabel(wwValue) {
  const raw = String(wwValue || '').trim();
  if (!raw) return '';
  const digits = raw.replace(/\D/g, '');
  const ww = digits.length >= 2 ? digits.slice(-2) : digits;
  return `WW${ww}`;
}

function populateWWSelect(sel, values, placeholder) {
  sel.innerHTML = `<option value="">${placeholder}</option>`;
  for (const v of values) {
    const opt = document.createElement('option');
    // Keep full value (e.g. 202617) for SQL filtering, show only WWxx in UI.
    opt.value = String(v);
    opt.textContent = formatWWLabel(v);
    sel.appendChild(opt);
  }
}

// ─── Area / Sub / Tool selectors ─────────────────────────────────────────────
filterArea.addEventListener('change', () => {
  const area = filterArea.value;
  areaBadge.textContent = area || '—';
  setActiveAreaTab(area);

  // Sub-area
  subAreaSection.style.display = 'none';
  filterSub.innerHTML = '<option value="">All</option>';

  if (area && AREA_MAP[area] && AREA_MAP[area].subs.length) {
    subAreaSection.style.display = '';
    for (const s of AREA_MAP[area].subs) {
      const opt = document.createElement('option');
      opt.value = opt.textContent = s;
      filterSub.appendChild(opt);
    }
  }

  if (coolantSection && filterCoolant) {
    if (area === 'HDMX') {
      coolantSection.style.display = '';
    } else {
      coolantSection.style.display = 'none';
      filterCoolant.value = '';
    }
  }

  if (hdmxLayoutSection) {
    hdmxLayoutSection.style.display = area === 'HDMX' ? '' : 'none';
  }

  if (hdmxBaggedSection) {
    hdmxBaggedSection.style.display = area === 'HDMX' ? '' : 'none';
    if (area !== 'HDMX' && filterBagged) filterBagged.value = '';
  }

  if (ppvLayoutSection) {
    ppvLayoutSection.style.display = area === 'PPV' ? '' : 'none';
  }

  populateToolList(area, '');
});

function setActiveAreaTab(area) {
  for (const btn of areaTabs) {
    const active = btn.dataset.area === (area || '');
    btn.classList.toggle('is-active', active);
    btn.setAttribute('aria-selected', active ? 'true' : 'false');
  }
}

async function handleAreaTabClick(area) {
  if (filterArea.value === area) {
    setActiveAreaTab(area);
    return;
  }

  filterArea.value = area;
  filterArea.dispatchEvent(new Event('change'));
  await fetchAndRender();
}

for (const btn of areaTabs) {
  btn.addEventListener('click', async () => {
    await handleAreaTabClick(btn.dataset.area || '');
  });
}

filterSub.addEventListener('change', async () => {
  populateToolList(filterArea.value, filterSub.value);
  await fetchAndRender();
});

if (filterBagged) {
  filterBagged.addEventListener('change', async () => { await fetchAndRender(); });
}

function populateToolList(area, sub) {
  toolCheckList.innerHTML = '';
  const tokenSet = new Set();
  const toolEntries = [];

  function addToolEntry(toolId, areaHint = '', subHint = '') {
    const token = formatToolTokenForFilter(toolId, areaHint);
    const key = normalise(token);
    if (!token || tokenSet.has(key)) return;

    const loc = areaHint
      ? { area: areaHint, sub: subHint || '' }
      : resolveArea(toolId);
    const group = getAllFilterToolGroup(loc.area, loc.sub);

    tokenSet.add(key);
    toolEntries.push({ value: token, label: token, group });
  }

  if (!area) {
    const ALLOWED_PREFIXES = ['HDBI', 'BDC', 'HDMX', 'HST', 'SST', 'PTC'];
    
    // Static order requested by user: HDBI, BDC, HDMX, HST, SST, PTC
    for (const toolId of (AREA_MAP.HDBI?.tools?.HDBI || [])) addToolEntry(toolId, 'HDBI', 'HDBI');
    for (const toolId of (AREA_MAP.HDBI?.tools?.BDC || [])) addToolEntry(toolId, 'HDBI', 'BDC');
    for (const toolId of (AREA_MAP.HDMX?.tools?.[''] || [])) addToolEntry(toolId, 'HDMX', '');
    for (const toolId of (AREA_MAP.PPV?.tools?.HST || [])) addToolEntry(toolId, 'PPV', 'HST');
    for (const toolId of (AREA_MAP.PPV?.tools?.SST || [])) addToolEntry(toolId, 'PPV', 'SST');
    for (const toolId of (AREA_MAP.PPV?.tools?.PTC || [])) addToolEntry(toolId, 'PPV', 'PTC');

    // Include dynamic tools discovered in cache, excluding OTHER.
    const tc = colMapping.toolCol;
    for (const row of baseDataCache) {
      const toolId = row[tc];
      if (!toolId) continue;
      const loc = resolveArea(toolId);
      const group = getAllFilterToolGroup(loc.area, loc.sub);
      if (group === 'OTHER') continue;
      addToolEntry(toolId, loc.area, loc.sub);
    }

    // Filter to only keep tools starting with allowed prefixes
    const filtered = toolEntries.filter(entry => {
      return ALLOWED_PREFIXES.some(prefix => entry.label.startsWith(prefix));
    });

    filtered.sort((a, b) => {
      const ai = ALL_TOOL_GROUP_ORDER.indexOf(a.group);
      const bi = ALL_TOOL_GROUP_ORDER.indexOf(b.group);
      if (ai !== bi) return ai - bi;
      return a.label.localeCompare(b.label);
    });

    toolEntries.length = 0;
    toolEntries.push(...filtered);
  } else {
    const areaData = AREA_MAP[area];
    if (!areaData) return;

    let toolGroups = [];
    if (sub && areaData.tools[sub]) {
      toolGroups = areaData.tools[sub];
    } else {
      // All subs or no subs
      toolGroups = Object.values(areaData.tools).flat();
      // deduplicate
      toolGroups = [...new Set(toolGroups)];
    }

    if (area === 'CLASS' || (area === 'PPV' && sub === 'PDC')) {
      const tc = colMapping.toolCol;
      const dynamic = [...new Set(baseDataCache
        .map(r => r[tc])
        .filter(Boolean)
        .filter(t => {
          const loc = resolveArea(t);
          return loc.area === area && String(loc.sub || '') === String(sub || '');
        }))];
      if (dynamic.length) toolGroups = dynamic;
    }

    for (const toolId of toolGroups) addToolEntry(toolId, area, sub);

    if (area === 'PPV' && !sub) {
      // PPV all-subs: order HST → SST → PTC, then alphabetically within each group
      const PPV_SUB_ORDER = { HST: 0, SST: 1, PTC: 2 };
      toolEntries.sort((a, b) => {
        const ai = PPV_SUB_ORDER[a.group] ?? 99;
        const bi = PPV_SUB_ORDER[b.group] ?? 99;
        if (ai !== bi) return ai - bi;
        return a.label.localeCompare(b.label);
      });
    } else {
      toolEntries.sort((a, b) => a.label.localeCompare(b.label));
    }
  }

  if (!toolEntries.length) {
    toolCheckList.innerHTML = '<span class="no-tools">No tools defined</span>';
    return;
  }

  const allLbl = document.createElement('label');
  const allCb  = document.createElement('input');
  allCb.type = 'checkbox';
  allCb.id = 'toolToggleAll';
  allCb.checked = true;
  allLbl.className = 'tool-toggle-all';
  allLbl.appendChild(allCb);
  allLbl.append(' Select / Deselect all');
  toolCheckList.appendChild(allLbl);

  const emptyLbl = document.createElement('label');
  const emptyCb  = document.createElement('input');
  emptyCb.type = 'checkbox';
  emptyCb.id = 'toolToggleEmpty';
  emptyCb.checked = true;
  emptyLbl.className = 'tool-toggle-empty';
  emptyLbl.appendChild(emptyCb);
  emptyLbl.append(' Tools Empty');
  // HDMX uses the dedicated Bagged Tools dropdown; no need for a separate empty toggle.
  if (area !== 'HDMX') toolCheckList.appendChild(emptyLbl);

  // Pre-build PPV sub-area token sets for display label resolution.
  // entry.group can be 'OTHER' when no sub-area filter is active, so we
  // look up the actual sub directly from AREA_MAP instead.
  const ppvHstTokens = new Set((AREA_MAP.PPV?.tools?.HST || []).map(t => normalise(formatToolTokenForFilter(t, 'PPV'))));
  const ppvPtcTokens = new Set((AREA_MAP.PPV?.tools?.PTC || []).map(t => normalise(formatToolTokenForFilter(t, 'PPV'))));

  for (const entry of toolEntries) {
    const lbl = document.createElement('label');
    const cb  = document.createElement('input');
    cb.type = 'checkbox'; cb.value = entry.value; cb.checked = true;
    cb.dataset.toolItem = '1';
    lbl.appendChild(cb);
    let displayLabel = entry.label;
    if (area === 'HDMX' && entry.label.startsWith('HDMX')) {
      displayLabel = entry.label.slice(4);
    } else if (area === 'PPV' && entry.label.startsWith('CR')) {
      const normVal = normalise(entry.value);
      if (ppvHstTokens.has(normVal))      displayLabel = `HST ${entry.label.slice(2)}`;
      else if (ppvPtcTokens.has(normVal)) displayLabel = `PTC ${entry.label.slice(2)}`;
    }
    lbl.append(` ${displayLabel}`);
    toolCheckList.appendChild(lbl);

    cb.addEventListener('change', syncToolToggleAllState);
  }

  allCb.addEventListener('change', () => {
    const items = toolCheckList.querySelectorAll('input[type=checkbox][data-tool-item="1"]');
    for (const cb of items) cb.checked = allCb.checked;
    const emptyCbEl = document.getElementById('toolToggleEmpty');
    if (emptyCbEl) emptyCbEl.checked = allCb.checked;
    syncToolToggleAllState();
  });

  const emptyCb2 = document.getElementById('toolToggleEmpty');
  if (emptyCb2) emptyCb2.addEventListener('change', syncToolToggleAllState);

  syncToolToggleAllState();
}

function selectedTools() {
  return [...toolCheckList.querySelectorAll('input[type=checkbox][data-tool-item="1"]:checked')]
    .map(cb => cb.value);
}

function includeToolsEmpty() {
  const emptyCb = document.getElementById('toolToggleEmpty');
  return emptyCb ? !!emptyCb.checked : true;
}

function filterOutEmptyOnlyTools(data) {
  const tc = colMapping.toolCol;
  const pc = colMapping.productCol;
  const cc = colMapping.cellCol;

  const rowsByTool = new Map();
  for (const row of data) {
    const toolId = row[tc];
    if (!toolId) continue;
    if (!rowsByTool.has(toolId)) rowsByTool.set(toolId, []);
    rowsByTool.get(toolId).push(row);
  }

  const toolsWithProducts = new Set();
  for (const [toolId, rows] of rowsByTool.entries()) {
    const loc = resolveArea(toolId);

    // For HDBI tools with a defined 2D layout, a cell only counts as occupied
    // if its cell ID actually matches a position in the layout. Otherwise the
    // card renders that cell as "Empty" regardless of the raw product value.
    let expectedCells = null;
    if (loc.area === 'HDBI') {
      const layoutKey = getHdbiLayoutKey(toolId);
      const layout = HDBI_LAYOUTS[layoutKey];
      if (layout) {
        expectedCells = new Set();
        for (const col of layout.columns) {
          for (const rowId of layout.rows) {
            expectedCells.add((col ? `${col}${rowId}` : rowId).toUpperCase());
          }
        }
      }
    }

    const hasAssignedProduct = rows.some((row) => {
      if (expectedCells) {
        const cellId = String(row[cc] || '').trim().toUpperCase();
        if (!expectedCells.has(cellId)) return false;
      }
      const product = normalizeProductForStats(row[pc], loc.area, toolId);
      return !isEmptyProductForStats(product);
    });
    if (hasAssignedProduct) toolsWithProducts.add(toolId);
  }

  return data.filter((row) => toolsWithProducts.has(row[tc]));
}

function isToolFilterActive() {
  const allCbs = [...toolCheckList.querySelectorAll('input[type=checkbox][data-tool-item="1"]')];
  if (!allCbs.length) return false;
  const checked = allCbs.filter(cb => cb.checked).length;
  // If all tools are checked, user is not restricting by tool.
  return checked > 0 && checked < allCbs.length;
}

function syncToolToggleAllState() {
  const allCb = document.getElementById('toolToggleAll');
  if (!allCb) return;

  const items = [...toolCheckList.querySelectorAll('input[type=checkbox][data-tool-item="1"]')];
  const emptyCbEl = document.getElementById('toolToggleEmpty');
  const allCheckboxes = emptyCbEl ? [...items, emptyCbEl] : items;

  if (!allCheckboxes.length) {
    allCb.checked = false;
    allCb.indeterminate = false;
    return;
  }

  const checked = allCheckboxes.filter(cb => cb.checked).length;
  allCb.checked = checked === allCheckboxes.length;
  allCb.indeterminate = checked > 0 && checked < allCheckboxes.length;
}

// ─── Apply Filters ────────────────────────────────────────────────────────────
applyBtn.addEventListener('click', async () => {
  if (!isMappingComplete(colMapping)) { openSettings(); return; }
  await fetchAndRender();
});

async function fetchAndRender() {
  const area = filterArea.value;
  const sub  = filterSub.value;
  let checkedTools = selectedTools();
  const toolFilterActive = isToolFilterActive();
  // For HDMX the Tools Empty checkbox is hidden; always show all tools by default.
  const showEmptyTools = area === 'HDMX' ? true : includeToolsEmpty();
  const tc = colMapping.toolCol;
  const productQuery = productSearch.value.trim();
  const coolantQuery = String(filterCoolant?.value || '').toUpperCase();
  const baggedQuery  = String(filterBagged?.value  || '');

  const ready = await ensureBaseDataLoaded(false);
  if (!ready) return;

  // When in HDMX and the "Tools Bagged" checkbox is unchecked with no explicit
  // dropdown selection, treat it as "Active only" — hide all bagged tools.
  // Exception: if the user explicitly selected individual tools from the list,
  // their selection takes priority and no implicit filter is applied.
  // NOTE: showEmptyTools (Tools Empty checkbox) for HDMX is handled separately
  // via filterOutEmptyOnlyTools — it does NOT feed into effectiveBaggedQuery.
  const effectiveBaggedQuery = baggedQuery;

  let visibleData = applyUiFilters(baseDataCache, area, sub, checkedTools, toolFilterActive);
  if (!area) {
    visibleData = visibleData.filter((row) => !isMachineHiddenInAll(row[tc] || ''));
  }

  if (area === 'HDMX' && coolantQuery) {
    visibleData = visibleData.filter((row) => getHdmxCoolant(row[tc]) === coolantQuery);
  }

  if (area === 'HDMX' && effectiveBaggedQuery) {
    if (effectiveBaggedQuery === 'bagged') {
      visibleData = visibleData.filter(r => isToolDisabled(r[tc]));
    } else if (effectiveBaggedQuery === 'active') {
      visibleData = visibleData.filter(r => !isToolDisabled(r[tc]));
    }
  }

  // For HDMX, bagged/active filtering replaces the empty-tool filter entirely.
  // For other areas, apply the empty-tool filter when the checkbox is unchecked.
  const isBaggedOnlyView = area === 'HDMX' && effectiveBaggedQuery === 'bagged';
  if (!showEmptyTools && !isBaggedOnlyView) {
    visibleData = filterOutEmptyOnlyTools(visibleData);
  }

  const productFilterResult = applyProductToolFilter(visibleData, productQuery);
  visibleData = productFilterResult.rows;

  highlightSet.clear();
  for (const key of productFilterResult.highlightKeys) highlightSet.add(key);

  currentData = visibleData;
  buildProductColorMap(visibleData);
  renderUsageStats(visibleData);
  // When a bagged filter is active, always allow static HDMX tool injection
  // so tools with no DB rows (truly empty) are still shown in the result.
  // Also force it when the user has specifically selected individual HDMX tools
  // so those cards always render even when they have no DB rows for the period.
  const allowStatic = (!productFilterResult.filtered && showEmptyTools)
    || (area === 'HDMX' && !!effectiveBaggedQuery)
    || (area === 'HDMX' && toolFilterActive);
  renderMachineGrid(
    visibleData,
    toolFilterActive ? checkedTools : [],
    allowStatic,
    area === 'HDMX' ? coolantQuery          : '',
    area === 'HDMX' ? effectiveBaggedQuery  : ''
  );
  recordCount.textContent = `${visibleData.length} rows`;
  const inForecast = isForecastPeriod(filterWW.value, filterDay.value);
  if (productFilterResult.filtered) {
    setStatus(visibleData.length
      ? `Showing tools with product "${productQuery}" (${productFilterResult.matchCount} matching cell(s)).`
      : `No tools found with product "${productQuery}" for the selected filters.`);
  } else if (!visibleData.length) {
    setStatus('No data for the selected filters.');
  } else if (!inForecast) {
    // Only clear status for normal (non-forecast) views; forecast status was
    // already set by buildForecastData and should remain visible.
    setStatus('');
  }
}

function applyProductToolFilter(data, query) {
  const q = String(query || '').trim();
  if (!q) return { rows: data, highlightKeys: new Set(), matchCount: 0, filtered: false };

  const tc = colMapping.toolCol;
  const cc = colMapping.cellCol;
  const pc = colMapping.productCol;
  const needle = q.toUpperCase();

  const matchingTools = new Set();
  const highlightKeys = new Set();
  let matchCount = 0;

  function canonicalToolToken(toolId) {
    const machine = String(toolId || '').trim();
    if (!machine) return '';
    const loc = resolveArea(machine);
    const token = formatToolTokenForFilter(machine, loc.area);
    return `${loc.area}||${loc.sub || ''}||${token}`;
  }

  for (const row of data) {
    const toolId = row[tc] || '';
    const cellId = row[cc] || '';
    const product = String(normalizeProductForCard(toolId, row[pc]) || '').toUpperCase();
    if (!product.includes(needle)) continue;
    matchingTools.add(canonicalToolToken(toolId));
    if (toolId && cellId) highlightKeys.add(`${toolId}||${cellId}`);
    matchCount += 1;
  }

  const rows = data.filter((row) => matchingTools.has(canonicalToolToken(row[tc] || '')));
  return { rows, highlightKeys, matchCount, filtered: true };
}

function normalizeProductForStats(value, area = '', toolId = '') {
  // Use the same normalization as the card cells so both always produce the
  // same canonical product name and therefore the same hash-based color.
  const canonical = normalizeProductForCard(toolId, value);

  // Stats-only: group all GNR-SP variants under one label for HDMX so the
  // stats table stays compact (does not affect card colors).
  if (String(area || '').toUpperCase() === 'HDMX' && /^GNR-SP/i.test(canonical)) {
    return 'GNR-SP';
  }

  return canonical;
}

function isEmptyProductForStats(product) {
  return /^empty$/i.test(String(product || '').trim());
}

function createEmptyStats() {
  // products: Map<productName, { total: number, byTool: Map<toolDisplayName, count> }>
  return { totalCells: 0, usedCells: 0, emptyCells: 0, products: new Map() };
}

function ensureProductStat(productsMap, productName) {
  let entry = productsMap.get(productName);
  if (!entry) {
    entry = { total: 0, byTool: new Map() };
    productsMap.set(productName, entry);
  }
  return entry;
}

function statsBucketKey(area, sub) {
  return `${area}||${sub || ''}`;
}

function buildStatsByBucket(data) {
  const tc = colMapping.toolCol;
  const cc = colMapping.cellCol;
  const pc = colMapping.productCol;

  const byTool = {};
  for (const row of data) {
    const toolId = row[tc];
    if (!toolId) continue;
    if (!byTool[toolId]) byTool[toolId] = [];
    byTool[toolId].push(row);
  }

  const statsByBucket = new Map();
  const activeCoolantFilter = String(filterCoolant?.value || '').toUpperCase();

  function ensureBucket(area, sub) {
    const key = statsBucketKey(area, sub);
    if (!statsByBucket.has(key)) statsByBucket.set(key, createEmptyStats());
    return statsByBucket.get(key);
  }

  for (const [toolId, rows] of Object.entries(byTool)) {
    const loc = resolveArea(toolId);
    if (!['HDBI', 'HDMX', 'PPV'].includes(loc.area)) continue;

    const byCell = new Map();
    for (const row of rows) {
      const cellId = String(row[cc] || '').trim().toUpperCase();
      if (!cellId) continue;

      const product = normalizeProductForStats(row[pc], loc.area, toolId);
      const prev = byCell.get(cellId);
      // Prefer a non-empty value if there are duplicate rows per cell.
      if (!prev || (isEmptyProductForStats(prev) && !isEmptyProductForStats(product))) {
        byCell.set(cellId, product);
      }
    }

    const fixedCells = fixedCellsForTool(loc.area, toolId);
    const cellsToCount = (fixedCells && fixedCells.length)
      ? fixedCells.map(c => String(c).trim().toUpperCase())
      : [...byCell.keys()];

    // For HDMX, bagged tools count all their cells as Empty in the statistics.
    const isBaggedHdmx = loc.area === 'HDMX' && isToolDisabled(toolId);

    for (const cellId of cellsToCount) {
      const product = isBaggedHdmx ? 'Empty' : (byCell.get(cellId) || 'Empty');
      const areaStats = ensureBucket(loc.area, loc.sub || '');
      areaStats.totalCells += 1;

      if (isEmptyProductForStats(product)) {
        areaStats.emptyCells += 1;
      } else {
        areaStats.usedCells += 1;
        const entry = ensureProductStat(areaStats.products, product);
        entry.total += 1;
        const toolLabel = formatMachineDisplayName(toolId) || String(toolId);
        entry.byTool.set(toolLabel, (entry.byTool.get(toolLabel) || 0) + 1);
      }
    }
  }

  // Keep HDMX stats aligned with the rendered cards by including tools from the
  // static HDMX list even when they have no active DB rows (all cells = Empty).
  // If product filter is active, do not inject empty tools because the grid also
  // hides tools without at least one matching product.
  const hasProductFilter = String(productSearch.value || '').trim().length > 0;
  if (!hasProductFilter && includeToolsEmpty()) {
    const selectedTokens = selectedTools();
    const toolFilterActive = isToolFilterActive();
    const existingHdmxDigits = new Set(
      Object.keys(byTool)
        .filter((t) => resolveArea(t).area === 'HDMX')
        .map((t) => String(t).replace(/\D/g, '').slice(-4))
    );

    const staticHdmxTools = AREA_MAP.HDMX?.tools?.[''] || [];
    for (const toolId of staticHdmxTools) {
      if (activeCoolantFilter && getHdmxCoolant(toolId) !== activeCoolantFilter) continue;
      const digits = String(toolId).replace(/\D/g, '').slice(-4);
      if (existingHdmxDigits.has(digits)) continue;
      if (toolFilterActive && !selectedTokens.some((token) => machineMatchesToken(toolId, token))) continue;

      const areaStats = ensureBucket('HDMX', '');
      const fixedCells = fixedCellsForTool('HDMX', toolId) || CLASS_CELLS;
      areaStats.totalCells += fixedCells.length;
      areaStats.emptyCells += fixedCells.length;
    }
  }

  return statsByBucket;
}

function renderAreaStatsCard(container, areaName, subArea, areaStats) {
  if (!container) return;

  const cls = areaName === 'HDBI'
    ? 'stats-card-hdbi'
    : areaName === 'HDMX'
      ? 'stats-card-hdmx'
      : 'stats-card-ppv';
  const utilization = areaStats.totalCells > 0
    ? ((areaStats.usedCells / areaStats.totalCells) * 100)
    : 0;
  const productRows = [...areaStats.products.entries()]
    .sort((a, b) => b[1].total - a[1].total || a[0].localeCompare(b[0]))
    .slice(0, 20)
    .map(([product, info]) => {
      const cls = productClass(product);
      const toolsText = [...info.byTool.entries()]
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
        .map(([tool, count]) => `${tool} (${count})`)
        .join(' - ');
      return `<tr>
        <td class="stats-color-cell ${cls}"></td>
        <td>${product}</td>
        <td>${info.total}</td>
        <td class="stats-tools-cell">${toolsText}</td>
      </tr>`;
    })
    .join('');

  const productContent = productRows
    ? `<table class="stats-product-table"><thead><tr><th>Color</th><th>Product</th><th>Total Cells</th><th class="th-tools">Tools (Cells)</th></tr></thead><tbody>${productRows}</tbody></table>`
    : '<div class="stats-empty">No used cells for the current filters.</div>';

  container.className = `stats-card ${cls}`;
  container.innerHTML = `
    <h4>${subArea ? `${areaName} - ${subArea}` : areaName}</h4>
    <div class="stats-kpi-row">
      <div class="stats-kpi">
        <span class="stats-kpi-label">Total Cells</span>
        <span class="stats-kpi-value">${areaStats.totalCells}</span>
      </div>
      <div class="stats-kpi">
        <span class="stats-kpi-label">In Use</span>
        <span class="stats-kpi-value">${areaStats.usedCells}</span>
      </div>
      <div class="stats-kpi">
        <span class="stats-kpi-label">Empty</span>
        <span class="stats-kpi-value">${areaStats.emptyCells}</span>
      </div>
      <div class="stats-kpi">
        <span class="stats-kpi-label">Utilization</span>
        <span class="stats-kpi-value">${utilization.toFixed(1)}%</span>
      </div>
    </div>
    <div class="stats-product-section">
      <div class="stats-products-title">Cells by Product</div>
      <div class="stats-product-table-wrapper">
        ${productContent}
      </div>
    </div>
  `;
}

function buildHdmxCoolantStats(data) {
  const tc = colMapping.toolCol;
  const cc = colMapping.cellCol;
  const pc = colMapping.productCol;

  const allHdmxTools = AREA_MAP.HDMX?.tools?.[''] || [];
  const selectedTokens = selectedTools();
  const toolFilterActive = isToolFilterActive();
  const coolantFilter = String(filterCoolant?.value || '').toUpperCase();
  const rawBaggedFilter = String(filterBagged?.value  || '');
  const baggedFilter = rawBaggedFilter;

  const scopedTools = allHdmxTools.filter((toolId) => {
    if (coolantFilter && getHdmxCoolant(toolId) !== coolantFilter) return false;
    if (baggedFilter === 'bagged' && !isToolDisabled(toolId)) return false;
    if (baggedFilter === 'active' &&  isToolDisabled(toolId)) return false;
    if (!toolFilterActive || !selectedTokens.length) return true;
    return selectedTokens.some((token) => machineMatchesToken(toolId, token));
  });

  const rowsByTool = new Map(scopedTools.map((toolId) => [toolId, []]));
  for (const row of (data || [])) {
    const rawTool = row[tc] || '';
    if (resolveArea(rawTool).area !== 'HDMX') continue;
    const toolDigits = String(rawTool).replace(/\D/g, '').slice(-4);
    if (!rowsByTool.has(toolDigits)) continue;
    rowsByTool.get(toolDigits).push(row);
  }

  const byCoolant = {
    HFE: { tools: 0, totalCells: 0, usedCells: 0, emptyCells: 0 },
    EGDI: { tools: 0, totalCells: 0, usedCells: 0, emptyCells: 0 },
  };

  for (const toolId of scopedTools) {
    const coolant = String(HDMX_COOLANT_BY_TOOL[toolId] || '').toUpperCase();
    if (!byCoolant[coolant]) continue;

    const fixedCells = fixedCellsForTool('HDMX', toolId) || CLASS_CELLS;
    const rows = rowsByTool.get(toolId) || [];

    const byCell = new Map();
    for (const row of rows) {
      const cellId = String(row[cc] || '').trim().toUpperCase();
      if (!cellId) continue;
      const product = normalizeProductForStats(row[pc], 'HDMX', toolId);
      const prev = byCell.get(cellId);
      if (!prev || (isEmptyProductForStats(prev) && !isEmptyProductForStats(product))) {
        byCell.set(cellId, product);
      }
    }

    let used = 0;
    for (const cellId of fixedCells) {
      const product = byCell.get(String(cellId).trim().toUpperCase()) || 'Empty';
      if (!isEmptyProductForStats(product)) used += 1;
    }

    const total = fixedCells.length;
    byCoolant[coolant].tools += 1;
    byCoolant[coolant].totalCells += total;
    byCoolant[coolant].usedCells += used;
    byCoolant[coolant].emptyCells += (total - used);
  }

  return byCoolant;
}

function renderHdmxCoolantStatsCard(container, coolantStats) {
  if (!container) return;

  const CELLS_PER_HDMX_TOOL = 10;

  const rows = ['HFE', 'EGDI']
    .map((coolant) => {
      const s = coolantStats[coolant] || { tools: 0, totalCells: 0, usedCells: 0, emptyCells: 0 };
      const totalCells = s.tools * CELLS_PER_HDMX_TOOL;
      const utilization = totalCells > 0 ? ((s.usedCells / totalCells) * 100) : 0;
      return `<tr>
        <td>${coolant}</td>
        <td>${s.tools}</td>
        <td>${totalCells}</td>
        <td>${s.usedCells}</td>
        <td>${s.emptyCells}</td>
        <td>${utilization.toFixed(1)}%</td>
      </tr>`;
    })
    .join('');

  container.className = 'stats-card stats-card-hdmx-coolant';
  container.innerHTML = `
    <h4>HDMX - Coolant Stats</h4>
    <div class="stats-products-title">By Coolant (HFE / EGDI)</div>
    <table class="stats-product-table">
      <thead>
        <tr><th>Coolant</th><th>Tools</th><th>Total Cells</th><th>Used Cells</th><th>Empty Cells</th><th>Utilization</th></tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

/**
 * Attach multi-select toggle-panel buttons to tabsEl/bodyEl.
 * panels: [{ id, label, render(containerEl) }]
 */
function attachTogglePanels(tabsEl, bodyEl, panels) {
  tabsEl.innerHTML = '';
  bodyEl.innerHTML = '';
  bodyEl.style.display = 'none';

  const openPanels = new Map(); // panelId → containerEl

  for (const panel of panels) {
    const btn = document.createElement('button');
    btn.className = 'area-tab stats-panel-tab';
    btn.type = 'button';
    btn.dataset.panelId = panel.id;
    btn.textContent = panel.label;

    btn.addEventListener('click', () => {
      if (openPanels.has(panel.id)) {
        openPanels.get(panel.id).remove();
        openPanels.delete(panel.id);
        btn.classList.remove('is-active');
        if (openPanels.size === 0) bodyEl.style.display = 'none';
      } else {
        btn.classList.add('is-active');
        const panelDiv = document.createElement('div');
        panelDiv.className = 'stats-panel-item';
        panelDiv.dataset.panelId = panel.id;
        openPanels.set(panel.id, panelDiv);
        // Re-insert in declaration order
        bodyEl.innerHTML = '';
        for (const p of panels) {
          if (openPanels.has(p.id)) bodyEl.appendChild(openPanels.get(p.id));
        }
        bodyEl.style.display = '';
        panel.render(panelDiv);
      }
    });

    tabsEl.appendChild(btn);
  }
}

function getStatsTargets(activeArea) {
  if (activeArea === 'HDBI') {
    return [
      { area: 'HDBI', sub: 'HDBI' },
      { area: 'HDBI', sub: 'BDC' },
    ];
  }

  if (activeArea === 'PPV') {
    return AREA_MAP.PPV.subs.map((sub) => ({ area: 'PPV', sub }));
  }

  if (activeArea === 'HDMX') {
    return [{ area: 'HDMX', sub: '' }];
  }

  return [
    { area: 'HDBI', sub: 'HDBI' },
    { area: 'HDBI', sub: 'BDC' },
    { area: 'HDMX', sub: '' },
    ...AREA_MAP.PPV.subs.map((sub) => ({ area: 'PPV', sub })),
  ];
}

function renderUsageStats(data) {
  if (!statsPanelTabs) return;

  // Panels are now rendered inline inside each section header.
  // Always hide the top-level stats section.
  if (statsSection) statsSection.style.display = 'none';
  statsPanelTabs.innerHTML = '';
  statsPanelBody.innerHTML = '';
}

// ─── HDMX inline recent-changes card ─────────────────────────────────────────
async function renderHdmxInlineChanges(container, appendMode = false) {
  let root;
  if (appendMode) {
    root = document.createElement('div');
    root.className = 'hdmx-inline-changes-section';
    container.appendChild(root);
  } else {
    container.className = 'stats-card stats-card-hdmx-changes';
    root = container;
  }

  root.innerHTML = `
    <h4>&#9719; Recent Allocation Changes &mdash; HDMX</h4>
    <div class="hdmx-changes-toolbar">
      <input type="text" class="filter-input changes-search-input hdmx-changes-search" placeholder="Filter by machine, cell or product…" />
    </div>
    <div class="hdmx-changes-body"><p class="changes-loading"><span class="spinner"></span> Loading…</p></div>
  `;

  const searchInput = root.querySelector('.hdmx-changes-search');
  const bodyDiv     = root.querySelector('.hdmx-changes-body');

  async function loadAndRender(filterText = '') {
    // Reuse global cache if already loaded, otherwise fetch
    if (!_changesData.length) {
      try {
        const params = mappingParams();
        const result = await apiFetch(`/api/changes?${params}`);
        _changesData = result.changes || [];
      } catch (err) {
        bodyDiv.innerHTML = `<p class="changes-error">Error loading changes: ${err.message}</p>`;
        return;
      }
    }

    // Keep only HDMX machines
    const hdmxChanges = _changesData.filter(ch => resolveArea(ch.machine || '').area === 'HDMX');
    const q = filterText.trim().toLowerCase();

    // Group by comparison period (same logic as modal)
    const groups = new Map();
    for (const ch of hdmxChanges) {
      const gKey = `${ch.fromWW}|${ch.fromDay}→${ch.toWW}|${ch.toDay}`;
      if (!groups.has(gKey)) {
        groups.set(gKey, { from: { ww: ch.fromWW, day: ch.fromDay }, to: { ww: ch.toWW, day: ch.toDay }, items: [] });
      }
      groups.get(gKey).items.push(ch);
    }

    let html = '';
    for (const [, group] of groups) {
      const items = q
        ? group.items.filter(c =>
            (c.machine    || '').toLowerCase().includes(q) ||
            (c.cell       || '').toLowerCase().includes(q) ||
            (c.oldProduct || '').toLowerCase().includes(q) ||
            (c.newProduct || '').toLowerCase().includes(q)
          )
        : group.items;
      if (!items.length) continue;

      const fromLabel = formatChangePeriod(group.from.ww, group.from.day);
      const toLabel   = formatChangePeriod(group.to.ww,   group.to.day);
      html += `<div class="changes-group">
        <div class="changes-group-header">
          <span class="changes-period-from">${fromLabel}</span>
          <span class="changes-arrow">&#8594;</span>
          <span class="changes-period-to">${toLabel}</span>
          <span class="changes-count">${items.length} change${items.length !== 1 ? 's' : ''}</span>
        </div>
        ${renderChangesTable(items)}
      </div>`;
    }

    if (!html) {
      bodyDiv.innerHTML = hdmxChanges.length
        ? '<p class="changes-empty">No changes match the current filter.</p>'
        : '<p class="changes-empty">No allocation changes detected in HDMX in the last 2 days.</p>';
    } else {
      bodyDiv.innerHTML = html;
    }
  }

  searchInput.addEventListener('input', () => loadAndRender(searchInput.value));
  await loadAndRender('');
}

// ─── HDBI inline recent-changes card ─────────────────────────────────────────
async function renderHdbiSubInlineChanges(container, sub) {
  container.innerHTML = `
    <h4>&#9719; Recent Allocation Changes &mdash; HDBI${sub && sub !== 'HDBI' ? ' ' + sub : ''}</h4>
    <div class="hdmx-changes-toolbar">
      <input type="text" class="filter-input changes-search-input hdbi-sub-changes-search" placeholder="Filter by machine, cell or product…" />
    </div>
    <div class="hdmx-changes-body"><p class="changes-loading"><span class="spinner"></span> Loading…</p></div>
  `;

  const searchInput = container.querySelector('.hdbi-sub-changes-search');
  const bodyDiv     = container.querySelector('.hdmx-changes-body');

  async function loadAndRender(filterText = '') {
    if (!_changesData.length) {
      try {
        const params = mappingParams();
        const result = await apiFetch(`/api/changes?${params}`);
        _changesData = result.changes || [];
      } catch (err) {
        bodyDiv.innerHTML = `<p class="changes-error">Error loading changes: ${err.message}</p>`;
        return;
      }
    }

    const subChanges = _changesData.filter(ch => {
      const loc = resolveArea(ch.machine || '');
      return loc.area === 'HDBI' && loc.sub === sub;
    });
    const q = filterText.trim().toLowerCase();

    const groups = new Map();
    for (const ch of subChanges) {
      const gKey = `${ch.fromWW}|${ch.fromDay}\u2192${ch.toWW}|${ch.toDay}`;
      if (!groups.has(gKey)) {
        groups.set(gKey, { from: { ww: ch.fromWW, day: ch.fromDay }, to: { ww: ch.toWW, day: ch.toDay }, items: [] });
      }
      groups.get(gKey).items.push(ch);
    }

    let html = '';
    for (const [, group] of groups) {
      const items = q
        ? group.items.filter(c =>
            (c.machine    || '').toLowerCase().includes(q) ||
            (c.cell       || '').toLowerCase().includes(q) ||
            (c.oldProduct || '').toLowerCase().includes(q) ||
            (c.newProduct || '').toLowerCase().includes(q)
          )
        : group.items;
      if (!items.length) continue;

      const fromLabel = formatChangePeriod(group.from.ww, group.from.day);
      const toLabel   = formatChangePeriod(group.to.ww,   group.to.day);
      html += `<div class="changes-group">
        <div class="changes-group-header">
          <span class="changes-period-from">${fromLabel}</span>
          <span class="changes-arrow">&#8594;</span>
          <span class="changes-period-to">${toLabel}</span>
          <span class="changes-count">${items.length} change${items.length !== 1 ? 's' : ''}</span>
        </div>
        ${renderChangesTable(items)}
      </div>`;
    }

    if (!html) {
      bodyDiv.innerHTML = subChanges.length
        ? '<p class="changes-empty">No changes match the current filter.</p>'
        : `<p class="changes-empty">No allocation changes detected in HDBI ${sub} in the last 2 days.</p>`;
    } else {
      bodyDiv.innerHTML = html;
    }
  }

  searchInput.addEventListener('input', () => loadAndRender(searchInput.value));
  await loadAndRender('');
}

async function renderHdbiInlineChanges(container, appendMode = false) {
  let root;
  if (appendMode) {
    root = document.createElement('div');
    root.className = 'hdbi-inline-changes-section';
    container.appendChild(root);
  } else {
    root = container;
  }

  root.innerHTML = `
    <h4>&#9719; Recent Allocation Changes &mdash; HDBI</h4>
    <div class="hdmx-changes-toolbar">
      <input type="text" class="filter-input changes-search-input hdmx-changes-search" placeholder="Filter by machine, cell or product…" />
    </div>
    <div class="hdmx-changes-body"><p class="changes-loading"><span class="spinner"></span> Loading…</p></div>
  `;

  const searchInput = root.querySelector('.hdmx-changes-search');
  const bodyDiv     = root.querySelector('.hdmx-changes-body');

  async function loadAndRender(filterText = '') {
    if (!_changesData.length) {
      try {
        const params = mappingParams();
        const result = await apiFetch(`/api/changes?${params}`);
        _changesData = result.changes || [];
      } catch (err) {
        bodyDiv.innerHTML = `<p class="changes-error">Error loading changes: ${err.message}</p>`;
        return;
      }
    }

    // Keep only HDBI machines (both HDBI and BDC sub-areas)
    const hdbiChanges = _changesData.filter(ch => resolveArea(ch.machine || '').area === 'HDBI');
    const q = filterText.trim().toLowerCase();

    const groups = new Map();
    for (const ch of hdbiChanges) {
      const gKey = `${ch.fromWW}|${ch.fromDay}→${ch.toWW}|${ch.toDay}`;
      if (!groups.has(gKey)) {
        groups.set(gKey, { from: { ww: ch.fromWW, day: ch.fromDay }, to: { ww: ch.toWW, day: ch.toDay }, items: [] });
      }
      groups.get(gKey).items.push(ch);
    }

    let html = '';
    for (const [, group] of groups) {
      const items = q
        ? group.items.filter(c =>
            (c.machine    || '').toLowerCase().includes(q) ||
            (c.cell       || '').toLowerCase().includes(q) ||
            (c.oldProduct || '').toLowerCase().includes(q) ||
            (c.newProduct || '').toLowerCase().includes(q)
          )
        : group.items;
      if (!items.length) continue;

      const fromLabel = formatChangePeriod(group.from.ww, group.from.day);
      const toLabel   = formatChangePeriod(group.to.ww,   group.to.day);
      html += `<div class="changes-group">
        <div class="changes-group-header">
          <span class="changes-period-from">${fromLabel}</span>
          <span class="changes-arrow">&#8594;</span>
          <span class="changes-period-to">${toLabel}</span>
          <span class="changes-count">${items.length} change${items.length !== 1 ? 's' : ''}</span>
        </div>
        ${renderChangesTable(items)}
      </div>`;
    }

    if (!html) {
      bodyDiv.innerHTML = hdbiChanges.length
        ? '<p class="changes-empty">No changes match the current filter.</p>'
        : '<p class="changes-empty">No allocation changes detected in HDBI in the last 2 days.</p>';
    } else {
      bodyDiv.innerHTML = html;
    }
  }

  searchInput.addEventListener('input', () => loadAndRender(searchInput.value));
  await loadAndRender('');
}

// ─── PPV inline recent-changes panel ────────────────────────────────────────
async function renderPpvInlineChanges(container, sub) {
  container.innerHTML = `
    <h4>&#9719; Recent Allocation Changes &mdash; PPV ${sub || ''}</h4>
    <div class="hdmx-changes-toolbar">
      <input type="text" class="filter-input changes-search-input ppv-changes-search" placeholder="Filter by machine, cell or product…" />
    </div>
    <div class="hdmx-changes-body"><p class="changes-loading"><span class="spinner"></span> Loading…</p></div>
  `;

  const searchInput = container.querySelector('.ppv-changes-search');
  const bodyDiv     = container.querySelector('.hdmx-changes-body');

  async function loadAndRender(filterText = '') {
    if (!_changesData.length) {
      try {
        const params = mappingParams();
        const result = await apiFetch(`/api/changes?${params}`);
        _changesData = result.changes || [];
      } catch (err) {
        bodyDiv.innerHTML = `<p class="changes-error">Error loading changes: ${err.message}</p>`;
        return;
      }
    }

    // Filter by PPV area and specific sub-area (HST / SST / PTC)
    const ppvChanges = _changesData.filter(ch => {
      const loc = resolveArea(ch.machine || '');
      return loc.area === 'PPV' && (!sub || loc.sub === sub);
    });
    const q = filterText.trim().toLowerCase();

    const groups = new Map();
    for (const ch of ppvChanges) {
      const gKey = `${ch.fromWW}|${ch.fromDay}→${ch.toWW}|${ch.toDay}`;
      if (!groups.has(gKey)) {
        groups.set(gKey, { from: { ww: ch.fromWW, day: ch.fromDay }, to: { ww: ch.toWW, day: ch.toDay }, items: [] });
      }
      groups.get(gKey).items.push(ch);
    }

    let html = '';
    for (const [, group] of groups) {
      const items = q
        ? group.items.filter(c =>
            (c.machine    || '').toLowerCase().includes(q) ||
            (c.cell       || '').toLowerCase().includes(q) ||
            (c.oldProduct || '').toLowerCase().includes(q) ||
            (c.newProduct || '').toLowerCase().includes(q)
          )
        : group.items;
      if (!items.length) continue;

      const fromLabel = formatChangePeriod(group.from.ww, group.from.day);
      const toLabel   = formatChangePeriod(group.to.ww,   group.to.day);
      html += `<div class="changes-group">
        <div class="changes-group-header">
          <span class="changes-period-from">${fromLabel}</span>
          <span class="changes-arrow">&#8594;</span>
          <span class="changes-period-to">${toLabel}</span>
          <span class="changes-count">${items.length} change${items.length !== 1 ? 's' : ''}</span>
        </div>
        ${renderChangesTable(items)}
      </div>`;
    }

    if (!html) {
      bodyDiv.innerHTML = ppvChanges.length
        ? '<p class="changes-empty">No changes match the current filter.</p>'
        : `<p class="changes-empty">No allocation changes detected in PPV ${sub || ''} in the last 2 days.</p>`;
    } else {
      bodyDiv.innerHTML = html;
    }
  }

  searchInput.addEventListener('input', () => loadAndRender(searchInput.value));
  await loadAndRender('');
}

// ─── Render machine grid ──────────────────────────────────────────────────────
// Groups machines by area/sub using DB data after applying UI filters.
function renderMachineGrid(data, activeToolTokens = [], allowStaticHdmxTools = true, hdmxCoolantFilter = '', hdmxBaggedFilter = '') {
  machineGrid.innerHTML = '';
  emptyState.style.display = 'none';

  const activeAreaFilter = filterArea.value;
  const toolFilterActive = activeToolTokens.length > 0;

  const hasStaticTools = allowStaticHdmxTools && (!activeAreaFilter || activeAreaFilter === 'HDMX') &&
    (AREA_MAP.HDMX?.tools?.[''] || []).length > 0;

  if (!data.length && !hasStaticTools) {
    emptyState.style.display = '';
    return;
  }

  const tc = colMapping.toolCol;

  // Group data rows by machine name
  const byTool = {};
  for (const row of data) {
    const t = row[tc];
    if (!byTool[t]) byTool[t] = [];
    byTool[t].push(row);
  }

  // Inject static HDMX tools that have no DB data (no active products) so they
  // render with all cells as 'Empty' using the fixed cell template.
  // When a tool filter is active, only inject tools that match the selected tokens.
  if (allowStaticHdmxTools && (!activeAreaFilter || activeAreaFilter === 'HDMX')) {
    const hdmxStaticTools = AREA_MAP.HDMX?.tools?.[''] || [];
    for (const toolId of hdmxStaticTools) {
      if (hdmxCoolantFilter && getHdmxCoolant(toolId) !== hdmxCoolantFilter) continue;
      if (hdmxBaggedFilter === 'bagged'  && !isToolDisabled(toolId)) continue;
      if (hdmxBaggedFilter === 'active'  &&  isToolDisabled(toolId)) continue;
      const digits = String(toolId).replace(/\D/g, '').slice(-4);
      const alreadyPresent = Object.keys(byTool).some(k => {
        const kDigits = String(k).replace(/\D/g, '').slice(-4);
        return kDigits === digits && resolveArea(k).area === 'HDMX';
      });
      if (alreadyPresent) continue;
      if (toolFilterActive && !activeToolTokens.some(t => machineMatchesToken(toolId, t))) continue;
      byTool[toolId] = [];
    }
  }

  // Determine which machines to show and in which order/section from filtered data.
  const sections = buildSectionPlanFromData(byTool);
  const activeArea = activeAreaFilter;
  const ppvStatsByBucket = activeArea === 'PPV' ? buildStatsByBucket(data) : null;
  const hdbiStatsByBucket = activeArea === 'HDBI' ? buildStatsByBucket(data) : null;

  let totalCards = 0;
  for (const section of sections) {
    // Safe ID for CSS (replace spaces)
    const secId = `${section.area}-${section.sub || 'all'}`.replace(/\s/g, '_');

    // Build section header using DOM methods so we can append tabs inside it after
    const secHeader = document.createElement('div');
    secHeader.className = `section-header area-bg-${section.area}`;
    if (section.sub) secHeader.dataset.sub = section.sub;
    const _areaSpan = document.createElement('span');
    _areaSpan.className = 'section-area';
    _areaSpan.textContent = section.area;
    secHeader.appendChild(_areaSpan);
    if (section.sub) {
      // Don't show the sub badge when it's the same text as the area (e.g. HDBI › HDBI)
      if (section.sub.toUpperCase() !== section.area.toUpperCase()) {
        const _subSpan = document.createElement('span');
        _subSpan.className = 'section-sub';
        _subSpan.textContent = section.sub;
        secHeader.appendChild(_subSpan);
      }
    }
    // WW / Day label  e.g. "– WW20.3"
    const _wwVal  = String(filterWW?.value  || '').replace(/\D/g, '').slice(-2);
    const _dayVal = String(filterDay?.value || '').trim();
    if (_wwVal) {
      const _wwSpan = document.createElement('span');
      _wwSpan.className = 'section-ww-label';
      _wwSpan.textContent = `WW${_wwVal}${_dayVal ? '.' + _dayVal : ''}`;
      secHeader.appendChild(_wwSpan);
    }
    const _countSpan = document.createElement('span');
    _countSpan.className = 'section-count';
    _countSpan.id = `sec-${secId}`;
    secHeader.appendChild(_countSpan);
    machineGrid.appendChild(secHeader);

    const secGrid = document.createElement('div');
    secGrid.className = 'section-cards';
    if (section.sub) secGrid.dataset.sub = section.sub;
    if (section.area === 'HDMX') {
      const activeCols = document.querySelector('.hdmx-cols-btn.is-active');
      if (activeCols) secGrid.dataset.hdmxCols = activeCols.dataset.cols;
    }
    if (section.area === 'PPV') {
      const activeCols = document.querySelector('.ppv-cols-btn.is-active');
      if (activeCols) secGrid.dataset.ppvCols = activeCols.dataset.cols;
    }

    // HDMX section: toggle tabs in header, body as first flex item (order:-1) in secGrid
    if (section.area === 'HDMX') {
      const tabsEl = document.createElement('div');
      tabsEl.className = 'stats-panel-tabs section-header-tabs';
      secHeader.appendChild(tabsEl);
      const bodyEl = document.createElement('div');
      bodyEl.className = 'section-stats-body';
      secGrid.appendChild(bodyEl);
      const _hdmxData = data;
      attachTogglePanels(tabsEl, bodyEl, [
        {
          id: 'usage',
          label: 'Cell & Product Usage',
          render: (container) => {
            const statsByBucket = buildStatsByBucket(_hdmxData || []);
            container.innerHTML = '';
            const targets = getStatsTargets('HDMX');
            for (const target of targets) {
              const card = document.createElement('article');
              container.appendChild(card);
              const stats = statsByBucket.get(statsBucketKey(target.area, target.sub)) || createEmptyStats();
              renderAreaStatsCard(card, target.area, target.sub, stats);
            }
          }
        },
        {
          id: 'coolant',
          label: 'HDMX Coolant Stats',
          render: (container) => {
            container.innerHTML = '';
            const coolantCard = document.createElement('article');
            container.appendChild(coolantCard);
            const coolantStats = buildHdmxCoolantStats(_hdmxData || []);
            renderHdmxCoolantStatsCard(coolantCard, coolantStats);
          }
        },
        {
          id: 'changes',
          label: 'Recent Allocation Changes',
          render: (container) => {
            container.innerHTML = '';
            const changesCard = document.createElement('div');
            container.appendChild(changesCard);
            renderHdmxInlineChanges(changesCard, false);
          }
        }
      ]);
    }

    // HDBI sub-areas: toggle tabs in header, body as first flex item (order:-1) in secGrid
    if (activeArea === 'HDBI' && section.area === 'HDBI' && section.sub) {
      const tabsEl = document.createElement('div');
      tabsEl.className = 'stats-panel-tabs section-header-tabs';
      secHeader.appendChild(tabsEl);
      const bodyEl = document.createElement('div');
      bodyEl.className = 'section-stats-body';
      secGrid.appendChild(bodyEl);
      const sub = section.sub;
      const stats = hdbiStatsByBucket.get(statsBucketKey('HDBI', sub)) || createEmptyStats();
      attachTogglePanels(tabsEl, bodyEl, [
        {
          id: 'usage',
          label: 'Cell & Product Usage',
          render: (container) => {
            container.innerHTML = '';
            const card = document.createElement('article');
            container.appendChild(card);
            renderAreaStatsCard(card, 'HDBI', sub, stats);
          }
        },
        {
          id: 'changes',
          label: 'Recent Allocation Changes',
          render: (container) => {
            container.innerHTML = '';
            const panel = document.createElement('div');
            container.appendChild(panel);
            renderHdbiSubInlineChanges(panel, sub);
          }
        }
      ]);
    }

    // PPV sub-areas: same pattern
    if (activeArea === 'PPV' && section.area === 'PPV' && section.sub) {
      const tabsEl = document.createElement('div');
      tabsEl.className = 'stats-panel-tabs section-header-tabs';
      secHeader.appendChild(tabsEl);
      const bodyEl = document.createElement('div');
      bodyEl.className = 'section-stats-body';
      secGrid.appendChild(bodyEl);
      const sub = section.sub;
      const stats = ppvStatsByBucket.get(statsBucketKey('PPV', sub)) || createEmptyStats();
      attachTogglePanels(tabsEl, bodyEl, [
        {
          id: 'usage',
          label: 'Cell & Product Usage',
          render: (container) => {
            container.innerHTML = '';
            const card = document.createElement('article');
            container.appendChild(card);
            renderAreaStatsCard(card, 'PPV', sub, stats);
          }
        },
        {
          id: 'changes',
          label: 'Recent Allocation Changes',
          render: (container) => {
            container.innerHTML = '';
            const panel = document.createElement('div');
            container.appendChild(panel);
            renderPpvInlineChanges(panel, sub);
          }
        }
      ]);
    }

    let sectionCards = 0;
    for (const toolId of section.tools) {
      // Definitive bagged-filter gate: same check as the BAGGED badge itself.
      if (section.area === 'HDMX' && hdmxBaggedFilter) {
        if (hdmxBaggedFilter === 'bagged' && !isToolDisabled(toolId)) continue;
        if (hdmxBaggedFilter === 'active' &&  isToolDisabled(toolId)) continue;
      }
      const rows = byTool[toolId] || [];
      secGrid.appendChild(buildMachineCard(toolId, rows, section.area, section.sub));
      sectionCards++;
      totalCards++;
    }
    // HDBI sub-area: inject a flex row-break so 4344 renders on its own row below 5082+4786
    if (section.area === 'HDBI' && section.sub === 'HDBI') {
      const rowBreak = document.createElement('div');
      rowBreak.className = 'hdbi-row-break';
      secGrid.appendChild(rowBreak);
    }
    machineGrid.appendChild(secGrid);

    const countEl = document.getElementById(`sec-${secId}`);
    if (countEl) countEl.textContent = `${sectionCards} machine${sectionCards !== 1 ? 's' : ''}`;
  }

  if (!totalCards) emptyState.style.display = '';
}

function applyUiFilters(data, area, sub, selectedToolIds, toolFilterActive = false) {
  const tc = colMapping.toolCol;
  const selected = selectedToolIds || [];

  if (!area) {
    if (!toolFilterActive || !selected.length) return data;
    return data.filter((row) => {
      const machine = row[tc] || '';
      return selected.some((token) => machineMatchesToken(machine, token));
    });
  }

  return data.filter((row) => {
    const machine = row[tc] || '';
    const loc = resolveArea(machine);
    if (loc.area !== area) return false;
    if (loc.area === 'PPV' && loc.sub === 'PDC') return false;
    if (sub && loc.sub !== sub) return false;
    if (!toolFilterActive || !selected.length) return true;
    return selected.some((token) => machineMatchesToken(machine, token));
  });
}

function machineMatchesToken(machineName, token) {
  const m = normalise(machineName);
  const t = normalise(token);
  if (!t) return false;
  if (m === t || m.includes(t)) return true;

  // Fallback: match by digits (e.g. BDC4944 should match CR03HBDC24944)
  const mDigits = m.replace(/\D/g, '');
  const tDigits = t.replace(/\D/g, '');
  if (tDigits && mDigits.includes(tDigits)) {
    if (t.includes('BDC')) return m.includes('BDC');
    if (t.includes('HDBI') || t.includes('HBI') || t.includes('MBI')) {
      return (m.includes('HDBI') || m.includes('HBI') || m.includes('MBI'));
    }
    return true;
  }

  return false;
}

/**
 * When no area filter: group machines returned by the DB into sections using
 * resolveArea(), preserving the canonical HDBI → HDMX → PPV order.
 */
function buildSectionPlanFromData(byTool) {
  // Collect machines per { area, sub } bucket
  const buckets = {}; // key: "area||sub"
  const order   = []; // insertion order for deterministic rendering

  // Define canonical order so sections always appear HDBI → HDMX → PPV
  const canonicalOrder = [
    { area: 'HDBI', sub: 'HDBI' },
    { area: 'HDBI', sub: 'BDC'  },
    { area: 'HDMX', sub: ''     },
    { area: 'CLASS',sub: ''     },
    { area: 'PPV',  sub: 'HST'  },
    { area: 'PPV',  sub: 'SST'  },
    { area: 'PPV',  sub: 'PTC'  },
    { area: 'OTHER',sub: ''     },
  ];
  for (const s of canonicalOrder) {
    buckets[`${s.area}||${s.sub}`] = { ...s, tools: [] };
  }

  for (const machineName of Object.keys(byTool)) {
    const { area, sub } = resolveArea(machineName);
    const key = `${area}||${sub}`;
    if (!buckets[key]) buckets[key] = { area, sub, tools: [] };
    buckets[key].tools.push(machineName);
  }

  // Sort tools: HDMX by numeric tool number, others alphabetically
  for (const bucket of Object.values(buckets)) {
    if (bucket.area === 'HDMX') {
      bucket.tools.sort((a, b) => {
        const numA = parseInt((String(a).match(/(\d{4,5})$/) || [])[1] || '0', 10);
        const numB = parseInt((String(b).match(/(\d{4,5})$/) || [])[1] || '0', 10);
        return numA - numB;
      });
    } else {
      bucket.tools.sort();
    }
  }

  return canonicalOrder
    .map(s => buckets[`${s.area}||${s.sub}`])
    .filter(s => s && s.tools.length > 0);
}

/**
 * When an area filter IS set: use the static AREA_MAP order but restrict
 * to machines that actually appear in byTool (i.e. have data).
 */
// Build a normalised lookup: stripped machine name → { area, sub }
// Strips spaces and converts to upper for fuzzy matching.
function buildMachineAreaLookup() {
  const lookup = new Map();
  for (const [area, aData] of Object.entries(AREA_MAP)) {
    for (const [sub, tools] of Object.entries(aData.tools)) {
      for (const t of tools) {
        lookup.set(normalise(t), { area, sub });
      }
    }
  }
  return lookup;
}
function normalise(s) { return String(s).replace(/\s+/g, '').toUpperCase(); }
let _machineAreaLookup = null;
function getMachineAreaLookup() {
  if (!_machineAreaLookup) _machineAreaLookup = buildMachineAreaLookup();
  return _machineAreaLookup;
}
function resolveArea(machineName) {
  const lk = getMachineAreaLookup();
  const norm = normalise(machineName);
  const direct = lk.get(norm);
  if (direct) return direct;

  if (norm.includes('CLASS')) {
    return { area: 'CLASS', sub: '' };
  }

  if (norm.startsWith('CR03DHHX')) {
    return { area: 'HDMX', sub: '' };
  }

  if (norm.startsWith('CR03DHST')) {
    return { area: 'PPV', sub: 'HST' };
  }

  if (norm.startsWith('CR03HPDC')) {
    return { area: 'OTHER', sub: '' };
  }

  if (norm.startsWith('CR03HPTC')) {
    return { area: 'PPV', sub: 'PTC' };
  }

  if (norm.startsWith('CR03TSST')) {
    return { area: 'PPV', sub: 'SST' };
  }

  // Heuristics for DB machine naming patterns with prefixes/suffixes.
  // Examples: CR03DHBI5082, CR03DHBI4786, CR03DMBI4786, CR03HBDC24944, CR03HBDC25257
  if ((norm.includes('HDBI') || norm.includes('HBI') || norm.includes('MBI')) && norm.includes('5082')) {
    return { area: 'HDBI', sub: 'HDBI' };
  }
  if ((norm.includes('HDBI') || norm.includes('HBI') || norm.includes('MBI')) && norm.includes('4344')) {
    return { area: 'HDBI', sub: 'HDBI' };
  }
  if ((norm.includes('HDBI') || norm.includes('HBI') || norm.includes('MBI')) && norm.includes('4786')) {
    return { area: 'HDBI', sub: 'HDBI' };
  }
  if (norm.includes('BDC') && norm.includes('4944')) {
    return { area: 'HDBI', sub: 'BDC' };
  }
  if (norm.includes('BDC') && norm.includes('5257')) {
    return { area: 'HDBI', sub: 'BDC' };
  }

  return { area: 'OTHER', sub: '' };
}

function formatMachineDisplayName(machineName) {
  const raw = String(machineName || '').trim();
  if (!raw) return raw;

  const upper = raw.toUpperCase();
  const tailDigits = (upper.match(/(\d{4,5})$/) || [])[1] || '';

  if (upper.startsWith('CR03DHST')) {
    const suffix = tailDigits ? tailDigits.slice(-4) : '';
    return suffix ? `HST${suffix}` : 'HST';
  }

  if (upper.startsWith('CR03HPDC')) {
    const suffix = tailDigits ? tailDigits.slice(-5) : '';
    return suffix ? `PDC${suffix}` : 'PDC';
  }

  if (upper.startsWith('CR03HPTC')) {
    const suffix = tailDigits ? tailDigits.slice(-4) : '';
    return suffix ? `PTC${suffix}` : 'PTC';
  }

  if (upper.startsWith('CR03TSST')) {
    const suffix = tailDigits ? tailDigits : '';
    return suffix ? `SST${suffix}` : 'SST';
  }

  // HDMX machines: CR03DHHX#### -> ####
  if (upper.startsWith('CR03DHHX')) {
    const suffix = tailDigits ? tailDigits.slice(-4) : '';
    return suffix || raw;
  }

  // HDBI/HBI/MBI machines: show canonical HDBI + last 4 digits (e.g. HDBI5082)
  if (upper.includes('HDBI') || upper.includes('HBI') || upper.includes('MBI')) {
    const suffix = tailDigits ? tailDigits.slice(-4) : '';
    return suffix ? `HDBI${suffix}` : 'HDBI';
  }

  // BDC machines: drop prefix and remove leading extra digit from 5-digit tails
  // e.g. CR03HBDC24944 -> BDC4944, CR03HBDC25257 -> BDC5257
  if (upper.includes('BDC')) {
    let suffix = tailDigits;
    if (suffix.length === 5) suffix = suffix.slice(1);
    if (suffix.length > 4) suffix = suffix.slice(-4);
    return suffix ? `BDC${suffix}` : 'BDC';
  }

  if (upper.includes('PDC')) {
    const suffix = tailDigits ? tailDigits.slice(-5) : '';
    return suffix ? `PDC${suffix}` : 'PDC';
  }

  if (upper.includes('PTC')) {
    const suffix = tailDigits ? tailDigits.slice(-4) : '';
    return suffix ? `PTC${suffix}` : 'PTC';
  }

  if (upper.includes('SST')) {
    const suffix = tailDigits ? tailDigits : '';
    return suffix ? `SST${suffix}` : 'SST';
  }

  return raw;
}

function buildHdbiLayoutCard(toolId, rows, area, sub, layoutKey = '') {
  // When a tool is disabled, treat all cells as Empty
  const toolDisabled = isToolDisabled(toolId);
  if (toolDisabled) rows = [];

  const card = document.createElement('div');
  card.className = `machine-card area-${area || 'HDBI'}`;
  card.dataset.tool = toolId;

  // Header
  const header = document.createElement('div');
  header.className = 'machine-card-header';

  const titleSpan = document.createElement('span');
  titleSpan.className = 'machine-card-title-center';
  titleSpan.textContent = formatMachineDisplayName(toolId);
  header.appendChild(titleSpan);
  if (toolDisabled) {
    const badge = document.createElement('span');
    badge.className = 'tool-disabled-badge';
    badge.textContent = 'BAGGED';
    header.appendChild(badge);
  }
  card.appendChild(header);

  // 2D Layout Table
  const cc = colMapping.cellCol;
  const pc = colMapping.productCol;

  const key = layoutKey || getHdbiLayoutKey(toolId) || toolId;
  card.dataset.layoutKey = key.replace(/\s+/g, '-');
  const layout = HDBI_LAYOUTS[key];
  if (!layout) {
    const fallback = document.createElement('div');
    fallback.style.cssText = 'padding:8px 10px;font-size:11px;color:#9ca3af;';
    fallback.textContent = 'No layout defined for ' + toolId;
    card.appendChild(fallback);
    return card;
  }

  // Build a map: cell → row for fast lookup
  const cellMap = new Map();
  for (const row of rows) {
    const cellId = String(row[cc] || '').trim().toUpperCase();
    if (cellId) cellMap.set(cellId, row);
  }

  // Render 2D table
  const table = document.createElement('table');
  table.className = 'hdbi-layout-table';

  // Header row (columns)
  const thead = document.createElement('thead');
  const hrow = document.createElement('tr');
  hrow.innerHTML = '<th></th>'; // Top-left corner cell
  for (const col of layout.columns) {
    const th = document.createElement('th');
    th.textContent = col || '';
    hrow.appendChild(th);
  }
  thead.appendChild(hrow);
  table.appendChild(thead);

  // Body rows (cells/products)
  const tbody = document.createElement('tbody');
  for (const rowId of layout.rows) {
    const tr = document.createElement('tr');
    const rowLbl = document.createElement('th');
    rowLbl.textContent = rowId;
    rowLbl.className = 'hdbi-row-label';
    tr.appendChild(rowLbl);

    for (const col of layout.columns) {
      const td = document.createElement('td');
      td.className = 'hdbi-cell';

      const cellKey = col ? `${col}${rowId}` : rowId;
      const dataRow = cellMap.get(cellKey.toUpperCase());
      const product = normalizeProductForCard(toolId, dataRow ? dataRow[pc] : '');

      td.innerHTML = `<div class="cell-slot ${productClass(product)}" data-tool="${toolId}" data-cell="${cellKey}" data-product="${product}">${product}</div>`;
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  card.appendChild(table);

  // Click detail modal (on table cells)
  card.addEventListener('click', (e) => {
    if (e.target.closest('.cell-slot')) showDetailModal(toolId, rows);
  });

  return card;
}

function buildMachineCard(toolId, rows, area, sub) {
  // When a tool is disabled, treat all cells as Empty
  const toolDisabled = isToolDisabled(toolId);
  if (toolDisabled) rows = [];

  // Use 2D layout for HDBI/BDC machines
  const hdbiLayoutKey = getHdbiLayoutKey(toolId);
  if (hdbiLayoutKey) {
    return buildHdbiLayoutCard(toolId, rows, area, sub, hdbiLayoutKey);
  }

  // Use matrix card design for PPV/HST tools (A-D columns, 601..101 rows)
  if (String(area || '').toUpperCase() === 'PPV' && String(sub || '').toUpperCase() === 'HST') {
    return buildHdbiLayoutCard(toolId, rows, area, sub, 'PPV HST');
  }

  // Use matrix card design for PPV/SST tools (A-E columns, 401..101 rows)
  if (String(area || '').toUpperCase() === 'PPV' && String(sub || '').toUpperCase() === 'SST') {
    return buildHdbiLayoutCard(toolId, rows, area, sub, 'PPV SST');
  }

  const card = document.createElement('div');
  card.className = `machine-card area-${area || 'HDMX'}`;
  card.dataset.tool = toolId;

  // Header
  const header = document.createElement('div');
  header.className = 'machine-card-header';

  if (String(area || '').toUpperCase() === 'HDMX') {
    // HDMX layout: [coolant badge left] [tool name truly centered] [spacer right]
    header.classList.add('machine-card-header-hdmx');

    const coolant = getHdmxCoolant(toolId);
    const leftBadge = document.createElement('span');
    leftBadge.className = 'badge-sub badge-coolant';
    leftBadge.textContent = coolant || '';

    const titleSpan = document.createElement('span');
    titleSpan.className = 'machine-card-title-center';
    titleSpan.textContent = formatMachineDisplayName(toolId);

    header.appendChild(leftBadge);
    header.appendChild(titleSpan);
    if (toolDisabled) {
      const badge = document.createElement('span');
      badge.className = 'tool-disabled-badge';
      badge.textContent = 'BAGGED';
      header.appendChild(badge);
    }
  } else {
    header.innerHTML = `<span class="machine-card-title-center">${formatMachineDisplayName(toolId)}</span>`;
    if (toolDisabled) {
      const badge = document.createElement('span');
      badge.className = 'tool-disabled-badge';
      badge.textContent = 'BAGGED';
      header.appendChild(badge);
    }
  }
  card.appendChild(header);

  // Cells
  const cellList = document.createElement('div');
  cellList.className = 'cell-list';

  // Determine ordered cells: prefer L05→L01 for HDMX, otherwise ascending
  const cc = colMapping.cellCol;
  const pc = colMapping.productCol;

  let cellRows = [...rows];

  const fixedCells = fixedCellsForTool(area, toolId);
  if (fixedCells) {
    const byCell = new Map();
    for (const row of rows) {
      const key = String(row[cc] || '').trim().toUpperCase();
      if (key) byCell.set(key, row);
    }
    cellRows = fixedCells.map((cell) => byCell.get(cell) || { [cc]: cell, [pc]: 'Empty' });
  }

  if (!cellRows.length) {
    // Show placeholder for machines with no data
    const emptyRow = document.createElement('div');
    emptyRow.style.cssText = 'padding:8px 10px;font-size:11px;color:#9ca3af;';
    emptyRow.textContent = 'No allocation data';
    cellList.appendChild(emptyRow);
  } else {
    const isHdmx = String(area || '').toUpperCase() === 'HDMX';
    if (isHdmx) {
      renderHdmxPairedRows(cellList, cellRows, toolId);
    } else {
      // Sort: numeric suffix descending (L05→L01), else alphabetical
      cellRows.sort((a, b) => {
        const na = parseInt((a[cc] || '').replace(/\D/g, ''), 10);
        const nb = parseInt((b[cc] || '').replace(/\D/g, ''), 10);
        if (!isNaN(na) && !isNaN(nb)) return nb - na;
        return String(a[cc]).localeCompare(String(b[cc]));
      });

      for (const row of cellRows) {
        cellList.appendChild(buildCellRow(row, toolId));
      }
    }
  }

  card.appendChild(cellList);

  // Click → detail modal
  card.addEventListener('click', (e) => {
    if (!e.target.closest('.cell-slot')) showDetailModal(toolId, rows);
  });

  return card;
}

function renderHdmxPairedRows(cellList, cellRows, toolId) {
  const cc = colMapping.cellCol;
  const pc = colMapping.productCol;

  // Group cells by base key removing the last digit: A501/A502 => A50
  const byPair = new Map();
  for (const row of cellRows) {
    const cellId = String(row[cc] || '').trim().toUpperCase();
    const pairKey = cellId.length > 1 ? cellId.slice(0, -1) : cellId;
    if (!byPair.has(pairKey)) byPair.set(pairKey, []);
    byPair.get(pairKey).push(row);
  }

  const sortedPairKeys = [...byPair.keys()].sort((a, b) => {
    const aNum = parseInt(a.replace(/\D/g, ''), 10);
    const bNum = parseInt(b.replace(/\D/g, ''), 10);
    if (!isNaN(aNum) && !isNaN(bNum) && aNum !== bNum) return bNum - aNum;
    return String(a).localeCompare(String(b));
  });

  for (const pairKey of sortedPairKeys) {
    const pairRows = [...(byPair.get(pairKey) || [])].sort((a, b) => {
      const aCell = String(a[cc] || '').toUpperCase();
      const bCell = String(b[cc] || '').toUpperCase();
      const aLast = parseInt(aCell.slice(-1), 10);
      const bLast = parseInt(bCell.slice(-1), 10);
      if (!isNaN(aLast) && !isNaN(bLast) && aLast !== bLast) return aLast - bLast;
      return aCell.localeCompare(bCell);
    });

    const pairRow = document.createElement('div');
    pairRow.className = 'cell-row-pair';

    // Ensure two slots for visual horizontal pairing (e.g. A501 next to A502)
    const slots = pairRows.slice(0, 2);
    while (slots.length < 2) {
      const missingCell = `${pairKey}${slots.length + 1}`;
      slots.push({ [cc]: missingCell, [pc]: 'Empty' });
    }

    for (const row of slots) {
      pairRow.appendChild(buildCellRow(row, toolId));
    }
    cellList.appendChild(pairRow);
  }
}

function buildCellRow(row, toolId) {
  const cc = colMapping.cellCol;
  const pc = colMapping.productCol;
  const oc = colMapping.tosCol;

  const cellId  = row[cc]  || '?';
  const product = normalizeProductForCard(toolId, row[pc]);
  const tos     = oc ? (row[oc] || '') : '';

  const key = `${toolId}||${cellId}`;
  const isHighlighted = highlightSet.has(key);

  const rowEl  = document.createElement('div');
  rowEl.className = 'cell-row';

  const idEl   = document.createElement('div');
  idEl.className = 'cell-id';
  idEl.textContent = cellId;

  const slotEl = document.createElement('div');
  slotEl.className = `cell-slot ${productClass(product)}${isHighlighted ? ' cell-highlight' : ''}`;
  slotEl.textContent = product;
  slotEl.dataset.tool    = toolId;
  slotEl.dataset.toolLabel = formatMachineDisplayName(toolId);
  slotEl.dataset.cell    = cellId;
  slotEl.dataset.product = product;
  slotEl.dataset.tos     = tos;

  slotEl.addEventListener('mouseenter', showTooltip);
  slotEl.addEventListener('mousemove',  moveTooltip);
  slotEl.addEventListener('mouseleave', hideTooltip);

  rowEl.appendChild(idEl);
  rowEl.appendChild(slotEl);
  return rowEl;
}

// ─── Tooltip ──────────────────────────────────────────────────────────────────
function showTooltip(e) {
  const d = e.currentTarget.dataset;
  const lines = [
    `Tool:    ${d.toolLabel || d.tool}`,
    `Cell:    ${d.cell}`,
    `Product: ${d.product || 'Empty'}`,
  ];
  if (d.tos) lines.push(`TOS:     ${d.tos}`);
  tooltip.textContent = lines.join('\n');
  tooltip.style.display = 'block';
  moveTooltip(e);
}
function moveTooltip(e) {
  const x = e.clientX + 14;
  const y = e.clientY + 14;
  tooltip.style.left = `${Math.min(x, window.innerWidth - 280)}px`;
  tooltip.style.top  = `${Math.min(y, window.innerHeight - 120)}px`;
}
function hideTooltip() { tooltip.style.display = 'none'; }

// ─── Detail modal ─────────────────────────────────────────────────────────────
function showDetailModal(toolId, rows) {
  detailTitle.textContent = `Machine: ${formatMachineDisplayName(toolId)}`;
  const cc = colMapping.cellCol;
  const pc = colMapping.productCol;
  const oc = colMapping.tosCol;
  const wc = colMapping.wwCol;
  const dc = colMapping.dayCol;

  let html = `<table class="detail-cell-table"><thead><tr>
    <th>Cell</th><th>Product</th>${oc ? '<th>TOS</th>' : ''}
    <th>WW</th><th>Day</th></tr></thead><tbody>`;

  let rowsToShow = rows;
  if (!rowsToShow.length) {
    rowsToShow = [{ [cc]: '--', [pc]: 'Empty', [wc]: '--', [dc]: '--', ...(oc ? { [oc]: '--' } : {}) }];
  }
  const fixedCells = fixedCellsForTool('', toolId);
  if (fixedCells) {
    const byCell = new Map();
    for (const row of rows) {
      const key = String(row[cc] || '').trim().toUpperCase();
      if (key) byCell.set(key, row);
    }
    rowsToShow = fixedCells.map((cell) => byCell.get(cell) || { [cc]: cell, [pc]: 'Empty', [wc]: '--', [dc]: '--', ...(oc ? { [oc]: '--' } : {}) });
  }

  for (const row of rowsToShow) {
    html += `<tr>
      <td>${row[cc] || '—'}</td>
      <td><strong>${row[pc] == null || row[pc] === '' ? 'Empty' : row[pc]}</strong></td>
      ${oc ? `<td>${row[oc] || '—'}</td>` : ''}
      <td>${row[wc] || '—'}</td>
      <td>${row[dc] || '—'}</td>
    </tr>`;
  }
  html += '</tbody></table>';

  detailBody.innerHTML = html;
  detailModal.style.display = 'flex';
}

closeDetailBtn.addEventListener('click', () => { detailModal.style.display = 'none'; });
detailModal.addEventListener('click', (e) => { if (e.target === detailModal) detailModal.style.display = 'none'; });

// ─── HDMX layout selector ────────────────────────────────────────────────────
for (const btn of document.querySelectorAll('.hdmx-cols-btn')) {
  btn.addEventListener('click', async () => {
    document.querySelectorAll('.hdmx-cols-btn').forEach(b => b.classList.remove('is-active'));
    btn.classList.add('is-active');
    // Re-apply data-hdmx-cols to all HDMX section-cards already in the DOM
    document.querySelectorAll('.area-bg-HDMX + .section-cards').forEach(grid => {
      grid.dataset.hdmxCols = btn.dataset.cols;
    });
  });
}

// ─── PPV layout selector ─────────────────────────────────────────────────────
for (const btn of document.querySelectorAll('.ppv-cols-btn')) {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.ppv-cols-btn').forEach(b => b.classList.remove('is-active'));
    btn.classList.add('is-active');
    document.querySelectorAll('.area-bg-PPV + .section-cards').forEach(grid => {
      grid.dataset.ppvCols = btn.dataset.cols;
    });
  });
}

// ─── Product search ───────────────────────────────────────────────────────────
productSearchBtn.addEventListener('click', doProductSearch);
productSearch.addEventListener('keydown', (e) => { if (e.key === 'Enter') doProductSearch(); });

async function doProductSearch() {
  if (!isMappingComplete(colMapping)) { openSettings(); return; }

  const ready = await ensureBaseDataLoaded(false);
  if (!ready) return;

  await fetchAndRender();
}

function rerenderHighlights() {
  // Re-apply highlight class without full re-render for performance
  document.querySelectorAll('.cell-slot').forEach(el => {
    const key = `${el.dataset.tool}||${el.dataset.cell}`;
    el.classList.toggle('cell-highlight', highlightSet.has(key));
  });
}

// ─── Clear ────────────────────────────────────────────────────────────────────
clearBtn.addEventListener('click', async () => {
  filterArea.value = '';
  setActiveAreaTab('');
  filterSub.innerHTML = '<option value="">All</option>';
  subAreaSection.style.display = 'none';
  if (coolantSection) coolantSection.style.display = 'none';
  if (filterCoolant) filterCoolant.value = '';
  if (hdmxBaggedSection) hdmxBaggedSection.style.display = 'none';
  if (filterBagged) filterBagged.value = '';
  if (hdmxLayoutSection) hdmxLayoutSection.style.display = 'none';
  if (ppvLayoutSection)  ppvLayoutSection.style.display  = 'none';
  productSearch.value = '';
  areaBadge.textContent = '—';
  populateToolList('', '');
  highlightSet.clear();
  currentData = [];
  // Re-render using cached base data for the active ww/day.
  await fetchAndRender();
});

// ─── Export ───────────────────────────────────────────────────────────────────
exportBtn.addEventListener('click', () => {
  if (!isMappingComplete(colMapping)) { openSettings(); return; }
  const tools = selectedTools();
  const ww    = filterWW.value;
  const day   = filterDay.value;

  const p = mappingParams();
  if (ww)   p.set('ww', ww);
  if (day)  p.set('day', day);
  for (const t of tools) p.append('tools', t);

  // Trigger download via hidden link
  const a = document.createElement('a');
  a.href = `/api/export?${p}`;
  a.download = '';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
});

// ─── Settings / Column Mapping modal ─────────────────────────────────────────
settingsBtn.addEventListener('click', openSettings);
cancelMappingBtn.addEventListener('click', () => { settingsModal.style.display = 'none'; });
settingsModal.addEventListener('click', (e) => { if (e.target === settingsModal) settingsModal.style.display = 'none'; });

function openSettings() {
  settingsModal.style.display = 'flex';
  buildMappingGrid();
}

function buildMappingGrid() {
  mappingGrid.innerHTML = '';
  for (const key of MAPPING_KEYS) {
    const row = document.createElement('div');
    row.className = 'mapping-row';

    const lbl = document.createElement('label');
    lbl.textContent = MAPPING_LABELS[key];
    lbl.htmlFor = `map_${key}`;

    const sel = document.createElement('select');
    sel.id = `map_${key}`;
    sel.dataset.key = key;
    sel.innerHTML = '<option value="">(not mapped)</option>';

    // Populate with discovered columns
    for (const col of discoveredCols) {
      const opt = document.createElement('option');
      opt.value = opt.textContent = col;
      if (colMapping[key] === col) opt.selected = true;
      sel.appendChild(opt);
    }
    // If current mapping not in discovered list, still show it
    if (colMapping[key] && !discoveredCols.includes(colMapping[key])) {
      const opt = document.createElement('option');
      opt.value = opt.textContent = colMapping[key];
      opt.selected = true;
      sel.appendChild(opt);
    }

    row.appendChild(lbl);
    row.appendChild(sel);
    mappingGrid.appendChild(row);
  }
}

discoverBtn.addEventListener('click', async () => {
  discoverStatus.innerHTML = '<span class="spinner"></span> Querying schema…';
  try {
    const { columns } = await apiFetch('/api/schema');
    discoveredCols = columns;
    discoverStatus.textContent = `Found ${columns.length} columns: ${columns.join(', ')}`;
    buildMappingGrid();
  } catch (err) {
    discoverStatus.textContent = `⚠ ${err.message}`;
  }
});

saveMappingBtn.addEventListener('click', async () => {
  const newMap = {};
  for (const sel of mappingGrid.querySelectorAll('select')) {
    newMap[sel.dataset.key] = sel.value;
  }
  colMapping = newMap;
  saveMapping(colMapping);
  settingsModal.style.display = 'none';

  if (isMappingComplete(colMapping)) {
    await loadFilters();
    baseDataCache = [];
    baseDataKey = '';
    await fetchAndRender();
    setStatus('Column mapping saved. Data loaded.');
  } else {
    setStatus('⚠ Mapping incomplete — some required columns are not mapped.');
  }
});

// ─── Utility ──────────────────────────────────────────────────────────────────
function setStatus(html) {
  statusMsg.innerHTML = html;
  const bar = statusMsg.closest('.status-bar');
  if (bar) bar.style.display = html ? '' : 'none';
}

// ─── Recent Changes modal ─────────────────────────────────────────────────────
const DAY_NAMES_CHANGES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
let _changesData = [];   // last loaded changes array (for export/filter)

function formatChangePeriod(ww, day) {
  const dayName = DAY_NAMES_CHANGES[Number(day)] || `Day ${day}`;
  return `${formatWWLabel(ww)} · ${dayName}`;
}

function renderChangesTable(items) {
  if (!items.length) return '<p class="changes-empty">No changes found for the current filter.</p>';
  let html = `<table class="changes-table">
    <thead><tr>
      <th>Machine</th><th>Cell</th><th>Previous Product</th><th>New Product</th><th>Type</th>
    </tr></thead><tbody>`;
  for (const ch of items) {
    const oldProd = stripMpePrefix(ch.oldProduct || '');
    const newProd = stripMpePrefix(ch.newProduct || '');
    const type = !ch.oldProduct ? 'Added' : !ch.newProduct ? 'Removed' : 'Changed';
    const cls  = type === 'Added' ? 'change-added' : type === 'Removed' ? 'change-removed' : 'change-modified';
    html += `<tr>
      <td>${formatMachineDisplayName(ch.machine) || '—'}</td>
      <td>${ch.cell || '—'}</td>
      <td class="change-old">${oldProd || '<em class="change-empty">Empty</em>'}</td>
      <td class="change-new">${newProd || '<em class="change-empty">Empty</em>'}</td>
      <td><span class="change-badge ${cls}">${type}</span></td>
    </tr>`;
  }
  html += '</tbody></table>';
  return html;
}

function renderChangesBody(filterText = '') {
  if (!_changesData.length) {
    changesBody.innerHTML = '<p class="changes-empty">No allocation changes detected in the last 2 days.</p>';
    return;
  }

  const q = filterText.trim().toLowerCase();

  // Group by comparison period
  const groups = new Map();
  for (const ch of _changesData) {
    const gKey = `${ch.fromWW}|${ch.fromDay}→${ch.toWW}|${ch.toDay}`;
    if (!groups.has(gKey)) groups.set(gKey, { from: { ww: ch.fromWW, day: ch.fromDay }, to: { ww: ch.toWW, day: ch.toDay }, items: [] });
    groups.get(gKey).items.push(ch);
  }

  let html = '';
  for (const [, group] of groups) {
    const items = q
      ? group.items.filter(c =>
          (c.machine    || '').toLowerCase().includes(q) ||
          (c.cell       || '').toLowerCase().includes(q) ||
          (c.oldProduct || '').toLowerCase().includes(q) ||
          (c.newProduct || '').toLowerCase().includes(q)
        )
      : group.items;
    if (!items.length) continue;

    const fromLabel = formatChangePeriod(group.from.ww, group.from.day);
    const toLabel   = formatChangePeriod(group.to.ww,   group.to.day);
    html += `<div class="changes-group">
      <div class="changes-group-header">
        <span class="changes-period-from">${fromLabel}</span>
        <span class="changes-arrow">&#8594;</span>
        <span class="changes-period-to">${toLabel}</span>
        <span class="changes-count">${items.length} change${items.length !== 1 ? 's' : ''}</span>
      </div>
      ${renderChangesTable(items)}
    </div>`;
  }

  changesBody.innerHTML = html || '<p class="changes-empty">No changes match the current filter.</p>';
}

async function openChanges() {
  if (!isMappingComplete(colMapping)) { openSettings(); return; }
  changesModal.style.display = 'flex';
  changesSearch.value = '';
  changesBody.innerHTML = '<p class="changes-loading">Loading changes\u2026</p>';
  _changesData = [];

  try {
    const params = mappingParams();
    const data   = await apiFetch(`/api/changes?${params}`);
    _changesData = data.changes || [];
    renderChangesBody('');
  } catch (err) {
    changesBody.innerHTML = `<p class="changes-error">Error loading changes: ${err.message}</p>`;
  }
}

changesBtn.addEventListener('click', openChanges);
closeChangesBtn.addEventListener('click', () => { changesModal.style.display = 'none'; });
disableToolsBtn.addEventListener('click', openDisableToolsModal);
changesModal.addEventListener('click', (e) => { if (e.target === changesModal) changesModal.style.display = 'none'; });

changesSearch.addEventListener('input', () => { renderChangesBody(changesSearch.value); });

changesExportBtn.addEventListener('click', () => {
  if (!_changesData.length) return;
  const q = changesSearch.value.trim().toLowerCase();
  const rows = q
    ? _changesData.filter(c =>
        (c.machine    || '').toLowerCase().includes(q) ||
        (c.cell       || '').toLowerCase().includes(q) ||
        (c.oldProduct || '').toLowerCase().includes(q) ||
        (c.newProduct || '').toLowerCase().includes(q)
      )
    : _changesData;
  const headers = ['Machine', 'Cell', 'Old Product', 'New Product', 'Type', 'From WW', 'From Day', 'To WW', 'To Day'];
  const csvLines = [headers.join(',')];
  for (const ch of rows) {
    const type = !ch.oldProduct ? 'Added' : !ch.newProduct ? 'Removed' : 'Changed';
    const dayNames = DAY_NAMES_CHANGES;
    csvLines.push([
      ch.machine    || '',
      ch.cell       || '',
      ch.oldProduct || '',
      ch.newProduct || '',
      type,
      ch.fromWW  || '',
      dayNames[Number(ch.fromDay)] || ch.fromDay || '',
      ch.toWW    || '',
      dayNames[Number(ch.toDay)]   || ch.toDay   || '',
    ].map(v => `"${String(v).replace(/"/g, '""')}"`).join(','));
  }
  const blob = new Blob([csvLines.join('\r\n')], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = 'allocation_changes.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

// ─── Planned Conversions ──────────────────────────────────────────────────────
const plannedBtn       = document.getElementById('plannedBtn');
const plannedModal     = document.getElementById('plannedModal');
const closePlannedBtn  = document.getElementById('closePlannedBtn');
const plannedForm      = document.getElementById('plannedForm');
const pArea            = document.getElementById('pArea');
const pSubField        = document.getElementById('pSubField');
const pSub             = document.getElementById('pSub');
const pTool            = document.getElementById('pTool');
const pCellList        = document.getElementById('pCellList');
const pProduct         = document.getElementById('pProduct');
const pDate            = document.getElementById('pDate');
const pNotes           = document.getElementById('pNotes');
const plannedFormError = document.getElementById('plannedFormError');
const plannedTableBody = document.getElementById('plannedTableBody');
const plannedExportBtn = document.getElementById('plannedExportBtn');

// Populate the product datalist with unique products from baseDataCache for the given area/sub.
function populatePlannedProductList(area, sub) {
  const datalist = document.getElementById('pProductList');
  if (!datalist) return;
  datalist.innerHTML = '';

  const tc = colMapping.toolCol;
  const pc = colMapping.productCol;
  if (!tc || !pc || !area) return;

  const products = new Set();
  for (const row of baseDataCache) {
    const toolId = row[tc] || '';
    const loc = resolveArea(toolId);
    if (loc.area !== area) continue;
    if (sub && loc.sub !== sub) continue;
    const p = normalizeProductForCard(toolId, row[pc]);
    if (p && p !== 'Empty' && !/^tlo$/i.test(p)) products.add(p);
  }

  const sorted = [...products].sort((a, b) => a.localeCompare(b));
  for (const p of sorted) {
    const opt = document.createElement('option');
    opt.value = p;
    datalist.appendChild(opt);
  }
}

// All cells by tool – derived from the same logic used by the dashboard cards.
function getCellsForPlannedTool(area, toolId) {
  const cells = fixedCellsForTool(area, toolId);
  return cells && cells.length ? cells : null;
}

// Populate sub-area dropdown when area changes inside the planned form.
pArea.addEventListener('change', () => {
  const area = pArea.value;

  // Sub-area
  pSub.innerHTML = '<option value="">Select…</option>';
  const areaData = AREA_MAP[area];
  if (area && areaData && areaData.subs && areaData.subs.length) {
    pSubField.style.display = '';
    for (const s of areaData.subs) {
      const opt = document.createElement('option');
      opt.value = opt.textContent = s;
      pSub.appendChild(opt);
    }
  } else {
    pSubField.style.display = 'none';
  }

  populatePlannedTools(area, '');
  populatePlannedCells(area, '');
  populatePlannedProductList(area, '');
});

pSub.addEventListener('change', () => {
  populatePlannedTools(pArea.value, pSub.value);
  populatePlannedCells(pArea.value, '');
  populatePlannedProductList(pArea.value, pSub.value);
});

pTool.addEventListener('change', () => {
  populatePlannedCells(pArea.value, pTool.value);
});

function populatePlannedTools(area, sub) {
  pTool.innerHTML = '';
  if (!area) {
    pTool.innerHTML = '<option value="">Select area first</option>';
    return;
  }
  const areaData = AREA_MAP[area];
  if (!areaData) { pTool.innerHTML = '<option value="">No tools</option>'; return; }

  let tools = [];
  if (sub && areaData.tools[sub]) {
    tools = areaData.tools[sub];
  } else {
    tools = Object.values(areaData.tools).flat();
    // de-duplicate
    tools = [...new Set(tools)];
  }

  if (!tools.length) {
    pTool.innerHTML = '<option value="">No tools defined</option>';
    return;
  }

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = 'Select tool…';
  pTool.appendChild(placeholder);

  for (const toolId of tools.sort((a, b) => a.localeCompare(b))) {
    const opt = document.createElement('option');
    opt.value = toolId;
    opt.textContent = toolId;
    pTool.appendChild(opt);
  }
}

function populatePlannedCells(area, toolId) {
  pCellList.innerHTML = '';
  delete pCellList.dataset.mode;

  if (!toolId) {
    pCellList.innerHTML = '<span class="pcell-placeholder">Select tool first</span>';
    return;
  }

  // ── HDMX: show level selectors (L1–L5) instead of individual cells ────────
  if (String(area || '').toUpperCase() === 'HDMX') {
    pCellList.dataset.mode = 'hdmx-levels';
    renderCheckboxList(HDMX_LEVELS.map(l => ({
      value: l.label,
      label: l.label,
    })));
    return;
  }

  const cells = getCellsForPlannedTool(area, toolId);
  if (!cells) {
    // No fixed layout — show a free-text input inside the list box
    const input = document.createElement('input');
    input.type = 'text';
    input.id = 'pCellFree';
    input.className = 'pcell-free-input';
    input.placeholder = 'e.g. LA101';
    input.maxLength = 32;
    pCellList.appendChild(input);
    return;
  }

  renderCheckboxList(cells.map(c => ({ value: c, label: c })));
}

// Renders the "Select All" toggle + one checkbox per item into pCellList.
// items: Array<{ value: string, label: string }>
function renderCheckboxList(items) {
  // ── Select All toggle row ─────────────────────────────────────────────────
  const toggleRow = document.createElement('label');
  toggleRow.className = 'pcell-toggle-row';
  const toggleCb = document.createElement('input');
  toggleCb.type = 'checkbox';
  toggleCb.id = 'pCellSelectAll';
  toggleRow.appendChild(toggleCb);
  toggleRow.appendChild(document.createTextNode('Select All'));
  pCellList.appendChild(toggleRow);

  function syncToggle() {
    const all     = pCellList.querySelectorAll('.pcell-item input[type="checkbox"]');
    const checked = pCellList.querySelectorAll('.pcell-item input[type="checkbox"]:checked');
    toggleCb.checked       = all.length > 0 && checked.length === all.length;
    toggleCb.indeterminate = checked.length > 0 && checked.length < all.length;
  }

  toggleCb.addEventListener('change', () => {
    pCellList.querySelectorAll('.pcell-item input[type="checkbox"]').forEach(cb => {
      cb.checked = toggleCb.checked;
    });
  });

  // ── One checkbox per item ─────────────────────────────────────────────────
  for (const item of items) {
    const lbl = document.createElement('label');
    lbl.className = 'pcell-item';
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.value = item.value;
    cb.addEventListener('change', syncToggle);
    lbl.appendChild(cb);
    lbl.appendChild(document.createTextNode(item.label));
    pCellList.appendChild(lbl);
  }
}

// Helper: returns selected cell values from pCellList.
// For HDMX level mode, each level is expanded to its constituent cells.
function getSelectedCells() {
  // Free-text mode
  const free = pCellList.querySelector('.pcell-free-input');
  if (free) return free.value.trim() ? [free.value.trim()] : [];

  const checked = [...pCellList.querySelectorAll('.pcell-item input[type="checkbox"]:checked')]
    .map(cb => cb.value);

  // HDMX level mode: expand L1 → [A101, A102], etc.
  if (pCellList.dataset.mode === 'hdmx-levels') {
    const expanded = [];
    for (const levelLabel of checked) {
      const level = HDMX_LEVELS.find(l => l.label === levelLabel);
      if (level) expanded.push(...level.cells);
    }
    return expanded;
  }

  return checked;
}

function showPlannedFormError(msg) {
  plannedFormError.textContent = msg;
  plannedFormError.style.display = msg ? '' : 'none';
}

async function loadPlannedConversions() {
  plannedTableBody.innerHTML = '<tr><td colspan="8" class="planned-empty">Loading…</td></tr>';
  try {
    const list = await apiFetch('/api/planned-conversions');
    renderPlannedTable(list);
  } catch (err) {
    plannedTableBody.innerHTML = `<tr><td colspan="8" class="planned-empty planned-error">Error: ${err.message}</td></tr>`;
  }
}

function renderPlannedTable(list) {
  if (!list.length) {
    plannedTableBody.innerHTML = '<tr><td colspan="8" class="planned-empty">No planned conversions yet.</td></tr>';
    return;
  }

  // Sort by date ascending so soonest changes appear first
  const sorted = [...list].sort((a, b) => a.date.localeCompare(b.date));

  let html = '';
  for (const entry of sorted) {
    const created = entry.createdAt
      ? new Date(entry.createdAt).toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' })
      : '—';
    const isPast = entry.date && entry.date < new Date().toISOString().slice(0, 10);
    html += `<tr class="${isPast ? 'planned-row-past' : ''}">
      <td>${escHtml(entry.area || '—')}</td>
      <td>${escHtml(entry.tool || '—')}</td>
      <td>${escHtml(entry.cell || '—')}</td>
      <td class="planned-product">${escHtml(entry.newProduct || '—')}</td>
      <td class="planned-date">${escHtml(isoDateToWwLabel(entry.date))}</td>
      <td>${escHtml(entry.notes || '')}</td>
      <td class="planned-created">${created}</td>
      <td><button class="btn btn-sm btn-delete-planned" data-id="${escHtml(entry.id)}" title="Delete">&#10005;</button></td>
    </tr>`;
  }
  plannedTableBody.innerHTML = html;

  // Bind delete buttons
  plannedTableBody.querySelectorAll('.btn-delete-planned').forEach(btn => {
    btn.addEventListener('click', async () => {
      if (!confirm('Delete this planned conversion?')) return;
      try {
        const res = await fetch(`/api/planned-conversions/${encodeURIComponent(btn.dataset.id)}`, { method: 'DELETE' });
        if (!res.ok) {
          const err = await res.json().catch(() => ({ error: res.statusText }));
          alert(`Error: ${err.error || res.statusText}`);
          return;
        }
        await loadPlannedConversions();
      } catch (err) {
        alert(`Error: ${err.message}`);
      }
    });
  });
}

function escHtml(str) {
  return String(str ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// Converts a stored ISO date (e.g. "2026-05-14") to Intel WW label (e.g. "WW20.5")
// Uses forecastSlots (next WW) and todayPeriod (current WW) for the lookup.
function isoDateToWwLabel(isoDate) {
  if (!isoDate) return '—';
  // Check next-WW slots
  if (forecastSlots.length) {
    const slot = forecastSlots.find(s => s.date === isoDate);
    if (slot) {
      const wwNum = Number(String(slot.ww).replace(/\D/g, '')) % 100;
      return `WW${wwNum}.${slot.day}`;
    }
  }
  // Check current WW by scanning Sun–Sat of the current week
  if (todayPeriod) {
    const todayWwNum = Number(String(todayPeriod.ww).replace(/\D/g, '')) % 100;
    const cursor = new Date();
    cursor.setDate(cursor.getDate() - cursor.getDay()); // rewind to Sunday
    for (let d = 0; d < 7; d++) {
      const iso = cursor.toISOString().slice(0, 10);
      if (iso === isoDate) return `WW${todayWwNum}.${cursor.getDay()}`;
      cursor.setDate(cursor.getDate() + 1);
    }
  }
  return isoDate; // fallback: raw ISO date
}

plannedForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  showPlannedFormError('');

  const area       = pArea.value.trim();
  const tool       = pTool.value.trim();
  const cells      = getSelectedCells();
  const newProduct = pProduct.value.trim();
  const date       = pDate.value.trim();
  const notes      = pNotes.value.trim();

  if (!area || !tool || !cells.length || !newProduct || !date) {
    showPlannedFormError('Please fill in all required fields (*) and select at least one cell.');
    return;
  }

  try {
    // POST one entry per selected cell
    for (const cell of cells) {
      const res = await fetch('/api/planned-conversions', {
        method : 'POST',
        headers: { 'Content-Type': 'application/json' },
        body   : JSON.stringify({ area, tool, cell, newProduct, date, notes }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        showPlannedFormError(`Cell ${cell}: ${err.error || res.statusText}`);
        return;
      }
    }
    // Reset form (keep area/sub/tool selection for quick successive adds)
    pProduct.value = '';
    pDate.value    = '';
    pNotes.value   = '';
    // Uncheck all cells
    pCellList.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.checked = false; cb.indeterminate = false; });
    await loadPlannedConversions();
  } catch (err) {
    showPlannedFormError(`Network error: ${err.message}`);
  }
});

// Export planned conversions as CSV
plannedExportBtn.addEventListener('click', async () => {
  try {
    const list = await apiFetch('/api/planned-conversions');
    if (!list.length) { alert('No planned conversions to export.'); return; }
    const headers = ['Area', 'Tool', 'Cell', 'New Product', 'Date', 'Notes', 'Registered'];
    const rows = list.map(e => [
      e.area, e.tool, e.cell, e.newProduct, e.date, e.notes || '',
      e.createdAt ? new Date(e.createdAt).toLocaleString('en-US') : '',
    ].map(v => `"${String(v ?? '').replace(/"/g,'""')}"`).join(','));
    const csv  = [headers.map(h => `"${h}"`).join(','), ...rows].join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = 'planned_conversions.csv';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (err) { alert(`Error: ${err.message}`); }
});

// Populate the Change Date dropdown with selectable Intel-calendar WW/day options:
// - Remaining days of the current WW that haven't passed yet (today excluded if before end of WW)
// - All 7 days of the next WW
// Each option shows "WWxx.d — Day Name (yyyy-mm-dd)" and its value is the ISO date.
function populatePlannedDates() {
  if (!todayPeriod || !forecastSlots.length) {
    pDate.innerHTML = '<option value="">Loading dates…</option>';
    return;
  }

  const todayDate   = new Date().toISOString().slice(0, 10);
  const todayWwFull = Number(String(todayPeriod.ww).replace(/\D/g, ''));  // e.g. 202619
  const todayWwNum  = todayWwFull % 100;                                   // e.g. 19
  const nextWwNum   = todayWwNum + 1;                                      // e.g. 20

  pDate.innerHTML = '<option value="">Select date…</option>';

  // ── Remaining days of the current WW (today included, past days excluded) ──
  const calCursor = new Date();
  calCursor.setDate(calCursor.getDate() - calCursor.getDay()); // rewind to Sunday
  for (let d = 0; d < 7; d++) {
    const iso = calCursor.toISOString().slice(0, 10);
    const dow = calCursor.getDay();
    if (iso >= todayDate) {
      const opt = document.createElement('option');
      opt.value = iso;
      opt.textContent = `WW${todayWwNum}.${dow}`;
      pDate.appendChild(opt);
    }
    calCursor.setDate(calCursor.getDate() + 1);
  }

  // ── All 7 days of the next WW (from forecastSlots) ──────────────────────────
  const nextWwSlots = forecastSlots.filter(s =>
    Number(String(s.ww).replace(/\D/g, '')) % 100 === nextWwNum
  );
  for (const slot of nextWwSlots) {
    const opt = document.createElement('option');
    opt.value = slot.date;
    opt.textContent = `WW${nextWwNum}.${slot.day}`;
    pDate.appendChild(opt);
  }
}

// Open/close modal
plannedBtn.addEventListener('click', () => {
  plannedModal.style.display = 'flex';
  populatePlannedDates();
  loadPlannedConversions();
});
closePlannedBtn.addEventListener('click', () => { plannedModal.style.display = 'none'; });
plannedModal.addEventListener('click', (e) => { if (e.target === plannedModal) plannedModal.style.display = 'none'; });
