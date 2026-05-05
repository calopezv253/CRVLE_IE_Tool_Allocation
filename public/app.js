/* ═══════════════════════════════════════════════════════════════════════════
   Tool Allocation Dashboard — app.js
   All DB interaction goes through the Node.js backend REST API.
   ═══════════════════════════════════════════════════════════════════════════ */

'use strict';

// ─── HDBI/BDC Layout definitions (2D physical layout) ────────────────────────
const HDBI_LAYOUTS = {
  'HDBI 4786': { columns: ['LA', 'RA'], rows: ['601', '501', '401', '301', '201', '101'] },
  'HDBI 5082': { columns: ['LA', 'LB', 'LC', 'RA', 'RB', 'RC'], rows: ['601', '501', '401', '301', '201', '101'] },
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
      HDBI: ['HDBI 5082', 'HDBI 4786'],
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
let colorIdx = 0;
const CLASS_CELLS = ['A101','A102','A201','A202','A301','A302','A401','A402','A501','A502'];
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
  'GMMLCC10T0210',
  'GMMLCC10T0287',
  'GMMLCC10T1336',
  'GMMLCC10T1567',
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

function normalizeHdmxProductForStats(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return 'Empty';

  const upper = raw.toUpperCase();
  if (['MPE_GNR-SP-LCC', 'MPE_GNR-SP-XCC', 'MPE_GNR-SP-HCC'].includes(upper)) {
    return 'MPE_GNR-SP';
  }
  if (HDMX_GMM_PRODUCTS.has(upper)) {
    return 'MPE_GMM';
  }

  return raw;
}

function isHdbiGmmTool(toolId) {
  const tag = String(formatMachineDisplayName(toolId) || '').replace(/\s+/g, '').toUpperCase();
  return tag === 'HDBI4786' || tag === 'HDBI5082';
}

function normalizeProductForCard(toolId, value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return 'Empty';

  const loc = resolveArea(toolId);
  const upper = raw.toUpperCase();

  if (loc.area === 'HDBI') {
    if (isHdbiGmmTool(toolId) && upper.startsWith('GMM')) return 'MPE_GMM';
    if (HDBI_GMM_PRODUCTS.has(upper)) return 'MPE_GMM';
  }

  if (loc.area === 'HDMX') {
    if (HDMX_GMM_PRODUCTS.has(upper)) return 'MPE_GMM';
    if (upper === 'MPE_NVL-S-28C') return 'MPE_NVL';
  }

  return raw;
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
  // Color assigned by buildProductColorMap() before render; fallback if called early
  if (!productColorMap[product]) {
    productColorMap[product] = PRODUCT_CLASSES[colorIdx % PRODUCT_CLASSES.length];
    colorIdx++;
  }
  return productColorMap[product];
}

// Pre-assign one unique color per product from sorted product list.
// Must be called before renderUsageStats / renderMachineGrid so both use
// the same stable, collision-free mapping.
function buildProductColorMap(data) {
  const pc = colMapping.productCol;
  const tc = colMapping.toolCol;
  if (!pc) return;

  // Gather all unique non-empty, non-idle, non-TLO product names
  const products = new Set();
  for (const row of data) {
    const p = String(row[pc] == null ? '' : row[pc]).trim();
    if (!p || /^(idle|empty)$/i.test(p) || /^tlo/i.test(p)) continue;
    // Use normalised name (same as what productClass receives)
    products.add(p);
  }

  // Sort alphabetically for deterministic ordering
  const sorted = [...products].sort((a, b) => a.localeCompare(b));

  // Reset map; re-assign sequentially — each product gets a unique slot
  Object.keys(productColorMap).forEach(k => { delete productColorMap[k]; });
  colorIdx = 0;
  for (const p of sorted) {
    productColorMap[p] = PRODUCT_CLASSES[colorIdx % PRODUCT_CLASSES.length];
    colorIdx++;
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
const statsGrid       = document.getElementById('statsGrid');
const changesBtn      = document.getElementById('changesBtn');
const changesModal    = document.getElementById('changesModal');
const changesBody     = document.getElementById('changesBody');
const closeChangesBtn = document.getElementById('closeChangesBtn');
const changesSearch   = document.getElementById('changesSearch');
const changesExportBtn= document.getElementById('changesExportBtn');

// ─── State ───────────────────────────────────────────────────────────────────
let colMapping    = loadMapping();
let discoveredCols= [];
let currentData   = [];   // raw rows from last /api/allocations call
let baseDataCache = [];   // raw rows for active ww/day (single DB load)
let baseDataKey   = '';   // cache key = "ww|day"
let highlightSet  = new Set();  // tool+cell keys to highlight

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

filterSub.addEventListener('change', () => {
  populateToolList(filterArea.value, filterSub.value);
});

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
    toolEntries.sort((a, b) => a.label.localeCompare(b.label));
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

  for (const entry of toolEntries) {
    const lbl = document.createElement('label');
    const cb  = document.createElement('input');
    cb.type = 'checkbox'; cb.value = entry.value; cb.checked = true;
    cb.dataset.toolItem = '1';
    lbl.appendChild(cb);
    lbl.append(` ${entry.label}`);
    toolCheckList.appendChild(lbl);

    cb.addEventListener('change', syncToolToggleAllState);
  }

  allCb.addEventListener('change', () => {
    const items = toolCheckList.querySelectorAll('input[type=checkbox][data-tool-item="1"]');
    for (const cb of items) cb.checked = allCb.checked;
    syncToolToggleAllState();
  });

  syncToolToggleAllState();
}

function selectedTools() {
  return [...toolCheckList.querySelectorAll('input[type=checkbox][data-tool-item="1"]:checked')]
    .map(cb => cb.value);
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
  if (!items.length) {
    allCb.checked = false;
    allCb.indeterminate = false;
    return;
  }

  const checked = items.filter(cb => cb.checked).length;
  allCb.checked = checked === items.length;
  allCb.indeterminate = checked > 0 && checked < items.length;
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
  const tc = colMapping.toolCol;
  const productQuery = productSearch.value.trim();
  const coolantQuery = String(filterCoolant?.value || '').toUpperCase();

  const ready = await ensureBaseDataLoaded(false);
  if (!ready) return;

  let visibleData = applyUiFilters(baseDataCache, area, sub, checkedTools, toolFilterActive);
  if (!area) {
    visibleData = visibleData.filter((row) => !isMachineHiddenInAll(row[tc] || ''));
  }

  if (area === 'HDMX' && coolantQuery) {
    visibleData = visibleData.filter((row) => getHdmxCoolant(row[tc]) === coolantQuery);
  }

  const productFilterResult = applyProductToolFilter(visibleData, productQuery);
  visibleData = productFilterResult.rows;

  highlightSet.clear();
  for (const key of productFilterResult.highlightKeys) highlightSet.add(key);

  currentData = visibleData;
  buildProductColorMap(visibleData);
  renderUsageStats(visibleData);
  renderMachineGrid(
    visibleData,
    toolFilterActive ? checkedTools : [],
    !productFilterResult.filtered,
    area === 'HDMX' ? coolantQuery : ''
  );
  recordCount.textContent = `${visibleData.length} rows`;
  if (productFilterResult.filtered) {
    setStatus(visibleData.length
      ? `Showing tools with product "${productQuery}" (${productFilterResult.matchCount} matching cell(s)).`
      : `No tools found with product "${productQuery}" for the selected filters.`);
  } else {
    setStatus(visibleData.length
      ? 'Data displayed from local cache.'
      : 'No data for the selected filters.');
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
  const p = String(value == null ? '' : value).trim();
  if (!p) return 'Empty';

  const upper = p.toUpperCase();

  if (String(area || '').toUpperCase() === 'HDBI') {
    if (isHdbiGmmTool(toolId) && upper.startsWith('GMM')) return 'MPE_GMM';
    if (HDBI_GMM_PRODUCTS.has(upper)) return 'MPE_GMM';
  }

  if (String(area || '').toUpperCase() === 'HDMX') {
    return normalizeHdmxProductForStats(p);
  }

  return p;
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

    for (const cellId of cellsToCount) {
      const product = byCell.get(cellId) || 'Empty';
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
  if (!hasProductFilter) {
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

  const scopedTools = allHdmxTools.filter((toolId) => {
    if (coolantFilter && getHdmxCoolant(toolId) !== coolantFilter) return false;
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

  const rows = ['HFE', 'EGDI']
    .map((coolant) => {
      const s = coolantStats[coolant] || { tools: 0, totalCells: 0, usedCells: 0, emptyCells: 0 };
      const utilization = s.totalCells > 0 ? ((s.usedCells / s.totalCells) * 100) : 0;
      return `<tr>
        <td>${coolant}</td>
        <td>${s.tools}</td>
        <td>${s.usedCells} cells</td>
        <td>${s.emptyCells} cells</td>
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
        <tr><th>Coolant</th><th>Tools</th><th>Used Cells</th><th>Empty Cells</th><th>Utilization</th></tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
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
  if (!statsGrid) return;

  const activeArea = filterArea.value;
  if (activeArea === 'PPV') {
    if (statsSection) statsSection.style.display = 'none';
    statsGrid.innerHTML = '';
    return;
  }

  if (!activeArea) {
    if (statsSection) statsSection.style.display = 'none';
    statsGrid.innerHTML = '';
    return;
  }

  const statsByBucket = buildStatsByBucket(data || []);
  const targets = getStatsTargets(activeArea);

  statsGrid.innerHTML = '';
  let lastCard = null;
  for (const target of targets) {
    const card = document.createElement('article');
    statsGrid.appendChild(card);
    const stats = statsByBucket.get(statsBucketKey(target.area, target.sub)) || createEmptyStats();
    renderAreaStatsCard(card, target.area, target.sub, stats);
    lastCard = card;
  }

  if (activeArea === 'HDBI' && lastCard) {
    // Append recent changes into the bottom of the HDBI-BDC card (last card)
    renderHdbiInlineChanges(lastCard, true);
  }

  if (activeArea === 'HDMX') {
    const coolantCard = document.createElement('article');
    statsGrid.appendChild(coolantCard);
    const coolantStats = buildHdmxCoolantStats(data || []);
    renderHdmxCoolantStatsCard(coolantCard, coolantStats);

    // Append recent changes into the bottom of the coolant card
    renderHdmxInlineChanges(coolantCard, true);
  }

  if (statsSection) statsSection.style.display = '';
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
    <p class="hdmx-changes-subtitle">Changes detected across the last 3 recorded work periods (up to 2 days).</p>
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
    <p class="hdmx-changes-subtitle">Changes detected across the last 3 recorded work periods (up to 2 days).</p>
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
    <p class="hdmx-changes-subtitle">Changes detected across the last 3 recorded work periods (up to 2 days).</p>
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
function renderMachineGrid(data, activeToolTokens = [], allowStaticHdmxTools = true, hdmxCoolantFilter = '') {
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

  let totalCards = 0;
  for (const section of sections) {
    if (activeArea === 'PPV' && section.area === 'PPV' && section.sub) {
      const wrapper = document.createElement('div');
      wrapper.className = 'ppv-stats-wrapper section-inline-stats';

      const statsCard = document.createElement('article');
      const stats = ppvStatsByBucket.get(statsBucketKey('PPV', section.sub)) || createEmptyStats();
      renderAreaStatsCard(statsCard, 'PPV', section.sub, stats);
      wrapper.appendChild(statsCard);

      const changesPanel = document.createElement('div');
      changesPanel.className = 'ppv-inline-changes-panel';
      wrapper.appendChild(changesPanel);
      renderPpvInlineChanges(changesPanel, section.sub);

      machineGrid.appendChild(wrapper);
    }

    // Safe ID for CSS (replace spaces)
    const secId = `${section.area}-${section.sub || 'all'}`.replace(/\s/g, '_');

    const secHeader = document.createElement('div');
    secHeader.className = `section-header area-bg-${section.area}`;
    secHeader.innerHTML = `
      <span class="section-area">${section.area}</span>
      ${section.sub ? `<span class="section-sub">${section.sub}</span>` : ''}
      <span class="section-count" id="sec-${secId}"></span>
    `;
    machineGrid.appendChild(secHeader);

    const secGrid = document.createElement('div');
    secGrid.className = 'section-cards';

    let sectionCards = 0;
    for (const toolId of section.tools) {
      const rows = byTool[toolId] || [];
      secGrid.appendChild(buildMachineCard(toolId, rows, section.area, section.sub));
      sectionCards++;
      totalCards++;
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

  // Sort tools inside each section alphabetically
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
  const card = document.createElement('div');
  card.className = `machine-card area-${area || 'HDBI'}`;
  card.dataset.tool = toolId;

  // Header
  const header = document.createElement('div');
  header.className = 'machine-card-header';
  header.innerHTML = `<span>${formatMachineDisplayName(toolId)}</span>`;
  if (sub || area) {
    const badge = document.createElement('span');
    badge.className = 'badge-sub';
    badge.textContent = sub || area || '';
    header.appendChild(badge);
  }
  card.appendChild(header);

  // 2D Layout Table
  const cc = colMapping.cellCol;
  const pc = colMapping.productCol;

  const key = layoutKey || getHdbiLayoutKey(toolId) || toolId;
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
      const product = dataRow ? (dataRow[pc] == null || dataRow[pc] === '' ? 'Empty' : dataRow[pc]) : 'Empty';

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
  header.innerHTML = `<span>${formatMachineDisplayName(toolId)}</span>`;
  if (sub || area) {
    const badgeGroup = document.createElement('div');
    badgeGroup.className = 'machine-card-badges';

    const areaBadge = document.createElement('span');
    areaBadge.className = 'badge-sub';
    areaBadge.textContent = sub || area || '';
    badgeGroup.appendChild(areaBadge);

    if (String(area || '').toUpperCase() === 'HDMX') {
      const coolant = getHdmxCoolant(toolId);
      if (coolant) {
        const coolantBadge = document.createElement('span');
        coolantBadge.className = 'badge-sub';
        coolantBadge.textContent = coolant;
        badgeGroup.appendChild(coolantBadge);
      }
    }

    header.appendChild(badgeGroup);
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
    const type = !ch.oldProduct ? 'Added' : !ch.newProduct ? 'Removed' : 'Changed';
    const cls  = type === 'Added' ? 'change-added' : type === 'Removed' ? 'change-removed' : 'change-modified';
    html += `<tr>
      <td>${ch.machine || '—'}</td>
      <td>${ch.cell || '—'}</td>
      <td class="change-old">${ch.oldProduct || '<em class="change-empty">Empty</em>'}</td>
      <td class="change-new">${ch.newProduct || '<em class="change-empty">Empty</em>'}</td>
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
