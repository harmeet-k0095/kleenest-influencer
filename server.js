require('dotenv').config();
const express = require('express');
const fetch = require('node-fetch');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Config (all secrets via environment variables) ───────────────────────────
const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

// Main live sheet (priyanka's drive)
const PRIYANKA_USER  = process.env.PRIYANKA_USER  || 'priyanka@kleenest.in';
const LIVE_FILE_ID   = process.env.LIVE_FILE_ID   || '01I5S75AMRWCD7BATFKZD3S2QT6KWNINU5';
const AGENCY_FILE_ID = process.env.AGENCY_FILE_ID || '01I5S75AMCVHK6UM4PXJHYMWSJNEFVJZYY';

// Individual order sheets (harmeet's drive)
const HARMEET_USER   = process.env.HARMEET_USER   || 'harmeet@kleenest.in';
const MOHIT_FILE_ID  = process.env.MOHIT_FILE_ID  || '01SI4WURI5JUACCCELQFF3YQHS5LEY5BUL';
const HARDEV_FILE_ID = process.env.HARDEV_FILE_ID || '01SI4WURPWNPRBPVDTHRGIQXZXI3RH2DRV';
const SATYAM_FILE_ID = process.env.SATYAM_FILE_ID || '01SI4WURNN6K7X3VGRBBA345MHU3WFADKB';
const APRIL_PLAN_ID  = process.env.APRIL_PLAN_ID  || '01SI4WUROWDSIEW7TMSBEJ6RXUIHMPUGKZ';

// ── Token Cache ──────────────────────────────────────────────────────────────
let tokenCache = { token: null, expiresAt: 0 };

async function getToken() {
  if (tokenCache.token && Date.now() < tokenCache.expiresAt) return tokenCache.token;
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      }),
    }
  );
  const data = await res.json();
  tokenCache = {
    token: data.access_token,
    expiresAt: Date.now() + (data.expires_in - 60) * 1000,
  };
  return tokenCache.token;
}

// ── Graph Helper ─────────────────────────────────────────────────────────────
async function graphGet(path) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  return res.json();
}

// ── Excel serial date → ISO string ──────────────────────────────────────────
function excelDateToISO(serial) {
  if (!serial || typeof serial !== 'number') return null;
  const utc = (serial - 25569) * 86400 * 1000;
  return new Date(utc).toISOString().split('T')[0];
}

// ── Name normalisers ─────────────────────────────────────────────────────────
function normBrand(b) {
  if (!b) return 'Unknown';
  const s = b.trim().toLowerCase();
  if (s === 'kleenest') return 'Kleenest';
  if (s === 'klenzmo')  return 'Klenzmo';
  return b.trim();
}
function normProduct(p) { return normAgencyProduct(p); }
function normPOC(p) {
  if (!p) return 'Unassigned';
  const m = { hardev: 'Hardev', 'hardev gill': 'Hardev', sattyam: 'Satyam', satyam: 'Satyam',
               priyanka: 'Priyanka', mohit: 'Mohit' };
  return m[p.trim().toLowerCase()] || p.trim();
}

// ── Excel error guard ────────────────────────────────────────────────────────
const EXCEL_ERRORS = ['#VALUE!', '#REF!', '#DIV/0!', '#NAME?', '#NULL!', '#NUM!', '#N/A'];
const isValid = v => {
  const s = String(v || '').trim();
  return s && !EXCEL_ERRORS.some(e => s.includes(e));
};

// ── Parse currency strings like "₹3,000.00" or plain numbers ────────────────
function parseAmount(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/[₹,\s]/g, '');
  return parseFloat(s) || 0;
}

// ── Normalise product names to the 11 canonical SKU names ────────────────────
const AGENCY_PRODUCT_NORM = {
  // ── Washing Machine range ────────────────────────────────────────────────
  'washing machine powder':          'Washing Machine Powder',
  'washing machine liquid':          'Washing Machine Liquid',
  'washing machine tablet':          'Washing Machine Tablet',
  'washing machine cleaner':         'Washing Machine Powder',  // legacy alias → powder
  'descale powder':                  'Washing Machine Powder',
  'descale liquid':                  'Washing Machine Liquid',
  'descale tablet':                  'Washing Machine Tablet',

  // ── Kitchen ──────────────────────────────────────────────────────────────
  'kitchen cleaner':                 'Kitchen Cleaner',

  // ── Magic Eraser ─────────────────────────────────────────────────────────
  'magic eraser':                    'Magic Eraser',

  // ── Bathroom Kits ────────────────────────────────────────────────────────
  'kleenest bathroom kit':           'Kleenest Bathroom Kit',
  'kleenest bathroom  kit':          'Kleenest Bathroom Kit',  // double-space variant
  'tc+bc kit':                       'Kleenest Bathroom Kit',
  'tc + bc kit':                     'Kleenest Bathroom Kit',
  'klenzmo bathroom kit':            'Klenzmo Bathroom Kit',

  // ── Tile & Floor ─────────────────────────────────────────────────────────
  'klenzmo floor & tile cleaner':    'Tile & Floor Cleaner',
  'klenzmo floor and tile cleaner':  'Tile & Floor Cleaner',
  'tile & floor cleaner':            'Tile & Floor Cleaner',
  'tile and floor cleaner':          'Tile & Floor Cleaner',

  // ── Copper / Metal ───────────────────────────────────────────────────────
  'copper and brass cleaner':        'Copper and Brass Cleaner',
  'metal cleaner':                   'Metal Cleaner Kit',
  'metal cleaner kit':               'Metal Cleaner Kit',

  // ── Trial Kit ────────────────────────────────────────────────────────────
  'cleaning trial kit':              'Cleaning Trial Kit',
  'cleaning trail kit':              'Cleaning Trial Kit',  // common typo
  'trial kit':                       'Cleaning Trial Kit',
};
function normAgencyProduct(name) {
  if (!name) return '';
  return AGENCY_PRODUCT_NORM[name.trim().toLowerCase()] || name.trim();
}

// ── Parse one agency sheet tab (TTC or INK REVENUE) ─────────────────────────
function parseAgencySheet(values, sourceName) {
  if (!values || values.length < 2) return [];
  const headers = values[0];
  const rows    = values.slice(1);

  const col         = kw => headers.findIndex(h => String(h).trim().toUpperCase() === kw.toUpperCase());
  const colIncludes = kw => headers.findIndex(h => String(h).trim().toUpperCase().includes(kw.toUpperCase()));

  const idx = {
    username:   col('USERNAME'),
    profileLink:col('PROFILE LINK'),
    brand:      col('BRAND NAME'),
    product:    col('PRODUCT NAME'),
    liveLink:   colIncludes('LIVE LINK'),
    liveDate:   colIncludes('DATE'),       // flexible: "LIVE DATE", "DATE", etc.
    productAmt: col('PRODUCT AMOUNT'),
    reelAmt:    col('REEL AMOUNT'),
    views:      col('VIEWS'),
    likes:      col('LIKES'),
    comments:   colIncludes('COMMENT'),
    cpv:        col('CPV'),
    language:   col('LANGUAGE'),
  };

  return rows
    .filter(r => r[idx.username] && String(r[idx.username]).trim())
    .map(r => {
      // Parse date: may be Excel serial number or already a string like "2026-04-05"
      let liveDate = null;
      if (idx.liveDate >= 0 && r[idx.liveDate]) {
        const raw = r[idx.liveDate];
        if (typeof raw === 'number') {
          liveDate = excelDateToISO(raw);
        } else {
          const s = String(raw).trim();
          // Accept YYYY-MM-DD; also handle DD/MM/YYYY
          if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
            liveDate = s;
          } else if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
            const [d, m, y] = s.split('/');
            liveDate = `${y}-${m}-${d}`;
          } else if (s) {
            liveDate = s;
          }
        }
      }
      return {
        source:     sourceName,
        poc:        'Agency',
        agency:     sourceName,
        username:   String(r[idx.username] || '').trim().replace(/\/$/, ''),
        profileLink:r[idx.profileLink] || '',
        brand:      normBrand(r[idx.brand]),
        product:    normAgencyProduct(r[idx.product]),
        liveLink:   r[idx.liveLink] || '',
        liveDate,
        productAmt: parseAmount(r[idx.productAmt]),
        reelAmt:    parseAmount(r[idx.reelAmt]),
        views:      Number(r[idx.views])    || 0,
        likes:      Number(r[idx.likes])    || 0,
        comments:   Number(r[idx.comments]) || 0,
        cpv:        Number(r[idx.cpv])      || 0,
        language:   String(r[idx.language] || '').trim(),
        total:      parseAmount(r[idx.reelAmt]) + parseAmount(r[idx.productAmt]),
      };
    });
}

// ── Count orders from an individual IS sheet ─────────────────────────────────
function countOrdersFromSheet(values, statusValues) {
  if (!values || values.length < 2) return 0;
  const headers = values[0];
  const rows    = values.slice(1);
  const statusIdx   = headers.findIndex(h => /ORDER STATUS|CREATOR STATUS/i.test(String(h)));
  const usernameIdx = headers.findIndex(h => /USER.?NAME|USERNAME/i.test(String(h)));
  if (statusIdx < 0 || usernameIdx < 0) return 0;
  return rows.filter(r =>
    r[usernameIdx] && String(r[usernameIdx]).trim() &&
    statusValues.some(sv => String(r[statusIdx]).trim().toLowerCase() === sv.toLowerCase())
  ).length;
}

// ── /api/debug/sheets — list all worksheet tabs in both key files ────────────
app.get('/api/debug/sheets', async (req, res) => {
  try {
    const prBase = `/users/${PRIYANKA_USER}/drive/items/${LIVE_FILE_ID}/workbook/worksheets`;
    const hmBase = `/users/${HARMEET_USER}/drive/items`;
    const [prSheets, aprilSheets] = await Promise.all([
      graphGet(prBase),
      graphGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets`),
    ]);
    res.json({
      priyanka_live_file: (prSheets.value || []).map(s => s.name),
      april_plan_file:    (aprilSheets.value || []).map(s => s.name),
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── /api/debug/peek — first 10 rows of candidate sheets ─────────────────────
app.get('/api/debug/peek', async (req, res) => {
  try {
    const prBase = `/users/${PRIYANKA_USER}/drive/items/${LIVE_FILE_ID}/workbook/worksheets`;
    const hmBase = `/users/${HARMEET_USER}/drive/items`;
    const enc = s => encodeURIComponent(s);
    const [agencySheet, targetUpdate, prSheet1, prSheet2, prSHeet2] = await Promise.all([
      graphGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('${enc('Agency Sheet ')}')/usedRange`),
      graphGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('${enc('Target Update')}')/usedRange`),
      graphGet(`${prBase}('Sheet1')/usedRange`),
      graphGet(`${prBase}('Sheet2')/usedRange`),
      graphGet(`${prBase}('${enc('SHeet 2')}')/usedRange`),
    ]);
    const head = raw => (raw.values || []).slice(0, 12);
    res.json({
      april_agency_sheet:   head(agencySheet),
      april_target_update:  head(targetUpdate),
      priyanka_sheet1:      head(prSheet1),
      priyanka_sheet2:      head(prSheet2),
      priyanka_SHeet2:      head(prSHeet2),
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── /api/debug/peek2 — check April plan Overview + Sheet1 ───────────────────
app.get('/api/debug/peek2', async (req, res) => {
  try {
    const hmBase = `/users/${HARMEET_USER}/drive/items`;
    const enc = s => encodeURIComponent(s);
    const [overview, sheet1, targetsPWise] = await Promise.all([
      graphGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('Overview')/usedRange`),
      graphGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('Sheet1')/usedRange`),
      graphGet(`/users/${PRIYANKA_USER}/drive/items/${LIVE_FILE_ID}/workbook/worksheets('${enc('Targets = P wise')}')/usedRange`),
    ]);
    const head = (raw, n=15) => (raw.values || []).slice(0, n);
    res.json({
      april_overview:    head(overview),
      april_sheet1:      head(sheet1),
      priyanka_targets_p_wise: head(targetsPWise),
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── /api/debug/resolve-share — find file ID from a SharePoint share URL ──────
app.get('/api/debug/resolve-share', async (req, res) => {
  try {
    const shareUrl = 'https://velocityeventures-my.sharepoint.com/:x:/g/personal/priyanka_kleenest_in/IQCRsIfwgmVWR7lqE_Ks1DadAUubLG1DFEYq5g6iifSKaxk?e=Bn3CRc';
    const encoded  = 'u!' + Buffer.from(shareUrl).toString('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
    const item = await graphGet(`/shares/${encoded}/driveItem`);
    // Also list priyanka's drive root to find the file
    const driveFiles = await graphGet(`/users/${PRIYANKA_USER}/drive/root/children`);
    res.json({
      shareItem: { id: item.id, name: item.name, error: item.error },
      driveFiles: (driveFiles.value || []).map(f => ({ id: f.id, name: f.name })),
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── /api/debug/list-folder — list Priyanka subfolder + search xlsx ───────────
app.get('/api/debug/list-folder', async (req, res) => {
  try {
    const PRIYANKA_FOLDER_ID = '01I5S75AMUPBOIRNQCENG26J54ZCCGPLZ3';
    const [allXlsx, searchLive, searchInhouse] = await Promise.all([
      graphGet(`/users/${PRIYANKA_USER}/drive/root/search(q='.xlsx')`),
      graphGet(`/users/${PRIYANKA_USER}/drive/root/search(q='live')`),
      graphGet(`/users/${PRIYANKA_USER}/drive/root/search(q='inhouse')`),
    ]);
    const fmt = arr => (arr.value || []).map(f => ({ id: f.id, name: f.name, path: f.parentReference?.path }));
    res.json({
      all_xlsx:       fmt(allXlsx),
      search_live:    fmt(searchLive),
      search_inhouse: fmt(searchInhouse),
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── /api/debug/raw-headers — first 3 rows of live sheet (tries all tab names) ──
app.get('/api/debug/raw-headers', async (req, res) => {
  try {
    const prBase = `/users/${PRIYANKA_USER}/drive/items/${LIVE_FILE_ID}/workbook/worksheets`;
    const LIVE_TAB_NAMES = ['Live Inhouse', 'Live In-house', 'Live I', 'Live Sheet', 'Sheet1', 'Sheet2', 'SHeet 2'];
    const results = {};
    for (const name of LIVE_TAB_NAMES) {
      const r = await graphGet(`${prBase}('${encodeURIComponent(name)}')/usedRange`).catch(e => ({ error: e.message }));
      const rows = (r.values || []).slice(0, 4);
      results[name] = {
        rowCount: (r.values || []).length,
        headers:  rows[0] || null,
        row1:     rows[1] || null,
      };
    }
    res.json(results);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── /api/dashboard ───────────────────────────────────────────────────────────
app.get('/api/dashboard', async (req, res) => {
  try {
    const prBase  = `/users/${PRIYANKA_USER}/drive/items/${LIVE_FILE_ID}/workbook/worksheets`;
    const hmBase  = `/users/${HARMEET_USER}/drive/items`;

    // Fetch all sheets in parallel
    // Helper: gracefully return empty if a sheet fetch fails (renamed/deleted tabs)
    const safeGet = path => graphGet(path).catch(() => ({ values: [] }));

    // Try multiple possible names for the live tab (it keeps getting renamed)
    const LIVE_TAB_NAMES = ['Live Inhouse', 'Live In-house', 'Live I', 'Live Sheet', 'Sheet1', 'Sheet2', 'SHeet 2'];
    const fetchLiveTab = async () => {
      for (const name of LIVE_TAB_NAMES) {
        const r = await safeGet(`${prBase}('${encodeURIComponent(name)}')/usedRange`);
        if (r.values && r.values.length > 1) return r;
      }
      return { values: [] };
    };

    const agBase = `/users/${PRIYANKA_USER}/drive/items/${AGENCY_FILE_ID}/workbook/worksheets`;

    const [
      liveRaw, targetRaw,
      mohitRaw, hardevRaw, satyamRaw, teamTargetsRaw, monthlyTargetRaw,
      overviewRaw, ttcRaw, inkRaw,
    ] = await Promise.all([
      fetchLiveTab(),
      safeGet(`${prBase}('Targets%20%3D%20P%20wise')/usedRange`),
      safeGet(`${hmBase}/${MOHIT_FILE_ID}/workbook/worksheets('Main%20Sheet')/usedRange`),
      safeGet(`${hmBase}/${HARDEV_FILE_ID}/workbook/worksheets('Main%20Sheet')/usedRange`),
      safeGet(`${hmBase}/${SATYAM_FILE_ID}/workbook/worksheets('April%20Creators%20')/usedRange`),
      safeGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('Team%20Targets')/usedRange`),
      safeGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('Monthy%20target')/usedRange`),
      graphGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('Overview')/usedRange`),
      safeGet(`${agBase}('TTC')/usedRange`),           // Agency: TTC
      safeGet(`${agBase}('INK%20REVENUE')/usedRange`), // Agency: Ink Revenue
    ]);

    // ── Parse Live Sheet ─────────────────────────────────────────────────────
    const mainValues = liveRaw.values || [];
    const headers    = mainValues[0] || [];
    const rows       = mainValues.slice(1);

    // Use case-insensitive, trimmed matching so renames don't break the parser
    const col = keyword => headers.findIndex(h => String(h).trim().toUpperCase() === keyword.toUpperCase());
    const colIncludes = keyword => headers.findIndex(h => String(h).trim().toUpperCase().includes(keyword.toUpperCase()));
    const idx = {
      poc:        col('POC'),
      liveDate:   col('LIVE DATE'),
      username:   col('USERNAME'),
      profileLink:col('PROFILE LINK'),
      brand:      col('BRAND NAME'),
      product:    col('PRODUCT NAME'),
      liveLink:   colIncludes('LIVE LINK'),
      approval:   col('APPROVAL'),
      productAmt: col('PRODUCT AMOUNT'),
      reelAmt:    col('REEL AMOUNT'),
      views:      col('VIEWS'),
      likes:      col('LIKES'),
      comments:   col('COMMENT'),
      cpv:        col('CPV'),
      total:      col('TOTAL'),
    };

    const EXCLUDED_POCS = ['dhruvika'];

    const influencers = rows
      .filter(r => {
        const poc      = normPOC(r[idx.poc]);
        const username = r[idx.username];
        if (EXCLUDED_POCS.includes(poc.toLowerCase())) return false;
        return isValid(username) && (isValid(r[idx.product]) || Number(r[idx.reelAmt]) > 0);
      })
      .map(r => ({
        poc:        normPOC(r[idx.poc]),
        liveDate:   excelDateToISO(r[idx.liveDate]),
        username:   String(r[idx.username]).trim().replace(/\/$/, ''),
        profileLink:r[idx.profileLink] || '',
        brand:      normBrand(r[idx.brand]),
        product:    normProduct(r[idx.product]),
        liveLink:   r[idx.liveLink] || '',
        approval:   String(r[idx.approval] || '').trim(),
        productAmt: Number(r[idx.productAmt]) || 0,
        reelAmt:    Number(r[idx.reelAmt])    || 0,
        views:      Number(r[idx.views])       || 0,
        likes:      Number(r[idx.likes])       || 0,
        comments:   Number(r[idx.comments])    || 0,
        cpv:        Number(r[idx.cpv])         || 0,
        total:      Number(r[idx.total])       || 0,
      }));

    // ── Reels gone live (have a live link) ───────────────────────────────────
    const reelsLive = rows.filter(r =>
      r[idx.liveLink] && String(r[idx.liveLink]).trim().startsWith('http')
    ).length;

    // ── Parse Reel Targets (Targets = P wise) ───────────────────────────────
    const tValues = targetRaw.values || [];
    const reelTargets = {};
    if (tValues.length >= 3) {
      tValues[1].forEach((h, i) => {
        if (h) reelTargets[String(h).trim()] = Number(tValues[2][i]) || 0;
      });
    }
    // keep old key for backward compat
    const productTargets = reelTargets;

    // ── Parse Order Targets (Monthly target sheet) ───────────────────────────
    const mtValues = monthlyTargetRaw.values || [];
    const orderTargets = {};
    let inProductSection = false;
    for (const row of mtValues) {
      const cell0 = String(row[0] || '').trim();
      const cell1 = String(row[1] || '').trim();
      if (cell0.includes('Product-wise Order Allocation')) { inProductSection = true; continue; }
      if (inProductSection) {
        if (cell0 === 'Product' || cell0 === '') continue;
        if (cell0.toLowerCase().includes('total')) break;
        const val = Number(row[1]);
        if (cell0 && !isNaN(val) && val > 0) orderTargets[cell0.trim()] = val;
      }
    }

    // ── Team Targets — parse from Overview sheet ─────────────────────────────
    // Overview rows 12-14 (0-indexed) contain: col0=name, col1=order target
    // col2=label (e.g. "Agency (Priyanka)"), col3=agency go-live target
    const overviewValues = overviewRaw.values || [];
    const teamTargets = {
      // Hardcoded fallbacks from Overview sheet "Targets given to team" section
      Mohit:    300,
      Satyam:   200,
      Hardev:   200,
      Priyanka: 660,
    };
    // Try to override with live sheet data if the structure matches
    overviewValues.forEach(row => {
      const name = normPOC(String(row[0] || '').trim());
      const num  = Number(row[1]);
      if (['Mohit','Hardev','Satyam'].includes(name) && num > 0) {
        teamTargets[name] = num;
      }
      // Agency (Priyanka) target from col2/col3
      const label3 = String(row[2] || '').trim().toLowerCase();
      const num3   = Number(row[3]);
      if (label3.includes('priyanka') && num3 > 0) {
        teamTargets['Priyanka'] = num3;
      }
    });

    // ── Orders placed per POC from individual sheets ──────────────────────────
    const ORDER_STATUSES = ['order place', 'order placed', 'live'];
    const pocOrders = {
      Mohit:    countOrdersFromSheet(mohitRaw.values,  ORDER_STATUSES),
      Hardev:   countOrdersFromSheet(hardevRaw.values, ORDER_STATUSES),
      Satyam:   countOrdersFromSheet(satyamRaw.values, ORDER_STATUSES),
      Priyanka: influencers.filter(i => i.poc === 'Priyanka').length,
    };

    // ── Aggregates ───────────────────────────────────────────────────────────
    const productMap = {};
    influencers.forEach(inf => {
      const p = inf.product || 'Unknown';
      if (!productMap[p]) productMap[p] = { reels: 0, approved: 0, totalPayout: 0, reelPayout: 0, brands: {}, pocs: {} };
      productMap[p].reels++;
      if (inf.approval.toLowerCase().includes('approved')) {
        productMap[p].approved++;
        productMap[p].totalPayout += inf.total;
        productMap[p].reelPayout  += inf.reelAmt;
      }
      productMap[p].brands[inf.brand] = (productMap[p].brands[inf.brand] || 0) + 1;
      productMap[p].pocs[inf.poc]     = (productMap[p].pocs[inf.poc] || 0) + 1;
    });

    const pocMap = {};
    influencers.forEach(inf => {
      const p = inf.poc;
      if (!pocMap[p]) pocMap[p] = { reels: 0, approved: 0, totalPayout: 0 };
      pocMap[p].reels++;
      if (inf.approval.toLowerCase().includes('approved')) {
        pocMap[p].approved++;
        pocMap[p].totalPayout += inf.total;
      }
    });

    const brandMap = {};
    influencers.forEach(inf => {
      const b = inf.brand;
      if (!brandMap[b]) brandMap[b] = { reels: 0, approved: 0, totalPayout: 0 };
      brandMap[b].reels++;
      if (inf.approval.toLowerCase().includes('approved')) {
        brandMap[b].approved++;
        brandMap[b].totalPayout += inf.total;
      }
    });

    // ── KPIs ─────────────────────────────────────────────────────────────────
    const totalOrders    = Object.values(pocOrders).reduce((s, v) => s + v, 0);
    const approved       = influencers.filter(i => i.approval.toLowerCase().includes('approved'));
    const totalPayout    = approved.reduce((s, i) => s + i.total, 0);
    const uniqueProducts = [...new Set(influencers.map(i => i.product).filter(Boolean))];

    // ── Live reels payout — computed from RAW rows (not filtered influencers) ──
    // Using raw rows ensures we include every row that has a live link,
    // even rows excluded from the influencers array (Excel errors, missing username, etc.)
    let liveReelsReelAmt    = 0;
    let liveReelsProductAmt = 0;
    rows.forEach(r => {
      const link = String(r[idx.liveLink] || '').trim();
      if (!link.startsWith('http')) return;
      liveReelsReelAmt    += Number(r[idx.reelAmt])    || 0;
      liveReelsProductAmt += Number(r[idx.productAmt]) || 0;
    });
    const liveReelsPayout = liveReelsReelAmt + liveReelsProductAmt;

    // ── SKU targets: merge reel + order targets with actuals ─────────────────
    // Canonical product name map (target sheet name → live sheet name)
    const SKU_MAP = {
      'Magic Eraser':   'Magic eraser',
      'Descale Powder': 'Washing Machine Powder',
      'Descale Liquid': 'Washing Machine Liquid',
      'Descale Tablet': 'Washing Machine Tablet',
      'TC+BC Kit':      'Kleenest Bathrooom Kit',
      'TC + BC Kit':    'Kleenest Bathrooom Kit',
      'Trial Kit':      'Cleaning Trail kit',
      'Metal Cleaner':  'Copper and Brass Cleaner',
      'Kitchen Cleaner':'Kitchen cleaner',
      'Klenzmo Bathroom Kit':          'Klenzmo Bathroom Kit',
      'Klenzmo floor & Tile Cleaner':  'Klenzmo Bathroom Kit',
    };

    // ── Parse Overview tab: TTC, Ink Revenue, and in-house SKU live targets ───
    // row2: product names at col 3+
    // row3: In-house targets
    // row4: TTC Agency Go-Live targets
    // row5: Ink Revenue Agency Go-Live targets
    const ovValues = overviewRaw.values || [];
    const ttcSkuLiveTargets   = {};   // product → TTC target
    const inkSkuLiveTargets   = {};   // product → INK REVENUE target
    const agencySkuLiveTargets= {};   // product → TTC + INK combined
    const inhouseSkuLiveTargets={};   // product → in-house target
    if (ovValues.length >= 6) {
      const productHeader = ovValues[2] || [];
      const inhouseRow    = ovValues[3] || [];
      const ttcRow        = ovValues[4] || [];
      const inkRow        = ovValues[5] || [];
      productHeader.slice(3).forEach((name, i) => {
        if (!name) return;
        const c   = i + 3;
        const ttc = Number(ttcRow[c]) || 0;
        const ink = Number(inkRow[c]) || 0;
        const key = String(name).trim();
        inhouseSkuLiveTargets[key] = Number(inhouseRow[c]) || 0;
        ttcSkuLiveTargets[key]     = ttc;
        inkSkuLiveTargets[key]     = ink;
        agencySkuLiveTargets[key]  = ttc + ink;
      });
    }

    // ── In-house actual live counts per product — derived from Live Sheet ────
    const inhouseSkuActuals = {};
    influencers.forEach(inf => {
      if (!inf.liveLink || !String(inf.liveLink).startsWith('http')) return;
      const p = inf.product;
      if (!p) return;
      inhouseSkuActuals[p] = (inhouseSkuActuals[p] || 0) + 1;
    });

    // ── Agency: parse TTC + INK REVENUE sheets ───────────────────────────────
    const ttcInfluencers = parseAgencySheet(ttcRaw.values,  'TTC');
    const inkInfluencers = parseAgencySheet(inkRaw.values,  'INK REVENUE');
    const agencyInfluencers = [...ttcInfluencers, ...inkInfluencers];

    const isLive = inf => inf.liveLink && String(inf.liveLink).trim().startsWith('http');

    // Per-agency live counts
    const ttcReelsLive    = ttcInfluencers.filter(isLive).length;
    const inkReelsLive    = inkInfluencers.filter(isLive).length;
    const agencyReelsLive = ttcReelsLive + inkReelsLive;

    // Per-agency SKU actuals
    const ttcSkuActuals    = {};
    const inkSkuActuals    = {};
    const agencySkuActuals = {};
    const addSku = (map, inf) => {
      const p = inf.product; if (!p) return;
      map[p] = (map[p] || 0) + 1;
      agencySkuActuals[p] = (agencySkuActuals[p] || 0) + 1;
    };
    ttcInfluencers.forEach(inf => { if (isLive(inf)) addSku(ttcSkuActuals, inf); });
    inkInfluencers.forEach(inf => { if (isLive(inf)) addSku(inkSkuActuals, inf); });

    // Per-agency payout
    let ttcReelsReelAmt = 0, ttcReelsProductAmt = 0;
    let inkReelsReelAmt = 0, inkReelsProductAmt = 0;
    ttcInfluencers.forEach(inf => { if (!isLive(inf)) return; ttcReelsReelAmt += inf.reelAmt; ttcReelsProductAmt += inf.productAmt; });
    inkInfluencers.forEach(inf => { if (!isLive(inf)) return; inkReelsReelAmt += inf.reelAmt; inkReelsProductAmt += inf.productAmt; });
    const ttcReelsPayout    = ttcReelsReelAmt    + ttcReelsProductAmt;
    const inkReelsPayout    = inkReelsReelAmt    + inkReelsProductAmt;
    const agencyReelsReelAmt    = ttcReelsReelAmt    + inkReelsReelAmt;
    const agencyReelsProductAmt = ttcReelsProductAmt + inkReelsProductAmt;
    const agencyReelsPayout     = agencyReelsReelAmt + agencyReelsProductAmt;

    // Count actual reels & orders per product from influencers
    const productActuals = {};
    influencers.forEach(inf => {
      const p = inf.product;
      if (!p) return;
      if (!productActuals[p]) productActuals[p] = { reels: 0 };
      productActuals[p].reels++;
    });

    // Normalise order target keys to match reel target keys
    const normSkuKey = k => k.replace(/\s+/g, ' ').trim()
      .replace('TC + BC Kit', 'TC+BC Kit');
    const normOrderTargets = {};
    Object.entries(orderTargets).forEach(([k, v]) => {
      normOrderTargets[normSkuKey(k)] = v;
    });

    // Build unified SKU targets array (deduplicated)
    const allSkuNames = [...new Set([
      ...Object.keys(reelTargets),
      ...Object.keys(normOrderTargets),
    ])];

    const skuTargets = allSkuNames.map(name => {
      const liveKey    = SKU_MAP[name] || name;
      const reelTarget  = reelTargets[name]      || 0;
      const orderTarget = normOrderTargets[name] || 0;
      const actualReels= productActuals[liveKey]?.reels || 0;
      return { name, liveKey, reelTarget, orderTarget, actualReels };
    });

    const overallReelTarget  = skuTargets.reduce((s, x) => s + x.reelTarget, 0);
    // Reel live targets: Agency 660 + In-house 400 = 1060
    const agencyReelTarget   = 660;
    const inhouseReelTarget  = 400;
    const totalReelTarget    = agencyReelTarget + inhouseReelTarget;
    // Order targets: Agency (Priyanka) 660 + In-house 400 = 1015
    const overallOrderTarget = 1015;
    const agencyOrderTarget  = 660;
    const inhouseOrderTarget = 400;
    const overallActualReels = influencers.length;

    res.json({
      kpis: {
        totalOrders,
        reelsLive,
        totalPayout,
        uniqueProducts: uniqueProducts.length,
        liveReelsPayout,
        liveReelsReelAmt,
        liveReelsProductAmt,
        agencyReelsLive,
        agencyReelsPayout,
        agencyReelsReelAmt,
        agencyReelsProductAmt,
        ttcReelsLive,
        ttcReelsPayout,
        ttcReelsReelAmt,
        ttcReelsProductAmt,
        inkReelsLive,
        inkReelsPayout,
        inkReelsReelAmt,
        inkReelsProductAmt,
        combinedReelsLive: reelsLive + agencyReelsLive,
        combinedPayout:    liveReelsPayout + agencyReelsPayout,
      },
      products:     productMap,
      pocs:         pocMap,
      brands:       brandMap,
      productTargets,
      reelTargets,
      orderTargets,
      skuTargets,
      agencySkuLiveTargets,
      ttcSkuLiveTargets,
      inkSkuLiveTargets,
      inhouseSkuLiveTargets,
      inhouseSkuActuals,
      agencySkuActuals,
      ttcSkuActuals,
      inkSkuActuals,
      agencyInfluencers,
      overallTargets: { reelTarget: overallReelTarget, agencyReelTarget, inhouseReelTarget, totalReelTarget, orderTarget: overallOrderTarget, agencyOrderTarget, inhouseOrderTarget, actualReels: overallActualReels, totalOrders },
      teamTargets,
      pocOrders,
      influencers,
      lastUpdated:  new Date().toISOString(),
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── POST /api/metrics/start — kick off an Apify run (returns instantly) ───────
// Body: { urls: [...] }
// Returns: { runId, datasetId }
app.post('/api/metrics/start', async (req, res) => {
  const { urls } = req.body || {};
  if (!Array.isArray(urls) || urls.length === 0) return res.json({ runId: null });

  const token = process.env.APIFY_TOKEN;
  if (!token) return res.status(500).json({ error: 'APIFY_TOKEN not configured' });

  try {
    const r = await fetch(
      `https://api.apify.com/v2/acts/apify~instagram-scraper/runs?token=${token}&memory=1024`,
      {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({
          directUrls:   urls,
          resultsType:  'posts',
          resultsLimit: urls.length,
        }),
      }
    );
    const j = await r.json();
    if (!j?.data?.id) return res.status(502).json({ error: 'Apify start failed', detail: j });
    res.json({ runId: j.data.id, datasetId: j.data.defaultDatasetId });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── GET /api/metrics/result/:runId — poll status; return items when done ──────
// Returns: { status: 'RUNNING'|'SUCCEEDED'|'FAILED', items?: [...] }
app.get('/api/metrics/result/:runId', async (req, res) => {
  const token = process.env.APIFY_TOKEN;
  if (!token) return res.status(500).json({ error: 'APIFY_TOKEN not configured' });

  try {
    const runRes = await fetch(
      `https://api.apify.com/v2/actor-runs/${req.params.runId}?token=${token}`
    );
    const run = await runRes.json();
    const status = run?.data?.status;

    if (status !== 'SUCCEEDED') return res.json({ status: status || 'UNKNOWN' });

    // Run finished — fetch dataset items
    const datasetId = run.data.defaultDatasetId;
    const itemsRes  = await fetch(
      `https://api.apify.com/v2/datasets/${datasetId}/items?token=${token}&clean=true`
    );
    const items = await itemsRes.json();

    const metrics = (Array.isArray(items) ? items : []).map(it => ({
      url:      (it.url || it.inputUrl || '').split('?')[0].replace(/\/$/, ''),
      views:    it.videoPlayCount ?? it.playCount ?? it.videoViewCount ?? 0,
      likes:    it.likesCount     ?? it.likeCount  ?? 0,
      comments: it.commentsCount  ?? it.commentCount ?? 0,
    }));

    res.json({ status: 'SUCCEEDED', items: metrics });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Start server locally; export for Vercel serverless ───────────────────────
if (require.main === module) {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => console.log(`Kleenest Dashboard running on http://localhost:${PORT}`));
}

module.exports = app;
