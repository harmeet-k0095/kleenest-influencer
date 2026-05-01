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
const PRIYANKA_USER   = process.env.PRIYANKA_USER   || 'priyanka@kleenest.in';
const LIVE_FILE_ID    = process.env.LIVE_FILE_ID    || '01I5S75AMRWCD7BATFKZD3S2QT6KWNINU5';
const AGENCY_FILE_ID  = process.env.AGENCY_FILE_ID  || '01I5S75AMCVHK6UM4PXJHYMWSJNEFVJZYY';
// Priyanka's dedicated order sheet (Tile & Floor Cleaner orders)
const PRIYANKA_TFC_ID = process.env.PRIYANKA_TFC_ID || '01I5S75AIEWB4M6G7HNJFYN6AIPX5OROKD';

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
  'kleenest bathrooom kit':          'Kleenest Bathroom Kit',  // 3-o typo variant
  'tc+bc kit':                       'Kleenest Bathroom Kit',
  'tc + bc kit':                     'Kleenest Bathroom Kit',
  'klenzmo bathroom kit':            'Klenzmo Bathroom Kit',

  // ── Tile & Floor ─────────────────────────────────────────────────────────
  'klenzmo floor & tile cleaner':    'Tile & Floor Cleaner',
  'klenzmo floor and tile cleaner':  'Tile & Floor Cleaner',
  'klenzmo tile cleaner new':        'Tile & Floor Cleaner',
  'tile & floor cleaner':            'Tile & Floor Cleaner',
  'tile& floor cleaner':             'Tile & Floor Cleaner',  // no-space variant
  'floor & tile cleaner':            'Tile & Floor Cleaner',
  'floor and tile cleaner':          'Tile & Floor Cleaner',
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
// Convert Excel date serial → JS Date (accounts for Excel's 1900 leap-year bug)
const excelSerialToDate = serial => {
  if (!serial || typeof serial !== 'number' || serial < 1) return null;
  return new Date(Math.round((serial - 25569) * 86400000));
};
// Returns true if the date serial falls in April 2026
const isApril2026 = serial => {
  const d = excelSerialToDate(serial);
  return d && d.getUTCFullYear() === 2026 && d.getUTCMonth() === 3; // month 3 = April (0-indexed)
};

function countOrdersFromSheet(values, statusValues) {
  if (!values || values.length < 2) return 0;
  const headers = values[0];
  const rows    = values.slice(1);
  const statusIdx   = headers.findIndex(h => /ORDER STATUS|CREATOR STATUS/i.test(String(h)));
  const usernameIdx = headers.findIndex(h => /USER.?NAME|USERNAME/i.test(String(h)));
  const dateIdx     = headers.findIndex(h => /ORDER.?DATE/i.test(String(h)));
  if (statusIdx < 0 || usernameIdx < 0) return 0;
  return rows.filter(r => {
    if (!r[usernameIdx] || !String(r[usernameIdx]).trim() || String(r[usernameIdx]).trim() === '#N/A') return false;
    if (!statusValues.some(sv => String(r[statusIdx]).trim().toLowerCase() === sv.toLowerCase())) return false;
    // If date column found, restrict to April 2026; rows with no date still counted
    if (dateIdx >= 0 && r[dateIdx]) return isApril2026(r[dateIdx]);
    return true;
  }).length;
}

// ── Count orders by SKU from an individual IS sheet ───────────────────────────
// Returns { productName: count } for rows matching the given statuses
function countOrdersBySkuFromSheet(values, statusValues) {
  if (!values || values.length < 2) return {};
  const headers = values[0];
  const rows    = values.slice(1);
  const statusIdx   = headers.findIndex(h => /ORDER STATUS|CREATOR STATUS/i.test(String(h)));
  const usernameIdx = headers.findIndex(h => /USER.?NAME|USERNAME/i.test(String(h)));
  const productIdx  = headers.findIndex(h => /PRODUCT.?NAME/i.test(String(h)));
  const dateIdx     = headers.findIndex(h => /ORDER.?DATE/i.test(String(h)));
  if (statusIdx < 0 || usernameIdx < 0 || productIdx < 0) return {};
  const skuMap = {};
  rows.forEach(r => {
    if (!r[usernameIdx] || !String(r[usernameIdx]).trim() || String(r[usernameIdx]).trim() === '#N/A') return;
    if (!statusValues.some(sv => String(r[statusIdx]).trim().toLowerCase() === sv.toLowerCase())) return;
    if (dateIdx >= 0 && r[dateIdx] && !isApril2026(r[dateIdx])) return;
    const raw = String(r[productIdx] || '').trim();
    const sku = normAgencyProduct(raw) || raw || 'Unknown';
    skuMap[sku] = (skuMap[sku] || 0) + 1;
  });
  return skuMap;
}

// ── /api/debug/inhouse-amounts — raw product/reel/total amounts from live sheet ──
app.get('/api/debug/inhouse-amounts', async (req, res) => {
  try {
    const prBase = `/users/${PRIYANKA_USER}/drive/items/${LIVE_FILE_ID}/workbook/worksheets`;
    const LIVE_TAB_NAMES = ['Live Inhouse', 'Live In-house', 'Live I', 'Live Sheet', 'Sheet1', 'Sheet2', 'SHeet 2'];
    let raw = { values: [] };
    for (const name of LIVE_TAB_NAMES) {
      const r = await graphGet(`${prBase}('${encodeURIComponent(name)}')/usedRange`).catch(() => ({}));
      if (r.values && r.values.length > 1) { raw = r; break; }
    }
    const values  = raw.values || [];
    const headers = values[0] || [];
    const rows    = values.slice(1);

    const col         = kw => headers.findIndex(h => String(h).trim().toUpperCase() === kw.toUpperCase());
    const colIncludes = kw => headers.findIndex(h => String(h).trim().toUpperCase().includes(kw.toUpperCase()));
    const liveIdx    = colIncludes('LIVE LINK');
    const userIdx    = col('USERNAME');
    const prodAmtIdx = col('PRODUCT AMOUNT');
    const reelAmtIdx = col('REEL AMOUNT');
    const totalIdx   = col('TOTAL');

    const isSocial = link => {
      const u = String(link || '').trim();
      return u.includes('instagram.com') || u.includes('youtube.com') || u.includes('youtu.be');
    };

    const liveRows = rows.filter(r => isSocial(r[liveIdx]));

    // All rows — collect raw values before any parsing
    const allRowsData = rows.map((r, i) => {
      const link = String(r[liveIdx] || '').trim();
      const isSocial = link.includes('instagram.com') || link.includes('youtube.com') || link.includes('youtu.be');
      const hasAnyLink = link.startsWith('http') || link.startsWith('https');
      return {
        row:        i + 2,
        username:   String(r[userIdx] || '').slice(0, 30),
        liveLink:   link.slice(0, 80),
        linkType:   !link ? 'NONE' : isSocial ? 'social' : hasAnyLink ? 'other_http' : 'non_http',
        rawProduct: r[prodAmtIdx],
        parsedProduct: parseAmount(r[prodAmtIdx]),
        rawReel:    r[reelAmtIdx],
        parsedReel: parseAmount(r[reelAmtIdx]),
      };
    });

    const socialRows   = allRowsData.filter(r => r.linkType === 'social');
    const otherHttpRows= allRowsData.filter(r => r.linkType === 'other_http');
    const nonHttpRows  = allRowsData.filter(r => r.linkType === 'non_http');
    const noLinkRows   = allRowsData.filter(r => r.linkType === 'NONE');

    const sumSocial    = socialRows.reduce((s, r) => s + r.parsedProduct, 0);
    const sumOtherHttp = otherHttpRows.reduce((s, r) => s + r.parsedProduct, 0);
    const sumAll       = allRowsData.reduce((s, r) => s + r.parsedProduct, 0);

    res.json({
      headers,
      totalRows:      rows.length,
      socialRows:     socialRows.length,
      otherHttpRows:  otherHttpRows.length,
      noLinkRows:     noLinkRows.length,
      sumSocialProduct:    sumSocial,
      sumOtherHttpProduct: sumOtherHttp,
      sumAllProduct:       sumAll,
      otherHttpDetails: otherHttpRows.map(r => ({ row: r.row, username: r.username, link: r.liveLink, product: r.parsedProduct, reel: r.parsedReel })),
      zeroProductSocialRows: socialRows.filter(r => r.parsedProduct === 0).map(r => ({ row: r.row, username: r.username, rawProduct: r.rawProduct })),
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

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

// ── /api/debug/ink — raw Ink Revenue sheet inspection ────────────────────────
app.get('/api/debug/ink', async (req, res) => {
  try {
    const agBase = `/users/${PRIYANKA_USER}/drive/items/${AGENCY_FILE_ID}/workbook/worksheets`;
    const raw = await graphGet(`${agBase}('INK%20REVENUE')/usedRange`);
    const values = raw.values || [];
    const headers = values[0] || [];
    const rows = values.slice(1);

    const colIncludes = kw => headers.findIndex(h => String(h).trim().toUpperCase().includes(kw.toUpperCase()));
    const col = kw => headers.findIndex(h => String(h).trim().toUpperCase() === kw.toUpperCase());
    const liveIdx = colIncludes('LIVE LINK');
    const userIdx = col('USERNAME');
    const productIdx = col('PRODUCT NAME');

    const allRows = rows.map((r, i) => ({
      row: i + 2,
      username: r[userIdx] || '',
      product: r[productIdx] || '',
      liveLink: r[liveIdx] || '',
      hasLink: String(r[liveIdx] || '').startsWith('http'),
    }));

    const withLink = allRows.filter(r => r.hasLink);
    const withLinkNoUser = allRows.filter(r => r.hasLink && !String(r.username).trim());

    res.json({
      totalRows: rows.length,
      headers,
      liveColIndex: liveIdx,
      userColIndex: userIdx,
      withLiveLink: withLink.length,
      withLinkButNoUsername: withLinkNoUser.length,
      withLinkRows: withLink.map(r => ({ row: r.row, username: r.username, product: r.product, liveLink: r.liveLink.slice(0,80) })),
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

// ── /api/debug/mohit — inspect Mohit's sheet headers + sample rows ───────────
app.get('/api/debug/mohit', async (req, res) => {
  try {
    const hmBase = `/users/${HARMEET_USER}/drive/items`;
    const raw = await graphGet(`${hmBase}/${MOHIT_FILE_ID}/workbook/worksheets('Main%20Sheet')/usedRange`).catch(e => ({ values: [], error: e.message }));
    const values = raw.values || [];
    res.json({
      totalRows: values.length,
      headers: values[0] || [],
      sample: values.slice(1, 6),
      lastRows: values.slice(-5),
    });
  } catch (err) { res.status(500).json({ error: err.message }); }
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

    const priyankaTfcBase = `/users/${PRIYANKA_USER}/drive/items/${PRIYANKA_TFC_ID}/workbook/worksheets`;

    const [
      liveRaw, targetRaw,
      mohitRaw, hardevRaw, satyamRaw, teamTargetsRaw, monthlyTargetRaw,
      overviewRaw, ttcRaw, inkRaw, priyankaTfcRaw,
    ] = await Promise.all([
      fetchLiveTab(),
      safeGet(`${prBase}('Targets%20%3D%20P%20wise')/usedRange`),
      safeGet(`${hmBase}/${MOHIT_FILE_ID}/workbook/worksheets('Main%20Sheet')/usedRange`),
      safeGet(`${hmBase}/${HARDEV_FILE_ID}/workbook/worksheets('Main%20Sheet')/usedRange`),
      safeGet(`${hmBase}/${SATYAM_FILE_ID}/workbook/worksheets('April%20master%20sheet')/usedRange`),
      safeGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('Team%20Targets')/usedRange`),
      safeGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('Monthy%20target')/usedRange`),
      graphGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('Overview')/usedRange`),
      safeGet(`${agBase}('TTC')/usedRange`),           // Agency: TTC
      safeGet(`${agBase}('INK%20REVENUE')/usedRange`), // Agency: Ink Revenue
      safeGet(`${priyankaTfcBase}('Sheet1')/usedRange`), // Priyanka's Tile & Floor Cleaner orders
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
        productAmt: parseAmount(r[idx.productAmt]),
        reelAmt:    parseAmount(r[idx.reelAmt]),
        views:      Number(r[idx.views])       || 0,
        likes:      Number(r[idx.likes])       || 0,
        comments:   Number(r[idx.comments])    || 0,
        cpv:        Number(r[idx.cpv])         || 0,
        total:      parseAmount(r[idx.total]),
      }));

    // ── Reels gone live (have a live link — Instagram or YouTube only) ────────
    const reelsLive = rows.filter(r => {
      const link = String(r[idx.liveLink] || '').trim();
      return link.includes('instagram.com') || link.includes('youtube.com') || link.includes('youtu.be');
    }).length;

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
    // Overview "Orders Placed - In-house Targets" section has:
    //   Mohit: 300, Satyam: 200, Hardev: 200, Priyanka: 100
    // NOTE: Priyanka = 100 orders placed target (not 660 — that's agency live reels, a different KPI)
    const overviewValues = overviewRaw.values || [];
    const teamTargets = {
      Mohit:    300,
      Satyam:   200,
      Hardev:   200,
      Priyanka: 100,
    };
    // Override with live values from Overview "Orders Placed - In-house Targets" section.
    // NOTE: Only read Mohit/Hardev/Satyam from the sheet — Priyanka's cell is a live formula
    // (=COUNTA) that returns her actual count, not her target of 100.
    let inOrdersSection = false;
    overviewValues.forEach(row => {
      const label = String(row[0] || '').trim();
      if (label.toLowerCase().includes('orders placed - in-house')) { inOrdersSection = true; return; }
      if (inOrdersSection) {
        const name = normPOC(label);
        const num  = Number(row[1]);
        if (['Mohit', 'Hardev', 'Satyam'].includes(name) && num > 0) {
          teamTargets[name] = num;
        }
      }
    });

    // ── Orders placed per POC from individual sheets ──────────────────────────
    const ORDER_STATUSES = ['order place', 'order placed', 'live'];
    const pocOrders = {
      Mohit:    countOrdersFromSheet(mohitRaw.values,      ORDER_STATUSES),
      Hardev:   countOrdersFromSheet(hardevRaw.values,     ORDER_STATUSES),
      Satyam:   countOrdersFromSheet(satyamRaw.values,     ORDER_STATUSES),
      Priyanka: countOrdersFromSheet(priyankaTfcRaw.values, ORDER_STATUSES),
    };

    // ── SKU-wise orders per POC ───────────────────────────────────────────────
    const pocSkuOrders = {
      Mohit:    countOrdersBySkuFromSheet(mohitRaw.values,       ORDER_STATUSES),
      Hardev:   countOrdersBySkuFromSheet(hardevRaw.values,      ORDER_STATUSES),
      Satyam:   countOrdersBySkuFromSheet(satyamRaw.values,      ORDER_STATUSES),
      Priyanka: countOrdersBySkuFromSheet(priyankaTfcRaw.values, ORDER_STATUSES),
    };

    // ── SKU Order Targets — from Overview rows 19-20 ──────────────────────────
    // row19: header row with SKU names, row20: corresponding target numbers
    // Canonical map from sheet names → canonical SKU names
    const SKU_ORDER_TARGET_NORM = {
      'magic eraser':              'Magic Eraser',
      'descale powder':            'Washing Machine Powder',
      'descale liquid':            'Washing Machine Liquid',
      'descale tablet':            'Washing Machine Tablet',
      'tc+bc kit':                 'Kleenest Bathroom Kit',
      'tc + bc kit':               'Kleenest Bathroom Kit',
      'trial kit':                 'Cleaning Trial Kit',
      'metal cleaner':             'Metal Cleaner Kit',
      'kitchen cleaner':           'Kitchen Cleaner',
      'klenzmo tile cleaner new':  'Tile & Floor Cleaner',
      'klenzmo floor & tile cleaner': 'Tile & Floor Cleaner',
    };
    const skuOrderTargets = {};   // canonical SKU name → target order count
    let skuHeaderRow = null, skuTargetRow = null;
    for (let ri = 0; ri < overviewValues.length; ri++) {
      const row = overviewValues[ri];
      const cell0 = String(row[0] || '').trim().toLowerCase();
      if (cell0 === 'magic eraser' || cell0 === 'descale powder') {
        // This is the SKU header row
        skuHeaderRow = row;
        skuTargetRow = overviewValues[ri + 1] || [];
        break;
      }
    }
    if (skuHeaderRow && skuTargetRow) {
      skuHeaderRow.forEach((name, i) => {
        if (!name) return;
        const key = String(name).trim().toLowerCase();
        if (key === 'total orders placed' || key === 'total') return;
        const canonical = SKU_ORDER_TARGET_NORM[key] || String(name).trim();
        const target    = Number(skuTargetRow[i]) || 0;
        if (target > 0) skuOrderTargets[canonical] = (skuOrderTargets[canonical] || 0) + target;
      });
    }

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

    // ── Live reels payout ────────────────────────────────────────────────────
    // Reel Amount        : rows with a social live link (Instagram / YouTube)
    // Live Product Amount: rows WITH a social live link (product sent + reel posted)
    // Advance Product Amt: rows WITHOUT any live link (product sent, reel not yet posted)
    let liveReelsReelAmt      = 0;
    let liveReelsProductAmt   = 0;   // product amount for rows that have gone live
    let advanceProductAmt     = 0;   // product amount for rows not yet live
    rows.forEach(r => {
      const link     = String(r[idx.liveLink] || '').trim();
      const isSocial = link.includes('instagram.com') || link.includes('youtube.com') || link.includes('youtu.be');
      const prodAmt  = parseAmount(r[idx.productAmt]);
      if (isSocial) {
        liveReelsReelAmt    += parseAmount(r[idx.reelAmt]);
        liveReelsProductAmt += prodAmt;
      } else if (prodAmt > 0) {
        advanceProductAmt   += prodAmt;   // product sent but no live link yet
      }
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

    // Count rows with Instagram or YouTube Shorts links as "live reels".
    // Amazon affiliate links and other non-social URLs in the Live Link column are excluded.
    const isLiveSocial = url => {
      const u = String(url || '').trim();
      return u.includes('instagram.com') || u.includes('youtube.com') || u.includes('youtu.be');
    };
    const isLive = inf => inf.liveLink && isLiveSocial(inf.liveLink);

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
    // Reel live targets: Agency 660 + In-house 450 = 1110
    const agencyReelTarget   = 660;
    const inhouseReelTarget  = 450;
    const totalReelTarget    = agencyReelTarget + inhouseReelTarget;
    // Order targets (from Overview): Mohit 300 + Satyam 200 + Hardev 200 + Priyanka 100 = 800 (in-house)
    const inhouseOrderTarget = Object.values(teamTargets).reduce((s, v) => s + v, 0);
    // Overall SKU order target = sum of product-wise targets (1020)
    const overallOrderTarget = Object.values(skuOrderTargets).reduce((s, v) => s + v, 0) || 1020;
    const agencyOrderTarget  = overallOrderTarget - inhouseOrderTarget;
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
        advanceProductAmt,
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
      pocSkuOrders,
      skuOrderTargets,
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
    res.set('Cache-Control', 'no-store');
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

    res.set('Cache-Control', 'no-store');
    if (status !== 'SUCCEEDED') return res.json({ status: status || 'UNKNOWN' });

    // Run finished — fetch dataset items
    const datasetId = run.data.defaultDatasetId;
    const itemsRes  = await fetch(
      `https://api.apify.com/v2/datasets/${datasetId}/items?token=${token}&clean=true`
    );
    const items = await itemsRes.json();

    const metrics = (Array.isArray(items) ? items : []).map(it => {
      const rawUrl = (it.url || it.inputUrl || '');
      // Normalise to shortcode key — Apify returns /p/ even for /reel/ originals
      const scMatch = rawUrl.match(/instagram\.com\/(?:p|reel|tv)\/([A-Za-z0-9_-]+)/);
      return {
        url:      scMatch ? `ig:${scMatch[1]}` : rawUrl.split('?')[0].replace(/\/$/, ''),
        views:    it.videoPlayCount ?? it.playCount ?? it.videoViewCount ?? 0,
        likes:    it.likesCount     ?? it.likeCount  ?? 0,
        comments: it.commentsCount  ?? it.commentCount ?? 0,
      };
    });

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
