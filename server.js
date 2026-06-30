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
// Month name → 2-digit number
const MONTH_NAME_TO_NUM = {
  january:'01',february:'02',march:'03',april:'04',may:'05',june:'06',
  july:'07',august:'08',september:'09',october:'10',november:'11',december:'12',
};

function parseAgencySheet(values, sourceName) {
  if (!values || values.length < 2) return [];
  const headers = values[0];
  const rows    = values.slice(1);

  const col         = kw => headers.findIndex(h => String(h).trim().toUpperCase() === kw.toUpperCase());
  const colIncludes = kw => headers.findIndex(h => String(h).trim().toUpperCase().includes(kw.toUpperCase()));

  const idx = {
    closingMonth: 0,               // always col 0 — "Closing Month"
    username:   col('USERNAME'),
    profileLink:col('PROFILE LINK'),
    brand:      col('BRAND NAME'),
    product:    col('PRODUCT NAME'),
    liveLink:   colIncludes('LIVE LINK'),
    liveDate:   colIncludes('DATE'),
    productAmt: col('PRODUCT AMOUNT'),
    reelAmt:    col('REEL AMOUNT'),
    views:      col('VIEWS'),
    likes:      col('LIKES'),
    comments:   colIncludes('COMMENT'),
    cpv:        col('CPV'),
    language:   col('LANGUAGE'),
  };

  // Forward-fill "Closing Month" — Excel merged cells show value only in first row
  let lastClosingMonth = null;
  const enriched = rows.map(r => {
    const raw = String(r[idx.closingMonth] || '').trim().toLowerCase();
    if (MONTH_NAME_TO_NUM[raw]) lastClosingMonth = raw;
    return { r, closingMonthName: lastClosingMonth };
  });

  return enriched
    .filter(({ r }) => r[idx.username] && String(r[idx.username]).trim())
    .map(({ r, closingMonthName }) => {
      // Parse live date
      let liveDate = null;
      if (idx.liveDate >= 0 && r[idx.liveDate]) {
        const raw = r[idx.liveDate];
        if (typeof raw === 'number') {
          liveDate = excelDateToISO(raw);
        } else {
          const s = String(raw).trim();
          if (/^\d{4}-\d{2}-\d{2}$/.test(s))      liveDate = s;
          else if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) { const [d,m,y]=s.split('/'); liveDate=`${y}-${m}-${d}`; }
          else if (s)                                liveDate = s;
        }
      }

      // Derive YYYY-MM from closing month name (use current year)
      const closingMonthYM = closingMonthName
        ? `${new Date().getFullYear()}-${MONTH_NAME_TO_NUM[closingMonthName]}`
        : null;

      return {
        source:        sourceName,
        poc:           'Agency',
        agency:        sourceName,
        username:      String(r[idx.username] || '').trim().replace(/\/$/, ''),
        profileLink:   r[idx.profileLink] || '',
        brand:         normBrand(r[idx.brand]),
        product:       normAgencyProduct(r[idx.product]),
        liveLink:      r[idx.liveLink] || '',
        liveDate,
        closingMonthYM,              // "2026-04", "2026-05", etc. from "Closing Month" column
        productAmt:    parseAmount(r[idx.productAmt]),
        reelAmt:       parseAmount(r[idx.reelAmt]),
        views:         Number(r[idx.views])    || 0,
        likes:         Number(r[idx.likes])    || 0,
        comments:      Number(r[idx.comments]) || 0,
        cpv:           Number(r[idx.cpv])      || 0,
        language:      String(r[idx.language] || '').trim(),
        total:         parseAmount(r[idx.reelAmt]) + parseAmount(r[idx.productAmt]),
      };
    });
}

// ── Count orders from an individual IS sheet ─────────────────────────────────
// Convert Excel date serial → JS Date (accounts for Excel's 1900 leap-year bug)
const excelSerialToDate = serial => {
  if (!serial || typeof serial !== 'number' || serial < 1) return null;
  return new Date(Math.round((serial - 25569) * 86400000));
};
// Returns true if a date serial falls in the given YYYY-MM (e.g. "2026-05")
const isInTargetMonth = (serial, targetYM) => {
  if (!targetYM) return true;
  const d = excelSerialToDate(serial);
  if (!d) return false;
  const [y, m] = targetYM.split('-').map(Number);
  return d.getUTCFullYear() === y && d.getUTCMonth() === m - 1;
};

// Count orders matching status values, filtered to a target month (YYYY-MM)
function countOrdersFromSheet(values, statusValues, targetYM) {
  if (!values || values.length < 2) return 0;
  const headers     = values[0];
  const rows        = values.slice(1);
  const statusIdx   = headers.findIndex(h => /ORDER STATUS|CREATOR STATUS/i.test(String(h)));
  const usernameIdx = headers.findIndex(h => /USER.?NAME|USERNAME/i.test(String(h)));
  const productIdx  = headers.findIndex(h => /PRODUCT.?NAME/i.test(String(h)));
  const dateIdx     = headers.findIndex(h => /ORDER.?DATE/i.test(String(h)));
  if (statusIdx < 0 || usernameIdx < 0) return 0;
  // Deduplicate by (username + product) to avoid double-counting when a row
  // progresses from "Order Placed" → "Live" and both rows exist in the sheet.
  const seen = new Set();
  let count = 0;
  rows.forEach(r => {
    const user = String(r[usernameIdx] || '').trim();
    if (!user || user === '#N/A') return;
    if (!statusValues.some(sv => String(r[statusIdx]).trim().toLowerCase() === sv.toLowerCase())) return;
    if (dateIdx >= 0 && (!r[dateIdx] || !isInTargetMonth(r[dateIdx], targetYM))) return;
    const product = productIdx >= 0 ? String(r[productIdx] || '').trim().toLowerCase() : '';
    const key = `${user.toLowerCase()}||${product}`;
    if (seen.has(key)) return;
    seen.add(key);
    count++;
  });
  return count;
}

// Count orders by SKU, filtered to a target month (YYYY-MM)
function countOrdersBySkuFromSheet(values, statusValues, targetYM) {
  if (!values || values.length < 2) return {};
  const headers     = values[0];
  const rows        = values.slice(1);
  const statusIdx   = headers.findIndex(h => /ORDER STATUS|CREATOR STATUS/i.test(String(h)));
  const usernameIdx = headers.findIndex(h => /USER.?NAME|USERNAME/i.test(String(h)));
  const productIdx  = headers.findIndex(h => /PRODUCT.?NAME/i.test(String(h)));
  const dateIdx     = headers.findIndex(h => /ORDER.?DATE/i.test(String(h)));
  if (statusIdx < 0 || usernameIdx < 0 || productIdx < 0) return {};
  const skuMap = {};
  rows.forEach(r => {
    if (!r[usernameIdx] || !String(r[usernameIdx]).trim() || String(r[usernameIdx]).trim() === '#N/A') return;
    if (!statusValues.some(sv => String(r[statusIdx]).trim().toLowerCase() === sv.toLowerCase())) return;
    if (dateIdx >= 0 && (!r[dateIdx] || !isInTargetMonth(r[dateIdx], targetYM))) return;
    const raw = String(r[productIdx] || '').trim();
    const sku = normAgencyProduct(raw) || raw || 'Unknown';
    skuMap[sku] = (skuMap[sku] || 0) + 1;
  });
  return skuMap;
}

// ── Parse overall reel targets from a targets tab ─────────────────────────────
// Row 3=In-house, Row 4=TTC, Row 5=Ink Revenue; col 1 = overall total
const parseTargetTabOverall = values => {
  if (!values || values.length < 4) return null;
  const inhouseReelTarget = Number((values[3] || [])[1]) || 0;
  const agencyTtcTarget   = Number((values[4] || [])[1]) || 0;
  const agencyInkTarget   = Number((values[5] || [])[1]) || 0;
  const agencyReelTarget  = agencyTtcTarget + agencyInkTarget;
  const totalReelTarget   = inhouseReelTarget + agencyReelTarget;
  if (!inhouseReelTarget && !agencyReelTarget) return null;
  return { inhouseReelTarget, agencyTtcTarget, agencyInkTarget, agencyReelTarget, totalReelTarget };
};

// ── Parse order targets from a targets tab ────────────────────────────────────
const ORDER_TARGET_ABBR_MAP = {
  'me':'Magic Eraser','magic eraser':'Magic Eraser',
  'dp':'Washing Machine Powder','descale powder':'Washing Machine Powder',
  'tc+bc':'Kleenest Bathroom Kit','tc+bc kit':'Kleenest Bathroom Kit','tc + bc kit':'Kleenest Bathroom Kit',
  'trial':'Cleaning Trial Kit','trial kit':'Cleaning Trial Kit',
  'mc':'Metal Cleaner Kit','metal cleaner':'Metal Cleaner Kit',
  'kc':'Kitchen Cleaner','kitchen cleaner':'Kitchen Cleaner',
  'dl':'Washing Machine Liquid','descale liquid':'Washing Machine Liquid',
  'dt':'Washing Machine Tablet','descale tablet':'Washing Machine Tablet',
  'klenzmo tile':'Tile & Floor Cleaner','klenzmo floor & tile cleaner':'Tile & Floor Cleaner',
};
function parseTargetTabOrders(tabValues) {
  if (!tabValues || tabValues.length < 2) return { skuOrderTargets:{}, teamTargets:{}, inhouseOrderTarget:0, overallOrderTarget:0 };

  // SKU order targets — "Overall Product-wise" section
  const skuOrderTargets = {};
  for (let ri = 0; ri < tabValues.length; ri++) {
    const cell0 = String(tabValues[ri][0] || '').trim().toLowerCase();
    if (cell0.includes('overall product-wise') || cell0.includes('product-wise order')) {
      const headerRow = tabValues[ri + 1] || [];
      const valRow    = tabValues[ri + 2] || [];
      headerRow.forEach((h, i) => {
        if (!h) return;
        const key = String(h).trim().toLowerCase();
        // Skip summary/total columns — not individual SKUs
        if (!key || key.includes('total') || key.includes('grand')) return;
        const canonical = ORDER_TARGET_ABBR_MAP[key] || String(h).trim();
        const val = Number(valRow[i]) || 0;
        if (val > 0) skuOrderTargets[canonical] = (skuOrderTargets[canonical] || 0) + val;
      });
      break;
    }
  }

  // Team POC targets — "Orders Placed - In-house" section
  const teamTargets = {};
  let inSection = false;
  tabValues.forEach(row => {
    const label = String(row[0] || '').trim();
    if (label.toLowerCase().includes('orders placed - in-house')) { inSection = true; return; }
    if (inSection) {
      const name = normPOC(label);
      const num  = Number(row[1]);
      if (['Mohit', 'Hardev', 'Satyam', 'Priyanka'].includes(name) && num > 0) teamTargets[name] = num;
    }
  });

  // overallOrderTarget = sum of per-team targets (the "Orders Placed - In-house Targets" section)
  // This gives April=800, May=1100 — NOT the SKU product-wise rows which are a separate planning view
  const inhouseOrderTarget = Object.values(teamTargets).reduce((s, v) => s + v, 0);
  const overallOrderTarget = inhouseOrderTarget;
  return { skuOrderTargets, teamTargets, inhouseOrderTarget, overallOrderTarget };
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

// ── /api/debug/harmeet-drive — list harmeet's drive + resolve share URL ───────
app.get('/api/debug/harmeet-drive', async (req, res) => {
  try {
    const hmBase   = `/users/${HARMEET_USER}/drive`;
    const shareUrl = 'https://velocityeventures-my.sharepoint.com/:x:/g/personal/harmeet_kleenest_in/IQDWHJBLfmyQSJ9G9EHY-hlZAaSEKDBqRIT9r0Eajk7NRjg?e=Tavl3t';
    const encoded  = 'u!' + Buffer.from(shareUrl).toString('base64').replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_');
    const [files, shareItem] = await Promise.all([
      graphGet(`${hmBase}/root/children`),
      graphGet(`/shares/${encoded}/driveItem`).catch(e => ({ error: e.message })),
    ]);
    res.json({
      shareItem: { id: shareItem.id, name: shareItem.name, error: shareItem.error },
      driveFiles: (files.value || []).map(f => ({ id: f.id, name: f.name })),
    });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── /api/debug/harmeet-may — peek at the May targets file tabs + first rows ───
app.get('/api/debug/harmeet-may', async (req, res) => {
  try {
    const fileId = req.query.id;
    if (!fileId) return res.status(400).json({ error: 'pass ?id=FILE_ID' });
    const base = `/users/${HARMEET_USER}/drive/items/${fileId}/workbook/worksheets`;
    const sheets = await graphGet(base);
    const tabs   = (sheets.value || []).map(s => s.name);
    // Peek at each tab (first 10 rows)
    const peeks  = {};
    for (const tab of tabs.slice(0, 10)) {
      const r = await graphGet(`${base}('${encodeURIComponent(tab)}')/usedRange`).catch(() => ({ values: [] }));
      peeks[tab] = (r.values || []).slice(0, 8);
    }
    res.json({ tabs, peeks });
  } catch (err) { res.status(500).json({ error: err.message }); }
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

// ── /api/debug/orders-april — detailed April order counts per POC ─────────────
app.get('/api/debug/orders-april', async (req, res) => {
  try {
    const hmBase         = `/users/${HARMEET_USER}/drive/items`;
    const priyankaTfcBase = `/users/${PRIYANKA_USER}/drive/items/${PRIYANKA_TFC_ID}/workbook/worksheets`;
    const safeGet = p => graphGet(p).catch(() => ({ values: [] }));
    const YM = '2026-04';
    const ORDER_STATUSES = ['order place', 'order placed', 'live'];

    const [mohit, hardev, satyamMay, satyamApr, priyanka] = await Promise.all([
      safeGet(`${hmBase}/${MOHIT_FILE_ID}/workbook/worksheets('Main%20Sheet')/usedRange`),
      safeGet(`${hmBase}/${HARDEV_FILE_ID}/workbook/worksheets('Main%20Sheet')/usedRange`),
      safeGet(`${hmBase}/${SATYAM_FILE_ID}/workbook/worksheets('May%20master%20sheet')/usedRange`),
      safeGet(`${hmBase}/${SATYAM_FILE_ID}/workbook/worksheets('April%20master%20sheet')/usedRange`),
      safeGet(`${priyankaTfcBase}('Sheet1')/usedRange`),
    ]);

    function detailedCount(values, label) {
      if (!values || values.length < 2) return { label, error: 'no data', rawCount: 0, dedupCount: 0 };
      const headers     = values[0];
      const rows        = values.slice(1);
      const statusIdx   = headers.findIndex(h => /ORDER STATUS|CREATOR STATUS/i.test(String(h)));
      const usernameIdx = headers.findIndex(h => /USER.?NAME|USERNAME/i.test(String(h)));
      const productIdx  = headers.findIndex(h => /PRODUCT.?NAME/i.test(String(h)));
      const dateIdx     = headers.findIndex(h => /ORDER.?DATE/i.test(String(h)));
      const allStatuses = [...new Set(rows.map(r => String(r[statusIdx]||'').trim()).filter(Boolean))];

      // All rows in target month with valid status
      const aprilRows = rows.filter(r => {
        const user = String(r[usernameIdx] || '').trim();
        if (!user || user === '#N/A') return false;
        if (!ORDER_STATUSES.some(sv => String(r[statusIdx]).trim().toLowerCase() === sv.toLowerCase())) return false;
        if (dateIdx >= 0) return r[dateIdx] && isInTargetMonth(r[dateIdx], YM);
        return true;
      });

      // Dedup by username only
      const uniqueByUser = new Set(aprilRows.map(r => String(r[usernameIdx]).trim().toLowerCase()));

      // Dedup by (username + product)
      const seenUserProduct = new Set();
      aprilRows.forEach(r => {
        const user = String(r[usernameIdx] || '').trim().toLowerCase();
        const product = productIdx >= 0 ? String(r[productIdx] || '').trim().toLowerCase() : '';
        seenUserProduct.add(`${user}||${product}`);
      });

      // Per-user breakdown: show users with multiple rows + their products
      const userProductMap = {};
      aprilRows.forEach(r => {
        const user    = String(r[usernameIdx] || '').trim().toLowerCase();
        const product = productIdx >= 0 ? String(r[productIdx] || '').trim() : 'N/A';
        const status  = String(r[statusIdx] || '').trim();
        if (!userProductMap[user]) userProductMap[user] = [];
        userProductMap[user].push({ product, status });
      });
      const multiRowUsers = Object.entries(userProductMap)
        .filter(([, entries]) => entries.length > 1)
        .slice(0, 15)
        .map(([user, entries]) => ({ user, entries }));

      return {
        label,
        totalRows: rows.length,
        rawCount: aprilRows.length,            // all matching rows (no dedup)
        dedupByUser: uniqueByUser.size,        // unique usernames
        dedupByUserProduct: seenUserProduct.size, // unique (username+product) — what dashboard uses
        headers: headers.slice(0, 15),
        statusColIdx: statusIdx, productColIdx: productIdx, dateColIdx: dateIdx,
        uniqueStatuses: allStatuses.slice(0, 20),
        multiRowUsers,  // users appearing more than once — shows their products/statuses
      };
    }

    res.json({
      mohit:     detailedCount(mohit.values,     'Mohit'),
      hardev:    detailedCount(hardev.values,     'Hardev'),
      satyamMay: detailedCount(satyamMay.values, 'Satyam-MaySheet'),
      satyamApr: detailedCount(satyamApr.values, 'Satyam-AprilSheet'),
      priyanka:  detailedCount(priyanka.values,  'Priyanka'),
    });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── /api/debug/priyanka-months — inspect Priyanka TFC Closing Month values ────
app.get('/api/debug/priyanka-months', async (req, res) => {
  try {
    const priyankaTfcBase = `/users/${PRIYANKA_USER}/drive/items/${PRIYANKA_TFC_ID}/workbook/worksheets`;
    const raw = await graphGet(`${priyankaTfcBase}('Sheet1')/usedRange`).catch(() => ({ values: [] }));
    const values = raw.values || [];
    if (values.length < 2) return res.json({ error: 'no data' });
    const headers = values[0];
    const rows = values.slice(1);
    const closingMonthIdx = headers.findIndex(h => /CLOSING.?MONTH/i.test(String(h)));
    const usernameIdx     = headers.findIndex(h => /USER.?NAME|USERNAME/i.test(String(h)));
    const statusIdx       = headers.findIndex(h => /ORDER STATUS|CREATOR STATUS/i.test(String(h)));
    const uniqueMonths = [...new Set(rows.map(r => String(r[closingMonthIdx]||'').trim()).filter(Boolean))];
    const sample = rows.slice(0, 10).map(r => ({
      closingMonth: r[closingMonthIdx],
      username: r[usernameIdx],
      status: r[statusIdx],
    }));
    res.json({ closingMonthIdx, uniqueMonths, sample, totalRows: rows.length });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── /api/debug/targets-raw — inspect raw April/May target tab structures ─────
app.get('/api/debug/targets-raw', async (req, res) => {
  try {
    const hmBase = `/users/${HARMEET_USER}/drive/items`;
    const safeGet = p => graphGet(p).catch(() => ({ values: [] }));
    const [aprilRaw, mayRaw] = await Promise.all([
      safeGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('April%20Targets')/usedRange`),
      safeGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('May%20Targets')/usedRange`),
    ]);
    const trim = raw => (raw.values || []).map((r, i) => ({ row: i, cells: r.slice(0, 8) })).filter(r => r.cells.some(c => String(c).trim()));
    res.json({ april: trim(aprilRaw), may: trim(mayRaw) });
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

    // Satyam keeps separate tabs per month; fetch both so we can count each month accurately
    const fetchSatyamTab = async () => {
      for (const name of ['May master sheet', 'April master sheet', 'Master sheet', 'Main Sheet']) {
        const r = await safeGet(`${hmBase}/${SATYAM_FILE_ID}/workbook/worksheets('${encodeURIComponent(name)}')/usedRange`);
        if (r.values && r.values.length > 1) return r;
      }
      return { values: [] };
    };

    const agBase         = `/users/${PRIYANKA_USER}/drive/items/${AGENCY_FILE_ID}/workbook/worksheets`;
    const priyankaTfcBase = `/users/${PRIYANKA_USER}/drive/items/${PRIYANKA_TFC_ID}/workbook/worksheets`;

    // Fetch April + May target tabs in parallel (both needed for month-aware targets)
    const [
      liveRaw, targetRaw,
      mohitRaw, hardevRaw, satyamRaw, satyamAprilRaw,
      aprilTargetRaw, mayTargetRaw, ttcRaw, inkRaw, priyankaTfcRaw,
    ] = await Promise.all([
      fetchLiveTab(),
      safeGet(`${prBase}('Targets%20%3D%20P%20wise')/usedRange`),
      safeGet(`${hmBase}/${MOHIT_FILE_ID}/workbook/worksheets('Main%20Sheet')/usedRange`),
      safeGet(`${hmBase}/${HARDEV_FILE_ID}/workbook/worksheets('Main%20Sheet')/usedRange`),
      fetchSatyamTab(),                                // May (or first available) Satyam sheet
      safeGet(`${hmBase}/${SATYAM_FILE_ID}/workbook/worksheets('April%20master%20sheet')/usedRange`),
      safeGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('April%20Targets')/usedRange`),
      safeGet(`${hmBase}/${APRIL_PLAN_ID}/workbook/worksheets('May%20Targets')/usedRange`),
      safeGet(`${agBase}('TTC')/usedRange`),
      safeGet(`${agBase}('INK%20REVENUE')/usedRange`),
      safeGet(`${priyankaTfcBase}('Sheet1')/usedRange`),
    ]);

    // ── Derive target month: prefer May if it has data, fall back to April ───
    const MONTH_NUMS = { January:'01',February:'02',March:'03',April:'04',May:'05',
                         June:'06',July:'07',August:'08',September:'09',
                         October:'10',November:'11',December:'12' };
    const mayHasData   = (mayTargetRaw.values   || []).length > 5;
    const aprilHasData = (aprilTargetRaw.values || []).length > 5;
    const _targetTabName = mayHasData ? 'May Targets' : aprilHasData ? 'April Targets' : '';
    const targetTabRaw   = mayHasData ? mayTargetRaw : aprilHasData ? aprilTargetRaw : { values: [] };
    const _monthName     = _targetTabName.replace(' Targets', '').trim();
    const _monthNum      = MONTH_NUMS[_monthName];
    const targetMonth    = _monthNum ? `${new Date().getFullYear()}-${_monthNum}` : null;

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

    // ── Helper: extract YYYY-MM from a raw row's liveDate cell ──────────────
    const rowMonth = r => {
      const raw = r[idx.liveDate];
      if (!raw) return null;
      if (typeof raw === 'number') { const iso = excelDateToISO(raw); return iso ? iso.slice(0,7) : null; }
      const s = String(raw).trim();
      const m = s.match(/(\d{4})-(\d{2})/);
      return m ? m[0] : null;
    };

    // ── Reels gone live — filtered to target month if known ──────────────────
    const reelsLive = rows.filter(r => {
      const link = String(r[idx.liveLink] || '').trim();
      const isSocial = link.includes('instagram.com') || link.includes('youtube.com') || link.includes('youtu.be');
      if (!isSocial) return false;
      if (targetMonth && rowMonth(r) !== targetMonth) return false;
      return true;
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

    // ── Parse Targets Tab (active month) ────────────────────────────────────
    const targetTabValues = targetTabRaw.values || [];

    // ── Parse order + reel targets for both April and May ────────────────────
    const _year      = new Date().getFullYear();
    const _aprOrders = parseTargetTabOrders(aprilTargetRaw.values);
    const _mayOrders = parseTargetTabOrders(mayTargetRaw.values);
    const _aprReel   = parseTargetTabOverall(aprilTargetRaw.values);
    const _mayReel   = parseTargetTabOverall(mayTargetRaw.values);

    // Active-month targets (drives server-side KPIs)
    const _activeOrders  = mayHasData ? _mayOrders : aprilHasData ? _aprOrders : { skuOrderTargets:{}, teamTargets:{}, inhouseOrderTarget:0, overallOrderTarget:0 };
    const orderTargets   = _activeOrders.skuOrderTargets;
    const teamTargets    = Object.keys(_activeOrders.teamTargets).length
      ? _activeOrders.teamTargets
      : { Mohit: 300, Satyam: 200, Hardev: 200, Priyanka: 100 };

    // ── Helper: build pocOrders + pocSkuOrders for a given month ────────────
    const ORDER_STATUSES = ['order place', 'order placed', 'live'];
    const _satyamFor = ym => {
      // Use April-specific sheet for April; fall back to main sheet otherwise
      const aprSheet = (satyamAprilRaw.values || []).length > 1 ? satyamAprilRaw.values : null;
      return ym === `${_year}-04` && aprSheet ? aprSheet : satyamRaw.values;
    };
    const _pocOrdersFor = ym => ({
      Mohit:    countOrdersFromSheet(mohitRaw.values,       ORDER_STATUSES, ym),
      Hardev:   countOrdersFromSheet(hardevRaw.values,      ORDER_STATUSES, ym),
      Satyam:   countOrdersFromSheet(_satyamFor(ym),        ORDER_STATUSES, ym),
      Priyanka: countOrdersFromSheet(priyankaTfcRaw.values, ORDER_STATUSES, ym),
    });
    const _pocSkuOrdersFor = ym => ({
      Mohit:    countOrdersBySkuFromSheet(mohitRaw.values,       ORDER_STATUSES, ym),
      Hardev:   countOrdersBySkuFromSheet(hardevRaw.values,      ORDER_STATUSES, ym),
      Satyam:   countOrdersBySkuFromSheet(_satyamFor(ym),        ORDER_STATUSES, ym),
      Priyanka: countOrdersBySkuFromSheet(priyankaTfcRaw.values, ORDER_STATUSES, ym),
    });

    // Orders for the active target month (used in server KPIs)
    const pocOrders    = _pocOrdersFor(targetMonth);
    const pocSkuOrders = _pocSkuOrdersFor(targetMonth);
    const skuOrderTargets = { ...orderTargets };

    // ── Per-month targets map (sent to client for month-aware display) ────────
    const monthlyTargets = {};
    if (_aprReel) monthlyTargets[`${_year}-04`] = {
      ..._aprReel, ..._aprOrders,
      pocOrders:    _pocOrdersFor(`${_year}-04`),
      pocSkuOrders: _pocSkuOrdersFor(`${_year}-04`),
    };
    if (_mayReel) monthlyTargets[`${_year}-05`] = {
      ..._mayReel, ..._mayOrders,
      pocOrders:    _pocOrdersFor(`${_year}-05`),
      pocSkuOrders: _pocSkuOrdersFor(`${_year}-05`),
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
      const inMonth  = !targetMonth || rowMonth(r) === targetMonth;
      if (isSocial && inMonth) {
        liveReelsReelAmt    += parseAmount(r[idx.reelAmt]);
        liveReelsProductAmt += prodAmt;
      } else if (prodAmt > 0 && !isSocial && inMonth) {
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

    // ── Parse SKU Live Targets from targets tab ──────────────────────────────
    // targetTabValues[2]: product names starting at col 3
    // targetTabValues[3]: In-house targets per SKU + total at col 1
    // targetTabValues[4]: TTC Agency Go-Live per SKU + total at col 1
    // targetTabValues[5]: Ink Revenue Agency Go-Live per SKU + total at col 1
    const ttcSkuLiveTargets   = {};
    const inkSkuLiveTargets   = {};
    const agencySkuLiveTargets= {};
    const inhouseSkuLiveTargets={};
    if (targetTabValues.length >= 6) {
      const productHeader = targetTabValues[2] || [];
      const inhouseRow    = targetTabValues[3] || [];
      const ttcRow        = targetTabValues[4] || [];
      const inkRow        = targetTabValues[5] || [];
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
    // All agency rows across all months — sent to client for table display / month filter browsing
    const _allAgencyInfluencers = [...ttcInfluencers, ...inkInfluencers];

    // Count rows with Instagram or YouTube Shorts links as "live reels".
    // Amazon affiliate links and other non-social URLs in the Live Link column are excluded.
    const isLiveSocial = url => {
      const u = String(url || '').trim();
      return u.includes('instagram.com') || u.includes('youtube.com') || u.includes('youtu.be');
    };
    const isLive = inf => inf.liveLink && isLiveSocial(inf.liveLink);

    // Filter agency row to target month: match closingMonthYM, falling back to liveDate
    const isAgencyInMonth = inf => {
      if (!targetMonth) return true;                              // no month set → include all
      const ym = inf.closingMonthYM || inf.liveDate?.slice(0, 7);
      return ym === targetMonth;
    };
    const isLiveInMonth = inf => isLive(inf) && isAgencyInMonth(inf);

    // Send ALL months to client — client-side month filter handles per-month browsing
    const agencyInfluencers = _allAgencyInfluencers;

    // Per-agency live counts (target-month filtered)
    const ttcReelsLive    = ttcInfluencers.filter(isLiveInMonth).length;
    const inkReelsLive    = inkInfluencers.filter(isLiveInMonth).length;
    const agencyReelsLive = ttcReelsLive + inkReelsLive;

    // Per-agency SKU actuals (target-month filtered)
    const ttcSkuActuals    = {};
    const inkSkuActuals    = {};
    const agencySkuActuals = {};
    const addSku = (map, inf) => {
      const p = inf.product; if (!p) return;
      map[p] = (map[p] || 0) + 1;
      agencySkuActuals[p] = (agencySkuActuals[p] || 0) + 1;
    };
    ttcInfluencers.forEach(inf => { if (isLiveInMonth(inf)) addSku(ttcSkuActuals, inf); });
    inkInfluencers.forEach(inf => { if (isLiveInMonth(inf)) addSku(inkSkuActuals, inf); });

    // Per-agency payout (target-month filtered)
    let ttcReelsReelAmt = 0, ttcReelsProductAmt = 0;
    let inkReelsReelAmt = 0, inkReelsProductAmt = 0;
    ttcInfluencers.forEach(inf => { if (!isLiveInMonth(inf)) return; ttcReelsReelAmt += inf.reelAmt; ttcReelsProductAmt += inf.productAmt; });
    inkInfluencers.forEach(inf => { if (!isLiveInMonth(inf)) return; inkReelsReelAmt += inf.reelAmt; inkReelsProductAmt += inf.productAmt; });
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
    // Reel live targets — use active parsed tab (fallback hardcodes)
    const _activeReel       = mayHasData ? _mayReel : aprilHasData ? _aprReel : null;
    const inhouseReelTarget = _activeReel?.inhouseReelTarget || 525;
    const agencyTtcTarget   = _activeReel?.agencyTtcTarget   || 0;
    const agencyInkTarget   = _activeReel?.agencyInkTarget   || 0;
    const agencyReelTarget  = _activeReel?.agencyReelTarget  || 710;
    const totalReelTarget   = inhouseReelTarget + agencyReelTarget;
    // Order targets — use parsed helper (no more double-counting "Total Orders Placed")
    const inhouseOrderTarget = _activeOrders.inhouseOrderTarget || Object.values(teamTargets).reduce((s, v) => s + v, 0);
    const overallOrderTarget = _activeOrders.overallOrderTarget || 1020;
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
      monthlyTargets,
      overallTargets: { reelTarget: overallReelTarget, agencyReelTarget, inhouseReelTarget, totalReelTarget, orderTarget: overallOrderTarget, agencyOrderTarget, inhouseOrderTarget, actualReels: overallActualReels, totalOrders },
      teamTargets,
      pocOrders,
      pocSkuOrders,
      skuOrderTargets,
      influencers,
      targetMonth,
      lastUpdated:  new Date().toISOString(),
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── Graph PATCH helper (write to Excel) ──────────────────────────────────────
async function graphPatch(path, body) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  return res.json();
}

function colLetter(index) {
  let col = '', n = index + 1;
  while (n > 0) { const rem = (n - 1) % 26; col = String.fromCharCode(65 + rem) + col; n = Math.floor((n - 1) / 26); }
  return col;
}

// ── GET /api/metrics/empty-rows — rows with Instagram links but no VIEWS ───────
// Scans: Live Inhouse (in-house), TTC + INK REVENUE (agency)
// Query: ?limit=N (default 50, max 200)  ?type=inhouse|agency (default: both)  ?month=YYYY-MM
// Returns: { rows: [{excelRow, username, url, viewsCol, likesCol, commentsCol, fileId, sheetName}], totalEmpty, inhouseEmpty, agencyEmpty }
app.get('/api/metrics/empty-rows', async (req, res) => {
  const limit     = Math.min(parseInt(req.query.limit) || 50, 200);
  const type      = req.query.type  || 'all';  // 'inhouse' | 'agency' | 'all'
  const monthFilter = req.query.month || '';   // 'YYYY-MM' or '' for all

  // Scan one sheet and append qualifying rows to `out`; increment total.count
  async function scanSheet(fileId, sheetName, out, totalRef) {
    const enc  = encodeURIComponent(sheetName);
    const base = `/users/${PRIYANKA_USER}/drive/items/${fileId}/workbook/worksheets`;
    let data;
    try { data = await graphGet(`${base}('${enc}')/usedRange`); } catch { return; }
    const values = data.values || [];
    if (values.length < 2) return;

    const headers  = values[0];
    const col      = kw => headers.findIndex(h => String(h).trim().toUpperCase() === kw.toUpperCase());
    const colInc   = kw => headers.findIndex(h => String(h).trim().toUpperCase().includes(kw.toUpperCase()));

    const idx = {
      username: col('USERNAME'),
      liveLink: colInc('LIVE LINK'),
      views:    col('VIEWS'),
      likes:    col('LIKES'),
      comments: colInc('COMMENT'),
      liveDate: colInc('LIVE DATE'),
    };
    if (idx.liveLink < 0) return; // sheet layout doesn't match — skip silently

    const viewsCol    = colLetter(idx.views);
    const likesCol    = colLetter(idx.likes);
    const commentsCol = colLetter(idx.comments);

    for (let i = 1; i < values.length; i++) {
      const row  = values[i];
      const link = String(row[idx.liveLink] || '').trim();
      if (!link.includes('instagram.com')) continue;
      // Skip bare domain URLs with no reel/post path — Apify rejects them
      const igPath = link.replace(/^https?:\/\/(www\.)?instagram\.com\/?/, '').split('?')[0].trim();
      if (!igPath || igPath.length < 3) continue;
      const views = row[idx.views];
      if (views && Number(views) > 0) continue;

      // Month filter — skip rows whose live date doesn't match selected month
      if (monthFilter && idx.liveDate >= 0) {
        const dateVal = row[idx.liveDate];
        const d = excelSerialToDate(dateVal);
        if (!d) continue; // no date — skip when month filter is active
        const rowYM = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
        if (rowYM !== monthFilter) continue;
      }

      totalRef.count++;
      if (out.length < limit) {
        out.push({
          excelRow:  i + 1,
          username:  String(row[idx.username] || '').trim(),
          url:       link.split('?')[0].replace(/\/$/, ''),
          viewsCol, likesCol, commentsCol,
          fileId, sheetName,
        });
      }
    }
  }

  try {
    const rows    = [];
    const inTotal = { count: 0 };
    const agTotal = { count: 0 };

    if (type !== 'agency') await scanSheet(LIVE_FILE_ID,   'Live Inhouse', rows, inTotal);
    if (type !== 'inhouse') {
      await scanSheet(AGENCY_FILE_ID, 'TTC',         rows, agTotal);
      await scanSheet(AGENCY_FILE_ID, 'INK REVENUE', rows, agTotal);
    }

    res.json({
      rows,
      totalEmpty:   inTotal.count + agTotal.count,
      inhouseEmpty: inTotal.count,
      agencyEmpty:  agTotal.count,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── POST /api/metrics/write-sheet — write scraped metrics to Excel cells ───────
// Body: { updates: [{ excelRow, views, likes, comments, viewsCol, likesCol, commentsCol, fileId, sheetName }] }
// fileId / sheetName default to LIVE_FILE_ID / 'Live Inhouse' for backward compat
app.post('/api/metrics/write-sheet', async (req, res) => {
  const { updates } = req.body || {};
  if (!Array.isArray(updates) || !updates.length) return res.json({ written: 0 });

  let written = 0;
  const errors = [];
  for (const u of updates) {
    const {
      excelRow, views, likes, comments,
      viewsCol, likesCol, commentsCol,
      fileId    = LIVE_FILE_ID,
      sheetName = 'Live Inhouse',
    } = u;
    const enc  = encodeURIComponent(sheetName);
    const base = `/users/${PRIYANKA_USER}/drive/items/${fileId}/workbook/worksheets`;
    try {
      await graphPatch(`${base}('${enc}')/range(address='${viewsCol}${excelRow}')`,    { values: [[views]]    });
      await graphPatch(`${base}('${enc}')/range(address='${likesCol}${excelRow}')`,    { values: [[likes]]    });
      await graphPatch(`${base}('${enc}')/range(address='${commentsCol}${excelRow}')`, { values: [[comments]] });
      written++;
    } catch (err) {
      errors.push({ excelRow, sheetName, error: err.message });
    }
  }
  res.json({ written, errors });
});

// ── POST /api/metrics/start — kick off an Apify run (returns instantly) ───────
// Body: { usernames: [...] }
// Returns: { runId, datasetId }
//
// IMPORTANT: apify/instagram-scraper never returns videoViewCount when scraping
// individual reel/post URLs via directUrls — that field is only populated when
// scraping a creator's profile (their recent posts feed). So instead of passing
// reel URLs directly, we scrape each unique creator's profile and match the
// returned posts back to our rows by shortcode (done client-side, same as before).
app.post('/api/metrics/start', async (req, res) => {
  const { usernames } = req.body || {};
  if (!Array.isArray(usernames) || usernames.length === 0) return res.json({ runId: null });

  const token = process.env.APIFY_TOKEN;
  if (!token) return res.status(500).json({ error: 'APIFY_TOKEN not configured' });

  const uniqueUsernames = [...new Set(usernames.map(u => String(u).trim()).filter(Boolean))];
  const profileUrls = uniqueUsernames.map(u => `https://www.instagram.com/${u}/`);
  // Per-profile post limit — bounds cost ($0.0023/result) while covering each
  // creator's recent reels. Raise if older posts are frequently missed.
  const RESULTS_LIMIT_PER_PROFILE = 25;

  try {
    const r = await fetch(
      `https://api.apify.com/v2/acts/apify~instagram-scraper/runs?token=${token}&memory=1024`,
      {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({
          directUrls:   profileUrls,
          resultsType:  'posts',
          resultsLimit: RESULTS_LIMIT_PER_PROFILE,
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
      // Extract shortcode — key on it so /p/ and /reel/ and /reels/ all match
      const scMatch = rawUrl.match(/instagram\.com\/(?:p|reels?|tv)\/([A-Za-z0-9_-]+)/);
      return {
        url:      scMatch ? `https://www.instagram.com/reel/${scMatch[1]}/` : rawUrl.split('?')[0].replace(/\/$/, ''),
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
