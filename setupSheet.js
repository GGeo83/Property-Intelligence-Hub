/**
 * PROPERTY INTELLIGENCE HUB — Google Apps Script
 *
 * FUNCTIONS IN THIS FILE:
 *   setupSheet()          — First-time setup (clears + seeds all tabs). Run ONCE.
 *   addNewProperties()    — Append only new properties/sources without touching activity data.
 *   addUpgradesToSheet()  — Safe migration: adds status column to activities, creates events tab.
 *   clearSeedData()       — Wipe demo/seed rows from activities, events, scan_log. Run after testing.
 *   sendWeeklyDigest()    — Send HTML email digest to each principal. Run manually or via trigger.
 *   setupWeeklyDigest()   — Installs a Monday-9am time trigger for sendWeeklyDigest. Run ONCE.
 *   doPost(e)             — Web App endpoint — handles all webhook POSTs from hub and scan skills.
 *
 * DEDUPLICATION STRATEGY (in doPost):
 *   Before inserting any activity, a fingerprint is computed:
 *     • If sourceUrl present  →  propId + actType + sourceUrl
 *     • Otherwise            →  propId + actType + first 120 chars of desc (lowercased)
 *   Existing rows from the last 30 days are checked against this fingerprint.
 *   If a match exists the new activity is skipped (returns duplicate:true).
 *   This prevents the same permit/article/complaint from re-appearing on every daily scan.
 */

// ─────────────────────────────────────────────────────────────────────────────
// CONFIGURATION — update email addresses before running sendWeeklyDigest
// ─────────────────────────────────────────────────────────────────────────────
const PRINCIPAL_EMAILS = {
  APJCJM: 'geovanny22@gmail.com',   // AC — update with actual recipient email
  ECJIV:  'geovanny22@gmail.com',   // EC — update with actual recipient email
  ELJ:    'geovanny22@gmail.com',   // EL — update with actual recipient email
  EBJ:    'geovanny22@gmail.com',   // EB — update with actual recipient email
};

const PRINCIPAL_DISPLAY = {
  APJCJM: 'AC', ECJIV: 'EC', ELJ: 'EL', EBJ: 'EB',
};

// propId → principalId (used for digest grouping)
const PROP_PRINCIPAL_MAP = {
  prop_56_beacon:'APJCJM', prop_482_island:'APJCJM', prop_milton_estate:'APJCJM',
  prop_35_hayride:'APJCJM', prop_92_blodgett:'APJCJM', prop_nantucket_estate:'APJCJM',
  prop_mandarin_boston:'APJCJM',
  prop_16_union:'ECJIV', prop_131_commonwealth:'ECJIV',
  prop_18_louisburg:'ELJ', prop_2929_winding_oak:'ELJ', prop_louisburg_farm_fl:'ELJ', prop_dover_estate:'ELJ',
  prop_1_charles_river:'EBJ', prop_3_charles_river:'EBJ',
};

const PROP_SHORT_NAMES = {
  prop_56_beacon:'56 Beacon St', prop_482_island:'482 Island Dr',
  prop_16_union:'16 Union Wharf', prop_131_commonwealth:'131 Commonwealth',
  prop_18_louisburg:'18 Louisburg Sq', prop_2929_winding_oak:'2929 Winding Oak', prop_louisburg_farm_fl:'Louisburg Farm FL',
  prop_1_charles_river:'1 Charles River Sq', prop_3_charles_river:'3 Charles River Sq',
  prop_milton_estate:'Milton Estate', prop_35_hayride:'35 Hayride Dr',
  prop_92_blodgett:'92 Blodgett Way', prop_dover_estate:'Dover Estate',
  prop_nantucket_estate:'Nantucket Estate',
  prop_mandarin_boston:'Mandarin Oriental',
};

const TYPE_COLORS_EMAIL = {
  permit:'#2563eb', complaint:'#d97706', crime:'#dc2626', legal:'#dc2626',
  court:'#9333ea', fire:'#dc2626', landmarks:'#7c3aed', planning:'#0891b2',
  rumor:'#ea580c', forsale:'#16a34a', news:'#6366f1', other:'#94a3b8',
};


// ─────────────────────────────────────────────────────────────────────────────
// 1. SETUP SHEET (run once — clears and re-seeds everything)
// ─────────────────────────────────────────────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const tabNames = ['properties', 'activities', 'sources', 'scan_log', 'events'];
  tabNames.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    else sheet.clearContents();
  });

  // PROPERTIES tab
  const propsSheet = ss.getSheetByName('properties');
  propsSheet.appendRow(['propId','principalId','addr','city','state','zip','status','permits','complaints','crime','landmarks','planning','rumors','notes']);
  const properties = [
    ['prop_56_beacon',        'APJCJM', '56 Beacon Street',        'Boston',     'MA', '02108', 'active', '', '', '', '', '', '', 'Beacon Hill Historic District / BLC oversight'],
    ['prop_482_island',       'APJCJM', '482 Island Drive',        'Palm Beach', 'FL', '33480', 'active', '', '', '', '', '', '', 'Palm Beach Island / PBC Planning Commission'],
    ['prop_16_union',         'ECJIV',  '16 Union Wharf',          'Boston',     'MA', '02109', 'active', '', '', '', '', '', '', 'North End / Waterfront / Boston Harbor'],
    ['prop_131_commonwealth', 'ECJIV',  '131 Commonwealth Ave',    'Boston',     'MA', '02116', 'active', '', '', '', '', '', '', 'Back Bay / Back Bay Architectural Commission'],
    ['prop_18_louisburg',     'ELJ',    '18 Louisburg Square',     'Boston',     'MA', '02108', 'active', '', '', '', '', '', '', 'Beacon Hill / Most exclusive private square in Boston'],
    ['prop_2929_winding_oak', 'ELJ',    '2929 Winding Oak Lane',   'Wellington', 'FL', '33414', 'active', '', '', '', '', '', '', 'Wellington / Palm Beach County equestrian'],
    ['prop_louisburg_farm_fl','ELJ',    'Louisburg Farm FL',       'Wellington', 'FL', '33414', 'active', '', '', '', '', '', '', '3261 & 3315 Old Hampton Dr — Wellington P&Z / Palm Beach County equestrian'],
    ['prop_1_charles_river',  'EBJ',    '1 Charles River Square',  'Boston',     'MA', '02114', 'active', '', '', '', '', '', '', 'Beacon Hill / Private historic square off Mount Vernon St'],
    ['prop_3_charles_river',  'EBJ',    '3 Charles River Square',  'Boston',     'MA', '02114', 'active', '', '', '', '', '', '', 'Beacon Hill / Private historic square off Mount Vernon St'],
    ['prop_milton_estate',    'APJCJM', 'Milton Estate',            'Milton',     'MA', '02186', 'active', '', '', '', '', '', '', 'Multi-parcel: 1134, 1150, 1196 Canton Ave — Milton Planning Board, Norfolk County'],
    ['prop_35_hayride',       'APJCJM', '35 Hayride Drive',        'Stowe',      'VT', '05672', 'active', '', '', '', '', '', '', 'Stowe / Vermont Act 250 / Lamoille County / Stowe Reporter monitored'],
    ['prop_92_blodgett',      'APJCJM', '92 Blodgett Way',         'Lake Placid','NY', '12946', 'active', '', '', '', '', '', '', 'Adirondack Park Agency (APA) oversight / Essex County / Mirror Lake'],
    ['prop_dover_estate',     'ELJ',    'Dover Estate',            'Dover',      'MA', '02030', 'active', '', '', '', '', '', '', 'Multi-parcel estate: 20,29,33,39,45,49,56,62,64,68,74 Farm St · 20,36,40 Pegan Ln — Norfolk County'],
    ['prop_nantucket_estate', 'APJCJM', 'Nantucket Estate',         'Nantucket',  'MA', '02554', 'active', '', '', '', '', '', '', '1 Sandy Dr · 32B & 29 Hulbert Ave — Nantucket HDC Historic District / Waterfront'],
    ['prop_mandarin_boston',  'APJCJM', 'Mandarin Oriental Boston', 'Boston',    'MA', '02199', 'active', '', '', '', '', '', '', '776 Boylston St W12A · W12B · 778 Boylston St APT 7G — Back Bay Architectural Commission oversight'],
  ];
  properties.forEach(row => propsSheet.appendRow(row));

  // ACTIVITIES tab — with status column (K)
  const actsSheet = ss.getSheetByName('activities');
  actsSheet.appendRow(['id','propId','principalId','actType','scope','nearAddr','desc','time','sourceUrl','createdAt','status']);

  // SOURCES tab
  const srcsSheet = ss.getSheetByName('sources');
  srcsSheet.appendRow(['propId','addr','sourceType','label','url','priority','notes']);
  const sources = _getAllSources();
  sources.forEach(row => srcsSheet.appendRow(row));

  // SCAN_LOG tab
  const logSheet = ss.getSheetByName('scan_log');
  logSheet.appendRow(['scanType','runDate','findingsCount','summary']);

  // EVENTS tab — upcoming hearings, meetings, deadlines
  const eventsSheet = ss.getSheetByName('events');
  eventsSheet.appendRow(['propId','principalId','eventType','eventDate','title','desc','url','createdAt']);

  // Format headers
  ['properties','activities','sources','scan_log','events'].forEach(name => {
    const sheet = ss.getSheetByName(name);
    sheet.getRange(1,1,1,sheet.getLastColumn()).setFontWeight('bold');
    sheet.setFrozenRows(1);
  });

  Logger.log('✅ SETUP COMPLETE — ' + properties.length + ' properties, ' + sources.length + ' sources seeded. Events tab ready.');
}


// ─────────────────────────────────────────────────────────────────────────────
// 2. ADD NEW PROPERTIES (safe — won't touch existing activity data)
// ─────────────────────────────────────────────────────────────────────────────
function addNewProperties() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const propsSheet = ss.getSheetByName('properties');
  const existingIds = new Set(propsSheet.getDataRange().getValues().map(r => r[0]));

  const newProperties = [
    ['prop_milton_estate',     'APJCJM', 'Milton Estate',             'Milton',     'MA','02186','active','','','','','','','Multi-parcel: 1134, 1150, 1196 Canton Ave — Milton Planning Board, Norfolk County'],
    ['prop_35_hayride',        'APJCJM', '35 Hayride Drive',          'Stowe',      'VT','05672','active','','','','','','','Stowe / Vermont Act 250 / Lamoille County / Stowe Reporter monitored'],
    ['prop_92_blodgett',       'APJCJM', '92 Blodgett Way',          'Lake Placid','NY','12946','active','','','','','','','Adirondack Park Agency (APA) oversight / Essex County / Mirror Lake'],
    ['prop_louisburg_farm_fl', 'ELJ',    'Louisburg Farm FL',          'Wellington', 'FL','33414','active','','','','','','','3261 & 3315 Old Hampton Dr — Wellington P&Z / Palm Beach County equestrian'],
    ['prop_dover_estate',      'ELJ',    'Dover Estate',              'Dover',      'MA','02030','active','','','','','','','Multi-parcel estate: 20,29,33,39,45,49,56,62,64,68,74 Farm St · 20,36,40 Pegan Ln — Norfolk County'],
    ['prop_nantucket_estate',  'APJCJM', 'Nantucket Estate',           'Nantucket',  'MA','02554','active','','','','','','','1 Sandy Dr · 32B & 29 Hulbert Ave — Nantucket HDC Historic District / Waterfront'],
    ['prop_mandarin_boston',   'APJCJM', 'Mandarin Oriental Boston',  'Boston',     'MA','02199','active','','','','','','','776 Boylston St W12A · W12B · 778 Boylston St APT 7G — Back Bay Architectural Commission oversight'],
  ];

  let propsAdded = 0;
  newProperties.forEach(row => {
    if (existingIds.has(row[0])) { Logger.log('⏭ Skipping ' + row[0]); return; }
    propsSheet.appendRow(row);
    propsAdded++;
    Logger.log('✅ Added: ' + row[0]);
  });

  const srcsSheet = ss.getSheetByName('sources');
  const existingSrcKeys = new Set(srcsSheet.getDataRange().getValues().map(r => r[0]+'|'+r[3]));
  const newSources = _getNewPropertySources();
  let srcsAdded = 0;
  newSources.forEach(row => {
    const key = row[0]+'|'+row[3];
    if (existingSrcKeys.has(key)) return;
    srcsSheet.appendRow(row);
    srcsAdded++;
  });

  Logger.log('✅ DONE — added ' + propsAdded + ' properties, ' + srcsAdded + ' sources. Activities/scan_log untouched.');
}


// ─────────────────────────────────────────────────────────────────────────────
// 3. ADD UPGRADES TO EXISTING SHEET (adds status col + events tab safely)
//
// HOW TO USE:
//   1. Open sheet > Extensions > Apps Script > paste this file > Save
//   2. Select "addUpgradesToSheet" in dropdown > Run
//   3. Check Execution Log — should show "UPGRADE COMPLETE"
//   After this: the report's resolve buttons and upcoming events section will work.
// ─────────────────────────────────────────────────────────────────────────────
function addUpgradesToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Add status column to activities (column K) if not present
  const actsSheet = ss.getSheetByName('activities');
  const headers = actsSheet.getRange(1, 1, 1, Math.max(actsSheet.getLastColumn(), 10)).getValues()[0];
  const statusColIdx = headers.indexOf('status');
  if (statusColIdx === -1) {
    const nextCol = actsSheet.getLastColumn() + 1;
    actsSheet.getRange(1, nextCol).setValue('status');
    actsSheet.getRange(1, nextCol).setFontWeight('bold');
    Logger.log('✅ Added "status" column to activities at column ' + nextCol);
  } else {
    Logger.log('⏭ "status" column already exists at position ' + (statusColIdx + 1));
  }

  // 2. Add events tab if not present
  let eventsSheet = ss.getSheetByName('events');
  if (!eventsSheet) {
    eventsSheet = ss.insertSheet('events');
    eventsSheet.appendRow(['propId','principalId','eventType','eventDate','title','desc','url','createdAt']);
    eventsSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    eventsSheet.setFrozenRows(1);
    Logger.log('✅ Created "events" tab with headers');
  } else {
    Logger.log('⏭ "events" tab already exists');
  }

  Logger.log('✅ UPGRADE COMPLETE — activities.status column ready, events tab ready. No data was touched.');
}


// ─────────────────────────────────────────────────────────────────────────────
// 4. SEND WEEKLY DIGEST (email per principal — filtered to their properties)
//
// HOW TO USE (manual):
//   Select "sendWeeklyDigest" in dropdown > Run
//
// HOW TO USE (automated trigger):
//   Run setupWeeklyDigest() ONCE to install a Monday 9am trigger.
//   After that it runs automatically every Monday.
// ─────────────────────────────────────────────────────────────────────────────
function sendWeeklyDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const actsSheet = ss.getSheetByName('activities');
  const allRows = actsSheet.getDataRange().getValues().slice(1); // skip header

  // Filter to last 7 days
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 7);

  const recent = allRows.filter(r => {
    const d = new Date(r[9] || r[7]); // createdAt (col J) or time (col H)
    return !isNaN(d) && d >= cutoff && r[6]; // has a description
  });

  if (recent.length === 0) {
    Logger.log('No recent activity in last 7 days — digest not sent.');
    return;
  }

  // Group by principal (resolve from propId if principalId missing)
  const byPrincipal = {};
  recent.forEach(r => {
    const pid = PROP_PRINCIPAL_MAP[r[1]] || r[2] || 'UNKNOWN';
    if (!byPrincipal[pid]) byPrincipal[pid] = [];
    byPrincipal[pid].push(r);
  });

  // Fetch upcoming events for the next 14 days too
  const eventsSheet = ss.getSheetByName('events');
  const upcomingEvents = eventsSheet ? _getUpcomingEvents(eventsSheet, 14) : [];

  // Send to each principal with activity
  const sent = [];
  Object.entries(byPrincipal).forEach(([pid, acts]) => {
    const email = PRINCIPAL_EMAILS[pid];
    if (!email) { Logger.log('⚠ No email configured for ' + pid); return; }

    const principalEvents = upcomingEvents.filter(ev => PROP_PRINCIPAL_MAP[ev[0]] === pid || ev[1] === pid);
    const html = _buildDigestHtml(pid, acts, principalEvents);
    const dateStr = new Date().toLocaleDateString('en-US', {month:'long', day:'numeric', year:'numeric'});

    MailApp.sendEmail({
      to: email,
      subject: `🏢 PIH Weekly Digest — ${PRINCIPAL_DISPLAY[pid] || pid} — ${dateStr}`,
      htmlBody: html,
    });
    sent.push(pid + ' → ' + email);
    Logger.log('✅ Sent to ' + pid + ' (' + email + '): ' + acts.length + ' activities');
  });

  Logger.log('✅ DIGEST COMPLETE — sent: ' + sent.join(', '));
}

function _getUpcomingEvents(eventsSheet, daysAhead) {
  const now = new Date();
  const horizon = new Date(); horizon.setDate(horizon.getDate() + daysAhead);
  return eventsSheet.getDataRange().getValues().slice(1).filter(r => {
    if (!r[3]) return false; // no date
    const d = new Date(r[3]);
    return !isNaN(d) && d >= now && d <= horizon;
  }).sort((a, b) => new Date(a[3]) - new Date(b[3]));
}

function _buildDigestHtml(pid, acts, events) {
  const pName = PRINCIPAL_DISPLAY[pid] || pid;
  const dateStr = new Date().toLocaleDateString('en-US', {weekday:'long', month:'long', day:'numeric', year:'numeric'});
  const critical = acts.filter(r => ['crime','legal','court','fire'].includes(r[3]));

  // Group activities by propId
  const byProp = {};
  acts.forEach(r => {
    const k = r[1] || 'unknown';
    if (!byProp[k]) byProp[k] = [];
    byProp[k].push(r);
  });

  const propRows = Object.entries(byProp).map(([propId, rows]) => {
    const propName = PROP_SHORT_NAMES[propId] || propId;
    const actRows = rows.map(r => {
      const tc = TYPE_COLORS_EMAIL[r[3]] || '#94a3b8';
      const tl = (r[3] || 'other').toUpperCase();
      const age = r[9] ? Math.floor((new Date() - new Date(r[9]))/86400000) : '?';
      return `<tr>
        <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;font-size:13px;color:#0f172a;line-height:1.4">${r[6]}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;white-space:nowrap">
          <span style="background:${tc}18;color:${tc};border:1px solid ${tc}40;padding:2px 7px;border-radius:10px;font-size:10px;font-weight:700">${tl}</span>
        </td>
        <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;font-size:11px;color:#94a3b8;white-space:nowrap">${age}d ago</td>
        <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;font-size:11px">
          ${r[8] ? `<a href="${r[8]}" style="color:#3b82f6;text-decoration:none">Source ↗</a>` : ''}
        </td>
      </tr>`;
    }).join('');

    return `<div style="margin-bottom:20px">
      <div style="font-size:13px;font-weight:700;color:#0f172a;padding:8px 12px;background:#f8fafc;border-left:3px solid #3b82f6;margin-bottom:0">${propName}</div>
      <table style="width:100%;border-collapse:collapse;border:1px solid #e2e8f0">${actRows}</table>
    </div>`;
  }).join('');

  const eventsHtml = events.length ? `
    <div style="margin-bottom:24px">
      <h3 style="font-size:13px;font-weight:700;color:#0f172a;margin:0 0 10px;text-transform:uppercase;letter-spacing:.05em">📅 Upcoming Events (Next 14 Days)</h3>
      ${events.map(ev => {
        const d = new Date(ev[3]);
        const dateLabel = d.toLocaleDateString('en-US',{weekday:'short',month:'short',day:'numeric'});
        return `<div style="padding:8px 12px;border-left:3px solid #f59e0b;background:#fffbeb;margin-bottom:6px;font-size:13px">
          <strong style="color:#92400e">${dateLabel}</strong> — ${PROP_SHORT_NAMES[ev[0]]||ev[0]}: ${ev[4]}
          ${ev[6] ? `<a href="${ev[6]}" style="color:#3b82f6;text-decoration:none;margin-left:8px">Details ↗</a>` : ''}
        </div>`;
      }).join('')}
    </div>` : '';

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;background:#f8fafc;padding:20px">
  <div style="background:#1e293b;padding:20px 24px;border-radius:8px 8px 0 0">
    <div style="font-size:18px;font-weight:800;color:#fff">📊 Property Intelligence Hub</div>
    <div style="font-size:12px;color:#94a3b8;margin-top:4px">Weekly Digest · ${pName} · ${dateStr}</div>
  </div>
  <div style="background:#fff;padding:20px 24px;border:1px solid #e2e8f0;border-top:none">
    ${critical.length ? `
    <div style="background:#fef2f2;border:1px solid #fecaca;border-radius:6px;padding:10px 14px;margin-bottom:20px">
      <strong style="color:#dc2626;font-size:13px">⚠ ${critical.length} Critical Alert${critical.length>1?'s':''} This Week</strong>
      <div style="font-size:12px;color:#7f1d1d;margin-top:3px">Crime · Legal · Fire activity detected — review immediately</div>
    </div>` : ''}
    <div style="display:flex;gap:16px;margin-bottom:24px">
      <div style="flex:1;background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:12px;text-align:center">
        <div style="font-size:24px;font-weight:800;color:#0f172a">${acts.length}</div>
        <div style="font-size:11px;color:#94a3b8;margin-top:2px">Total Alerts</div>
      </div>
      <div style="flex:1;background:${critical.length?'#fef2f2':'#f0fdf4'};border:1px solid ${critical.length?'#fecaca':'#bbf7d0'};border-radius:6px;padding:12px;text-align:center">
        <div style="font-size:24px;font-weight:800;color:${critical.length?'#dc2626':'#16a34a'}">${critical.length}</div>
        <div style="font-size:11px;color:#94a3b8;margin-top:2px">Critical</div>
      </div>
      <div style="flex:1;background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:12px;text-align:center">
        <div style="font-size:24px;font-weight:800;color:#0f172a">${Object.keys(byProp).length}</div>
        <div style="font-size:11px;color:#94a3b8;margin-top:2px">Properties Active</div>
      </div>
    </div>
    ${eventsHtml}
    <h3 style="font-size:13px;font-weight:700;color:#0f172a;margin:0 0 12px;text-transform:uppercase;letter-spacing:.05em">🚨 This Week's Alerts</h3>
    ${propRows}
    <div style="margin-top:20px;padding-top:16px;border-top:1px solid #e2e8f0;font-size:11px;color:#94a3b8;text-align:center">
      Property Intelligence Hub · Alerts are auto-generated by scan agents<br>
      To manage alerts and view the live dashboard, open the PIH hub.
    </div>
  </div>
</body></html>`;
}


// ─────────────────────────────────────────────────────────────────────────────
// 5. SETUP WEEKLY TRIGGER (run ONCE — installs Monday 9am automatic trigger)
// ─────────────────────────────────────────────────────────────────────────────
function setupWeeklyDigest() {
  // Remove any existing sendWeeklyDigest triggers first
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendWeeklyDigest') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('sendWeeklyDigest')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  Logger.log('✅ Weekly digest trigger installed — will run every Monday at 9am.');
  Logger.log('To remove it: go to Apps Script > Triggers (clock icon) and delete.');
}


// ─────────────────────────────────────────────────────────────────────────────
// 6. WEB APP — doPost (handles all webhook POSTs from hub and scan skills)
// ─────────────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ok = s => ContentService.createTextOutput(JSON.stringify(s)).setMimeType(ContentService.MimeType.JSON);

    // ── addProperty ──────────────────────────────────────────────────────────
    if (data.action === 'addProperty') {
      ss.getSheetByName('properties').appendRow([
        data.propId || ('prop_' + Date.now()), data.principalId||'', data.addr||'',
        data.city||'', data.state||'', data.zip||'', data.status||'active',
        '','','','','','', data.notes||''
      ]);
      return ok({ok:true,action:'addProperty'});
    }

    // ── addActivity (single activity post — deduped) ──────────────────────────
    if (data.action === 'addActivity') {
      const actsSheet = ss.getSheetByName('activities');
      const result = _insertActivityIfNew(actsSheet, data);
      return ok({ok:true, action:'addActivity', duplicate:result.duplicate, id:result.id});
    }

    // ── logScan ───────────────────────────────────────────────────────────────
    if (data.action === 'logScan') {
      ss.getSheetByName('scan_log').appendRow([
        data.scanType||'daily',
        data.runDate||new Date().toISOString().split('T')[0],
        data.findingsCount||0,
        data.summary||''
      ]);
      if (data.activities && data.activities.length > 0) {
        const actsSheet = ss.getSheetByName('activities');
        let inserted = 0, dupes = 0;
        data.activities.forEach(act => {
          const result = _insertActivityIfNew(actsSheet, act);
          result.duplicate ? dupes++ : inserted++;
        });
        Logger.log('logScan: ' + inserted + ' new, ' + dupes + ' duplicates skipped');
      }
      return ok({ok:true, action:'logScan'});
    }

    // ── resolveActivity (mark as reviewed or actioned from report) ────────────
    if (data.action === 'resolveActivity') {
      const sheet = ss.getSheetByName('activities');
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.actId) {
          // Ensure status column exists (col K = index 10)
          sheet.getRange(i + 1, 11).setValue(data.status || 'reviewed');
          return ok({ok:true,action:'resolveActivity',actId:data.actId,status:data.status});
        }
      }
      return ok({ok:false,error:'Activity not found: ' + data.actId});
    }

    // ── addEvent (log upcoming hearing/meeting/deadline from scan) ────────────
    if (data.action === 'addEvent') {
      const sheet = ss.getSheetByName('events');
      if (!sheet) return ok({ok:false,error:'events tab not found — run addUpgradesToSheet() first'});
      // Check for duplicate (same propId + eventDate + title)
      const existing = sheet.getDataRange().getValues().slice(1);
      const isDupe = existing.some(r => r[0]===data.propId && r[3]===data.eventDate && r[4]===data.title);
      if (!isDupe) {
        sheet.appendRow([
          data.propId||'', data.principalId||PROP_PRINCIPAL_MAP[data.propId]||'',
          data.eventType||'hearing', data.eventDate||'',
          data.title||'', data.desc||'', data.url||'',
          new Date().toISOString()
        ]);
      }
      return ok({ok:true,action:'addEvent',duplicate:isDupe});
    }

    // ── sendDigest (trigger digest on demand from report) ─────────────────────
    if (data.action === 'sendDigest') {
      sendWeeklyDigest();
      return ok({ok:true,action:'sendDigest'});
    }

    // ── updateProperty (patch one or more columns on a matching propId row) ─────
    if (data.action === 'updateProperty') {
      if (!data.propId) return ok({ok:false,error:'propId required'});
      const sheet = ss.getSheetByName('properties');
      const rows = sheet.getDataRange().getValues();
      const headers = rows[0]; // ['propId','principalId','addr','city','state','zip','status',...]
      let updated = 0;
      const fields = {addr:2,city:3,state:4,zip:5,status:6,notes:13};
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.propId) {
          Object.entries(fields).forEach(([key, col]) => {
            if (data[key] !== undefined) sheet.getRange(i + 1, col + 1).setValue(data[key]);
          });
          updated++;
        }
      }
      return ok({ok:true, action:'updateProperty', propId:data.propId, updated});
    }

    // ── deleteProperty (remove all rows matching propId from properties tab) ───
    if (data.action === 'deleteProperty') {
      if (!data.propId) return ok({ok:false,error:'propId required'});
      const sheet = ss.getSheetByName('properties');
      const rows = sheet.getDataRange().getValues();
      // Collect row indices to delete (iterate bottom-up to preserve indices)
      const toDelete = [];
      for (let i = rows.length - 1; i >= 1; i--) {
        if (rows[i][0] === data.propId) toDelete.push(i + 1); // 1-based
      }
      toDelete.forEach(r => sheet.deleteRow(r));
      return ok({ok:true, action:'deleteProperty', propId:data.propId, deletedRows:toDelete.length});
    }

    // ── deleteProperties (batch — array of propIds) ────────────────────────────
    if (data.action === 'deleteProperties') {
      if (!data.propIds || !data.propIds.length) return ok({ok:false,error:'propIds array required'});
      const sheet = ss.getSheetByName('properties');
      let totalDeleted = 0;
      data.propIds.forEach(pid => {
        const rows = sheet.getDataRange().getValues();
        for (let i = rows.length - 1; i >= 1; i--) {
          if (rows[i][0] === pid) { sheet.deleteRow(i + 1); totalDeleted++; }
        }
      });
      return ok({ok:true, action:'deleteProperties', propIds:data.propIds, deletedRows:totalDeleted});
    }

    return ok({ok:false,error:'Unknown action: ' + data.action});

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// 7. CLEAR SEED DATA (run once after testing — wipes demo rows, keeps headers)
//
// HOW TO USE:
//   Select "clearSeedData" in the function dropdown > Run
//   This clears the activities, events, and scan_log tabs completely (headers kept).
//   Use this to remove seedSampleData() demo rows before real scans start.
// ─────────────────────────────────────────────────────────────────────────────
function clearSeedData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cleared = [];

  ['activities', 'events', 'scan_log'].forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) { Logger.log('⚠ Tab not found: ' + tabName); return; }
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) { Logger.log('⏭ ' + tabName + ' already empty (header only)'); return; }
    // Delete all data rows (keep row 1 = header)
    sheet.deleteRows(2, lastRow - 1);
    cleared.push(tabName + ' (' + (lastRow - 1) + ' rows removed)');
    Logger.log('✅ Cleared ' + (lastRow - 1) + ' rows from ' + tabName);
  });

  if (cleared.length) {
    Logger.log('✅ SEED CLEAR COMPLETE — ' + cleared.join(', '));
    Logger.log('The PIH dashboard will now show empty until real scans post data.');
  } else {
    Logger.log('Nothing to clear — all tabs were already empty.');
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// PRIVATE HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Compute a deduplication fingerprint for an activity.
 * Priority: propId + actType + sourceUrl (most precise)
 * Fallback:  propId + actType + first 120 chars of desc (lowercased, trimmed)
 */
function _actFingerprint(act) {
  const url = (act.sourceUrl || act.url || '').trim();
  if (url) return (act.propId||'') + '|' + (act.actType||act.type||'') + '|' + url;
  const desc = (act.desc || act.description || '').toLowerCase().trim().substring(0, 120);
  return (act.propId||'') + '|' + (act.actType||act.type||'') + '|' + desc;
}

/**
 * Insert an activity row only if no matching fingerprint exists within the last 30 days.
 * Returns {duplicate: bool, id: string}
 */
function _insertActivityIfNew(actsSheet, act) {
  const fp = _actFingerprint(act);
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);

  // Read existing rows and check for fingerprint match within 30-day window
  const existing = actsSheet.getDataRange().getValues().slice(1); // skip header
  const isDupe = existing.some(r => {
    // r[9] = createdAt, r[8] = sourceUrl, r[6] = desc, r[1] = propId, r[3] = actType
    const rowDate = new Date(r[9] || r[7]);
    if (isNaN(rowDate) || rowDate < cutoff) return false; // too old, ignore for dedup
    const rowFp = _actFingerprint({
      propId: r[1], actType: r[3], sourceUrl: r[8], desc: r[6]
    });
    return rowFp === fp;
  });

  if (isDupe) return {duplicate: true, id: null};

  const id = 'act_' + Date.now() + '_' + Math.random().toString(36).substr(2,5);
  actsSheet.appendRow([
    id,
    act.propId||'',
    act.principalId || PROP_PRINCIPAL_MAP[act.propId] || '',
    act.actType || act.type || 'other',
    act.scope || 'on_property',
    act.nearAddr || '',
    act.desc || act.description || '',
    act.time || act.eventDate || new Date().toISOString(),
    act.sourceUrl || act.url || '',
    new Date().toISOString(), // createdAt = now (when scan found it, not when event happened)
    'new'
  ]);
  return {duplicate: false, id};
}


function _getAllSources() {
  return [
    ['prop_56_beacon','56 Beacon Street','landmarks','Boston Landmarks Commission','https://www.boston.gov/departments/landmarks-commission',1,'Mandatory login: ggordillo@lthill.com / LTHill2024!'],
    ['prop_56_beacon','56 Beacon Street','news','Beacon Hill Civic Association','https://www.beaconhillcivic.org',1,'Neighborhood alerts, BLC hearing notices'],
    ['prop_56_beacon','56 Beacon Street','news','Patch Beacon Hill','https://patch.com/massachusetts/beaconhill',2,''],
    ['prop_56_beacon','56 Beacon Street','news','Boston.com Beacon Hill','https://www.boston.com/tag/beacon-hill/',2,''],
    ['prop_56_beacon','56 Beacon Street','permits','Boston Inspectional Services','https://www.boston.gov/departments/inspectional-services',1,''],
    ['prop_56_beacon','56 Beacon Street','crime','BPD Crime Incident Reports','https://data.boston.gov/dataset/crime-incident-reports',2,''],
    ['prop_56_beacon','56 Beacon Street','complaints','Boston 311','https://311.boston.gov',2,''],
    ['prop_482_island','482 Island Drive','news','Palm Beach Daily News','https://www.palmbeachdailynews.com',1,'PRIMARY'],
    ['prop_482_island','482 Island Drive','planning','PBC Planning Commission','https://discover.pbc.gov/pzb/planning/Pages/Planning-Commission-Agendas-Minutes.aspx',1,'PDF agendas — check weekly'],
    ['prop_482_island','482 Island Drive','permits','Town of Palm Beach Building','https://townofpalmbeach.com/174/Building-Zoning',1,''],
    ['prop_482_island','482 Island Drive','news','Palm Beach Post','https://www.palmbeachpost.com',2,''],
    ['prop_482_island','482 Island Drive','crime','PBSO','https://www.pbso.org',2,''],
    ['prop_482_island','482 Island Drive','rumors','Dirt.com','https://www.dirt.com',2,'Luxury transactions, off-market rumors'],
    ['prop_16_union','16 Union Wharf','news','North End Waterfront','https://northendwaterfront.com',1,'PRIMARY'],
    ['prop_16_union','16 Union Wharf','planning','Boston Harbor Association','https://www.tbha.org',1,''],
    ['prop_16_union','16 Union Wharf','planning','BPDA','https://www.bostonplans.org',1,''],
    ['prop_16_union','16 Union Wharf','permits','Boston Inspectional Services','https://www.boston.gov/departments/inspectional-services',1,''],
    ['prop_16_union','16 Union Wharf','news','Boston.com North End','https://www.boston.com/tag/north-end/',2,''],
    ['prop_131_commonwealth','131 Commonwealth Ave','landmarks','Back Bay Architectural Commission','https://www.boston.gov/departments/back-bay-architectural-commission',1,'MANDATORY'],
    ['prop_131_commonwealth','131 Commonwealth Ave','news','Back Bay Sun','https://www.backbaysun.com',1,'PRIMARY'],
    ['prop_131_commonwealth','131 Commonwealth Ave','news','Back Bay Association','https://www.backbayassociation.org',2,''],
    ['prop_131_commonwealth','131 Commonwealth Ave','landmarks','Boston Landmarks Commission','https://www.boston.gov/departments/landmarks-commission',1,''],
    ['prop_131_commonwealth','131 Commonwealth Ave','permits','Boston Inspectional Services','https://www.boston.gov/departments/inspectional-services',1,''],
    ['prop_18_louisburg','18 Louisburg Square','landmarks','Boston Landmarks Commission','https://www.boston.gov/departments/landmarks-commission',1,'Mandatory login'],
    ['prop_18_louisburg','18 Louisburg Square','news','Beacon Hill Civic Association','https://www.beaconhillcivic.org',1,''],
    ['prop_18_louisburg','18 Louisburg Square','news','Patch Beacon Hill','https://patch.com/massachusetts/beaconhill',1,''],
    ['prop_18_louisburg','18 Louisburg Square','permits','Boston Inspectional Services','https://www.boston.gov/departments/inspectional-services',1,''],
    ['prop_2929_winding_oak','2929 Winding Oak Lane','planning','Wellington Planning & Zoning','https://www.wellingtonfl.gov/373/Planning-Zoning',1,''],
    ['prop_2929_winding_oak','2929 Winding Oak Lane','planning','Wellington Village Council','https://www.wellingtonfl.gov/agendacenter',1,''],
    ['prop_2929_winding_oak','2929 Winding Oak Lane','news','Wellington The Magazine','https://wellingtonthemagazine.com',1,'PRIMARY'],
    ['prop_2929_winding_oak','2929 Winding Oak Lane','news','Palm Beach Post','https://www.palmbeachpost.com',2,''],
    // Louisburg Farm FL — 3261 & 3315 Old Hampton Dr, Wellington
    ['prop_louisburg_farm_fl','Louisburg Farm FL','planning','Wellington Planning & Zoning','https://www.wellingtonfl.gov/373/Planning-Zoning',1,''],
    ['prop_louisburg_farm_fl','Louisburg Farm FL','planning','PBC Building Division','https://pbcgov.com/pzb/building/',1,''],
    ['prop_louisburg_farm_fl','Louisburg Farm FL','planning','PBC Property Appraiser','https://www.pbcgov.com/papa/',1,''],
    ['prop_louisburg_farm_fl','Louisburg Farm FL','news','Wellington The Magazine','https://wellingtonthemagazine.com',1,'PRIMARY'],
    ['prop_louisburg_farm_fl','Louisburg Farm FL','news','Palm Beach Post','https://www.palmbeachpost.com',2,''],
    ['prop_louisburg_farm_fl','Louisburg Farm FL','crime','PBSO','https://www.pbso.org',2,''],
    ['prop_1_charles_river','1 Charles River Square','landmarks','Boston Landmarks Commission','https://www.boston.gov/departments/landmarks-commission',1,''],
    ['prop_1_charles_river','1 Charles River Square','news','Beacon Hill Civic Association','https://www.beaconhillcivic.org',1,''],
    ['prop_1_charles_river','1 Charles River Square','permits','Boston Inspectional Services','https://www.boston.gov/departments/inspectional-services',1,''],
    ['prop_3_charles_river','3 Charles River Square','landmarks','Boston Landmarks Commission','https://www.boston.gov/departments/landmarks-commission',1,''],
    ['prop_3_charles_river','3 Charles River Square','news','Beacon Hill Civic Association','https://www.beaconhillcivic.org',1,''],
    ['prop_3_charles_river','3 Charles River Square','permits','Boston Inspectional Services','https://www.boston.gov/departments/inspectional-services',1,''],
    ['prop_35_hayride','35 Hayride Drive','planning','Vermont Act 250','https://anr.vermont.gov/land-use/act-250',1,'PRIMARY'],
    ['prop_35_hayride','35 Hayride Drive','planning','Stowe Planning & Zoning','https://www.townofstowevt.org/planning-zoning',1,''],
    ['prop_35_hayride','35 Hayride Drive','news','Stowe Reporter','https://www.stowereporter.com',1,'PRIMARY'],
    ['prop_35_hayride','35 Hayride Drive','news','VTDigger','https://vtdigger.org',2,''],
    ['prop_35_hayride','35 Hayride Drive','permits','Town of Stowe','https://www.townofstowevt.org',1,''],
    ['prop_milton_estate','Milton Estate','news','Milton Times','https://www.miltontimes.com/',1,'PRIMARY'],
    ['prop_milton_estate','Milton Estate','planning','Milton Planning Board','https://www.miltonma.gov/AgendaCenter/Planning-Board-39',1,''],
    ['prop_milton_estate','Milton Estate','permits','Milton Building Inspector','https://www.miltonma.gov/190/Building-Inspector',1,''],
  ].concat(_getNewPropertySources());
}

function _getNewPropertySources() {
  return [
    ['prop_92_blodgett','92 Blodgett Way','landmarks','Adirondack Park Agency (APA)','https://apa.ny.gov',1,'MANDATORY'],
    ['prop_92_blodgett','92 Blodgett Way','planning','Essex County Planning','https://www.essexcountyny.gov/planning',1,''],
    ['prop_92_blodgett','92 Blodgett Way','permits','North Elba / Lake Placid Bldg','https://www.northelba.org',1,''],
    ['prop_92_blodgett','92 Blodgett Way','news','Adirondack Daily Enterprise','https://www.adirondackdailyenterprise.com',1,'PRIMARY'],
    ['prop_92_blodgett','92 Blodgett Way','news','Adirondack Explorer','https://www.adirondackexplorer.org',2,''],
    ['prop_92_blodgett','92 Blodgett Way','crime','Essex County Sheriff','https://www.essexcountyny.gov/sheriff',2,''],
    ['prop_dover_estate','Dover Estate','planning','Dover Planning Board','https://www.doverma.org/planning-board',1,''],
    ['prop_dover_estate','Dover Estate','permits','Dover Building Dept','https://www.doverma.org/building-department',1,''],
    ['prop_dover_estate','Dover Estate','planning','Dover Conservation Comm.','https://www.doverma.org/conservation-commission',1,'HIGH PRIORITY wetlands'],
    ['prop_dover_estate','Dover Estate','permits','Norfolk County Deeds','https://www.norfolkdeeds.org',1,''],
    ['prop_dover_estate','Dover Estate','news','Wicked Local Dover','https://dover.wickedlocal.com',1,'PRIMARY'],
    ['prop_dover_estate','Dover Estate','news','Patch Dover/Sherborn','https://patch.com/massachusetts/dover',2,''],
    // Nantucket Estate — 1 Sandy Dr · 32B & 29 Hulbert Ave
    ['prop_nantucket_estate','Nantucket Estate','landmarks','Nantucket HDC','https://www.nantucket-ma.gov/189/Historic-District-Commission',1,'MANDATORY'],
    ['prop_nantucket_estate','Nantucket Estate','permits','Nantucket Building Dept','https://www.nantucket-ma.gov/187/Building',1,''],
    ['prop_nantucket_estate','Nantucket Estate','planning','Nantucket Planning Board','https://www.nantucket-ma.gov/197/Planning-Board',1,''],
    ['prop_nantucket_estate','Nantucket Estate','news','ACK.net','https://www.ack.net',1,'PRIMARY'],
    ['prop_nantucket_estate','Nantucket Estate','news','Nantucket Current','https://nantucketcurrent.com',2,''],
    // Mandarin Oriental Boston — 776 Boylston W12A·W12B · 778 Boylston APT 7G (Back Bay)
    ['prop_mandarin_boston','Mandarin Oriental Boston','landmarks','Back Bay Architectural Commission','https://www.boston.gov/departments/back-bay-architectural-commission',1,'MANDATORY'],
    ['prop_mandarin_boston','Mandarin Oriental Boston','landmarks','Boston Landmarks Commission','https://www.boston.gov/departments/landmarks-commission',1,''],
    ['prop_mandarin_boston','Mandarin Oriental Boston','permits','Boston Inspectional Services','https://www.boston.gov/departments/inspectional-services',1,''],
    ['prop_mandarin_boston','Mandarin Oriental Boston','planning','BPDA','https://www.bostonplans.org',1,''],
    ['prop_mandarin_boston','Mandarin Oriental Boston','complaints','Boston 311','https://311.boston.gov',2,''],
    ['prop_mandarin_boston','Mandarin Oriental Boston','crime','BPD Crime Incident Reports','https://data.boston.gov/dataset/crime-incident-reports',2,''],
    ['prop_mandarin_boston','Mandarin Oriental Boston','news','Back Bay Sun','https://www.backbaysun.com',1,'PRIMARY'],
  ];
}
