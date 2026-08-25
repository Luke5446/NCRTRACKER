/**
 * NCR DASHBOARD BUILDER  (Google Apps Script)
 * ---------------------------------------------------------------------------
 * Adds a self-updating "Dashboard" tab to your NCR register.
 * It only CREATES a tab — it never changes your data or the register.
 *
 * This file also contains:
 *   - applyNewNcrDefaults  (trigger: On form submit — sets Status/Owner/wrap)
 *   - backfillOwnersOnce   (run manually once if needed)
 *   - doGet                (Web app API feeding the manager dashboard HTML)
 * ---------------------------------------------------------------------------
 */

// ===== EDIT IF NEEDED ======================================================
const SRC = 'Form Responses 1';   // the exact name of your register tab (headers row 1, data row 2)
const FIRST = 2;                  // first row of real data (headers on row 1, data begins row 2)
const YEAR = 2026;                // year shown in the month-by-department breakdown

// Register column letters — update these if columns ever move again:
const COL = {
  ts:     'A',   // Timestamp
  dept:   'C',   // Department
  type:   'D',   // Type
  issue:  'E',   // Issue category
  cost:   'K',   // Est. cost (£)
  status: 'M',   // Status
  sev:    'N',   // Severity
  root:   'P',   // Root cause
  target: 'U',   // Target Complete
  days:   'W',   // Days open
};
// ===========================================================================

const DEPTS = ['Sales', 'Production', 'IT & Digital', 'Stores', 'Accounts', 'Suppliers', 'Procurement',
               'EMB', 'Logistics', 'DPD', 'H&B', 'ISO/QC', 'Other'];
const SEVS  = ['Critical', 'Major', 'Minor'];
const TYPES = ['Non Conformance', 'Customer Complaint', 'Observation'];
const CATS  = ['Input error', 'Process error', 'Pricing error', 'Missing item',
               'Training issue', 'Supplier fault', 'Transit/delivery damage', 'Other'];

const NAVY = '#1F3864', BLUE = '#2E5496', GREEN = '#538135', RED = '#C00000', AMBER = '#BF8F00';
const LBLUE = '#DDEBF7', LGREY = '#F2F2F2';

function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let src = ss.getSheetByName(SRC);
  if (!src) src = ss.getSheets().find(s => s.getName() !== 'Dashboard'); // fallback
  const S = "'" + src.getName() + "'!";
  const R = c => `${S}${c}${FIRST}:${c}`;   // builds e.g. 'Sheet'!M3:M from a column letter

  let sh = ss.getSheetByName('Dashboard');
  if (sh) sh.clear(); else sh = ss.insertSheet('Dashboard', 0);
  sh.setHiddenGridlines(true);

  const F = (a1, formula) => sh.getRange(a1).setFormula(formula);
  const V = (a1, val) => sh.getRange(a1).setValue(val);

  // ---- Title ----
  sh.getRange('B1:K1').merge();
  V('B1', 'NCR Dashboard');
  sh.getRange('B1').setFontSize(20).setFontWeight('bold').setFontColor(NAVY);
  sh.getRange('B2:K2').merge();
  V('B2', 'Live summary — updates automatically as the register changes');
  sh.getRange('B2').setFontColor('#808080').setFontStyle('italic');

  // ---- KPI cards ----
  const kpis = [
    ['Total NCRs', `=COUNTA(${R(COL.ts)})`, NAVY,  'B'],
    ['Open',        `=COUNTIF(${R(COL.status)},"Open")`, RED, 'D'],
    ['In progress', `=COUNTIF(${R(COL.status)},"In Progress")`, AMBER, 'F'],
    ['Closed',      `=COUNTIF(${R(COL.status)},"Closed")`, GREEN, 'H'],
    ['Overdue',     `=COUNTIFS(${R(COL.status)},"<>Closed",${R(COL.target)},"<"&TODAY(),${R(COL.target)},"<>")`, RED, 'J'],
  ];
  kpis.forEach(([lab, f, col, c]) => {
    const c2 = String.fromCharCode(c.charCodeAt(0) + 1);
    sh.getRange(`${c}4:${c2}4`).merge(); sh.getRange(`${c}5:${c2}5`).merge();
    V(`${c}4`, lab); F(`${c}5`, f);
    sh.getRange(`${c}4`).setFontColor('#808080').setFontSize(10);
    sh.getRange(`${c}5`).setFontWeight('bold').setFontSize(24).setFontColor(col);
    sh.getRange(`${c}4:${c2}5`).setBackground(LGREY);
  });

  // ---- Secondary metrics ----
  const mets = [
    ['Open cost (£)', `=SUMIFS(${R(COL.cost)},${R(COL.status)},"<>Closed")`, '£#,##0', 'B'],
    ['Avg days open', `=IFERROR(ROUND(AVERAGEIFS(${R(COL.days)},${R(COL.status)},"<>Closed"),1),0)`, '0.0', 'D'],
    ['Longest open (days)', `=IFERROR(MAXIFS(${R(COL.days)},${R(COL.status)},"<>Closed"),0)`, '0', 'F'],
    ['Total cost (£)', `=SUM(${R(COL.cost)})`, '£#,##0', 'H'],
  ];
  mets.forEach(([lab, f, fmt, c]) => {
    const c2 = String.fromCharCode(c.charCodeAt(0) + 1);
    sh.getRange(`${c}7:${c2}7`).merge(); sh.getRange(`${c}8:${c2}8`).merge();
    V(`${c}7`, lab); F(`${c}8`, f);
    sh.getRange(`${c}7`).setFontColor('#808080').setFontSize(10);
    sh.getRange(`${c}8`).setFontWeight('bold').setFontSize(16).setFontColor(NAVY).setNumberFormat(fmt);
  });

  // ---- helper to draw a header band ----
  const band = (a1, text) => { sh.getRange(a1).merge();
    sh.getRange(a1.split(':')[0]).setValue(text).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(NAVY); };
  const colHead = (cells) => cells.forEach(([a1, txt]) =>
    sh.getRange(a1).setValue(txt).setFontWeight('bold').setBackground(LBLUE));

  // ---- By department ----
  band('B10:E10', 'By department');
  colHead([['B11', 'Department'], ['C11', 'Total'], ['D11', 'Open'], ['E11', 'Overdue']]);
  DEPTS.forEach((d, i) => { const r = 12 + i;
    V(`B${r}`, d);
    F(`C${r}`, `=COUNTIF(${R(COL.dept)},B${r})`);
    F(`D${r}`, `=COUNTIFS(${R(COL.dept)},B${r},${R(COL.status)},"Open")`);
    F(`E${r}`, `=COUNTIFS(${R(COL.dept)},B${r},${R(COL.status)},"<>Closed",${R(COL.target)},"<"&TODAY(),${R(COL.target)},"<>")`);
  });
  const lastDept = 11 + DEPTS.length;     // last department row
  const deptTotal = lastDept + 1;         // department totals row
  V(`B${deptTotal}`, 'Total'); sh.getRange(`B${deptTotal}`).setFontWeight('bold');
  ['C', 'D', 'E'].forEach(c => { F(`${c}${deptTotal}`, `=SUM(${c}12:${c}${lastDept})`); sh.getRange(`${c}${deptTotal}`).setFontWeight('bold'); });

  const row2 = deptTotal + 2;             // band row for severity (left) + type (right)
  // ---- By severity ----
  band(`B${row2}:D${row2}`, 'By severity');
  colHead([[`B${row2 + 1}`, 'Severity'], [`C${row2 + 1}`, 'Total'], [`D${row2 + 1}`, 'Open']]);
  SEVS.forEach((s, i) => { const r = row2 + 2 + i;
    V(`B${r}`, s);
    F(`C${r}`, `=COUNTIF(${R(COL.sev)},B${r})`);
    F(`D${r}`, `=COUNTIFS(${R(COL.sev)},B${r},${R(COL.status)},"<>Closed")`);
  });

  // ---- By type ----
  band(`G${row2}:I${row2}`, 'By type');
  colHead([[`G${row2 + 1}`, 'Type'], [`H${row2 + 1}`, 'Total']]);
  TYPES.forEach((tp, i) => { const r = row2 + 2 + i;
    V(`G${r}`, tp);
    F(`H${r}`, `=COUNTIF(${R(COL.type)},G${r})`);
  });

  const row3 = row2 + Math.max(SEVS.length, TYPES.length) + 4;  // band row for category + root causes
  // ---- By issue category (auto-listed from the data) ----
  band(`B${row3}:D${row3}`, 'By issue category');
  V(`B${row3 + 1}`, 'Auto-listed, most frequent first'); sh.getRange(`B${row3 + 1}`).setFontColor('#808080').setFontStyle('italic').setFontSize(9);
  F(`B${row3 + 2}`, `=IFERROR(QUERY(${R(COL.issue)},"select Col1, count(Col1) where Col1 is not null group by Col1 order by count(Col1) desc limit 12",0),"No categories recorded yet")`);

  // ---- Most common root causes (recurring) ----
  band(`G${row3}:I${row3}`, 'Most common root causes (recurring)');
  V(`G${row3 + 1}`, 'Auto-listed, most frequent first'); sh.getRange(`G${row3 + 1}`).setFontColor('#808080').setFontStyle('italic').setFontSize(9);
  F(`G${row3 + 2}`, `=IFERROR(QUERY(${R(COL.root)},"select Col1, count(Col1) where Col1 is not null group by Col1 order by count(Col1) desc limit 8",0),"No root causes recorded yet")`);

  // ---- Highlight overdue counts in red when above zero ----
  const ruleDept = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setFontColor(RED).setBold(true)
    .setRanges([sh.getRange(`E12:E${lastDept}`)]).build();
  sh.setConditionalFormatRules([ruleDept]);

  // ---- Widths & tidy ----
  sh.setColumnWidth(1, 24);
  [['B', 168], ['C', 86], ['D', 92], ['E', 92], ['F', 96], ['G', 168], ['H', 86], ['I', 86], ['J', 96], ['K', 96]]
    .forEach(([c, w]) => sh.setColumnWidth(c.charCodeAt(0) - 64, w));
  sh.getRange('B1:K60').setFontFamily('Arial');
  buildMonthlyByDept(ss, src, S);
  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi && Logger.log('Dashboard built on tab: ' + sh.getName());
}


/**
 * Builds the "Monthly by Dept" tab — a count of NCRs per department per month.
 * Called automatically by buildDashboard().
 */
function buildMonthlyByDept(ss, src, S) {
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const colLetter = c => { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; };

  let sh = ss.getSheetByName('Monthly by Dept');
  if (sh) sh.clear(); else sh = ss.insertSheet('Monthly by Dept');
  sh.setHiddenGridlines(true);

  sh.getRange('B1:O1').merge();
  sh.getRange('B1').setValue('NCRs by month and department — ' + YEAR)
    .setFontSize(16).setFontWeight('bold').setFontColor('#1F3864');

  // header row
  sh.getRange('B3').setValue('Department');
  MONTHS.forEach((mn, i) => sh.getRange(3, 3 + i).setValue(mn));
  sh.getRange(3, 15).setValue('Total');
  sh.getRange(3, 2, 1, 14).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#1F3864').setHorizontalAlignment('center');
  sh.getRange('B3').setHorizontalAlignment('left');

  // department rows
  DEPTS.forEach((d, di) => {
    const r = 4 + di;
    sh.getRange(r, 2).setValue(d);
    for (let m = 1; m <= 12; m++) {
      sh.getRange(r, 2 + m).setFormula(
        `=COUNTIFS(${S}${COL.dept}${FIRST}:${COL.dept},B${r},${S}${COL.ts}${FIRST}:${COL.ts},">="&DATE(${YEAR},${m},1),${S}${COL.ts}${FIRST}:${COL.ts},"<"&EDATE(DATE(${YEAR},${m},1),1))`);
    }
    sh.getRange(r, 15).setFormula(`=SUM(C${r}:N${r})`);
  });

  // totals row
  const tr = 4 + DEPTS.length;
  sh.getRange(tr, 2).setValue('Total').setFontWeight('bold');
  for (let c = 3; c <= 15; c++) {
    const L = colLetter(c);
    sh.getRange(tr, c).setFormula(`=SUM(${L}4:${L}${tr - 1})`).setFontWeight('bold');
  }

  // formatting
  sh.getRange(4, 15, DEPTS.length + 1, 1).setFontWeight('bold');        // total column
  sh.getRange(4, 3, DEPTS.length + 1, 13).setNumberFormat('0;-0;"–"').setHorizontalAlignment('center');
  sh.setColumnWidth(1, 24); sh.setColumnWidth(2, 150);
  for (let c = 3; c <= 14; c++) sh.setColumnWidth(c, 46);
  sh.setColumnWidth(15, 62);
  sh.setFrozenRows(3);
  sh.getRange(1, 2, tr, 14).setFontFamily('Arial');
}


// ===== Auto-defaults for new NCRs (added Aug 2026) =====
// Trigger: From spreadsheet -> On form submit -> applyNewNcrDefaults
const NCR_STATUS_COL = 13; // Column M — Status
const NCR_OWNER_COL  = 15; // Column O — Owner

const NCR_OWNERS = {
  'Sales':        'Duncan',
  'Production':   'Gail',
  'EMB':          'Sharon',
  'Stores':       'Ste',
  'Accounts':     'Vanessa',
  'DPD':          'Duncan',
  'H&B':          'Duncan',
  'Logistics':    'Duncan',
  'Suppliers':    'Duncan',
  'Procurement':  'Duncan',
  'IT & Digital': 'Ben',
  'Other':        'Luke H',
  'ISO/QC':       'Yvonne'
};

function applyNewNcrDefaults(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  // Status -> Open
  const statusCell = sheet.getRange(row, NCR_STATUS_COL);
  if (statusCell.getValue() === '') statusCell.setValue('Open');

  // Owner -> department manager
  const dept = String(sheet.getRange(row, 3).getValue()).trim(); // Column C — Department
  const ownerCell = sheet.getRange(row, NCR_OWNER_COL);
  if (ownerCell.getValue() === '' && NCR_OWNERS[dept]) {
    ownerCell.setValue(NCR_OWNERS[dept]);
  }

  // Wrap text + top-align the new row
  sheet.getRange(row, 1, 1, sheet.getLastColumn())
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment('top');
}

function backfillOwnersOnce() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  const last = sheet.getLastRow();
  for (let row = 2; row <= last; row++) { // data starts row 2
    const dept = String(sheet.getRange(row, 3).getValue()).trim();
    const ownerCell = sheet.getRange(row, NCR_OWNER_COL);
    const status = sheet.getRange(row, NCR_STATUS_COL).getValue();
    if (ownerCell.getValue() === '' && NCR_OWNERS[dept] && status !== 'Closed') {
      ownerCell.setValue(NCR_OWNERS[dept]);
    }
  }
}


// ===== NCR Dashboard API (added Aug 2026) =====
// Deploy: Deploy > Manage deployments > pencil > Version: New version > Deploy
//   - Execute as: Me
//   - Who has access: Anyone
// The Web app URL goes into API_URL at the top of ncr-dashboard.html.

const DASH_SHEET_NAME = 'Form Responses 1';

function doGet(e) {
  if (e && e.parameter && e.parameter.photo) return servePhoto_(e.parameter.photo, e.parameter.key);  // photo request — see Photos.gs

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASH_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Find the header row: first of the top 5 rows containing "Timestamp"
  let headerRow = 1;
  const top = sheet.getRange(1, 1, Math.min(5, lastRow), lastCol).getValues();
  for (var h = 0; h < top.length; h++) {
    if (top[h].map(String).indexOf('Timestamp') > -1) { headerRow = h + 1; break; }
  }
  if (lastRow <= headerRow) return jsonOut_({ rows: [] });

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(function(x){ return String(x).trim(); });
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();

  function col(name) { return headers.indexOf(name); }
  const c = {
    ts: col('Timestamp'), email: col('Email address'),
    dept: col('Department'), type: col('NCR Type'),
    cat: col('Issue category'), what: col('What happened?'),
    suspected: col('Suspected cause?'),
    immediate: col('Any immediate action already taken'),
    cust: col('Customer/Supplier'), order: col('Order number'),
    status: col('Status'), sev: col('Severity'), owner: col('Owner'),
    root: col('Root Cause'), contain: col('Containment Action'),
    corrective: col('Corrective Action'), verified: col('Verified by'),
    notes: col('Resolution notes'),
    target: col('Target Complete Date'), closedDate: col('Date closed'),
    ref: col('NCR ref')
  };

  const rows = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (!r[c.ts]) continue;                                     // skip blank rows
    var what = String(r[c.what] || '').trim();
    if (what.toLowerCase() === 'test') continue;                // skip test entries
    rows.push({
      row: headerRow + 1 + i,                                   // sheet row for deep links
      ref: String(r[c.ref] || ''),
      ts: r[c.ts] instanceof Date ? r[c.ts].toISOString() : String(r[c.ts]),
      email: String(r[c.email] || '').trim(),                   // who raised it, column B
      dept: String(r[c.dept] || ''),
      type: String(r[c.type] || ''),
      cat: String(r[c.cat] || ''),
      what: what.slice(0, 1000),
      suspected: String(r[c.suspected] || ''),
      immediate: String(r[c.immediate] || ''),
      photos: photosForRow_(headers, r),                        // photos from column L — see Photos.gs
      cust: String(r[c.cust] || ''),
      order: String(r[c.order] || ''),
      status: String(r[c.status] || ''),
      sev: String(r[c.sev] || ''),
      owner: String(r[c.owner] || ''),
      root: String(r[c.root] || ''),
      contain: String(r[c.contain] || ''),
      corrective: String(r[c.corrective] || ''),
      verified: String(r[c.verified] || ''),
      notes: String(r[c.notes] || ''),
      target: r[c.target] instanceof Date ? r[c.target].toISOString() : String(r[c.target] || ''),
      closed: r[c.closedDate] instanceof Date ? r[c.closedDate].toISOString() : String(r[c.closedDate] || '')
    });
  }

  return jsonOut_({
    generated: new Date().toISOString(),
    sheetUrl: ss.getUrl(),
    gid: sheet.getSheetId(),
    rows: rows
  });
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ===== Phase 2: save NCR completions from the app (added Aug 2026) =====
// After pasting: Deploy > Manage deployments > pencil > Version: New version > Deploy
const APP_KEY = 'tibard-ncr-2026';   // must match APP_KEY in ncr-dashboard.html

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const req = JSON.parse(e.postData.contents);
    if (req.key !== APP_KEY) return jsonOut_({ ok: false, error: 'Not authorised' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(DASH_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // Find the header row (same logic as doGet)
    let headerRow = 1;
    const top = sheet.getRange(1, 1, Math.min(5, lastRow), lastCol).getValues();
    for (var h = 0; h < top.length; h++) {
      if (top[h].map(String).indexOf('Timestamp') > -1) { headerRow = h + 1; break; }
    }
    const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
      .map(function(x){ return String(x).trim(); });
    const refCol = headers.indexOf('NCR ref') + 1;
    if (refCol === 0) return jsonOut_({ ok: false, error: 'NCR ref column not found' });

    // Locate the row: trust req.row if the ref matches, otherwise search by ref
    let row = Number(req.row) || 0;
    const refWanted = String(req.ref || '').trim();
    if (!(row > headerRow && String(sheet.getRange(row, refCol).getValue()).trim() === refWanted)) {
      row = 0;
      const refs = sheet.getRange(headerRow + 1, refCol, lastRow - headerRow, 1).getValues();
      for (var i = 0; i < refs.length; i++) {
        if (String(refs[i][0]).trim() === refWanted) { row = headerRow + 1 + i; break; }
      }
    }
    if (!row) return jsonOut_({ ok: false, error: 'NCR ' + refWanted + ' not found in the register' });

    // Only these fields can be written from the app
    const FIELD_HEADERS = {
      status: 'Status', sev: 'Severity', owner: 'Owner',
      root: 'Root Cause', contain: 'Containment Action',
      corrective: 'Corrective Action',
      target: 'Target Complete Date', verified: 'Verified by',
      notes: 'Resolution notes'
    };

    const u = req.updates || {};

    // ---- Mandatory fields, enforced server-side ----
    // In Progress needs a Severity; Closed needs everything.
    if (u.status === 'Closed' || u.status === 'In Progress') {
      const need = u.status === 'Closed'
        ? { sev: 'Severity', owner: 'Owner', root: 'Root Cause',
            contain: 'Containment Action', corrective: 'Corrective Action',
            target: 'Target Complete Date', verified: 'Verified by',
            notes: 'Resolution notes' }
        : { sev: 'Severity' };
      const missing = [];
      Object.keys(need).forEach(function(k){
        const colIdx = headers.indexOf(need[k]) + 1;
        const existing = colIdx > 0 ? String(sheet.getRange(row, colIdx).getValue()).trim() : '';
        const incoming = (k in u) ? String(u[k]).trim() : '';
        if (!incoming && !existing) missing.push(need[k]);
      });
      if (missing.length) {
        return jsonOut_({ ok: false,
          error: (u.status === 'Closed' ? 'Cannot close — fill in: ' : 'Cannot set In Progress — fill in: ') + missing.join(', ') });
      }
    }
    let written = 0;
    Object.keys(FIELD_HEADERS).forEach(function(k){
      if (!(k in u)) return;
      const colIdx = headers.indexOf(FIELD_HEADERS[k]) + 1;
      if (colIdx === 0) return;
      let val = u[k];
      if (k === 'target' && val) val = new Date(val);   // date input arrives as yyyy-mm-dd
      sheet.getRange(row, colIdx).setValue(val);
      written++;
    });

    // Closing an NCR stamps Date closed if it's empty
    if (u.status === 'Closed') {
      const closedCol = headers.indexOf('Date closed') + 1;
      if (closedCol > 0 && sheet.getRange(row, closedCol).getValue() === '') {
        sheet.getRange(row, closedCol).setValue(new Date());
      }
    }

    return jsonOut_({ ok: true, ref: refWanted, row: row, fieldsWritten: written });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
