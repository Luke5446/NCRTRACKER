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
    ref: col('NCR ref'),
    origDept: col(REASSIGN_FROM_HEADER), reassigned: col(REASSIGN_LOG_HEADER)   // see "Reassigning an NCR" below
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
      closed: r[c.closedDate] instanceof Date ? r[c.closedDate].toISOString() : String(r[c.closedDate] || ''),
      origDept: c.origDept > -1 ? String(r[c.origDept] || '') : '',        // department it was first raised against
      reassigned: c.reassigned > -1 ? String(r[c.reassigned] || '') : ''   // who moved it, when and why
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

    // Moving the NCR to another department — see "Reassigning an NCR" below
    if (req.action === 'reassign') return reassignNcr_(sheet, headers, row, refWanted, req);

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
        : { sev: 'Severity', target: 'Target Complete Date' };
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


// ===== Closing an NCR by hand in the sheet (added Sep 2026) =====
// The app will not close an NCR until every one of these is filled in, but
// the sheet itself would let anyone pick "Closed" from the dropdown. This runs
// on every hand edit to the register: if Status is set to Closed on a row that
// is still missing any of them, the previous status is put back and a note on
// the cell says what is missing. Saves made from the app come in through
// doPost and are checked there, not here.
// Nothing to install — a function called onEdit runs on its own.
const CLOSE_NEEDS = ['Severity', 'Owner', 'Root Cause', 'Containment Action',
  'Corrective Action', 'Target Complete Date', 'Verified by', 'Resolution notes'];

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== DASH_SHEET_NAME) return;

    const lastCol = sheet.getLastColumn();
    let headerRow = 1;
    const top = sheet.getRange(1, 1, Math.min(5, sheet.getLastRow()), lastCol).getValues();
    for (var h = 0; h < top.length; h++) {
      if (top[h].map(String).indexOf('Timestamp') > -1) { headerRow = h + 1; break; }
    }
    const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
      .map(function(x){ return String(x).trim(); });
    // Department changed by hand in the sheet: keep the same record a move from
    // the app would, so the NCR still says where it came from.
    const deptCol = headers.indexOf('Department') + 1;
    if (deptCol && e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 &&
        e.range.getColumn() === deptCol && e.range.getRow() > headerRow &&
        e.oldValue !== undefined && e.oldValue !== null) {
      const from = String(e.oldValue).trim(), to = String(e.value || '').trim();
      if (from && to && from !== to) {
        const who = (e.user && e.user.getEmail && e.user.getEmail()) || 'edited in the sheet';
        noteReassignment_(sheet, headers, e.range.getRow(), from, to, who, 'changed in the sheet');
      }
    }

    const statusCol = headers.indexOf('Status') + 1;
    if (!statusCol) return;

    // only edits that touch the Status column matter
    const c1 = e.range.getColumn(), c2 = c1 + e.range.getNumColumns() - 1;
    if (statusCol < c1 || statusCol > c2) return;

    const r1 = e.range.getRow(), n = e.range.getNumRows();
    const single = n === 1 && e.range.getNumColumns() === 1;
    for (var i = 0; i < n; i++) {
      const row = r1 + i;
      if (row <= headerRow) continue;
      const cell = sheet.getRange(row, statusCol);
      if (String(cell.getValue()).trim() !== 'Closed') continue;

      const blank = function(name){
        const c = headers.indexOf(name) + 1;
        return !c || String(sheet.getRange(row, c).getValue()).trim() === '';
      };
      const missing = CLOSE_NEEDS.filter(blank);
      if (!missing.length) { cell.clearNote(); continue; }

      // put back what it was; for a pasted block there is no "was", so make a sensible guess
      let back = single && e.oldValue !== undefined && e.oldValue !== null ? String(e.oldValue) : '';
      if (!back || back === 'Closed') back = (!blank('Severity') && !blank('Target Complete Date')) ? 'In Progress' : 'Open';
      cell.setValue(back);
      cell.setNote('Not closed \u2014 still needs: ' + missing.join(', ') + '.\n' +
        'An NCR can only be closed once every one of these is filled in, the same as in the app.');
    }
  } catch (err) {
    // never get in the way of someone editing the sheet
  }
}


// ===== Reassigning an NCR to another department (added Oct 2026) =====
// A manager can move an open NCR that was raised against the wrong department.
// The row itself moves: Department (column C) becomes the new department, so
// every count — the Dashboard tab, Monthly by Dept, and the app's department
// and monthly figures — picks it up under the new department and drops it from
// the old one. The manager moving it picks the Issue category again. The
// investigation written so far (Severity, Target date, Root cause, Containment,
// Corrective action, Verified by, Resolution notes) is cleared so the new
// department's manager fills it in afresh; what was there is kept in the
// Reassignment history line. Status goes back to Open and Owner becomes the new
// department's manager. A Closed NCR cannot be moved — reopen it first.
//
// Two columns are added on the right of the register the first time a move
// happens (after Notified, column AA). Existing columns are never moved:
//   "Original department"  — the department it was first raised against, set once
//   "Reassignment history" — one line per move: when, from → to, by whom, why,
//                             plus the investigation that was cleared
// After pasting: Deploy > Manage deployments > pencil > Version: New version > Deploy
const REASSIGN_FROM_HEADER = 'Original department';
const REASSIGN_LOG_HEADER  = 'Reassignment history';
const REASSIGN_CLEARS = ['Severity', 'Target Complete Date', 'Root Cause', 'Containment Action',
  'Corrective Action', 'Verified by', 'Resolution notes', 'Date closed'];

function reassignNcr_(sheet, headers, row, ref, req) {
  const to = String(req.to || '').trim();
  const by = String(req.by || '').trim();
  const reason = String(req.reason || '').trim();
  const cat = String(req.cat || '').trim();
  if (DEPTS.indexOf(to) === -1) return jsonOut_({ ok: false, error: 'Unknown department: ' + to });
  if (!cat) return jsonOut_({ ok: false, error: 'Pick an issue category' });
  if (!by) return jsonOut_({ ok: false, error: 'Say who is moving it' });
  if (!reason) return jsonOut_({ ok: false, error: 'Give a reason for the move' });

  const colOf = function(name){ return headers.indexOf(name) + 1; };
  const deptCol = colOf('Department');
  if (!deptCol) return jsonOut_({ ok: false, error: 'Department column not found' });
  const from = String(sheet.getRange(row, deptCol).getValue()).trim();
  if (from === to) return jsonOut_({ ok: false, error: ref + ' is already against ' + to });

  const statusCol = colOf('Status');
  const status = statusCol ? String(sheet.getRange(row, statusCol).getValue()).trim() : '';
  if (status === 'Closed') return jsonOut_({ ok: false, error: ref + ' is closed. Reopen it (set Status to Outstanding and save), then move it.' });

  // keep what the old department wrote, then clear it for the new one
  const kept = [];
  REASSIGN_CLEARS.forEach(function(name){
    const c = colOf(name); if (!c) return;
    const cell = sheet.getRange(row, c);
    const v = cell.getValue();
    const text = v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'd MMM yyyy') : String(v == null ? '' : v).trim();
    if (text) kept.push(name + ': ' + text);
    cell.clearContent();
  });
  const detail = reason + (kept.length ? ' | Cleared on move (' + from + '): ' + kept.join(' | ') : '');

  sheet.getRange(row, deptCol).setValue(to);
  const catCol = colOf('Issue category');
  if (catCol) sheet.getRange(row, catCol).setValue(cat);
  if (statusCol) sheet.getRange(row, statusCol).setValue('Open');
  noteReassignment_(sheet, headers, row, from, to, by, detail);

  // Owner follows the department
  const ownerCol = colOf('Owner');
  let owner = '';
  if (ownerCol) {
    owner = NCR_OWNERS[to] || '';
    sheet.getRange(row, ownerCol).setValue(owner);
  }
  return jsonOut_({ ok: true, ref: ref, row: row, from: from, to: to, owner: owner });
}

// Writes the record of a move. Adds the two columns if the register does not
// have them yet. headers is the header row as read by the caller; a column
// added here is appended after the last one.
function noteReassignment_(sheet, headers, row, from, to, by, reason) {
  const fromCol = ensureColumn_(sheet, headers, REASSIGN_FROM_HEADER);
  const logCol  = ensureColumn_(sheet, headers, REASSIGN_LOG_HEADER);

  const origCell = sheet.getRange(row, fromCol);
  if (String(origCell.getValue()).trim() === '') origCell.setValue(from);

  const when = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy HH:mm');
  const line = when + ' \u2014 ' + from + ' \u2192 ' + to + ' by ' + by + ': ' + reason;
  const logCell = sheet.getRange(row, logCol);
  const existing = String(logCell.getValue()).trim();
  logCell.setValue(existing ? existing + '\n' + line : line)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('top');
}

// Column number of a header, adding it at the right-hand end if it is missing.
function ensureColumn_(sheet, headers, name) {
  const idx = headers.indexOf(name);
  if (idx > -1) return idx + 1;
  const lastCol = sheet.getLastColumn();
  sheet.getRange(findHeaderRow_(sheet), lastCol + 1).setValue(name).setFontWeight('bold');
  headers.push(name);               // so a second call in the same request sees it
  return lastCol + 1;
}

function findHeaderRow_(sheet) {
  const lastCol = sheet.getLastColumn();
  const top = sheet.getRange(1, 1, Math.min(5, sheet.getLastRow()), lastCol).getValues();
  for (var h = 0; h < top.length; h++) {
    if (top[h].map(String).indexOf('Timestamp') > -1) return h + 1;
  }
  return 1;
}
