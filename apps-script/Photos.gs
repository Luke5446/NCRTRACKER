/**
 * NCR PHOTOS  (Google Apps Script — add this as a new file in the "Dashboard NCR" project)
 * ---------------------------------------------------------------------------
 * Serves the photos managers attach to the NCR form (register column L, "Image")
 * to the dashboard app, so the photo can be seen inside the NCR itself instead
 * of opening the Google Sheet and then Drive.
 *
 * Two jobs:
 *   photosForRow_()  — called from doGet, adds the photo ids to each NCR row
 *   servePhoto_()    — called from doGet when the app asks for one photo,
 *                      returns the image itself as a data: URI
 *
 * The form uploads land in Drive owned by the form owner and are NOT shared
 * with anyone. Rather than making them public, the script (which runs as the
 * owner) reads the file and hands the bytes to the app. Nothing in Drive has to
 * be re-shared, and only file ids that actually appear in the register can be
 * served — see isPhotoInRegister_().
 *
 * Wiring up: see apps-script/README.md in the NCRTRACKER repo — two one-line
 * edits in Code.gs, then Deploy > Manage deployments > New version.
 * ---------------------------------------------------------------------------
 */

// ===== EDIT IF NEEDED ======================================================
const PHOTO_HEADER       = 'Image';   // header of the photo column in the register
const PHOTO_COL_FALLBACK = 12;        // column L — used only if that header is missing
const PHOTO_MAX_WIDTH    = 1400;      // px — photos are resized before they are sent
const PHOTO_MAX_BYTES    = 6 * 1024 * 1024;   // refuse anything bigger than this
const PHOTO_CACHE_SECS   = 21600;     // 6 hours — how long known photo ids stay cached
// ===========================================================================


/**
 * The photos on one register row, as [{id: '...'}] for Drive uploads or
 * [{url: '...'}] for anything already on the open web.
 * Deliberately does no Drive calls — doGet runs this for every row.
 */
function photosForRow_(headers, row) {
  var idx = headers.indexOf(PHOTO_HEADER);
  if (idx === -1) idx = PHOTO_COL_FALLBACK - 1;
  return parsePhotoCell_(row[idx]);
}


/** Splits the cell text (Forms writes one link per upload, comma separated). */
function parsePhotoCell_(cell) {
  var text = String(cell == null ? '' : cell).trim();
  if (!text) return [];
  var parts = text.split(/[\s,;]+/);
  var out = [], seen = {};
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim();
    if (!part) continue;
    var id = driveIdFrom_(part);
    var key = id || part;
    if (seen[key]) continue;
    seen[key] = true;
    if (id) out.push({ id: id });
    else if (/^https?:\/\//i.test(part)) out.push({ url: part });
  }
  return out;
}


/** Pulls the Drive file id out of any of the link shapes Google uses. */
function driveIdFrom_(url) {
  var s = String(url || '').trim();
  var m = s.match(/[?&]id=([-\w]{20,})/);   if (m) return m[1];
  m = s.match(/\/file\/d\/([-\w]{20,})/);   if (m) return m[1];
  m = s.match(/\/d\/([-\w]{20,})/);         if (m) return m[1];
  if (/^[-\w]{25,}$/.test(s)) return s;     // a bare file id
  return '';
}


/**
 * GET ?photo=<file id>&key=<APP_KEY>
 * Returns { ok:true, dataUri:'data:image/jpeg;base64,...' } for the app to drop
 * straight into an <img>, or { ok:false, error:'...' } so it can show a link.
 */
function servePhoto_(fileId, key) {
  try {
    if (key !== APP_KEY) return jsonOut_({ ok: false, error: 'Not authorised' });

    var id = driveIdFrom_(fileId) || String(fileId || '').trim();
    if (!id) return jsonOut_({ ok: false, error: 'No photo reference given' });
    if (!isPhotoInRegister_(id)) return jsonOut_({ ok: false, error: 'That photo is not on the NCR register' });

    var blob  = photoBlob_(id);
    var type  = String(blob.getContentType() || '');
    if (type.indexOf('image/') !== 0) {
      return jsonOut_({ ok: false, notImage: true, error: 'The attachment is not an image (' + (type || 'unknown type') + ')' });
    }

    var bytes = blob.getBytes();
    if (bytes.length > PHOTO_MAX_BYTES) {
      return jsonOut_({ ok: false, tooBig: true, error: 'Photo is too large to show here' });
    }

    return jsonOut_({
      ok: true,
      id: id,
      name: blob.getName() || '',
      type: type,
      dataUri: 'data:' + type + ';base64,' + Utilities.base64Encode(bytes)
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}


/**
 * Reads the file as the sheet owner. Asks Drive for a resized copy first so a
 * 5MB phone photo doesn't get sent whole over 4G, and so iPhone HEIC uploads
 * come back as something a browser can actually display.
 */
function photoBlob_(id) {
  try {
    var res = UrlFetchApp.fetch(
      'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w' + PHOTO_MAX_WIDTH,
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() === 200) {
      var resized = res.getBlob();
      if (String(resized.getContentType() || '').indexOf('image/') === 0) return resized;
    }
  } catch (err) {
    // fall through and send the original file
  }
  return DriveApp.getFileById(id).getBlob();
}


/**
 * Only photos that appear in the register can be served — otherwise the web app
 * would hand out any Drive file the owner can see to anyone who guessed an id.
 */
function isPhotoInRegister_(id) {
  var cache = CacheService.getScriptCache();
  if (cache.get('ncrphoto:' + id) === 'y') return true;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASH_SHEET_NAME);
  if (!sheet) return false;
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) return false;

  var headerRow = photoHeaderRow_(sheet, lastRow, lastCol);
  if (lastRow <= headerRow) return false;
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(function (x) { return String(x).trim(); });
  var idx = headers.indexOf(PHOTO_HEADER);
  var colNum = idx === -1 ? PHOTO_COL_FALLBACK : idx + 1;

  var values = sheet.getRange(headerRow + 1, colNum, lastRow - headerRow, 1).getValues();
  var known = {}, found = false;
  for (var i = 0; i < values.length; i++) {
    var photos = parsePhotoCell_(values[i][0]);
    for (var j = 0; j < photos.length; j++) {
      if (!photos[j].id) continue;
      known['ncrphoto:' + photos[j].id] = 'y';
      if (photos[j].id === id) found = true;
    }
  }
  if (found) cache.putAll(known, PHOTO_CACHE_SECS);
  return found;
}


/** Same header-row hunt doGet uses: first of the top 5 rows holding "Timestamp". */
function photoHeaderRow_(sheet, lastRow, lastCol) {
  var top = sheet.getRange(1, 1, Math.min(5, lastRow), lastCol).getValues();
  for (var h = 0; h < top.length; h++) {
    if (top[h].map(String).indexOf('Timestamp') > -1) return h + 1;
  }
  return 1;
}
