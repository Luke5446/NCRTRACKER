# Apps Script — the register API behind the dashboard

The dashboard (`index.html`) is a static page. Everything it reads and writes goes
through one Google Apps Script web app bound to the **NCR Register (Live)** sheet
(script project: **Dashboard NCR**).

That script is not stored in this repo — it lives in Google. This folder holds the
one file that needs adding to it, plus the two one-line edits that hook it up.

## What this adds

Photos. Column **L (`Image`)** of the register holds the Drive links to whatever
was attached when the NCR was raised. `Photos.gs` sends those photos to the app so
they appear inside the NCR, instead of a manager having to open the sheet and then
Drive.

Suspected cause (**G**) and Any immediate action already taken (**I**) are already
mapped in `doGet` — they just need the deployment refreshed (step 4 below) to reach
the app.

## Wiring it up

**1. Add the file.** Apps Script editor → **Files ▸ +  ▸ Script**, name it `Photos`,
delete the stub `myFunction`, paste in the whole of `Photos.gs`.

**2. Let `doGet` answer photo requests.** In `Code.gs`, add one line at the very top
of `doGet`:

```js
function doGet(e) {
  if (e && e.parameter && e.parameter.photo) return servePhoto_(e.parameter.photo, e.parameter.key);   // <-- add
  const ss = SpreadsheetApp.getActiveSpreadsheet();
```

**3. Put the photo ids on every NCR row.** Still in `Code.gs`, inside the
`rows.push({ ... })` block, add one line after the `immediate:` line:

```js
      immediate: String(r[c.immediate] || ''),
      photos: photosForRow_(headers, r),          // <-- add
      cust: String(r[c.cust] || ''),
```

**4. Redeploy.** **Deploy ▸ Manage deployments ▸ ✏️ ▸ Version: New version ▸ Deploy**.
Nothing reaches the app until this is done — the `/exec` URL keeps serving the last
deployed version, which is why suspected cause and immediate action may not have
shown up yet.

**5. Re-authorise when asked.** Reading photos needs two permissions the script
didn't use before — Drive (to open the uploaded file) and external requests (to ask
Drive for a resized copy). Google will prompt once on the next run; accept as
`luke.harrison@tibard.co.uk`, the account that owns the register.

## Notes

- **Nothing in Drive gets re-shared.** The script runs as the sheet owner and hands
  the image bytes to the app itself. Form uploads stay private.
- **Only register photos can be served.** `servePhoto_` refuses any file id that
  isn't in column L of the register, so the endpoint can't be used to pull other
  Drive files.
- Photos are resized to 1400px wide before sending, so phone photos load quickly on
  the shop floor. Anything still over 6MB, or an attachment that isn't an image, is
  shown in the app as a link to Drive instead.
- If the `Image` header is ever renamed, the script falls back to column L.
