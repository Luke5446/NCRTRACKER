# Apps Script — the register API behind the dashboard

The dashboard (`index.html`) is a static page. Everything it reads and writes goes
through one Google Apps Script web app bound to the **NCR Register (Live)** sheet
(script project: **Dashboard NCR**).

This folder holds the two script files that project needs. Both are complete files
— copy each one in whole, so there is no line-by-line editing to get wrong.

| File | What to do with it |
| --- | --- |
| `Code.gs` | Replaces the existing `Code.gs`. Same file you already have, plus two added lines (both marked `see Photos.gs`). |
| `Photos.gs` | A new file. Nothing to replace. |

`CaseNumbers.gs` in the script project is untouched — leave it exactly as it is.

## What this adds

Photos. Column **L (`Image`)** of the register holds the Drive links to whatever was
attached when the NCR was raised. `Photos.gs` sends those photos to the app so they
appear inside the NCR, instead of a manager having to open the sheet and then Drive.

Suspected cause (**G**) and Any immediate action already taken (**I**) were already
mapped in `doGet` — they only need the deployment refreshed (step 3) to reach the app.

## Wiring it up

**1. Add `Photos.gs`.** In the Apps Script editor: **Files ▸ + ▸ Script**, name it
`Photos`, delete the `myFunction` stub it creates, then paste in the whole of
`Photos.gs`.

**2. Replace `Code.gs`.** Open `Code.gs`, select all (Ctrl+A), delete, then paste in
the whole of `Code.gs` from this folder. Save.

**3. Redeploy.** **Deploy ▸ Manage deployments ▸ ✏️ ▸ Version: New version ▸ Deploy.**
Nothing reaches the app until this is done — the `/exec` URL keeps serving the last
deployed version, which is why suspected cause and immediate action have not shown up
yet.

**4. Re-authorise when asked.** Reading photos needs two permissions the script didn't
use before — Drive (to open the uploaded file) and external requests (to ask Drive for
a resized copy). Google will prompt once; accept as `luke.harrison@tibard.co.uk`, the
account that owns the register.

**5. Check it.** Open **NCR-2026-007** (Suppliers & DPD) in the app — it has two photos
in column L, so both should appear inside the NCR.

## The two added lines in Code.gs

For reference, this is all that differs from the version already in the project.

At the top of `doGet`:

```js
if (e && e.parameter && e.parameter.photo) return servePhoto_(e.parameter.photo, e.parameter.key);
```

And one field inside the `rows.push({ ... })` block:

```js
photos: photosForRow_(headers, r),
```

## September 2026: closing an NCR by hand in the sheet

`Code.gs` gained one function at the bottom, `onEdit`. The app has always refused to
close an NCR until every field is filled in, but the sheet itself would let anyone pick
"Closed" from the dropdown with the row half empty — which is how 181 of the closed NCRs
came to be missing things like *Verified by* and *Resolution notes*. Now, if Status is
set to Closed on a row that is missing any of the eight required fields, the previous
status is put straight back and a note appears on the cell saying what is missing.

**To apply it:** open `Code.gs`, select all, delete, paste in the whole of `Code.gs`
from this folder, and **Save**. That is all — `onEdit` runs from the saved script, not
from the deployment, so **no redeploy is needed** this time, and it does not ask for any
new permissions.

The eight fields, the same as in the app: Severity, Owner, Root Cause, Containment
Action, Corrective Action, Target Complete Date, Verified by, Resolution notes.

The app now also points out the older NCRs that were closed with gaps — on the card,
in the NCR window, and on the monthly summary — so they can be finished off.

## Notes

- **Nothing in Drive gets re-shared.** The script runs as the sheet owner and hands the
  image bytes to the app itself. Form uploads stay private.
- **Only register photos can be served.** `servePhoto_` refuses any file id that isn't
  in column L of the register, so the endpoint can't be used to pull other Drive files.
- Photos are resized to 1400px wide before sending, so they load quickly on the shop
  floor. Anything still over 6MB, or an attachment that isn't an image, is shown in the
  app as a link to Drive instead.
- If the `Image` header is ever renamed, the script falls back to column L.
