# Gfiles_batch_converter
Google Apps Script to batch convert Google Docs, Sheets and Slides into Office formats (.docx, .xlsx, .pptx). Maintains incremental updates, versioning, logs, and a permanent mirrored folder structure in Drive.

Here's a crisp, GitHub-ready **README.md** in English you can drop in as-is.

---

# Permanent Mirror — Google → Office (Apps Script)

Maintain a **per-job permanent folder** (a “label”) that contains **only Office conversions** of Google Drive files:

* Google Docs → `.docx`
* Google Sheets → `.xlsx`
* Google Slides → `.pptx`

The script runs **incrementally** (only reconverts changed sources), supports a **global force/replace** mode, writes a **single run log** per execution, and adds a **Last Update marker** file.

---

## Features

* **Permanent label folder** per job (no date-stamped snapshot folders)
* **Incremental export** based on each source file’s last modified time
* **Force mode**: reconvert everything and replace previous versions
* **Sparse tree**: only creates destination folders that actually receive at least one converted file
* **Versioned outputs** with run timestamp appended to filenames
  `MyDoc_03-10-2025__16-45.docx`
* **Run marker** file (always one):
  `ze_Last_Update__DD-MM-YYYY__HH-mm.txt`
* **Execution log** (always one per run):
  `Execution_Log__DD-MM-YYYY__HH-mm.txt`
  Includes a summary and a line for each **oversized** source the API refused to export (path + URL + ID)

---

## How it works (high-level)

1. For each selected **job**:

   * Resolve/create the **label** folder (destination root for that job).
     If the label folder didn’t exist, this run is **forced (replace)** for that job.
   * Walk the source tree; for each Google file type:

     * If **changed** since last run → export to Office and write into the mirrored path (created lazily).
     * If **unchanged** → skip.
     * If **export fails due to size (HTTP 403)** → skip and log a detailed line (relative path + URL + ID).
   * Update the label’s **description** with the human timestamp.
   * Write the **Execution_Log__*.txt** for this run (and remove previous logs).
   * Create/replace the **ze_Last_Update__*.txt** marker file.

Incremental state is stored in `ScriptProperties` (`mtime:<fileId>`).

---

## Requirements

* A Google account with access to the source Drive folders.
* **Google Apps Script** (V8 runtime) project.
* First run will prompt for authorization (Drive access).

No advanced services needed (the script uses `UrlFetchApp` with an OAuth token for Drive v3 exports).

---

## Installation

1. Open **script.google.com** → New project.
2. Paste the script into `Code.gs`.
3. Ensure **V8** runtime (Project Settings → Script runtime).
4. Save.

---

## Configuration

In the **CONFIG** section:

```js
const BACKUP_PARENT_FOLDER_ID = ''; // Destination parent (or '' for My Drive root)

const JOBS = {
  Admin: { sourceId: 'SOURCE_ID_ADMIN', label: 'Administration', parentId: '' },
  Prod:  { sourceId: 'SOURCE_ID_PROD',  label: 'Production',     parentId: '' },
  Com:   { sourceId: 'SOURCE_ID_COM',   label: 'Communication',  parentId: '' },
};

const RUN = ['Admin', 'Prod']; // Which jobs to run this execution
const FORCE = false;           // true = reconvert everything and replace
const KEEP_LATEST_VERSIONS = 1;// keep N most recent versions per file (0 = keep all)
```

### Finding a Google Drive folder ID

Open the folder in your browser:

```
https://drive.google.com/drive/folders/XXXXXXXXXXXX
```

Use the string after `/folders/` as the ID: `XXXXXXXXXXXX`.

---

## Running

* From the Apps Script editor, run:
  `runPermanentMirrorVersioned()`
* On the first run, grant the requested permissions.

You can add a **time-based trigger** (cron-like) if you want scheduled runs.

---

## Outputs & Structure

```
LABEL/
  ├─ …/ (subfolders only created if at least one conversion lands there)
  ├─ MyDoc_03-10-2025__16-45.docx
  ├─ MySheet_03-10-2025__16-45.xlsx
  ├─ MySlides_03-10-2025__16-45.pptx
  ├─ ze_Last_Update__03-10-2025__16-45.txt
  └─ Execution_Log__03-10-2025__16-45.txt
```

* **Filenames**: timestamp appended with `DD-MM-YYYY__HH-mm`.
* **Marker file**: a single `ze_Last_Update__…txt` (replaced every run).
* **Execution log**: a single `Execution_Log__…txt` per run; older logs are removed.

**Log contents example:**

```
[03-10-2025__16-45] Job=Admin | Folders=5 | Created=12 | Skipped=4 | Errors=0 | Force=false
[03-10-2025__16-45] SKIPPED TOO LARGE | Path=Projects/BigDeck/SlidesMaster | URL=https://drive.google.com/open?id=... | Id=...
```

---

## Incremental logic

* The script stores `mtime:<fileId> = <lastModifiedMs>` in `ScriptProperties`.
* A file is reconverted only if `sourceLastModified > storedLastModified`, **unless**:

  * `FORCE = true`, or
  * the **label** folder was just created (first run or re-created) → **forced replace** for that job.

---

## Force mode

* `FORCE = true` → every selected job runs in **replace** mode:

  * existing converted versions for a given basename are **deleted** before writing the new one.
* When `FORCE = false`:

  * only **changed** files are converted,
  * optional pruning keeps only `KEEP_LATEST_VERSIONS` per basename.

---

## Handling “file too large to be exported” (HTTP 403)

When Google’s API refuses an export due to size, the script:

* **skips** the item and
* writes a **detailed log line** with the **relative path**, **URL**, and **file ID** in `Execution_Log__…txt`.

> Note: Google enforces size limits for exporting Google Docs/Sheets/Slides via API. For very large files, consider manual export or breaking content into smaller parts.

---

## Permissions / Scopes

Apps Script will request Drive access on first run (to read sources and create files/folders in the destination). `UrlFetchApp` uses the script’s OAuth token to call the Drive v3 export endpoint.

---

## Troubleshooting

* **403 “This file is too large to be exported.”**
  Expected for very large Docs/Sheets/Slides. See **Handling “file too large…”** above; check the execution log for the exact path and URL.
* **Nothing appears in the destination**
  Confirm you configured `RUN`, `JOBS`, and the destination parent IDs, and that the account has access.
* **Backticks look “red” in the editor**
  Ensure the project uses **V8** runtime and that you used **` (backticks)**, not typographic quotes.

---

## License

MIT
