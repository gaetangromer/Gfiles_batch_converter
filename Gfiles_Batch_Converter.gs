/************ CONFIG (placeholders) ************/
// Destination parent for all jobs (or '' to use My Drive root)
const BACKUP_PARENT_FOLDER_ID = ''; // e.g., 'DEST_PARENT_FOLDER_ID' or ''

// Define source jobs: sourceId (Drive folder), label (destination root name), optional parentId override
const JOBS = {
  Admin: { sourceId: 'SOURCE_ID_ADMIN', label: 'Administration', parentId: '' },
  Prod:  { sourceId: 'SOURCE_ID_PROD',  label: 'Production',     parentId: '' },
  Com:   { sourceId: 'SOURCE_ID_COM',   label: 'Communication',  parentId: '' },
};

// Select which jobs to run this execution
const RUN = ['Admin'];

// Force reconversion for all selected jobs (replace existing versions)
const FORCE = false;

// Keep only the N latest versions per base name in incremental mode (0 = keep all)
const KEEP_LATEST_VERSIONS = 1;

/************ CONSTANTS *************/
const MT = {
  GDOC:   'application/vnd.google-apps.document',
  GSHEET: 'application/vnd.google-apps.spreadsheet',
  GSLIDE: 'application/vnd.google-apps.presentation',
  DOCX:   'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  XLSX:   'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  PPTX:   'application/vnd.openxmlformats-officedocument.presentationml.presentation'
};
const EXPORT_MAP = {
  [MT.GDOC]:   { mime: MT.DOCX,  ext: '.docx'  },
  [MT.GSHEET]: { mime: MT.XLSX,  ext: '.xlsx'  },
  [MT.GSLIDE]: { mime: MT.PPTX,  ext: '.pptx'  }
};

// Incremental memo (per fileId → last seen mtime in ms)
const PROP_PREFIX = 'mtime:'; // key = mtime:<fileId> → <timestamp_ms>
let PROPS;

/************ ENTRYPOINT ************/
function runPermanentMirrorVersioned() {
  if (!Array.isArray(RUN) || RUN.length === 0) {
    throw new Error('RUN is empty. Example: const RUN = ["Admin","Prod"];');
  }

  PROPS = PropertiesService.getScriptProperties();
  // One human timestamp for this run (used in filenames, marker, description, and log)
  const humanTs = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy__HH-mm');

  RUN.forEach(alias => {
    const job = JOBS[alias];
    if (!job) { Logger.log(`Unknown alias: ${alias} — skipped`); return; }

    // Resolve destination parent: job.parentId > BACKUP_PARENT_FOLDER_ID > My Drive root
    let destParent;
    try {
      if (job.parentId && job.parentId.trim()) {
        destParent = DriveApp.getFolderById(job.parentId.trim());
      } else if (BACKUP_PARENT_FOLDER_ID && BACKUP_PARENT_FOLDER_ID.trim()) {
        destParent = DriveApp.getFolderById(BACKUP_PARENT_FOLDER_ID.trim());
      } else {
        destParent = DriveApp; // My Drive root
      }
    } catch (e) { throw new Error(`Destination parent not accessible for "${alias}": ${e}`); }

    // Ensure LABEL folder and detect if it was just created
    const labelName = sanitizeLabel_(job.label || alias);
    const { folder: jobRoot, created: labelJustCreated } = getOrCreateChildFolder_(destParent, labelName);

    // Effective force mode for this job (global FORCE or newly created label)
    const force = !!(FORCE || labelJustCreated);

    const srcRoot  = DriveApp.getFolderById(job.sourceId);
    const counters = { folders: 0, created: 0, skipped: 0, errors: 0 };
    const jobLog = []; // detailed lines (e.g., too large to export)

    processFolderVersioned_(srcRoot, jobRoot, [], humanTs, counters, force, jobLog);

    // Prepend summary as first line in the log
    jobLog.unshift(
      `[${humanTs}] Job=${alias} | Folders=${counters.folders} | Created=${counters.created} | ` +
      `Skipped=${counters.skipped} | Errors=${counters.errors} | Force=${force}`
    );

    // Update label description + write single Last Update marker
    try {
      jobRoot.setDescription(`Last Office export: ${humanTs}`);
      createOrReplaceLastUpdateMarker_(jobRoot, humanTs);
    } catch (e) {
      Logger.log(`Unable to set LastUpdate marker for ${alias}: ${e}`);
    }

    // Write execution log file (unique per run; removes older Execution_Log__*.txt)
    writeExecutionLog_(jobRoot, humanTs, jobLog);

    Logger.log(
      `[${alias}] OK | Folders(created): ${counters.folders} | Created: ${counters.created} | ` +
      `Skipped: ${counters.skipped} | Errors: ${counters.errors} | force=${force}`
    );
  });
}

/************ CORE (lazy tree + incremental/force + logging) ************/
function processFolderVersioned_(srcFolder, jobRoot, pathParts, humanTs, counters, force, jobLog) {
  // 1) Files (convertible Google types only)
  const files = srcFolder.getFiles();
  while (files.hasNext()) {
    const f  = files.next();
    const mt = f.getMimeType();
    const cfg = EXPORT_MAP[mt];
    if (!cfg) { counters.skipped++; continue; }

    const fileId = f.getId();
    const srcUpdatedTs = f.getLastUpdated().getTime();
    const lastSeen = Number(PROPS.getProperty(PROP_PREFIX + fileId)) || 0;

    // Export if forced OR source changed since last run
    const shouldExport = force || (srcUpdatedTs > lastSeen);
    if (!shouldExport) { counters.skipped++; continue; }

    try {
      // Lazily create destination path only when we actually write
      const targetFolder = getFolderAtPath_(jobRoot, pathParts, /*createIfMissing=*/true, counters);
      const baseName = f.getName();
      const outName  = `${baseName}_${humanTs}${cfg.ext}`;

      if (force) {
        // Replace mode: remove all older versions for this base file name
        cleanupOldVersions_(targetFolder, baseName, cfg.ext, /*keepN=*/0);
      }

      const blob = exportViaDriveV3_(fileId, cfg.mime).setName(outName);
      targetFolder.createFile(blob);
      counters.created++;

      // Record last seen mtime (useful even in force mode to resume incrementally next time)
      PROPS.setProperty(PROP_PREFIX + fileId, String(srcUpdatedTs));

      // In incremental mode, keep only N latest versions if configured
      if (!force && KEEP_LATEST_VERSIONS > 0) {
        cleanupOldVersions_(targetFolder, baseName, cfg.ext, KEEP_LATEST_VERSIONS);
      }
    } catch (e) {
      // Detect “too large to export” (HTTP 403)
      const msg = String(e);
      const isTooLarge = (msg.indexOf('This file is too large to be exported') !== -1) ||
                         (msg.indexOf('(403)') !== -1 && msg.toLowerCase().indexOf('too large') !== -1);

      if (isTooLarge) {
        const relPath = formatRelPath_(pathParts, f.getName());
        const line = `[${humanTs}] SKIPPED TOO LARGE | Path=${relPath} | URL=${f.getUrl()} | Id=${f.getId()}`;
        if (jobLog) jobLog.push(line);
        counters.skipped++;
      } else {
        counters.errors++;
        Logger.log(`Error: ${f.getName()} (${fileId}): ${e}`);
      }
      Utilities.sleep(150);
    }
  }

  // 2) Subfolders (depth-first; create lazily on write)
  const sub = srcFolder.getFolders();
  while (sub.hasNext()) {
    const s = sub.next();
    processFolderVersioned_(s, jobRoot, pathParts.concat([s.getName()]), humanTs, counters, force, jobLog);
  }
}

/************ HELPERS ************/
// Ensure a single child folder; return { folder, created }
function getOrCreateChildFolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return { folder: it.next(), created: false };
  return { folder: parentFolder.createFolder(name), created: true };
}

// Resolve/create nested path under root; increment counters.folders on creations
function getFolderAtPath_(rootFolder, pathParts, createIfMissing, counters) {
  let cur = rootFolder;
  for (let i = 0; i < pathParts.length; i++) {
    const seg = pathParts[i];
    const it = cur.getFoldersByName(seg);
    if (it.hasNext()) {
      cur = it.next();
    } else {
      if (!createIfMissing) return null;
      cur = cur.createFolder(seg);
      if (counters) counters.folders++;
    }
  }
  return cur;
}

// Remove older “base_*.ext” versions, keep `keepN` newest ones
function cleanupOldVersions_(folder, baseName, ext, keepN) {
  const prefix = baseName + '_';
  const cand = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const n = f.getName();
    if (n.startsWith(prefix) && n.endsWith(ext)) cand.push(f);
  }
  cand.sort((a, b) => b.getLastUpdated().getTime() - a.getLastUpdated().getTime());
  const start = Math.max(keepN, 0);
  for (let i = start; i < cand.length; i++) cand[i].setTrashed(true);
}

// Build relative path "Sub1/Sub2/FileName"
function formatRelPath_(pathParts, fileName) {
  return (pathParts && pathParts.length ? pathParts.join('/') + '/' : '') + fileName;
}

// Write single execution log for this run (removes older Execution_Log__*.txt)
function writeExecutionLog_(folder, humanTs, lines) {
  // Remove previous logs
  const all = folder.getFiles();
  while (all.hasNext()) {
    const f = all.next();
    const n = f.getName();
    if (n.startsWith('Execution_Log__') && n.endsWith('.txt')) f.setTrashed(true);
  }
  // Create this run's log
  const logName = `Execution_Log__${humanTs}.txt`;
  const content = (lines || []).join('\n') + '\n';
  folder.createFile(logName, content);
}

// Single marker file: ze_Last_Update__DD-MM-YYYY__HH-mm.txt
function createOrReplaceLastUpdateMarker_(folder, humanTs) {
  const markerName = `ze_Last_Update__${humanTs}.txt`;
  // Remove older markers
  const all = folder.getFiles();
  const toTrash = [];
  while (all.hasNext()) {
    const f = all.next();
    const n = f.getName();
    if (n.startsWith('ze_Last_Update__') && n.endsWith('.txt')) toTrash.push(f);
  }
  toTrash.forEach(f => f.setTrashed(true));
  // Create the new one
  folder.createFile(markerName, `Last export: ${humanTs}`);
}

// Export a Google file via Drive API v3 to the requested MIME → Blob
function exportViaDriveV3_(fileId, mime) {
  const url = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=${encodeURIComponent(mime)}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code < 200 || code > 299) {
    throw new Error(`Drive export v3 failed (${code}): ${res.getContentText().slice(0,200)}`);
  }
  return res.getBlob();
}

// Sanitize label for Drive folder naming
function sanitizeLabel_(s) {
  return String(s).replace(/[\/\\:\*\?"<>\|]/g, ' ').replace(/\s{2,}/g, ' ').trim();
}

/************ UTIL ************/
// Clear incremental cache (forces a full incremental recalculation next runs)
function resetIncrementCache() {
  PROPS = PropertiesService.getScriptProperties();
  const all = PROPS.getProperties();
  Object.keys(all).forEach(k => { if (k.startsWith(PROP_PREFIX)) PROPS.deleteProperty(k); });
}
