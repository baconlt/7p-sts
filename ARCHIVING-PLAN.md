# Season-End Archiving Plan (back-pocket)

> **Status:** NOT STARTED — hold until the season ends and all pay periods before the
> cutoff are locked/paid. Nothing here runs automatically.
> **Owner split:** 🧑 = Mike does it · 🤖 = Claude does it (ask me to when the time comes).

---

## Why we're doing this

Every read in the app goes through `sheetToObjects()` → `getDataRange().getValues()`, which
loads the **entire** sheet into memory and filters in JS. Cost scales with *total rows ever
created*, not rows needed. `Shifts` (~4k/season) and `TimeRecords` (~9k/season) grow forever
and are read in full in 24 and 22 places respectively. Archiving moves closed-season rows out
of the live tabs into `<Name>_Archive` tabs **in the same spreadsheet**. The live read paths
only ever touch the live tab, so once it's small again, every read is fast again. Total file
stays well under the 10M-cell limit for a decade+.

This is the durable fix (finding #1 from the DB audit). Per-request caching (finding #2) is a
separate, smaller change we can do anytime.

---

## What gets archived

| Sheet | Date column | Grows per season | Notes |
|---|---|---|---|
| `Shifts` | `date` | ~4,000+ | ✅ always archive |
| `TimeRecords` | `date` | ~9,000 | ✅ always archive |
| `Availability` | `date` | per guard/day | ✅ always archive |
| `ShiftStats` | `date` | 1/completed shift | ⚠️ **decision below** |
| `Notifications` | `sent_at` | append-only, never read | 🗑 prune, don't archive |

`Sessions` is already self-cleaning (90-day trigger) — leave it alone.

---

## Decisions to confirm before running (🧑 + 🤖)

- [ ] **Cutoff date** — archive everything with `date` **strictly before** this `YYYY-MM-DD`.
      Use the day after the last day of the last fully-closed/paid pay period. Set it in
      `ARCHIVE_CUTOFF` in the script.
- [ ] **Archive `ShiftStats`?** Metric badges & all-time stat leaderboards
      (`computeBadgeLeaderboard_`, admin stats) recompute from `ShiftStats`. If we archive it,
      those views only reflect the **current** live season going forward. Manually-awarded
      badges (Treasure Hunter, etc.) live in the `Badges` sheet and are **safe** either way.
      - If the system is treated as per-season (fresh each year): **archive it.**
      - If you want all-time stats to keep counting historical seasons: **leave `ShiftStats`
        live for now** (remove it from `ARCHIVE_SHEETS`) and revisit later.
- [ ] **Historical reporting** — after archiving, payroll/stat exports for *old* periods no
      longer see those rows in the live sheet. Current + future periods are unaffected. If you
      ever need an old-season report, we read the `_Archive` tab explicitly (small follow-up).

---

## Step-by-step checklist

### Phase 0 — Pre-flight (🧑)
- [ ] **Full backup #1:** Google Sheet → `File ▸ Make a copy` (name it
      `7P STS backup pre-archive <date>`). This is your rollback.
- [ ] Confirm the season is over and every pay period before the cutoff is **locked/paid**.
- [ ] Confirm no upcoming/active shifts exist before the cutoff date (they'd get archived).

### Phase 1 — Install the script (🤖 finalize → 🧑 paste)
- [ ] 🤖 Confirm the script below still matches the schema (Claude re-checks before you run).
- [ ] 🧑 Apps Script editor → open `Code.gs` → paste the **entire ARCHIVE block** (bottom of
      this doc) at the end of the file.
- [ ] 🧑 Edit the two config lines at the top of the block: set `ARCHIVE_CUTOFF` and, if you
      chose to keep stats live, remove `'ShiftStats'` from `ARCHIVE_SHEETS`.
- [ ] 🧑 `Cmd+S` to save. **No deployment needed** — these are editor-run functions, not part
      of the web app. (You do *not* need to touch Manage Deployments for this.)

### Phase 2 — Dry run (🧑, verify with 🤖)
- [ ] 🧑 In the editor's function dropdown, select **`archiveStatus`** → Run. Read
      `View ▸ Logs`. Note the current row counts.
- [ ] 🧑 Select **`archiveDryRun`** → Run. It moves **nothing**; it just logs how many rows
      *would* be archived per sheet.
- [ ] 🤖 Sanity-check those counts against expectations (roughly one season of rows).
- [ ] ⛔ If a sheet reports "⚠ no date column" or a header-mismatch error, **stop** and ping
      Claude before proceeding.

### Phase 3 — Execute (🧑)
- [ ] 🧑 Select **`archiveRun`** → Run. (First run will ask for authorization — allow it.)
- [ ] 🧑 Read the log: each sheet should say `archived N rows → <Name>_Archive (live now M rows)`.
- [ ] 🧑 Run **`archiveStatus`** again → confirm live sheets shrank and `_Archive` tabs hold
      the moved rows. Live + archive counts should equal the original totals.

### Phase 4 — Verify the app (🧑, with 🤖 if anything's off)
- [ ] Open the app → **Full Schedule** for the current period renders normally.
- [ ] A guard dashboard / My Schedule loads (uses `getShiftsForGuard`).
- [ ] **Public Board** (`?board=1`) shows today correctly.
- [ ] Run a payroll CSV export for the **current** period — totals look right.

### Phase 5 — Clean up & record (🧑)
- [ ] 🧑 Run **`pruneNotifications`** (deletes `Notifications` rows older than 30 days).
- [ ] 🧑 **Backup #2:** make another `File ▸ Make a copy` of the now-archived sheet.
- [ ] 🧑 Note here what you did: cutoff used = `__________`, date run = `__________`,
      ShiftStats archived? `Y / N`.

---

## Rollback

- **Fastest:** restore from the **Phase 0 backup copy** (it has every row, pre-archive).
- The moved rows also still exist in the `_Archive` tabs, so nothing is destroyed by the
  process itself — worst case we copy rows back. The script also verifies the archive write
  succeeded *before* it removes anything from the live sheet, and holds a script lock so it
  can't run twice at once.

---

## The ARCHIVE block (ready to paste into `Code.gs`)

```javascript
// ============================================================
// SEASON-END ARCHIVING  (run manually from the editor)
// See ARCHIVING-PLAN.md. Moves rows with date < ARCHIVE_CUTOFF
// out of live sheets into "<Name>_Archive" tabs in the same file.
// Safe: append→verify→rewrite, under a script lock, idempotent by id.
// ============================================================

var ARCHIVE_CUTOFF = '2027-01-01';   // ← EDIT: archive rows with date STRICTLY BEFORE this (YYYY-MM-DD)
var ARCHIVE_SHEETS = ['Shifts','TimeRecords','Availability','ShiftStats']; // remove 'ShiftStats' to keep all-time stats live

function archiveStatus()  {           // report row counts, changes nothing
  const ss = SS();
  const lines = ['ARCHIVE STATUS'];
  ARCHIVE_SHEETS.concat(ARCHIVE_SHEETS.map(n => n + '_Archive')).forEach(n => {
    const sh = ss.getSheetByName(n);
    lines.push('  ' + n + ': ' + (sh ? Math.max(0, sh.getLastRow() - 1) + ' rows' : '(none)'));
  });
  const msg = lines.join('\n'); Logger.log(msg); return msg;
}

function archiveDryRun()  { return archiveSeason_(true);  }   // logs what WOULD move, moves nothing
function archiveRun()     { return archiveSeason_(false); }   // actually archives

function archiveSeason_(dryRun) {
  const cutoff = String(ARCHIVE_CUTOFF).slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(cutoff))
    throw new Error('Set ARCHIVE_CUTOFF to a real YYYY-MM-DD date first.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const report = ['ARCHIVE ' + (dryRun ? 'DRY RUN' : 'RUN') + '  (date < ' + cutoff + ')'];
  try {
    const ss = SS();
    ARCHIVE_SHEETS.forEach(name => {
      const live = ss.getSheetByName(name);
      if (!live) { report.push('  ' + name + ': no live sheet, skipped'); return; }
      const data = live.getDataRange().getValues();
      if (data.length < 2) { report.push('  ' + name + ': empty'); return; }
      const headers = data[0].map(h => String(h).trim());
      const width   = headers.length;
      const dateCol = headers.indexOf('date');
      const idCol   = headers.indexOf('id');
      if (dateCol < 0) { report.push('  ' + name + ': ⚠ no "date" column, skipped'); return; }

      // Partition live rows by the cutoff (single snapshot).
      const moveRows = [], keepRows = [headers];
      for (let i = 1; i < data.length; i++) {
        const d = toYMD(data[i][dateCol]);
        if (d && d < cutoff) moveRows.push(data[i]); else keepRows.push(data[i]);
      }
      if (!moveRows.length) { report.push('  ' + name + ': 0 rows before cutoff'); return; }
      if (dryRun) { report.push('  ' + name + ': would archive ' + moveRows.length +
                                ' rows (live would keep ' + (keepRows.length - 1) + ')'); return; }

      // 1) Ensure archive tab with identical headers.
      const archName = name + '_Archive';
      let arch = ss.getSheetByName(archName);
      if (!arch) {
        arch = ss.insertSheet(archName);
        arch.getRange(1, 1, 1, width).setValues([headers])
            .setBackground('#4b5563').setFontColor('#ffffff').setFontWeight('bold');
        arch.setFrozenRows(1);
      } else {
        const aHead = arch.getRange(1, 1, 1, arch.getLastColumn()).getValues()[0].map(h => String(h).trim());
        if (aHead.slice(0, width).join(' ') !== headers.join(' '))
          throw new Error(archName + ' headers differ from ' + name + ' — align them before archiving.');
      }

      // 2) Idempotent append: skip ids already in the archive.
      let existing = new Set();
      if (idCol >= 0 && arch.getLastRow() > 1) {
        const av = arch.getRange(2, idCol + 1, arch.getLastRow() - 1, 1).getValues();
        av.forEach(r => existing.add(String(r[0])));
      }
      const toAppend = (idCol < 0) ? moveRows
                     : moveRows.filter(r => !existing.has(String(r[idCol])));
      if (toAppend.length) {
        const start = arch.getLastRow() + 1;
        arch.getRange(start, 1, toAppend.length, width).setValues(toAppend);
        SpreadsheetApp.flush();
        if (arch.getLastRow() !== start - 1 + toAppend.length)
          throw new Error(name + ': archive append size mismatch — aborting BEFORE touching live.');
      }

      // 3) Rewrite live with only retained rows (one clear + one write — no per-row deletes).
      live.clearContents();
      live.getRange(1, 1, keepRows.length, width).setValues(keepRows);
      live.getRange(1, 1, 1, width).setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
      SpreadsheetApp.flush();

      report.push('  ' + name + ': archived ' + moveRows.length + ' rows (' +
                  toAppend.length + ' new) → ' + archName +
                  ' | live now ' + (keepRows.length - 1) + ' rows');
    });
  } finally {
    lock.releaseLock();
  }
  const msg = report.join('\n') + '\nDONE.';
  Logger.log(msg);
  return msg;
}

function pruneNotifications() {        // delete Notifications older than 30 days
  const sheet = SH(SHEETS.NOTIFICATIONS);
  if (!sheet || sheet.getLastRow() < 2) return 'Notifications: nothing to prune.';
  const data = sheet.getDataRange().getValues();
  const head = data[0].map(h => String(h).trim());
  const tsCol = head.indexOf('sent_at');
  if (tsCol < 0) return 'Notifications: no sent_at column.';
  const cutoff = new Date(Date.now() - 30 * 864e5);
  let removed = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const t = data[i][tsCol] ? new Date(data[i][tsCol]) : null;
    if (t && t < cutoff) { sheet.deleteRow(i + 1); removed++; }
  }
  const msg = 'Notifications: pruned ' + removed + ' rows older than 30 days.';
  Logger.log(msg); return msg;
}
```

---

## Follow-ups to consider later (not part of this run)
- **Per-request caching** (audit finding #2): memoize `sheetToObjects` for the life of one
  server call — invisible, low-risk, speeds up multi-read handlers immediately.
- **Batch `updateById`** into one `setValues` per row + an `updateManyById` for loop callers
  (audit finding #3).
- **Archive reader** helper if/when an old-season report is needed.
