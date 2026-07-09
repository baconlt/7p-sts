# 7 Presidents STS — Lifeguard Scheduling System

## Architecture
- Google Apps Script web app (Code.gs + index.html + Board.html)
- Google Sheets backend
- PWA wrapper hosted on GitHub Pages (baconlt.github.io/7p-sts)

## Apps Script files (apps-script/)
- `Code.gs` — all server-side logic, auth, CRUD
- `index.html` — entire frontend SPA (HTML + CSS + JS in one file)
- `Board.html` — public no-auth beach assignments display page (served at `?board=1`)

## ⚠️ Deployment — no clasp, manual upload required
Changes are NOT live until manually uploaded. After any edit:
1. Open the Apps Script editor
2. Paste updated file(s) into the corresponding tab
3. Save (Cmd+S)
4. Deploy → Manage Deployments → edit the existing deployment → Deploy
- If you added/changed triggers in Code.gs, also run `installTriggers()` once from the editor
- Board.html is a separate file/tab in the editor — create it if it doesn't exist yet

Clipboard shortcut: `! cat apps-script/<file> | pbcopy`

## PWA Wrapper (pwa/)
- Thin iframe wrapper for PWA install + session persistence
- Deploy: push to GitHub repo baconlt/7p-sts (GitHub Pages)
- Bump CACHE version in sw.js on each deploy

## Auth
- Password-only (no Google OAuth)
- Sessions stored in Sessions sheet (90-day expiry)
- PWA wrapper stores token in localStorage, passes to iframe via postMessage
- `Board.html` is intentionally public — no auth check in `getBoardData()`

## Key conventions
- Admin server calls use `arun()` (prepends session token as first arg)
- Guard server calls use `runG()` (prepends session token as first arg)
- Unauthenticated server calls use `run()` (no token)
- All dates stored as YYYY-MM-DD strings
- Work week is Saturday–Friday
- Pay periods follow county schedule
- After editing a shift, re-render the active panel: `if(PANEL==='schedule') renderSchedGrid(); else if(PANEL==='full-schedule-admin') R['full-schedule-admin']();`

## Admin panels (R[panel] pattern)
- `dashboard` — admin home
- `schedule` — RETIRED: removed from nav/entry points (Full Schedule is used for scheduling). `R['schedule']`/`renderSchedGrid` remain as dormant, unreachable code; the `if(PANEL==='schedule')` refresh hooks now fall through to `full-schedule-admin`
- `full-schedule-admin` — by-post or by-guard full schedule; toggle via `FSA_VIEW`
- `board` — preview + copy URL for the public Board display
- `shifts`, `time-records`, `guards`, `posts`, `periods`, `locations`, `config`

## Guard panels
- `dashboard` — guard home; shifts override day-off availability entries in the list
- `my-avail`, `my-schedule`, `full-schedule` (read-only grid)

## Public Board (`?board=1`)
- Served by `doGet` when `e.parameter.board === '1'`
- Calls `getBoardData(dateStr)` — no token required
- Shows officers (Lifeguard Supervisor / Crew Captain / Captain / Beach Captain) + post cards
- Admin can get the URL via the 📺 Board nav tab → "Copy URL" button

## Sheets
Guards, Posts, ShiftTemplates, Shifts, Availability, PayPeriods,
Config, Notifications, TimeRecords, ShiftStats, Sessions, Locations,
Badges, Announcements, GuardNotes, ScheduleState, CacheHunt
- `Announcements` (id, message, created_at, created_by) is auto-created on first read via `ensureAnnouncementsSheet_()` — no manual setup needed. Admin-posted feed shown on all guard dashboards.
- `GuardNotes` (id, guard_id, date, tags, note, created_by, created_at) is auto-created via `ensureGuardNotesSheet_()`. Admin-only conduct/attendance log per guard (tags: LATE, NO SHOW, COMMENDATION, WARNING, DISCIPLINARY, NOTE + freeform). Reached via the 📝 Notes button on each Guard Roster row. Never surfaced to guards.
- `ScheduleState` (week_start, status, published_at, published_by) is auto-created via `ensureScheduleStateSheet_()`. One row per Saturday work-week; `status` is `draft` (default/missing) or `published`. Drives **draft/publish visibility**: guards/public see a shift only if its Saturday week is published; admins always see everything.
- `Shifts.sick` (`'true'`/`''`) — sick-day flag on an assigned shift. Column added on demand by `ensureShiftColumns_()` (called from `updateShift` whenever a `sick` value is written); also in the `Shifts` schema for fresh installs. Admin-only — guard/public views don't render it.
- `CacheHunt` (guard_id, last_ping_date, pings, best_ft, found_at) is auto-created via `ensureCacheHuntSheet_()`. One row per guard, tracking the **daily distance-reading gate** (`last_ping_date`) + `best_ft`. `found_at` is now unused (winning is code-based, tracked via the badge) — kept for schema stability.

## Cache Hunt (GPS treasure hunt easter egg)
- Silent-rollout geocaching game. Admin hides a physical cache holding a **secret code** and records its GPS coords. Guards get **one distance reading per day** (distance-only hint, no direction) from their dashboard to close in, then physically find the cache and **enter the code to win**. GPS never awards anything — the code is the sole win condition.
- **Winning is code-only + single-winner-per-round.** First guard to submit the correct code wins the one-time 🗺️ **Treasure Hunter** badge (a `manual:true`+`secret:true` BADGE_DEF — event-driven, not metric-computed; hidden from the badge grid until earned via the `secret` flag filter in `renderBadgeGrid`). `getBadgeState`/`computeBadgeLeaderboard_` read manual badges straight from the persisted `Badges` sheet, never calling `earned()`. **The badge row IS the claim** — `cacheHuntWinner_()` finds the sole `cache_hunt` badge holder; once it exists the round is "claimed" and further correct entries are rejected (a `LockService` script lock guards against a simultaneous double-claim). Reset removes the badge → round reopens.
- **Entry point** = a `.cache-card` on the guard dashboard (both `renderGuardDash` + `renderMobGuardDash`, after `weatherCardHtml()`). Card is **silent unless a hunt is live** (active toggle + coords + code set). States: hunt open / you-won / claimed-by-crewmate. Hidden during GV/impersonation.
- **Config keys** (`Config` sheet): `cachehunt_active` (`'true'`/`''`), `cachehunt_lat`, `cachehunt_lng`, `cachehunt_code` (secret code; matched case-insensitively via `normalizeCacheCode_`), `cachehunt_prize` (message shown to the winner). Managed via the **🧭 Cache Hunt card in the admin Config panel** — "📍 Use my location" button (stand at the cache, tap to capture coords), the current winner, and "↻ Reset round".
- **Server** (`Code.gs`): `clientCacheHuntStatus` (guard — active/iWon/claimed/winnerName/pingedToday), `clientCacheHuntPing(token,lat,lng)` (guard — daily-gated distance reading via `distanceFeet_` haversine; awards nothing), `clientCacheHuntClaim(token,code)` (guard — validates code under a script lock, awards the badge to the first correct claimer), `clientCacheHuntAdmin`/`clientCacheHuntSave(token,active,lat,lng,code,prize)`/`clientCacheHuntReset` (admin). `awardManualBadge_(gid,key,earnedAt)` is the idempotent event-badge writer. No migration needed — sheet + config keys auto-create.
- **Client** (`index.html`): `CACHEHUNT` global loaded via `loadCacheHunt()` in guard init; `cacheHuntCardHtml()`, `cacheHuntOpen()` (reading button + inline code-entry form), `cacheHuntGetReading()`→`cacheHuntShowReading()` (renders distance inline, keeping the code form), `cacheHuntClaim()`→`cacheHuntClaimResult()`.

## Sick days
- Marked per-shift-instance (never propagated to a recurring series). Admin edits an assigned shift in Full Schedule (desktop popover `editCellShift` / mobile `showEditShiftBSheet`) → "🤒 Mark as sick day" checkbox → Save. Shared save path: `_commitShiftEdit(shiftId,postId,tmplCode,sick)` in index.html.
- On turning sick **on**, the save first calls `clientShiftHasRecordedHours` (server `shiftHasRecordedHours_` — matches a TimeRecord by `shift_id` or same guard+date with logged/open time). If hours exist, a "Save as SICK anyway?" modal gates the commit; otherwise it saves straight through.
- Grid + shift-list rows show a 🤒 marker (and "SICK" label in the click popover) for flagged shifts.
- **Reporting**: `exportTimeRecordsCSV` (the payroll CSV) labels the Notes column `SICK` — it prefixes any matching time-record row, and emits a projected-hours row (scheduled `paid_hours`) for sick days that have **no** time record so the hours still surface.

## Draft / Publish (per work-week visibility)
- Default = **draft** (hidden from guards). A week becomes visible only when published.
- Guard/public reads filter to published weeks: `getBoardData`, `clientGetShifts` (tokenless), `getShiftsForGuard` (via `filterPublishedShifts_`). Admin reads do **not** filter — admin Full Schedule loads via `clientGetShiftsAdmin` (token, requireAdmin) and shows drafts.
- `filterPublishedShifts_` hides only **assigned** shifts in unpublished weeks (the in-progress roster). **Open/unassigned shells stay visible** so guards can still mark availability for event slots (e.g. Mega/Tournament) in weeks that aren't published yet — `availOptionsFor` (index.html) derives event options from scheduled shifts.
- Server helpers (`Code.gs`): `weekStartFor_` (Saturday key, mirrors `weekKey` in index.html), `publishedWeeks_`, `isWeekPublished_`, `setWeekPublished_`, `filterPublishedShifts_`, `getWeekStatesForPeriod_`.
- Actions: `clientPublishWeek` (existing 📢 flow — now also marks the week published, then emails), `clientSetWeekPublished(ws, bool)` (Unpublish), `clientClearWeek` (Discard draft — **refused** if the week is published; unpublish first).
- Full Schedule per-week bar shows a DRAFT/PUBLISHED chip + ✅ Publish Week / 🗑 Discard (draft) or ↩ Unpublish (published). Frontend caches state in `D.weekStates` via `clientGetWeekStates` (`loadWeekStates(pid)`).
- Edits to an already-published week go **live immediately** (no re-publish needed); use ↩ Unpublish to pull a week back to draft for heavy edits.
- ⚠️ **One-time after first deploy**: run `migratePublishExistingWeeks()` once from the editor — it marks every week that already has assigned shifts as published, so the current live schedule isn't suddenly hidden. (New/empty weeks stay draft.) Idempotent.
