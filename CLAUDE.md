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
- `schedule` — guard×date grid with availability (renderSchedGrid)
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
Config, Notifications, TimeRecords, ShiftStats, Sessions, Locations
