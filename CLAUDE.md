# 7 Presidents STS — Lifeguard Scheduling System

## Architecture
- Google Apps Script web app (Code.gs + index.html)
- Google Sheets backend
- PWA wrapper hosted on GitHub Pages (baconlt.github.io/7p-sts)

## Apps Script (apps-script/)
- Code.gs: all server-side logic, auth, CRUD
- index.html: entire frontend SPA (HTML + CSS + JS in one file)
- Deploy: paste into Apps Script editor, Deploy → Manage Deployments → edit existing
- URL: [your apps script URL]

## PWA Wrapper (pwa/)
- Thin iframe wrapper that enables PWA install + session persistence
- Deploy: push to GitHub repo baconlt/7p-sts, GitHub Pages serves it
- Bump CACHE version in sw.js on each deploy

## Auth
- Password-only (no Google OAuth)
- Sessions stored in Sessions sheet (90-day expiry)
- PWA wrapper stores token in localStorage, passes to iframe via postMessage

## Key conventions
- Admin functions use arun() which prepends auth token
- Guard functions use runG() which prepends auth token
- All dates stored as YYYY-MM-DD strings
- Work week is Saturday–Friday
- Pay periods follow county schedule

## Sheets
Guards, Posts, ShiftTemplates, Shifts, Availability, PayPeriods, 
Config, Notifications, TimeRecords, ShiftStats, Sessions
