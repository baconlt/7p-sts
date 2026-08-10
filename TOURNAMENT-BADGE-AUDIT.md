# Tournament Badge Audit — Craig Hummer 🏆 / Competitor Supreme 🥇

**Date:** 2026-08-07 · **Trigger:** Abby Malek reported working 5 tournament shifts without receiving Competitor Supreme.
**Data:** full `Seven Presidents STS Data` workbook — 4,397 Shifts, 1,385 TimeRecords, 42 guards, season to date.

---

## Verdict: Abby is right — and so is most of the roster

She has **5 non-cancelled `TOURN` assignments**, and the badge counted **0 of them**.

| Date | TOURN shift | Status | Her time record that day |
|---|---|---|---|
| 2026-07-07 | `S-1783260246373-VZLPK` | cancelled | — (doesn't count) |
| **2026-07-09** | `S-1783532537208-01WEP` | filled | none — away meet, no punch |
| **2026-07-13** | `S-1783530893936-KLVZ8` | filled | 5:56–9:01 PM, linked to her **8HR** shift |
| 2026-07-14 | `S-1783530902246-OR0DZ` | cancelled | — (doesn't count) |
| **2026-07-15** | `S-1783885468681-C30TV` | filled ("Belmar Tournament Rescheduled") | none — away meet, no punch |
| **2026-07-20** | `S-1780065248005-7MLQH` | filled | 6:00–8:45 PM, linked to her **8HR** shift |
| **2026-07-22** | `S-1780065255600-O5BN1` | filled | 5:54–8:33 PM, linked to her **8HR** shift |

She qualifies for **Craig Hummer 🏆 as of 2026-07-15** and **Competitor Supreme 🥇 as of 2026-07-22**. She held neither.

---

## The bug

`computeBadgeMetrics_` counted tournaments by joining TimeRecords → Shifts:

```js
const tournCompleted = completed.filter(r => {
  const shift = r.shift_id ? shiftsById[r.shift_id] : null;
  return shift && shift.template_code === 'TOURN';
});
const tournShifts = tournCompleted.length;
```

A record's `shift_id` is **whichever shift the guard punched in on** — and for an evening tournament that is
essentially never the tournament:

1. **Home meets.** A day guard clocks in at 8:45 AM against their `8HR` shift, clocks out at 5:15 PM, then clocks
   back in for the 6:00 PM tournament. The second punch inherits the **same `8HR` `shift_id`** — every one of
   Abby's three tournament punches is linked to her day shift. The `TOURN` row is never touched.
2. **Away meets.** Belmar, Manasquan and the like are off-site and out of geofence, so travelling guards often
   never punch at all — Abby has **no time record whatsoever** on 07-09 or 07-15, the two away dates.

Across the season the join found **17 tournament punches against 64 tournament assignments actually worked** —
it was measuring "did your punch happen to land on the tournament row," which nothing in the app arranges.

### Blast radius: the badges were unreachable for the entire roster

| Guard | Tournaments worked | Old count | Badges (old → new) |
|---|---|---|---|
| Abigail Malek | 5 | **0** | none → 🏆 + 🥇 |
| Grace Montanari | 5 | 1 | none → 🏆 + 🥇 |
| Stephen (Craig) Mcgrouther | 4 | **0** | none → 🏆 |
| Shane Toohey | 4 | **0** | none → 🏆 |
| Devyn Ford | 4 | **0** | none → 🏆 |
| Jordan Hom | 4 | 2 | none → 🏆 |
| John Donohoe | 4 | 2 | none → 🏆 |
| Michael Ford | 4 | 1 | none → 🏆 |
| Mia Ryan | 4 | 1 | none → 🏆 |
| Isabella Montanari | 4 | 1 | none → 🏆 |
| Michael Black | 3 | 1 | none → 🏆 |
| Andrei Vergara | 3 | 1 | none → 🏆 |
| Tyler Terhune | 3 | 2 | none → 🏆 |
| *(10 more at 1–2)* | 1–2 | 0–2 | none |

**Nobody on the roster held `tourn_3` or `tourn_5`.** The highest count the old metric ever produced was 2 —
one short of the first badge — so both awards had been dead since the day they shipped. The same root cause is
already documented on `team_player` 🙌 (`adminGrantBadge_`'s doc comment: "worked a tournament but clocked into a
non-TOURN shift, so Team Player never fired"); it was hand-granted to all 31 attendees on 2026-07-24 instead of
being fixed.

---

## The fix

Event-shift badges now credit off the **schedule**, matching how `bayshore_late` 🛟 already works.

- **`BADGE_SHIFT_CODES_ = ['TOURN']`** — the registry of template codes that badges count by assignment.
- **`scheduledEventDates_(allShifts, guardId)`** → `{ TOURN: [ymd, …] }` for every non-cancelled shift carrying
  the guard's id. Fed into `data.eventDates` by `loadBadgeData_`, and bucketed per guard in
  `computeBadgeLeaderboard_` so the leaderboard and each guard's own grid agree.
- **`computeBadgeMetrics_`** unions those dates with any completed record that *is* linked to a shift of that code
  (keeps RJ Williams' 07-07 punch on a since-cancelled tournament), then clamps to **today** — so a tournament
  booked for next week can't pre-credit a badge — and to `cutoff`, so the historical replay stays honest.
  Exposes `ctx.eventDates`; `tournShifts` and `mcpsTournament2026` read from it.
- **`computeEarnedDates_`** adds event dates to its replay candidates. Without this, a badge won on an unpunched
  away-meet day would be back-dated to some unrelated later shift.

### Why assignment, not attendance

Requiring a same-day time record — the extra condition `bayshore_late` carries — would still have given Abby only
3 of 5, because the away meets have no punch by design. There is no clock evidence to require. A non-cancelled
assignment is the only record the away meets leave, and the admin removes guards who don't travel (7 of Abby's
tournament rows were assigned; 2 were cancelled and correctly don't count).

### Guarding against the next one

`scheduledEventDates_` carries a ⚠️ in its doc comment, and `CLAUDE.md` now states the rule directly: **adding a
"worked N shifts of type X" badge means adding `X` to `BADGE_SHIFT_CODES_` and reading `ctx.eventDates.X` — never
filtering time records by their linked shift's `template_code`.**

---

## Deploying

1. Paste `apps-script/Code.gs` into the Apps Script editor, save, redeploy.
2. Run **`auditEventShiftBadges()`** once from the editor (Run → View → Logs).

The audit writes the missing `Badges` rows now instead of waiting for each guard to open their own stats page —
that page is the only place a computed badge's row gets written, so without it the leaderboard and 13 guards'
grids would disagree for weeks. It is additive and idempotent: it never deletes, never downgrades, and back-dates
`earned_at` to the day each badge actually tipped over.

**Expected result: 15 badge rows written** — 🏆 Craig Hummer to 13 guards, 🥇 Competitor Supreme to Abigail Malek
and Grace Montanari.
