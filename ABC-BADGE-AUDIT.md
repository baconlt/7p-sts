# ABC Badge Audit — ABC: Always Be Closing 🌇

**Date:** 2026-08-09 · **Trigger:** Abby Malek reported closing the beach well past 10 times without receiving ABC.
**Data:** full `Seven Presidents STS Data` workbook — 4,398 Shifts, 1,434 TimeRecords, 42 guards, season to date.

---

## Verdict: Abby is right — and the badge had never been awarded to anyone

She has **20 completed records clocking out in the 6:45–8:00 PM closing window**, and the badge counted **0 of them**.

| Date | Clock-out | Linked shift code | Auto? | Counts under fix? |
|---|---|---|---|---|
| 2026-06-07 | 7:15 PM | `8HR` | **auto** | no — auto clock-out |
| 2026-06-13 | 7:15 PM | `8HR` | — | no — no LS assignment |
| **2026-06-14** | 7:14 PM | `8HR` | — | ✅ |
| 2026-06-15 | 7:15 PM | `8HR` | **auto** | no — auto clock-out |
| **2026-06-16** | 7:15 PM | `8HR` | — | ✅ |
| **2026-06-17** | 7:11 PM | `8HR` | — | ✅ |
| **2026-06-30** | 7:15 PM | `8HR` | — | ✅ |
| 2026-07-01 | 7:15 PM | `LS` | **auto** | no — auto clock-out |
| **2026-07-02** | 7:15 PM | `8HR` | — | ✅ |
| **2026-07-05** | 7:15 PM | `8HR` | — | ✅ |
| **2026-07-06** | 7:15 PM | `8HR` | — | ✅ |
| **2026-07-08** | 7:15 PM | `8HR` | — | ✅ |
| 2026-07-11 | 7:15 PM | `8HR` | — | no — no LS assignment |
| **2026-07-17** | 7:17 PM | `8HR` | — | ✅ |
| **2026-07-27** | 7:16 PM | `8HR` | — | ✅ **(10th — badge earned)** |
| **2026-07-28** | 7:02 PM | `8HR` | — | ✅ |
| 2026-07-30 | 7:15 PM | `8HR` | **auto** | no — auto clock-out |
| **2026-08-03** | 7:15 PM | `8HR` | — | ✅ |
| **2026-08-04** | 7:15 PM | `8HR` | — | ✅ |
| 2026-08-05 | 7:15 PM | `8HR` | **auto** | no — auto clock-out |

**19 of her 20 closings are linked to her `8HR` shift.** The one linked to `LS` was auto-clocked-out, so even that
one failed. She independently holds **22 non-cancelled `LS` assignments** on the schedule, which line up with the
punches almost date for date.

She qualifies for **ABC 🌇 as of 2026-07-27** with **13 closes**. She held it not at all.

---

## The bug

`computeBadgeMetrics_` decided which shift a guard was on by joining TimeRecords → Shifts:

```js
const closingShifts = completed.filter(r => {
  if (String(r.auto_clocked_out) === 'true') return false;
  const shift = r.shift_id ? shiftsById[r.shift_id] : null;
  if (!shift || !LATE_CODES.has(shift.template_code)) return false;   // <- here
  const out = clockMinutesLocal_(r.clock_out);
  return out !== null && out >= CLOSE_MIN && out <= CLOSE_MAX;
}).length;
```

A record's `shift_id` is **whichever shift the guard punched in on**. A guard who works the day *and* closes
punches in **once**, at 8:45 AM, against their `8HR` shift, and out at 7:15 PM — one record, one `shift_id`, and
it is the day shift. The `LS` row (5:15–7:15 PM) is never touched by the clock at all. So the condition that was
meant to ask *"were you on the late shift?"* was really asking *"did your morning punch happen to land on the LS
row,"* which nothing in the app arranges.

This is the **same root cause** as the tournament badges (`TOURNAMENT-BADGE-AUDIT.md`), just from the other
direction — there the event punch inherited the day shift's id; here the day punch is the only punch there is.

### Blast radius: the badge was dead on arrival for the whole roster

`always_be_closing` appears in **zero rows of the `Badges` sheet**. Not one guard has ever held it.

| Guard | Real closes | Old count | Badge (old → new) |
|---|---|---|---|
| Bryan Mejia | 34 | 1 | none → 🌇 (2026-06-29) |
| Stephen (Craig) Mcgrouther | 33 | **0** | none → 🌇 (2026-06-25) |
| Mia Ryan | 14 | **0** | none → 🌇 (2026-07-27) |
| Savannah Ginda | 13 | **0** | none → 🌇 (2026-07-29) |
| Abigail Malek | 13 | **0** | none → 🌇 (2026-07-27) |
| Sarkis Marrin | 12 | 2 | none → 🌇 (2026-08-01) |
| Thomas Lamonia | 11 | 2 | none → 🌇 (2026-08-03) |
| Leeza Bernhaut | 9 | 0 | none (1 short) |
| Jordan Hom | 9 | 0 | none (1 short) |
| Michael Ford | 8 | 0 | none |
| *(11 more at 1–6)* | 1–6 | 0–2 | none |

The highest count the old metric ever produced across the entire season was **2** — against a target of 10. The
badge had been unreachable since the day it shipped.

---

## The fix

The **schedule** says which nights the guard was on the late shift; the **clock** says they actually closed.
A date counts only when both agree.

- **`BADGE_SHIFT_CODES_ = ['TOURN', 'LS', 'LS8']`** — the late codes now flow through the same
  `scheduledEventDates_` machinery the tournament fix introduced, so `ctx.eventDates.LS` / `.LS8` carry every
  non-cancelled late assignment (unioned with any record that *is* linked to one, clamped to today and to
  `cutoff`).
- **`closingShifts`** requires a late assignment that date **AND** a completed, non-auto-clocked-out punch
  landing in `CLOSE_MIN`–`CLOSE_MAX` (6:45 PM–8:00 PM). Moved below the `eventDates` block so it can read it.
- Counted by **distinct date** rather than by record — two punches on one evening is one beach close. This is
  what drops Thomas Lamonia from 12 records to 11 closes.
- `computeEarnedDates_` already folds `data.eventDates` into its replay candidates, so the back-dating picks up
  late-shift dates for free.

### Why both sources, not either one

**Schedule alone** would credit a guard who was rostered on the late shift and went home at 6 PM — the assignment
is a plan, not attendance. **Clock alone** cannot tell a beach close from a meet ending: on 2026-07-15 seven
guards clocked out at 7:30 PM off `TOURN` shifts (Michael Black, Tyler Terhune, John Donohoe, Devyn Ford, Jordan
Hom, Andrei Vergara, Craig Mcgrouther) — a tournament, not a close. Requiring the late assignment excludes all
seven correctly. Across the season, 41 in-window punches have no late assignment behind them and are rightly
not counted.

This differs from the tournament fix, which credits on the assignment *alone* — necessarily, because away meets
at Belmar and Manasquan leave no punch at all. Closing the beach always leaves a punch, so here the clock
evidence is available and worth requiring.

### Auto clock-outs still don't count — deliberately

Abby has 4 more nights (06-15, 07-01, 07-30, 08-05) where she was assigned the late shift and the system
auto-clocked her out at 7:15 PM. Those are almost certainly real closes, but an auto clock-out's `clock_out` is a
duration cap rather than a punch, so it is not evidence of anything — the same reasoning the late-shift and
10-hour badges already use, and the flip side of the `sleepyhead` 😴 demerit. She clears 10 without them.

### Guarding against the next one

`BADGE_SHIFT_CODES_` carries a ⚠️ note covering both directions of this failure, the `closingShifts` block
explains why it reads the schedule, and `CLAUDE.md` documents the rule.

---

## Deploying

1. Paste `apps-script/Code.gs` into the Apps Script editor, save, redeploy.
2. Run **`auditClosingBadge()`** once from the editor (Run → View → Logs).

The audit writes the missing `Badges` rows now instead of waiting for each guard to open their own stats page —
that page is the only place a computed badge's row gets written, so without it the leaderboard and seven guards'
grids would disagree for weeks. It is additive and idempotent: it never deletes, never downgrades, and back-dates
`earned_at` to the night each guard's 10th close landed.

**Expected result: 7 badge rows written** — 🌇 ABC to Bryan Mejia, Craig Mcgrouther, Mia Ryan, Savannah Ginda,
Abigail Malek, Sarkis Marrin and Thomas Lamonia.
