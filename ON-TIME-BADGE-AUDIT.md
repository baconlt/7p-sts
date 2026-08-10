# On-Time Badge Audit — Sharpshooter 🎯 / Sure Shot 🏹

**Date:** 2026-08-07 · **Trigger:** Craig McGrouther reported clocking in/out on time repeatedly without ever receiving either badge.
**Data:** all 1,385 completed TimeRecords, 2026-04-13 → 2026-08-07, 42 guards.

---

## Verdict: Craig is right

He has **19 genuine, never-overridden clock-in punches**, and **17 of those 19 landed exactly on 8:45 AM.**
He should hold **Sure Shot 🏹 Gold (17)**. He held nothing.

| Date | Genuine 8:45 punch-in |
|---|---|
| 05-20, 05-26, 06-18, 06-25, 06-28, 06-29, 06-30 | ✓ (7) |
| 07-03, 07-05, 07-06, 07-07, 07-11, 07-12, 07-18, 07-21, 07-26 | ✓ (9) |
| 08-01 | ✓ (1) |

He does **not** qualify for Sharpshooter, and that part is correct rather than a bug: Sharpshooter needs *both*
punches real and exact, and Craig hand-enters his clock-out every single day (0 of 67 records have an untouched
clock-out). Typical pattern on a late shift — he confirms clock-out at ~7:00–7:12 PM and types `19:15`.
That's a timesheet-accuracy question worth a separate conversation; it is not what broke the badges.

---

## Two bugs, both real

### Bug 1 — every record from April through June was flagged `edited`

The clock-out modal always sent `clock_out_override` (it pins the displayed minute so GPS lag can't drift the
punch). The server treated *any* supplied clock-out as a hand edit, so `edited='true'` was written even when the
guard changed nothing.

| Month | Records | `edited=FALSE` |
|---|---|---|
| 2026-04 | 11 | **0** |
| 2026-05 | 131 | **0** |
| 2026-06 | 344 | **0** |
| 2026-07 | 739 | 247 |
| 2026-08 | 160 | 51 |

Since the badges skip any `edited` record, **no punch in the first ~3 months of the season could ever count.**
The `clock_out_unedited` flag fixed this on **2026-07-01** — the first `edited=FALSE` record in the season.
486 pre-July records remain mis-flagged.

### Bug 2 — one hand-typed punch disqualified the other (still live until this change)

`computeBadgeMetrics_` gated both punches on the *record-level* `edited` flag:

```js
const isRealPunch = r =>
  String(r.auto_clocked_out) !== 'true' && String(r.edited) !== 'true';
```

`edited` goes true when **either** time is touched. So a guard who nailed 8:45:03 in the morning and then
corrected a forgotten clock-out at night lost credit for the morning punch too. This is exactly Craig's case,
and it kept hurting him after the July fix — every one of his 17 exact clock-ins was wiped out by his
hand-typed clock-out on the same row.

---

## The fix

**Provenance is now judged per punch, not per record** — `punchProvenance_(r) → { inReal, outReal }` in `Code.gs`.

*Going forward*, two new TimeRecords columns record the answer explicitly (`ensureTimeRecordColumns_`, added on
demand — no migration):

| Column | `'true'` means |
|---|---|
| `clock_in_edited` | the clock-in was typed/overridden, not punched |
| `clock_out_edited` | the clock-out was typed/overridden or auto-clocked-out |

`clockOut()` already computed `clockInChanged` and `clockOutChanged` separately — it just collapsed them into one
flag. Now both are written. Every other write path sets them too: `clockIn()` (`in=false`), auto-clock-out
(`out=true`), admin manual entry (both `true`), and the two edit paths via
`onTimePunchProvenanceAfterEdit_()`, which only demotes the punch that actually moved.

*For the 1,385 legacy rows*, provenance is inferred from evidence already on the row — deterministic, not guessed:

- **clock-in is real ⟺ `clock_in === created_at` to the millisecond.** `clockIn()` stamps both from the same
  `now`; every override path rewrites `clock_in` and leaves `created_at` alone. Equality proves the punch is untouched.
- **clock-out is real ⟺ not auto-clocked-out, and either never flagged `edited`, or `edited_at − clock_out` ∈ [0, 150s].**
  That window is the pre-July modal: an untouched punch lands a few seconds before the confirm timestamp;
  a typed time lands minutes away, or in the future (negative) — both rejected.

The 150s window is not arbitrary. The gap distribution over the 850 clock-out-path edits is cleanly bimodal:

| Gap (`edited_at − clock_out`) | Records |
|---|---|
| 0–59s | 154 |
| 60–119s | 26 |
| 120–149s | 8 |
| **150–299s** | 17 |
| 300–899s | 39 |
| ≥900s | 114 |
| negative (future-dated) | 492 |

`edited` itself is unchanged — it still means "some time on this row was changed," which is what the payroll
CSV and the pay-period change report need.

> The old note on `resetOnTimeBadges()` said these badges could never be granted retroactively. That was true
> only while provenance came from the record-level flag. It isn't anymore.

---

## Roster-wide impact

**35 of 42 guards were under-credited. 0 guards lose anything** — verified at both record and guard level; the new
rule is a strict superset of the old, and badges are sticky regardless. **30 guards change tier.**

Season totals: qualifying Sharpshooter days 128 (was ~103), Sure Shot days **272 (was ~78)**.
Guards holding Sharpshooter 26, Sure Shot **39**.

| Guard | 🎯 was → now | 🏹 was → now |
|---|---|---|
| **Stephen (Craig) Mcgrouther** | 0 — → 0 — | **0 — → 17 Gold** |
| Mia Ryan | 0 — → 1 Bronze | 5 Silver → 15 Gold |
| Lawrence (Finn) Carton V | 0 — → 0 — | 1 Bronze → 14 Gold |
| Abigail Malek | 6 Silver → 6 Silver | 7 Silver → 14 Gold |
| Leeza Bernhaut | 2 Bronze → 2 Bronze | 3 Bronze → 14 Gold |
| Grace Montanari | 11 Gold → 11 Gold | 0 — → 13 Gold |
| Michael Black | 2 Bronze → 2 Bronze | 1 Bronze → 12 Gold |
| Tyler Terhune | 10 Gold → 10 Gold | 6 Silver → 12 Gold |
| Michael Ford | 4 Bronze → **5 Silver** | 2 Bronze → 11 Gold |
| Michelle Tomaino | 9 Silver → 9 Silver | 5 Silver → 11 Gold |
| Slater Richardson | 5 Silver → 5 Silver | 2 Bronze → 9 Silver |
| Isaac Marcano | 4 Bronze → 4 Bronze | 2 Bronze → 9 Silver |
| Armando Rodriguez | 12 Gold → 13 Gold | 1 Bronze → 8 Silver |
| Sarkis Marrin | 3 Bronze → 3 Bronze | 0 — → 8 Silver |
| Tom Wicklund | 10 Gold → 10 Gold | 3 Bronze → 8 Silver |
| Devyn Ford | 6 Silver → 6 Silver | 3 Bronze → 8 Silver |
| Savannah Ginda | 0 — → 0 — | 1 Bronze → 7 Silver |
| RJ Williams | 3 Bronze → **4 Bronze** | 0 — → 6 Silver |
| Oliver Glenn | 5 Silver → 5 Silver | 1 Bronze → 6 Silver |
| Michael West | 0 — → 0 — | 6 Silver → 6 Silver |
| Liam Pollock | 0 — → **1 Bronze** | 0 — → 5 Silver |
| Michael Butler | 4 Bronze → 4 Bronze | 1 Bronze → 5 Silver |
| Dominika Tomaino | 4 Bronze → 4 Bronze | 2 Bronze → 5 Silver |
| Jordan Hom | 2 Bronze → 2 Bronze | 3 Bronze → 5 Silver |
| Marco Burgos | 0 — → 0 — | 4 Bronze → 5 Silver |
| John Donohoe | 10 Gold → **11 Gold** | 2 Bronze → 3 Bronze |
| Micaela Gomez | 1 Bronze → 1 Bronze | 0 — → 4 Bronze |
| Mike Tomaino | 0 — → 0 — | 1 Bronze → 4 Bronze |
| Greyson Delatush | 0 — → 0 — | 1 Bronze → 4 Bronze |
| Thomas Lamonia | 0 — → 0 — | 4 Bronze → 4 Bronze |
| Madeline Shorts | 4 Bronze → 4 Bronze | 0 — → 0 — |
| Bryan Mejia | 0 — → 0 — | 0 — → 3 Bronze |
| Edward Olsen | 0 — → 0 — | 0 — → 3 Bronze |
| Benjamin Jenkins | 3 Bronze → 3 Bronze | 0 — → 3 Bronze |
| Shane Toohey | 0 — → 0 — | 2 Bronze → 3 Bronze |
| Brian Rooney | 0 — → 0 — | 2 Bronze → 3 Bronze |
| Andrei Vergara | 1 Bronze → 1 Bronze | 0 — → 2 Bronze |
| Elizabeth Keeshen | 1 Bronze → 1 Bronze | 0 — → 1 Bronze |
| Isabella Montanari | 0 — → 0 — | 1 Bronze → 1 Bronze |
| Max Destefano | 0 — → 0 — | 1 Bronze → 1 Bronze |

---

## To apply

1. Upload the updated `Code.gs`, save, and redeploy.
2. In the Apps Script editor, run **`auditOnTimeBadges()`** once → View → Logs.
   Idempotent, never downgrades; credits every guard above at their true tier.
3. `resetOnTimeBadges()` is *not* needed — the audit only adds.
