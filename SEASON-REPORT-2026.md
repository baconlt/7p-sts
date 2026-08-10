# Seven Presidents STS — The 2026 Season in Data

**Seven Presidents Oceanfront Park · Long Branch, NJ**
Everything logged between **April 13 and August 6, 2026**.

Source: the live `Seven Presidents STS Data` workbook, pulled Aug 6, 2026.
1,342 time records · 1,046 stat cards · 4,376 scheduled shifts · 2,781 availability entries · 598 badge awards.

---

## The headline numbers

| | |
|---|---|
| Guard-hours on the sand | **10,257** |
| Preventive actions | **17,082** |
| Water rescues | **189** |
| Bather assists | **257** |
| First aid responses | **120** |
| Days with a guard on duty | **94** |
| Guards who worked | **41** |
| Stat cards filed | **1,046** |

Rescue breakdown: 170 rescue can · 8 paddleboard · 7 SeaBob · 4 can-and-line.

Two ratios worth sitting with:

- **90 preventive actions for every 1 rescue.** The overwhelming majority of this job is the whistle, the wave-off, and the walk down the beach. The rescue is the exception the preventives are buying.
- **1.84 rescues per 100 guard-hours** — one rescue for roughly every 54 hours a guard spent on a stand.

---

## 1. July was not a month. It was the season.

The workload isn't spread across the summer. It's stacked into July — and the concentration is far more extreme in the incident data than in the payroll.

| Month | Hours | Preventives | Rescues | First aid | Cards |
|---|---:|---:|---:|---:|---:|
| April | 23 | 0 | 0 | 0 | 0 |
| May | 881 | 73 | 0 | 0 | 63 |
| June | 2,694 | 2,643 | 41 | 20 | 271 |
| **July** | **5,706** | **13,159** | **144** | **93** | **627** |
| Aug 1–6 | 953 | 1,207 | 4 | 7 | 85 |

- July took **55.6% of every hour** worked this season…
- …but **77% of the preventive actions** and **76% of the rescues**.

A July hour on the stand carried **2.4× the preventive workload of a June hour** (2.31 preventives/hour vs 0.98). Staffing scaled up; demand scaled up faster.

---

## 2. The big finding: sunshine predicts rescues, surf doesn't

This is the most counterintuitive thing in the data, and it isn't close.

**Surf condition barely moves the rescue rate at all:**

| Surf | Cards | Rescues/card | Total rescues |
|---|---:|---:|---:|
| Rough | 91 | 0.198 | 18 |
| Calm | 675 | 0.181 | 122 |
| Moderate | 203 | 0.177 | 36 |

From flat to rough, the rescue rate moves **less than 12%**. A calm sea is very nearly as dangerous as a rough one.

**Weather moves it by a factor of ten:**

| Weather | Cards | Rescues/card | Preventives/card | Total rescues |
|---|---:|---:|---:|---:|
| **Sunny** | 619 | **0.225** | 21.5 | **139** |
| Foggy | 13 | 0.077 | 31.2 | 1 |
| Thunder/Lightning | 79 | 0.038 | 6.6 | 3 |
| Cloudy | 137 | 0.022 | 4.8 | 3 |
| Showers | 56 | 0.018 | 3.8 | 1 |

**139 of the season's 189 rescues happened on sunny cards.** A sunny shift is roughly **10× more likely** to produce a rescue than a cloudy one.

> The ocean isn't what makes the day dangerous — the **crowd** is. Sun fills the water, and a full water column is what generates rescues. Surf is a hazard multiplier, not the trigger.

**Operational reading:** staff to the forecast's sun, not the surf report. A flat, blazing Saturday is a heavier day than a head-high, overcast one.

---

## 3. The 85° rule

Air temperature is the cleanest single predictor in the dataset, and it has a visible **threshold at 85°F** rather than a smooth slope.

| Air temp | Cards | Preventives/card | Rescues/card |
|---|---:|---:|---:|
| 90°+ | 127 | 32.6 | **0.386** |
| 85–89° | 96 | 16.9 | **0.312** |
| 80–84° | 377 | 14.8 | 0.127 |
| 75–79° | 186 | 17.6 | 0.161 |
| 70–74° | 83 | 8.4 | 0.036 |
| under 70° | 27 | 1.3 | 0.000 |

Under 85°, the crew averaged **0.12 rescues and 14.2 preventives** per card (673 cards). At 85° and above: **0.35 rescues and 25.9 preventives** (223 cards) — **2.9× the rescue rate and 1.8× the preventive load.**

*(Note: guards report temperature in round numbers — 80° appears on 275 cards, 85° on 66 — so band boundaries matter. These bands are lower-inclusive: "85–89" means 85 ≤ temp < 90.)*

**It's a medical threshold too.** All four heat-emergency incident cards this season occurred at **85°F or higher** — 85°, 89°, 95°, 102°. Three of the four were on July 4 alone.

Hottest card of the season: **105°F on July 4**. Season mean: 80.8°F.

If there's one number to hang a staffing trigger on, it's 85.

---

## 4. Friday is the sleeper

Everyone braces for Saturday and Sunday. The hardest day per guard on the beach was **Friday**.

**How to read "rescues per shift":** it's total rescues divided by the number of stat cards (guard-shifts) filed that weekday. **0.442 means a Friday shift produced a rescue roughly every other time out** — one rescue for every 2.3 guard-shifts worked. The "1 in N" column says the same thing the intuitive way.

| Day | Shifts | Rescues | Rescues/shift | ≈ 1 rescue every | Preventives/shift |
|---|---:|---:|---:|---:|---:|
| **Friday** | 138 | 61 | **0.442** | **2.3 shifts** | 23.0 |
| Sunday | 214 | 49 | 0.229 | 4.4 shifts | 19.6 |
| Saturday | 214 | 43 | 0.201 | 5.0 shifts | 21.6 |
| Thursday | 113 | 15 | 0.133 | 7.5 shifts | 11.8 |
| Monday | 124 | 12 | 0.097 | 10.3 shifts | 9.6 |
| Tuesday | 123 | 7 | 0.057 | 17.6 shifts | 10.8 |
| Wednesday | 120 | 2 | 0.017 | 60 shifts | 10.4 |

Friday's rescue rate is **double Saturday's** and **26× Wednesday's** — and Friday produced **61 rescues, more than Saturday and Sunday's 43 and 49**, despite far fewer shifts worked.

Two forces are probably stacking: the weekend crowd arrives while the weekday staffing pattern is still in place.

**And this cuts against how the crew asks for time off:**

| Day-off requests | Thu | Tue | Wed | Mon | Fri | Sat | Sun |
|---|---:|---:|---:|---:|---:|---:|---:|
| Count | 190 | 184 | 184 | 176 | **174** | 99 | 89 |

The crew self-polices the weekend — Saturday and Sunday draw roughly half the off-requests of a weekday. But **Friday is treated as a weekday**, drawing 174 requests, while performing like the busiest day of the week.

**Actionable:** Friday is the one day where the crew's time-off instinct and the incident data point in opposite directions. Worth protecting Friday staffing the way Saturday already is.

---

## 5. Nine posts, nine different jobs

| Post | Cards | Preventives | Per card | Rescues | Bather assists | First aid |
|---|---:|---:|---:|---:|---:|---:|
| Kiernan Central | 242 | 6,042 | 25.0 | 28 | 70 | **44** |
| Joline | 212 | 4,696 | 22.2 | **79** | 81 | 20 |
| Atlantic | 207 | 1,519 | 7.3 | 19 | 29 | 16 |
| Patrol & Response | 151 | 760 | 5.0 | 14 | 8 | 13 |
| Kiernan North | 118 | 1,766 | 15.0 | 14 | 28 | 14 |
| Late Shift | 45 | 298 | 6.6 | 9 | 4 | 7 |
| Sea View | 26 | 914 | **35.2** | 21 | 24 | 4 |
| Kiernan South | 24 | 766 | 31.9 | 2 | 12 | 1 |
| Avenel | 21 | 321 | 15.3 | 3 | 1 | 1 |

- **Joline is the rescue beach** — 42% of the season's rescues from 20% of the cards.
- **Kiernan Central is the workhorse** — most cards, most preventives, and 44 first-aid responses (37% of the season's total).
- **Sea View and Kiernan South** are low-volume, high-intensity: few shifts, but the highest preventive rates on the park.
- Atlantic and Patrol & Response sit at the other end (7.3 and 5.0 per card). Those are structurally different assignments, not less diligent guards.

---

## 6. The four days that defined the summer

| | |
|---|---|
| **51 rescues** — July 3 | More than a quarter of the entire season's rescues in one day |
| **1,951 preventives** — July 26 | The heaviest preventive day on record, 11% of the season |
| **286 guard-hours** — July 4 | 29 guards, 105°F, 21 first-aid responses |
| **33 guards on duty** — July 24 | Largest single-day roster of the season (282.8 hours) |

### July 3 — 51 saves in one afternoon

Nine guards filed rescues. Surf was logged **Calm** and the flag was **Green** on nearly every card — which is exactly the point of finding #2. It was 95°F.

| Guard | Post | Rescues | Preventives | First aid |
|---|---|---:|---:|---:|
| John Donohoe | Joline | **15** | 250 | 4 |
| RJ Williams | Joline | 10 | 105 | 0 |
| Leeza Bernhaut | Atlantic | 7 | 10 | 1 |
| Michael Ford | Late Shift | 6 | 0 | 0 |
| Devyn Ford | Joline | 4 | 80 | 1 |
| Savannah Ginda | Sea View | 4 | 50 | 0 |
| Sarkis Marrin | Patrol & Response | 3 | 1 | 0 |
| Stephen (Craig) Mcgrouther | Late Shift | 1 | 0 | 1 |
| Isabella Montanari | Atlantic | 1 | 30 | 0 |

**29 of the 51 rescues came off Joline** — one post, three guards, one afternoon. Donohoe's 15-rescue card is the largest single card in the dataset by half again.

### July 4 — the longest day

105°F. Twenty-nine guards on. Seven of them logged shifts over **13 hours**, three over 13½ — clocking in at 8:45 AM and out between 10:40 and 10:50 PM. Twenty-one first-aid responses (the season's highest) and three of the season's four heat emergencies. 286.2 guard-hours, the biggest single day of the year.

---

## 7. Who carried it

### Hours worked — top 18

| Guard | Rank | Hours | Days |
|---|---|---:|---:|
| Stephen (Craig) Mcgrouther | Beach Captain | **604.2** | 64 |
| Bryan Mejia | Lieutenant | 552.0 | 60 |
| Sarkis Marrin | Vet | 511.9 | 58 |
| Devyn Ford | Lieutenant | 415.8 | 52 |
| Abigail Malek | Vet | 409.1 | 45 |
| Mike Tomaino | Lifeguard Supervisor | 408.5 | 73 |
| Michael Ford | Vet | 407.0 | 47 |
| Liam Pollock | Guard | 383.9 | 49 |
| Savannah Ginda | Lieutenant | 368.4 | 41 |
| Grace Montanari | Vet | 367.6 | 45 |
| Tyler Terhune | Lieutenant | 360.1 | 46 |
| Jordan Hom | Guard | 357.5 | 44 |
| Michelle Tomaino | Captain | 353.2 | 47 |
| Tom Wicklund | Crew Captain | 347.7 | 49 |
| Lawrence (Finn) Carton V | Vet | 337.6 | 42 |
| Brian Rooney | Vet | 315.7 | 41 |
| Mia Ryan | Beach Captain | 315.4 | 35 |
| Shane Toohey | Vet | 265.4 | 35 |

Craig Mcgrouther worked **604 hours over 64 days** — 5.9% of every hour the crew logged, and 52 hours more than the next guard.

### Overtime

**685.6 OT hours** across **103 guard-weeks** that broke 40 (weeks run Saturday–Friday). Twenty-four guards logged at least one OT week.

| Guard | OT weeks | OT hours | Biggest week |
|---|---:|---:|---:|
| Stephen (Craig) Mcgrouther | 9 | 124.3 | 61.0 |
| Bryan Mejia | 9 | 82.0 | 51.0 |
| Sarkis Marrin | 7 | 64.3 | 54.3 |
| Abigail Malek | 8 | 50.8 | 50.1 |
| Michael Ford | 6 | 49.4 | 56.3 |
| Mia Ryan | 5 | 40.1 | 51.3 |
| Savannah Ginda | 6 | 34.7 | 48.0 |
| Jordan Hom | 5 | 31.4 | 48.6 |
| Grace Montanari | 4 | 29.6 | 50.2 |
| Michelle Tomaino | 4 | 26.3 | 51.0 |

Mcgrouther ran **three consecutive 60-hour weeks** in mid-June (61.0, 60.0, 59.98) — the heaviest sustained stretch anyone pulled.

### Longest consecutive-day streaks

| Guard | Days straight | Run |
|---|---:|---|
| Bryan Mejia | **11** | May 18 – May 28 |
| Michael Butler | 10 | Jul 6 – Jul 15 |
| Mike Tomaino | 10 | Jul 10 – Jul 19 |
| Grace Montanari | 9 | Jul 20 – Jul 28 |
| Abigail Malek | 9 | Jun 30 – Jul 8 |
| Stephen (Craig) Mcgrouther | 9 | May 18 – May 26 |
| Michael Ford | 9 | Jul 11 – Jul 19 |
| Armando Rodriguez | 9 | Jul 6 – Jul 14 |

### The rescue board

Rescues cluster hard — they're a function of *which post on which day* far more than who's sitting on it.

| Guard | Rescues | Cards with a rescue |
|---|---:|---:|
| Leeza Bernhaut | 23 | 7 |
| John Donohoe | 20 | 5 |
| Grace Montanari | 16 | 6 |
| RJ Williams | 15 | 5 |
| Michael Ford | 14 | 6 |
| Andrei Vergara | 13 | 5 |
| Stephen (Craig) Mcgrouther | 12 | 9 |
| Savannah Ginda | 12 | 5 |
| Liam Pollock | 11 | 9 |
| Michelle Tomaino | 8 | 6 |

**Leeza Bernhaut logged 23 rescues from only 7 cards** — the highest per-card rate on the crew. **24 of the 41 guards** logged at least one rescue this season.

### Preventive-action leaders

| Guard | Preventives | Cards |
|---|---:|---:|
| Stephen (Craig) Mcgrouther | 2,305 | 71 |
| Michelle Tomaino | 1,484 | 46 |
| Jordan Hom | 957 | 38 |
| RJ Williams | 927 | 25 |
| John Donohoe | 905 | 27 |
| Devyn Ford | 825 | 43 |
| Savannah Ginda | 806 | 47 |
| Abigail Malek | 744 | 34 |

---

## 8. The crew is almost never late

Across **1,219 day-shift punches** measured against the 8:45 AM start:

| | |
|---|---|
| Clocked in at or before 8:45 | **92.6%** (1,129 of 1,219) |
| More than 5 minutes late | **18** punches (1.5%) |
| More than 15 minutes late | **10** punches, all season |
| Median clock-in | **8:45** — exactly on the bell |

The distribution is remarkable on its own: the modal clock-in is **8:45:00 sharp (817 punches)**, with a tight shoulder of early arrivals at 8:41–8:44 and only a thin late tail. Whatever the stand culture is, it's holding.

### The one soft spot: clocking out

The system auto-closed **118 records (8.8%)** because a guard never clocked out — roughly one shift in eleven ending with a forgotten punch and a correction later. That's the single largest source of payroll edits: **1,057 of 1,342 records were edited** after the fact.

Auto clock-outs generated 128 notification emails this season — the third-most common message the system sends, behind schedule publishes (284) and payroll reminders (187).

**Suggestion:** a clock-out nudge at 5:00 PM would likely eliminate most of them.

---

## 9. Late shifts — 342 of them, in a dozen people

The 5:15–7:15 PM late shift ran **342 times** through Aug 6. Twenty-five guards worked at least one; the top five absorbed nearly half.

| Guard | Late shifts |
|---|---:|
| Bryan Mejia | 40 |
| Stephen (Craig) Mcgrouther | 39 |
| Sarkis Marrin | 30 |
| Abigail Malek | 26 |
| Thomas Lamonia | 22 |
| Leeza Bernhaut | 22 |
| Mia Ryan | 20 |
| Michael Ford | 20 |
| Savannah Ginda | 17 |
| Jordan Hom | 16 |
| Armando Rodriguez | 15 |
| Isaac Marcano | 11 |

Mejia and Mcgrouther together covered **79 late shifts — 23% of every late shift run this season**. Both also top the overtime table, which is not a coincidence: the late shift is where the OT hours come from.

Why the pool stays small: guards volunteered for a late shift **349 times out of 2,781 availability entries (12.5%)**. Demand for the slot outruns supply, and the schedule leans on the same names.

---

## 10. The radio log — what else happened out there

Beyond water rescues: **20 incidents**, **89 ten-codes**, 18 patron-harassment reports, 9 acts of vandalism.

| Code | Count |
|---|---:|
| 10-24 — alcohol / substance | **49** |
| 10-55 — intoxicated patron | 32 |
| 10-34 — disturbance / fight | 8 |

The narrative fields are where the season actually lives. Verbatim from the cards:

> **10-24** — "Man with bottle of beer. Rangers sent EVERYONE including an apache helicopter"

> **Incident** — "Heat exhaustion elevated BP. Medstar transported patient to MMC"

> **Incident** — "Victim was experiencing what she called a panic attack… We administered oxygen. She gave herself an inhaler. We cooled the victim down. Victim called her significant other and denied care. He drove her home."

> **10-34** — "Young children hit a ball at a woman — lifeguards went to woman and kids and had an apology session."

> **10-34** — "Lady upset youths were playing music. Rangers contacted. Rangers arrived. Woman still upset about the Rangers not doing enough. Lifeguards instructed to focus on water."

> **Harassment** — "Fisherman paranoid about illegal fish was harassing [guards]. Rangers were called and situation was deescalated"

> **Vandalism** — "The buoy was cut off of the lifeguard rope behind the shed"

> **Incident** — "10-99 heat stroke"

### What the flags said

| Flag | Cards logging it | Share |
|---|---:|---:|
| Green | 711 | 68.0% |
| Yellow | 201 | 19.2% |
| Red | 114 | 10.9% |

*(A card may list more than one flag.)* Two of every three shifts flew green. Red went up on about one shift in nine — and **60 cards flew the yellow-over-red combination**, meaning conditions turned mid-shift.

---

## 11. Cadets — the next crew, in reps

Nine cadets worked **96 shifts and 483 hours**, logging physical training rep by rep on their own version of the stat card. (532 cadet slots were scheduled across the season; 101 were assigned.)

| Exercise | Total reps |
|---|---:|
| Push-ups | 1,445 |
| Squats | 1,315 |
| Sit-ups | 850 |
| Burpees | 152 |
| Out-and-ins | 74 |
| Pull-ups | 50 |
| Hill sprints | 20 |
| Box swims | 15 |
| Dune laps | 9 |

| Cadet | Total logged reps |
|---|---:|
| Marco Burgos | **1,354** |
| Greyson Delatush | 944 |
| Benjamin Jenkins | 478 |
| Dominika Tomaino | 447 |
| Oliver Glenn | 387 |
| Madeline Shorts | 332 |

Marco Burgos logged more reps than the next two cadets combined — including all 50 of the crew's recorded pull-ups.

### Formal training

**12 trainings** between July 17 and August 3, filling 139 attendee-slots:

Vehicle Safety Stand-down · Professionalism and Hierarchy · Passive victim rescues · Free Torp Rescue Skills · Bullying, Hazing & provocateurs · Rescues and Water Entry Intelligence · Tournament Lane Judging · 808 SeaBob operation and rescue · Code-X · 808 SeaBob · Heavy surf rescues · Splicing line

> "I learned the Seven Presidents, the lifeguard signals, some of the 10 codes, and how far fishermen should be from the bathing area."
> — Cadet stat card, July 7

---

## 12. Engagement with the system itself

The app isn't just recording the season — the crew is living in it. **1,786 login sessions**, 2,781 availability entries, 646 notifications sent.

| | |
|---|---|
| Badges earned | **598** across 52 types |
| Stat-card filing rate | **77%** (988 of 1,283 eligible work-days) |
| Ballots cast | **123** across 4 award elections |
| Poll responses | 27 (crew photo date) |

### Award voting turnout

All four secret-ballot rounds ran to completion with **30–32 voters** each — turnout the final rounds actually *improved* on (32 for Lifeguard of the Year, 31 for Rookie of the Year, vs 30 in each primary).

*Results are deliberately omitted from this document — ballots are anonymous by design and the tallies are held for the end-of-season banquet.*

### Most-earned badges

| Badge | Holders |
|---|---:|
| Stats Streak 5 | 34 |
| Storm Shift | 34 |
| Radio Cert | 33 |
| 100 Hours | 31 |
| Nomad | 31 |
| Team Player | 31 |
| Stats Streak 10 | 29 |
| Holiday Hero | 29 |
| Sure Shot | 28 |
| First Save | 25 |
| Sharpshooter | 25 |
| Late Shift 10hr | 24 |

**Badge board:** Craig Mcgrouther 30 · Michelle Tomaino 28 · Michael Ford 28 · Abigail Malek 22 · Grace Montanari 21.

The rarest awards went to single holders — including the 🗺️ **Treasure Hunter**, claimed by the one guard who found the physical cache.

### The Cache Hunt

Eight guards took a distance reading. Mia Ryan got closest without finding it — **451 feet**. Craig Mcgrouther (530 ft), Mike Tomaino (534 ft) and Tom Wicklund (583 ft) were all within striking distance.

John Donohoe took his reading from **207,927 feet away** — about 39 miles — which is its own kind of achievement.

---

## Method, and what to distrust

Everything here is a straight count off the live workbook. No estimates, no modeling.

**Caveats worth knowing:**

1. **Nine duplicate stat cards were found and removed** (same guard, same date, same beach, same numbers, submitted twice). They inflated the raw preventive total by 522 and the rescue total by 3. Every figure in this document is deduplicated: **1,046 cards / 17,082 preventives / 189 rescues**. Raw sheet totals are 1,056 / 17,604 / 192.

2. **The season isn't over.** August runs through the 6th only, so every monthly comparison involving August is partial.

3. **Stat cards are self-reported** and the filing rate is 77%, so roughly a quarter of worked days have no card. Preventive-action counts in particular are shift-end estimates by guards — treat the *ratios* as sound and the *absolute totals* as a floor.

4. **Rescue rates are given per card, not per day,** because staffing changed dramatically across the season. Comparing raw daily totals between June and July would mostly measure headcount.

5. **Punctuality covers day-shift punches only** (clock-ins between 7:00 and 10:00 AM), measured against the 8:45 template start. Late shifts and patrol assignments are excluded.

6. **Named lateness is not reported here.** The aggregate punctuality picture is included; the per-guard conduct log (`GuardNotes`) is admin-only and stays that way.

7. **A handful of early records carry malformed `auto_clocked_out` values** (timestamps in a boolean column, 7 rows in April). They're excluded from the auto-clock-out count, which uses strict `True` matches only.

---

*Generated from the Seven Presidents STS Data workbook · data through August 6, 2026.*
