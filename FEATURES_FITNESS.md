# Fitness & Weight Loss Tracker — Etsy feature list

**Product F.** Working title for the listing:

> *Fitness & Weight Loss Tracker Spreadsheet, Workout Planner and Log,
> Weight and Body Measurement Tracker, Habit and Macro Tracker for
> Excel and Google Sheets*

Price band **$12.99 – $19.99**. Audience: people starting or restarting
a routine, lifters who want a log that adds up, walkers and runners,
coaches who hand clients a tracker, anyone leaving a paid app.

The thing that has to be true for it to win at that price: **it is a
complete tracking system that does the arithmetic for you and never
breaks**, not a pretty weight log. Everything below is written as
"what the buyer gets", because that is how it has to read in the
listing.

---

## 1. Workbook structure — 16 tabs

Exactly the 16 you listed, in this order. The numbering is on the tab
names so the file opens in a sensible order and the navigation strip at
the top of every sheet links across.

| # | Tab | What is on it |
|---|-----|---------------|
| 1 | **Start Here** | Welcome, five-step quick start, the colour code (cream = type here, grey = calculated), how to clear the demo data, where the data lives (your own file), FAQ, licence, and the health disclaimer |
| 2 | **Profile & Goals** | Height, age, starting weight, goal weight, target date, units (kg/lb, cm/in), activity level, up to six independent goals (weight, strength, consistency, distance, habit, custom) each with its own target and deadline |
| 3 | **Dashboard** | KPI tiles, goal progress bars, this week versus last week, weight trend sparkline, workout and habit completion rates, streaks, "what to do next" line, and an honest empty state before any data is typed |
| 4 | **Workout Planner** | Weekly plan by day, training split (PPL, upper/lower, full body, custom), planned exercises, sets, target reps, planned minutes, planned versus completed |
| 5 | **Workout Log** | One row per set or per exercise: date, workout, exercise, sets, reps, weight, RPE, rest, duration, notes, done flag — with volume, estimated 1RM and personal-record detection calculated |
| 6 | **Exercise Library** | 150+ exercises shipped: name, muscle group, secondary muscles, equipment, movement pattern, difficulty, unilateral flag, how to do it, two substitutions. Fully editable, and it feeds every dropdown |
| 7 | **Cardio & Activity** | Walking, running, cycling, swimming, rowing, classes: date, type, distance, duration, pace, steps, heart rate, perceived effort, optional calories (user-entered or estimated from MET and body weight) |
| 8 | **Weight Tracker** | Daily or weekly weigh-ins, change since start, change since last, 7-day and 28-day moving average, weekly rate of change, projection to goal at the current rate, milestone markers |
| 9 | **Body Measurements** | Waist, hips, chest, shoulders, arms, forearms, thighs, calves, neck, plus four custom sites; change since start and since last, waist-to-hip and waist-to-height ratio, optional Navy-method body-fat estimate |
| 10 | **Habit Tracker** | Up to twelve habits across a month grid, daily ticks, current and best streak, completion rate by habit and by week, rest days that do not break a streak |
| 11 | **Meals & Macros** | Food log by meal, user-entered calories and macros, a 60-row food reference list that auto-fills a logged item, daily totals against targets, weekly averages, protein per kg of body weight |
| 12 | **Water & Sleep** | Water by glass or by volume against a target, bedtime, wake time, hours slept, sleep quality, mood and energy, weekly summaries |
| 13 | **Records & Milestones** | Automatic personal records per exercise (heaviest, best estimated 1RM, best volume), longest distance, fastest pace, consistency milestones, non-scale victories, reward notes |
| 14 | **Weekly & Monthly Review** | Twelve weekly review blocks and twelve monthly ones: adherence numbers pulled in automatically, then wins, barriers, lessons, and next period's three goals |
| 15 | **Reports & Charts** | Weight and moving average, workout frequency and volume by week, minutes by activity, measurement trends, habit consistency, macro averages, a month-by-month table, print-friendly layout |
| 16 | **Settings & Lists** | Every dropdown list, units and conversion factors, habit names, goal types, activity multipliers, calorie and macro targets, the review period, colour key, and the ID prefixes |

Hidden helper columns do the ranking and streak work so no tab needs an
array formula. No 17th tab: everything else earns its place on one of
these.

---

## 2. The 90 features, grouped as the listing will sell them

### A. Dashboard and automation (10)

1. KPI tiles that fill themselves: current weight, change, workouts this
   week, minutes, habit rate, streak, calories average, goal progress.
2. Automatic weekly summary for the last eight weeks.
3. Automatic monthly summary for twelve months.
4. Starting versus current versus goal weight, with the gap in both
   directions.
5. Goal progress as a percentage and as "x to go", per goal.
6. Workout completion rate: planned against logged.
7. Habit completion rate with current and best streak.
8. Trend sparklines and data bars that redraw as rows are typed.
9. Conditional formatting for milestones, records, missed days and
   targets met.
10. A navigation strip on every tab, plus clear empty states so a blank
    file does not look broken.

### B. Workout planning and logging (12)

11. Weekly planner with a training split and rest days.
12. Daily workout log with sets, reps, weight, RPE, rest and notes.
13. Strength and cardio logged separately but reported together.
14. Dropdowns driven by the exercise library, filtered by muscle group
    and equipment.
15. Volume per set, per exercise, per session and per week.
16. Estimated 1RM (Epley, with the formula shown and editable).
17. Automatic personal-record flag the moment a row beats the old best.
18. Planned versus completed, per exercise and per week.
19. Substitutions and custom exercises, added in one place.
20. Injury and modification notes that travel with the exercise.
21. Session duration, average rest and total work time.
22. Warm-up sets excluded from volume by a single flag.

### C. Weight and body progress (12)

23. Daily, weekly or ad-hoc weigh-ins — no fixed schedule.
24. 7-day and 28-day moving averages so a single bad morning means
    nothing.
25. Change since start, since last entry and over the last 4 weeks.
26. Weekly rate of change in kg or lb and as a percentage of body
    weight.
27. Projection: at this rate, goal reached on this date — stated as an
    estimate, not a promise.
28. Milestone markers every 5 lb / 2 kg, logged with the date.
29. Eleven measurement sites plus four custom ones.
30. Change and percentage change per site, since start and since last.
31. Waist-to-hip and waist-to-height ratios.
32. Optional Navy-method body-fat estimate, clearly labelled an
    estimate.
33. Progress photo log: date, pose, where the file lives, note (no
    photos stored in the file).
34. Non-scale victories: clothes, energy, lifts, sleep, mood.

### D. Habits, wellness and routine (10)

35. Twelve customisable habits, named by the buyer.
36. Month grid with one tick per habit per day.
37. Current streak, best streak and completion rate per habit.
38. Rest days that count as planned, not as failures.
39. Water intake against an editable target.
40. Sleep duration from bedtime and wake time, crossing midnight
    correctly.
41. Sleep quality, mood and energy on a 1–5 scale.
42. Stretching and mobility minutes.
43. Weekly consistency scorecard across every habit.
44. A custom routine checklist for morning or evening blocks.

### E. Meals and nutrition (10)

45. Breakfast, lunch, dinner, snacks and drinks logged by row.
46. Calories, protein, carbohydrate and fat typed by the buyer — no
   hidden database pretending to know what a "medium apple" is.
47. A 60-item food reference list that fills the macros when a logged
    item matches, and is fully editable.
48. Daily totals against daily targets, with the gap shown.
49. Weekly and monthly averages, not just daily numbers.
50. Protein per kilogram (or pound) of body weight.
51. Optional fibre, sugar, sodium and two custom nutrient columns.
52. Editable targets — calories, macros and the macro split — with an
    optional TDEE estimator the buyer can ignore or overwrite.
53. Water pulled alongside nutrition for the day.
54. Incomplete rows are flagged as incomplete rather than silently
    counted as zero.

### F. Reports and review (10)

55. Weekly progress report, printable on one page.
56. Monthly progress report, printable on one page.
57. Workout frequency by week and by month.
58. Exercise performance over time for any exercise picked from a
    dropdown.
59. Weight and measurement trends side by side.
60. Habit consistency comparison across habits.
61. Activity minutes and distance by type.
62. Macro averages against targets.
63. Goal review with the next period's targets.
64. Wins, barriers and lessons, with the adherence numbers already
    filled in.

### G. Usability and trust (14)

65. Works in Microsoft Excel 2010 and later, and in Google Sheets —
    stated plainly, and tested in both.
66. Beginner instructions on Start Here and a one-line help note on
    every tab.
67. Colour-coded cells: cream to type in, grey calculated, and a key on
    two tabs.
68. Dropdowns and data validation everywhere a typo would break a
    lookup.
69. Editable exercise, habit, food, goal and activity lists.
70. Metric and imperial throughout, switched in one cell, including the
    conversion of everything already logged.
71. Frozen headers, filters and sensible column widths on every
    register.
72. Consistent formatting, one colour palette, no stray fonts.
73. Blank-safe formulas: an empty row shows an empty cell, never
    `#DIV/0!`, `#VALUE!` or a misleading zero.
74. Every formula cell locked, every sheet protected with a password
    that is written in the instructions.
75. Sample data clearly marked, listed on Start Here, and removable in
    one documented pass.
76. A second, completely blank copy of the workbook in the same
    download.
77. Readable fonts, accessible contrast, no red-on-green as the only
    signal.
78. Print areas set on every page that is meant to be printed.

### H. Packaging and support (12)

79. Professional cover sheet with the shop name and the version.
80. Listing copy that states compatibility, file types and what is
    included.
81. Screenshots of the real dashboard and three real tabs, not mockups.
82. One-page quick-start PDF in the download.
83. Formula and troubleshooting guide (what to do when a dropdown is
    empty, when protection asks for a password, when Sheets converts a
    date).
84. An example completed week so the buyer can see the file working.
85. The blank template for immediate use.
86. Digital-download and "how to make your own copy" instructions for
    Google Sheets.
87. FAQ covering the ten questions buyers actually ask.
88. Privacy note: the file lives on the buyer's machine or Drive,
    nothing is sent anywhere.
89. Personal-use licence and usage terms.
90. **Health and fitness disclaimer** on Start Here, in the listing and
    at the foot of the dashboard.

---

## 3. What I would add to your list to actually win the category

These are the gaps against the trackers already selling well at
$15–20. None of them are hard once the engine is there.

| Addition | Why it sells |
|----------|--------------|
| **Estimated 1RM and a plate calculator** | Lifters compare trackers on this one feature; the plate calculator (bar + plates per side) is the thing they screenshot |
| **Progressive-overload prompt** | "Last time: 3×8 at 60 kg" shown next to today's row. Nobody wants to scroll back |
| **Pace, split and distance maths for cardio** | min/km and min/mile both ways, so runners and walkers both get a usable log |
| **MET-based calorie estimate for activity** | Optional, editable, clearly an estimate — but buyers expect a number |
| **BMI, BMR and TDEE with the activity multiplier** | Standard in paid apps. Ship it as an estimate the buyer can overwrite, never as advice |
| **Deficit / surplus arithmetic** | Calories in minus estimated out, weekly, with the honest caveat that it is an estimate |
| **Six independent goals, not one** | Weight is only one of the things people track, and single-goal files get returned |
| **Rest days that protect a streak** | The single most common complaint about habit trackers |
| **Start-date agnostic calendars** | Works whether the buyer starts in January or on a random Tuesday; week start Monday or Sunday |
| **A 12-week programme block** | Phases, deload weeks, and a start date that drives every planner row |
| **Demo mode with one-click removal** | A filled file sells; a filled file the buyer cannot clear gets a one-star review |
| **Printable month report** | The buyer prints it, sticks it on the fridge, tells a friend |

And three things to deliberately **not** do: no calorie database (it
dates and it is a licensing problem), no automatic "you should eat X"
recommendation, no promise about how fast anyone will lose weight.

---

## 4. Build order

**Must-have, version 1** — Start Here, Profile & Goals, Dashboard,
Workout Planner, Workout Log, Exercise Library, Weight Tracker, Body
Measurements, Habit Tracker, Settings & Lists, plus sample data,
protected formulas and the disclaimer.

**High value, same version if time allows** — Cardio & Activity, Meals
& Macros, Water & Sleep, Reports & Charts, the weekly/monthly review.

**Differentiators** — records and 1RM, plate calculator, progressive
overload prompt, six goals, metric/imperial switch, printable reports,
demo mode.

**Polish** — troubleshooting guide, print areas, screenshot set,
Google Sheets pass, listing copy.

---

## 5. Quality gates before it is listed

The same five the other five products in this repo go through:

1. `pyflakes` clean on the generator.
2. Bracket and parenthesis balance scan on every formula in the file.
3. Structural validator: tab order, locked formulas, unlocked inputs,
   protection, document author, no `#REF!`, no unknown functions, no
   undefined names, no precedence slips.
4. Sample-data checker: every ID in the demo data joins to a real row,
   every total adds up, the demo week is internally consistent.
5. Full recalculation of both the filled and the blank workbook — every
   formula must return a real answer, in the blank file as well. Blank
   files are where trackers break.

Then: open it in Google Sheets, check the dropdowns and the conditional
formatting survived, and print one report.

---

## 6. Disclaimer that goes in the file and the listing

> This spreadsheet is a personal organisation and tracking tool. It is
> not medical advice, a diagnosis, a treatment plan or a substitute for
> professional care. Calorie, macro, body-fat and energy-expenditure
> figures are estimates produced from the numbers you type in, and all
> targets are yours to set and edit. Speak to a doctor or a registered
> professional before starting a new exercise or nutrition programme,
> particularly if you are pregnant, managing a health condition, or
> taking medication. No outcome of any kind is promised.

Short version for the Start Here banner and the dashboard footer:
*"For personal tracking only — not medical advice. All targets are
yours to set."*
