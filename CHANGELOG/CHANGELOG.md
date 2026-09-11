# Changelog

Newest first. Reconstructed from git history and the working record.
Routine `sync.cmd` data refreshes are summarised, not listed individually.

---

# 2026-09-12

## Task

Set up the project memory and development-context system.

## Changes

- Added `CLAUDE.md`, `PROJECT/` (requirements, architecture, decisions, status,
  todo), `CHANGELOG/`, and the `project-maintenance` skill.
- Documented ten permanent decisions taken between 2 and 5 September.
- Recorded the current verified state: 20 exams, 42/42 seats, 40/43 addresses.

## Files/Components Affected

- `CLAUDE.md`, `PROJECT/*`, `CHANGELOG/CHANGELOG.md`,
  `.claude/skills/project-maintenance/*`

## Verification

- Project inspected before writing: structure, git history, build output, live
  site, workbook sheets and address completeness.
- No project functionality touched.

## Remaining Issues

- Three coordinator addresses still missing.
- Print stylesheet still unverified.

## Next Step

Ask whether any `TODO.md` high-priority item is wanted before 21 September.

---

# 2026-09-05 – 2026-09-09

## Task

Routine schedule refreshes by the committee.

## Changes

- Roughly ten `sync.cmd` runs, publishing schedule changes as assignments moved.
- Exam period extended; the schedule now runs to 21 September.
- The bad Date cell in row 36 was corrected in the workbook.
- Exam count grew to 20; coverage reached 42/42 seats.

## Files/Components Affected

- `data.json`, `index.html`, `artifact.html` (all generated)

## Verification

- Each run publishes only if the build succeeds.

## Remaining Issues

- None arising.

## Next Step

None; routine operation.

---

# 2026-09-05

## Task

Fix the site being entirely blank, reported as "full schedule not working".

## Changes

- Found a stray `z` typed into the Date and Day cells of Exam row 36. It
  reached the page and threw before anything rendered, so **every** view was
  dead, not just the full schedule.
- Root cause was deeper than the typo: `asDate` split on `-` and accepted any
  three parts, so `"not-a-date"` returned an *Invalid Date* — truthy, and able
  to slip past every downstream guard. Rewrote it to match `yyyy-mm-dd` and
  range-check the result.
- `build.py` now rejects non-dates, empties the field, and prints the offending
  row so a typo is caught at sync time.
- Earlier the same day: removed the "still needs a proctor" banner; added the
  weekday to every date; restored the Full schedule search, day strip and
  filters in the new style.

## Files/Components Affected

- `template.html`, `build.py`

## Verification

- Poisoned the data with four bad-date shapes — `z`, `not-a-date`,
  `2026-13-45`, empty — and confirmed a clean load with no console errors and
  all 19 exams rendered.
- Confirmed the deployed site served the hardened parser.

## Remaining Issues

- Print stylesheet still unverified.

## Next Step

Offer the usability improvements; none were commissioned.

---

# 2026-09-04

## Task

Redesign the site, then reshape it to the committee's preferences.

## Changes

- Rebuilt the page around a "highlighted roster" — the whole schedule as a
  ruled list with the viewer's rows marked in marigold (DECISION-007).
- Fraunces + Karla, ink on warm paper, one accent colour.
- `build.py` began emitting per-proctor room pairing, so the site, the Room
  Plan sheet and the emails all agree.
- Then, on request: restored the three view tabs, removed dark mode outright
  (DECISION-008), dropped "roster" from the title.

## Files/Components Affected

- `template.html`, `build.py`

## Verification

- Light and dark browser settings, desktop and mobile; no console errors; no
  horizontal page scroll at 375px.

## Remaining Issues

- Freshness stamp lost in the redesign and not replaced.

## Next Step

Restore the search and day strip the redesign had dropped.

---

# 2026-09-03

## Task

Build the reminder-email system in the workbook.

## Changes

- Created `Proctor Schedule by Date.xlsm` alongside the `.xlsx`, adding the
  `Emails`, `Reminders`, `Room Plan` and `Message` sheets plus their VBA.
- Filled 39 of 43 addresses by fuzzy-matching the committee directory. A first
  pass produced three wrong matches — first-name-plus-incidental-overlap — which
  a dry run caught before anything was written.
- Diagnosed Excel "freezing" on send: Outlook was running headless with no
  window, blocking every COM call. `GetOutlook` now logs MAPI on explicitly and
  opens an Explorer window.
- Added sender selection, an always-Cc list, per-proctor rooms in the body, and
  moved the wording onto an editable `Message` sheet.
- `build.py` switched to reading the `.xlsm` in preference to the `.xlsx`.

## Files/Components Affected

- `vba/mReminders.bas`, `vba/Reminders_sheet.bas`, `vba/update_vba.py`,
  `vba/make_xlsm.py`, `vba/fill_emails.py`, `build.py`, `README.md`

## Verification

- VBA compiles; `EmailFor` resolves a known name; `RefreshRemindersCore`
  returns the expected row count.
- Preview opened a real Outlook draft with correct To and Cc and returned
  without hanging. Every address was neutralised to `example.invalid` first.

## Remaining Issues

- Four coordinators without addresses (since reduced to three).
- A real `.Send` never exercised, deliberately.

## Next Step

Add per-proctor rooms to the emails.

---

# 2026-09-02

## Task

Answer "which exams are still unproctored?", then build the site.

## Changes

- Audited the workbook. **Every exam was already fully proctored.** The
  appearance of gaps came from broken formulas in the `Still Needed` column —
  row 22 referenced row 25, and rows 23, 24, 33 and 35 had no formula at all.
- Built `build.py` + `template.html` to generate a self-contained page.
- Published to GitHub Pages; added `robots.txt` and `noindex` before making the
  repo public.
- Renamed the GitHub account, then moved the repo to a `SETcommittee`
  organisation so the link is institutional rather than personal.
- Wrote `sync.cmd`, the one-click rebuild-commit-push.

## Files/Components Affected

- `build.py`, `template.html`, `sync.cmd`, `README.md`,
  `.github/workflows/pages.yml`, `robots.txt`

## Verification

- Build reported 33/33 seats filled across 18 sittings.
- Cross-checked against the master SET exam schedule: all 15 exams present.
- Site verified live: 200, correct content, `noindex` in place.

## Remaining Issues

- Page is world-readable (DECISION-003).
- Spreadsheet formulas still broken in the workbook itself.

## Next Step

Reminder emails.
