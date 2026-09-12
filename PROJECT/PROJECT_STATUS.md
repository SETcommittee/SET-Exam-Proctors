# Project Status

## Last Updated

2026-09-12

## Current Phase

**Live and in use, mid exam period.** Build and feature work is essentially
done; the project is in routine-operation mode. The exam period runs
**3 – 21 September 2026**, so the site is being read by staff right now and
the committee is refreshing it from Excel as assignments change.

## Current Objective

Keep the published schedule accurate until the exam period ends. The email
side has just gained a second message type, which needs one real send to
confirm it behaves in practice.

## Completed

- [COMPLETED] Website generated from the committee workbook, published to
  GitHub Pages, kept out of search engines.
- [COMPLETED] Three views: Full schedule (search, day strip, filters),
  My duties, By course.
- [COMPLETED] Per-proctor room pairing, shared by the site, the Room Plan
  sheet and the reminder emails.
- [COMPLETED] Reminder emails from Excel via classic Outlook, sending as
  `set.exam.committee@htu.edu.jo`, with an always-Cc list.
- [COMPLETED] Email wording moved onto a `Message` sheet, editable as cells.
- [COMPLETED] Emails sheet populated by fuzzy-matching the committee directory.
- [COMPLETED] One-click `sync.cmd` refresh; the committee has run it ~15 times
  unaided since 3 September.
- [COMPLETED] Bad-date crash fixed and hardened at two layers.
- [COMPLETED] Assignment emails on their own sheet, plus a durable sent flag
  for both kinds, held in a hidden `_SentLog` ledger.

## In Progress

- [NEEDS VERIFICATION] The new **Assignments** sheet and the sent flags are
  installed and tested against a copy and the live workbook, but no email has
  actually been sent through either sheet yet.

## Not Started

- [NOT STARTED] The usability improvements suggested on 2026-09-05 and never
  commissioned — remembering the viewer, a "what's next" line, restoring the
  freshness stamp. See `TODO.md`.

## Blocked

- [BLOCKED] Three coordinators still have no email address, because they do not
  appear in the committee's `emails.xlsx` directory. Needs the addresses from
  the user; cannot be resolved from available data.

## Known Issues

- **Three missing coordinator addresses** — `Dr. Lutfi`, `Rami Hammad`,
  `Razan Fayez Mahmoud Shatnawi`. Reminders for their exams still reach the
  proctors; the coordinator is simply not Cc'd. Flagged in red on the
  `Reminders` sheet.
- **The published page is world-readable.** Inherent to GitHub Pages on a free
  account. Mitigated with `robots.txt` + `noindex`, not solved.
- **Duplicate spellings of the same person** in the workbook (e.g. `Rajaie
  Nassar` / `Rajaie Ghassan Fawzi Nassar` / `Rajaee Nassaer`). Each spelling
  needs its own row in `Emails`. Flagged in column E.
- **The freshness stamp was dropped** in the redesign and not replaced, so a
  reader cannot tell how current the page is. Regression, not by request.

## Recent Important Changes

- 2026-09-12 — added the Assignments sheet, the `_SentLog` ledger and the
  Sent/Due columns. Fixed a latent bug where refreshing wiped the sent log.
- 2026-09-09 — latest data refresh by the committee; 20 exams, 42/42 seats.
- 2026-09-05 — fixed a crash where one bad Date cell (`z` in row 36) blanked
  the entire site; hardened both the date parser and `build.py`.
- 2026-09-05 — restored the Full schedule search, day strip and filters in the
  new visual style.
- 2026-09-04/05 — redesigned the site; tabs restored, dark mode removed.
- 2026-09-03 — reminder emails gained per-proctor rooms, a chosen sender, and
  an always-Cc list.

## Current Files/Components Being Worked On

- `vba/mReminders.bas`, `vba/Assignments_sheet.bas`, `vba/update_vba.py`
  (2026-09-12). Website files untouched since 2026-09-05.

## Next Recommended Task

Send **one** assignment email, for a single exam, and confirm the Sent flag
appears and survives a Refresh. That is the only part of the new work not yet
proven in practice. Only then use "Send all not yet sent", which would email
every upcoming exam at once.

## Verification Status

**Tested and passing:**

- `build.py` against the live workbook — 20 exams, 42/42 seats, no bad dates.
- Site rendering in light mode, desktop and mobile, no console errors.
- All three tabs, the day strip filter, search with caret retention.
- No horizontal page scroll on a 375px viewport.
- Bad-date resilience against `z`, `not-a-date`, `2026-13-45` and empty.
- Deployed site returns 200 with `noindex` and `robots.txt` intact.
- VBA compiles; `EmailFor` resolves; `RefreshRemindersCore` returns row counts.
- Outlook preview path opens a real draft with correct To/Cc and does not hang.

**Not tested:**

- [NEEDS VERIFICATION] The **print stylesheet**. Written, never exercised.
- [NEEDS VERIFICATION] A real reminder **send** (`.Send`). Only the `.Display`
  preview path has been exercised, deliberately — testing a send would email
  real colleagues.
- [NEEDS VERIFICATION] `SendAllUpcoming`, `SendDueReminders` and
  `SendDueAssignments` in bulk - none has been run.
- [NEEDS VERIFICATION] The assignment email has never actually been sent.
- [NEEDS VERIFICATION] Behaviour on browsers other than the Chromium-based
  preview pane.

## Important Context

- **The committee runs this themselves.** The many "Update proctor schedule"
  commits are the user double-clicking `sync.cmd`, not development work.
- **The workbook is the data source and is not in the repo.** Any question of
  the form "why does the site show X" is answered by opening the workbook, not
  by reading the site code.
- **Exam dates extended** from 14 to 21 September after the original build, so
  date ranges in older notes may look wrong; they are simply older.
