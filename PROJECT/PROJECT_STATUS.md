# Project Status

## Last Updated

2026-09-25

## Current Phase

**Closed.** The exam period ran **3 – 21 September 2026** and is over. On
25 September the site was stopped: the schedule was replaced with a closing
notice, the roster data was removed from the public web, and automatic
publishing was switched off (DECISION-013).

Nothing is running and nothing is scheduled. The next work on this project
would be reopening it for a new exam period.

## Current Objective

None. The project is dormant by request.

To reopen: set `WINDOW_START` in `build.py` to the first day of the new period,
delete the guard at the top of `sync.cmd`, and restore the `push` trigger in
`.github/workflows/pages.yml`. All three are commented in place.

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

- [COMPLETED] Site closed at the end of the period: closing notice published,
  roster data removed, publishing stopped.

## In Progress

Nothing.

Left unfinished when the period ended, and now moot unless the project reopens:

- [NEEDS VERIFICATION] The **Assignments** sheet and the sent flags are
  installed and tested against a copy and the live workbook, but no email was
  ever actually sent through either sheet.
- Note: `_SentLog` may still **not exist in the workbook**. It is created the
  first time a list is refreshed or an email sent. Its absence is normal, not
  a failed installation.

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
  account. Mitigated with `robots.txt` + `noindex`, not solved. *Moot while
  closed* — the live page now carries no names, rooms or times. It becomes live
  again the moment the site reopens.
- **Duplicate spellings of the same person** in the workbook (e.g. `Rajaie
  Nassar` / `Rajaie Ghassan Fawzi Nassar` / `Rajaee Nassaer`). Each spelling
  needs its own row in `Emails`. Flagged in column E.
- **The freshness stamp was dropped** in the redesign and not replaced, so a
  reader cannot tell how current the page is. Regression, not by request.

## Recent Important Changes

- 2026-09-25 — **site closed.** Closing notice published, `data.json` and
  `artifact.html` deleted, push-to-deploy removed, `sync.cmd` guarded
  (DECISION-013). Answers Q-01.
- 2026-09-16 and 2026-09-17 — the committee's last two `sync.cmd` refreshes of
  the period.
- 2026-09-12 11:57 — added the Assignments sheet, the `_SentLog` ledger and
  the Sent/Due columns. Fixed a latent bug where refreshing wiped the sent log.
- 2026-09-12 00:45 and 09:44 — two routine `sync.cmd` refreshes by the
  committee. The live site is current as of 09:44, 20 exams, 42/42 seats.
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

None. Do not restart work here without being asked.

If the committee returns for a new exam period, reopen as described under
*Current Objective*, then the first real task is still the one that was never
finished: send **one** assignment email for a single exam and confirm the Sent
flag survives a Refresh, before using "Send all not yet sent".

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
