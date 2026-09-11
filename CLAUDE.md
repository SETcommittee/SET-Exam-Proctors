# CLAUDE.md

Permanent working rules for this project. Read this first, every session.

---

## Project Overview

**What it is.** Two things that share one data source:

1. A public website showing exam proctoring duties for staff at the School of
   Engineering & Technology, Al Hussein Technical University (HTU).
2. A macro-enabled Excel workbook the exam committee uses to assign proctors,
   plan rooms, and send reminder emails.

**Purpose.** A proctor opens the site on a phone and sees where they must be
and when. A course coordinator sees who is covering their exam. The committee
assigns duties in Excel and pushes both the website and the reminder emails
from there.

**Platform / environment.**

| | |
|---|---|
| OS | Windows 11 |
| Python | 3.13 via the `py -3` launcher, with `openpyxl` and `pywin32` |
| Office | Excel desktop and **classic** Outlook desktop (Office16) |
| Hosting | GitHub Pages from `SETcommittee/SET-Exam-Proctors` (public repo) |
| Live site | https://setcommittee.github.io/SET-Exam-Proctors/ |
| Private copy | https://claude.ai/code/artifact/68b0fc6b-ff72-4104-a828-407373d9416c |

**Constraints that shape everything:**

- The source workbook lives in **OneDrive, outside this repository**, so the
  build only runs on the committee's machine. Nothing rebuilds in CI.
- GitHub Pages is **public even though it serves a private-looking roster**.
  The page carries staff names, rooms and student counts. `robots.txt` and
  `noindex` keep it out of search engines; that is the only protection.
- VBA automation reaches **classic Outlook only**. The new Outlook app
  (`olk.exe`) has no COM interface and cannot be used.
- The site is **light-only by explicit request**. Do not reintroduce dark mode.
- This repo is deployed to real staff mid-exam-period. A broken push is visible
  to ~21 people within a minute.

---

## Source of Truth

In order of authority:

1. **Actual project files** — authoritative for how things are implemented.
2. **`Proctor Schedule by Date.xlsm`** (OneDrive) — authoritative for exam,
   proctor, room and email **data**. It is not in this repo.
3. `PROJECT/REQUIREMENTS.md` — agreed requirements.
4. `PROJECT/ARCHITECTURE.md` — current architecture and the file map.
5. `PROJECT/DECISIONS.md` — permanent decisions and why.
6. `PROJECT/PROJECT_STATUS.md` — current state.
7. `PROJECT/TODO.md` — unfinished work.
8. `CHANGELOG/CHANGELOG.md` — history.

If a document disagrees with the files or the workbook, **investigate the
conflict**. Do not assume the document is right, and do not assume it is wrong.

---

## Context Loading Rules

**Do not reread the whole project for every task.**

On starting a session:

1. Read `CLAUDE.md` (this file).
2. Read `PROJECT/PROJECT_STATUS.md`.
3. Read `PROJECT/TODO.md`.
4. Read the most recent `CHANGELOG/CHANGELOG.md` entries.
5. Work out which requirements and decisions bear on the request.
6. Use the **Relevant File Map** in `ARCHITECTURE.md` to find the files that
   matter.
7. Inspect only those files.

Open unrelated or finished areas only when there is a specific reason to.

---

## Editing Rules

- Make the smallest change that correctly solves the task.
- Do not modify unrelated files.
- Preserve existing behaviour unless a change was asked for.
- Do not redesign the architecture without being asked.
- Do not add a second way to do something that already exists.
- Do not assume a change works. Verify it.
- Do not invent requirements.
- Ask when a decision is genuinely the user's to make.
- Match the surrounding code and prose style.
- Check who depends on a shared component before changing it.

**Specific to this project:**

- **Never edit `index.html`, `artifact.html` or `data.json` by hand.** They are
  generated. Edit `template.html` or `build.py` and rebuild.
- **Never rebuild the workbook with `vba/make_xlsm.py`** unless the `.xlsm` is
  gone. It rebuilds from the `.xlsx` and would erase the email addresses and
  wording entered by hand. Use `vba/update_vba.py` to change VBA in place.
- The workbook must be **closed in Excel** before any script writes to it.
- Treat the workbook as the committee's live data. Do not write to its cells
  without asking, even to correct an obvious typo.

---

## Documentation Rules

After meaningful work:

- Update `PROJECT/PROJECT_STATUS.md`.
- Update `PROJECT/TODO.md` if the task list moved.
- Add a `CHANGELOG/CHANGELOG.md` entry.
- Update `PROJECT/ARCHITECTURE.md` if the architecture actually changed.
- Update `PROJECT/DECISIONS.md` when a lasting decision is taken.

Skip all of this for trivial changes that would add no useful information.
A routine `sync.cmd` data refresh is not a documentation event.

---

## Resume Protocol

When the user says "continue the project", or returns after a gap:

1. Load context in the order under **Context Loading Rules**.
2. Read the current status.
3. Review unfinished tasks.
4. Review recent changes.
5. Pick the most relevant next task, or ask if it is ambiguous.
6. Inspect only the files that task needs.
7. Continue from where the project actually is.

Do not re-analyse the project from scratch.

---

## Verification Expectations

Say plainly which of these is true:

- "I implemented it" — the change is written.
- "I implemented and verified it" — the change is written **and** checked.

For this project, verification usually means one of:

- `py -3 build.py` runs clean and reports the expected exam and seat counts.
- The page loads in a browser with **no console errors** and the three tabs work.
- The deployed site returns 200 and contains the expected change.
- A VBA change compiles and a dialog-free function returns the expected value.

Never report work as finished on the strength of the code looking right.
