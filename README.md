# Exam Proctor Duty — SET

A one-page site so staff can look up their own proctoring duties, and course
coordinators can see who is covering their exams. Built from the Exam
Committee's `Proctor Schedule by Date.xlsx`.

Covers exams from **3 September 2026** onwards.

## For staff

Open the site and pick a tab:

- **My duties** — start typing your name (or hit *Browse all*) and pick yourself
  from the list. You get your duty count, hours, next duty, a day strip with
  your days highlighted, and an "Add to calendar" link per duty. Once you are
  selected the address bar holds your own link (`…/#p=Dana%20Slehat`) — copy it
  and send it to that person so it opens straight to their duties.
- **By course** — coordinators type a course name (or their own name) and see
  who is proctoring each sitting, the rooms, and the student count.
- **Full schedule** — everything. Tap a day in the strip to jump to it, or
  search by course, coordinator, room or proctor name, filtered by department.

Names and course lists stay hidden until someone searches or asks to browse, so
the page does not open onto a wall of chips.

Each card lists its proctors as a numbered roster — the number is the proctor
slot, so a card can be checked straight against the spreadsheet's
`Proctor 1`–`Proctor 10` columns. Unfilled slots show as **Not assigned yet**.

A banner at the top always states whether every exam is fully covered, and names
any exam that is short.

## Updating it after the spreadsheet changes

Double-click **`sync.cmd`**. It rebuilds the page from the workbook, commits,
and pushes. GitHub Pages redeploys in about a minute.

To rebuild without publishing:

```bash
py -3 build.py
```

Then open `index.html`.

## How it fits together

| File | What it does |
|---|---|
| `build.py` | Reads the workbook, writes `data.json` and `index.html` |
| `template.html` | The page — markup, styles and behaviour, with a `/*__DATA__*/` slot |
| `index.html` | Generated. The whole site, data baked in — no server needed |
| `artifact.html` | Generated. The same page without the `<html>`/`<head>` wrapper |
| `data.json` | Generated. The extracted schedule, if you want it elsewhere |
| `sync.cmd` | Rebuild + commit + push, in one double-click |
| `.github/workflows/pages.yml` | Deploys to GitHub Pages on every push to `main` |

`index.html` is fully self-contained, so it also works over email, from a USB
stick, or opened straight off disk.

### Where the source lives

`build.py` defaults to:

```
C:\Users\Faris.Alsalem\OneDrive - AL-Hussien bin Abdullah Technical University\Faris HTU\Exam Committee\Proctor Schedule by Date.xlsx
```

Override it without editing the script:

```bash
set PROCTOR_XLSX=D:\some\other\Proctor Schedule by Date.xlsx
set WINDOW_START=2026-09-03
py -3 build.py
```

Or pass the path as the first argument.

### Which sheets it reads

- **`Exam`** — header on row 5, data from row 6. Columns A–V: course,
  coordinator, department, date, day, start, finish, room, students, required
  proctors, then `Proctor 1`–`Proctor 10`.
- **`Proctors`** — the roster, names in column A from row 6. Used so that people
  with no duties in the window still appear in the name list (showing `0`).

The `Proctor Lookup` and `Course Lookup` sheets are not read — the site replaces
what they do.

> **Proctor counts are recalculated, not copied.** The workbook's `Still Needed`
> column cannot be trusted: on row 22 the formula points at row 25, and rows 23,
> 24, 33 and 35 have no formula at all. `build.py` counts the filled
> `Proctor 1`–`Proctor 10` cells instead and compares that to `# of Proctors`.
> Worth fixing in the spreadsheet too — see the note at the bottom.

## First-time setup

1. Create an empty repository on GitHub (private is fine — see the note below).
2. From this folder:

   ```bash
   git init -b main
   git add -A
   git commit -m "Exam proctor duty site"
   git remote add origin https://github.com/SETcommittee/SET-Exam-Proctors.git
   git push -u origin main
   ```

3. On GitHub: **Settings → Pages → Source → GitHub Actions**.
4. The site appears at `https://setcommittee.github.io/SET-Exam-Proctors/`.

Set your git identity once if you have not already:

```bash
git config --global user.name "Faris Alsalem"
git config --global user.email "faris.alsalem@htu.edu.jo"
```

### A note on privacy

GitHub Pages serves the site **publicly** even from a private repository — the
repo stays private, the published page does not. The page lists staff names,
room allocations and student counts. If that should not be on the open web,
either publish it somewhere behind the university login instead, or ask the
committee first.

## Reminder emails

The committee's workbook now has a macro-enabled twin,
`Proctor Schedule by Date.xlsm`, in the same folder. **That is the file to
edit from now on** — `build.py` reads it in preference to the `.xlsx`.

It adds two sheets:

**`Emails`** — all 43 people who appear as a proctor or a coordinator, with a
column for their address. The macro reads only column C. Column D holds a
`firstname.lastname@htu.edu.jo` guess to save typing; check each one before
copying it across. Column E flags people who appear under more than one
spelling (e.g. `Rajaie Nassar` / `Rajaie Ghassan Fawzi Nassar` /
`Rajaee Nassaer` are one person) — give each spelling the same address.

**`Reminders`** — one row per exam, with recipients resolved from the `Emails`
sheet. Click any row and the box at the top shows exactly what would be sent.
The coordinator is put in Cc, not addressed in the body. Then either double-click that row's **Send** cell, or use the buttons:

| Button | What it does |
|---|---|
| Refresh list | Rebuilds the rows from the `Exam` sheet. Run after any change. |
| Send for selected row | Sends the reminder for whichever row the cursor is on |
| Preview in Outlook | Opens the email in Outlook without sending |
| Send all upcoming | Sends every future exam's reminder, one email each |

The **No address for** column lists anyone whose address is still missing, in
red. Those people are silently left off the email, so keep that column empty.

### Sending behaviour

Two constants at the top of the `mReminders` module control this:

```vba
Public Const CONFIRM_BEFORE_SEND As Boolean = True
Public Const SEND_MODE As String = "SEND"
```

`CONFIRM_BEFORE_SEND` shows a dialog listing every recipient before anything
leaves. Set it to `False` for true one-click sending once you trust the
addresses. `SEND_MODE` can be `"DISPLAY"` instead of `"SEND"` to make every
button open the message in Outlook rather than send it.

Outlook must be installed and signed in — it is used to send, so the messages
land in your Sent Items as normal.

### Rebuilding the workbook

If the `.xlsm` is lost or the macros need changing, edit the `.bas` files in
`vba/` and run:

```bash
py -3 vba/make_xlsm.py
```

That regenerates the `.xlsm` from the `.xlsx` plus the VBA. It never modifies
the `.xlsx`.

## Fixing the spreadsheet

The site works around the broken formulas, but the workbook itself still shows
wrong numbers. In the `Exam` sheet:

- **V22** should be `=IF(J22="","",MAX(0,J22-U22))` — it currently reads `J25`.
- **U23, U24, U33, U35** need `=COUNTA(K<row>:T<row>)`.
- **V23, V24, V33, V35** need `=IF(J<row>="","",MAX(0,J<row>-U<row>))`.
