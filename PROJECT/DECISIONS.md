# Decisions

Permanent decisions and the reasoning behind them. A superseded decision is
kept and marked, never deleted, so the shape of the project stays explicable.

---

## DECISION-001

### Date
2026-09-02

### Decision
Generate a self-contained static page from the workbook, rather than building
an application that reads the spreadsheet live.

### Reason
The data changes a few times a day at most, is edited by one person in Excel,
and is read by ~21 people who mostly want one fact each. A generated page has
no backend, no hosting cost, no authentication to maintain, and still works
from disk or an email attachment.

### Alternatives Considered
- A live app reading the workbook — needs a server with access to OneDrive.
- Sharing the workbook itself — unreadable on a phone.
- SharePoint/OneDrive sharing — genuinely private, but serves `.html` as a
  download rather than a page.

### Impact
`build.py` + `template.html` produce everything. No runtime dependencies.

### Status
ACTIVE

---

## DECISION-002

### Date
2026-09-02

### Decision
Recompute proctor shortfalls from the filled `Proctor 1..10` cells instead of
reading the workbook's `Still Needed` column.

### Reason
That column's formulas are broken: row 22 referenced row 25, and rows 23, 24,
33 and 35 had no formula at all. It reported a shortfall that did not exist and
missed rows entirely — which is what prompted the original request.

### Alternatives Considered
Fixing the spreadsheet formulas alone — would leave the site trusting a column
that a stray edit can break again.

### Impact
The site's coverage figures are independent of spreadsheet formula health.
The workbook's own display can still be wrong; `README.md` records the fix.

### Status
ACTIVE

---

## DECISION-003

### Date
2026-09-02

### Decision
Host on public GitHub Pages with `robots.txt` and `noindex`, accepting that the
page is world-readable.

### Reason
GitHub Pages does not serve privately on a free account, and GitHub Pro does
not change that — Pro hides the source, not the site. Only Enterprise Cloud
offers private Pages. The user was told this and chose to proceed.

### Alternatives Considered
- OneDrive/SharePoint org-only sharing — real access control, but downloads
  the file instead of rendering it.
- Emailing `index.html` — fully private, no live updates.
- GitHub Pro — costs money, solves nothing here.

### Impact
Staff names, rooms and student counts are publicly reachable by URL. Search
engines are excluded, which closes the realistic exposure route (someone
googling a colleague's name).

### Status
ACTIVE

---

## DECISION-004

### Date
2026-09-03

### Decision
Reminder emails send through classic Outlook COM, with a confirmation dialog
on by default.

### Reason
The committee asked for one-click sending. A wrong email to twenty colleagues
cannot be recalled, so each send lists its recipients first. `CONFIRM_BEFORE_SEND`
and `SEND_MODE` are constants the committee can flip once they trust it.

### Alternatives Considered
- `.Display` only — safe but not what was asked for.
- Graph API / SMTP — no desktop dependency, but needs credentials and
  app registration.

### Impact
Sending requires classic Outlook running. New Outlook cannot drive it.

### Status
ACTIVE

---

## DECISION-005

### Date
2026-09-03

### Decision
Pair proctor *n* with room *n*, positionally.

### Reason
It is the convention the committee already uses. Verified against their own
room-assignment PDF across every multi-room exam — Engineering Maths,
Engineering Design, Engineering Science and others all matched exactly.

### Alternatives Considered
- A separate per-proctor room column — more accurate, but more to maintain and
  a change to how the committee already works.
- Not showing per-proctor rooms — leaves staff guessing.

### Impact
The site, the Room Plan sheet and the emails all derive rooms the same way and
cannot disagree. If the committee ever lists rooms in a different order, all
three are wrong together and silently.

### Status
ACTIVE

---

## DECISION-006

### Date
2026-09-03

### Decision
Change VBA with `update_vba.py`, which replaces code inside the existing
`.xlsm`. Keep `make_xlsm.py` only for first creation.

### Reason
`make_xlsm.py` rebuilds from the `.xlsx`, which would erase the 40 email
addresses and the edited message wording. Those were entered by hand and are
not recoverable from the source.

### Alternatives Considered
Rebuilding and re-importing the data each time — fragile and slow.

### Impact
`make_xlsm.py` is effectively a disaster-recovery tool. Anything that touches
VBA goes through `update_vba.py`.

### Status
ACTIVE

---

## DECISION-007

### Date
2026-09-04

### Decision
Redesign the site around a "highlighted roster": show the whole schedule as a
ruled list and mark the viewer's own rows in marigold.

### Reason
The previous design was a competent but generic instrument panel. The physical
equivalent of this task is a printed list on a noticeboard that people
highlight their own name on — grounding the design there gave it a specific
identity and put the viewer's duties in the context of the day.

### Alternatives Considered
Keeping the card-based dashboard; filtering to only the viewer's rows.

### Impact
Fraunces + Karla, ink on warm paper, one marigold accent. Structural devices
carry meaning (dots = exams that day, red = a gap) rather than decorate.

### Status
ACTIVE

---

## DECISION-008

### Date
2026-09-04

### Decision
Light theme only. No dark mode, and no `prefers-color-scheme` handling at all.

### Reason
Explicitly requested. Removed outright rather than hidden, so there is no
half-present theme to drift.

### Alternatives Considered
Keeping dark mode behind a toggle — rejected; the request was unambiguous.

### Impact
Every colour is stated on `:root` with `color-scheme: light`, so the page holds
its appearance whatever ground the host paints behind it.

### Status
ACTIVE

---

## DECISION-009

### Date
2026-09-05

### Decision
Validate dates in two places — reject bad values in `build.py`, and make the
page's `asDate` return null for anything that is not a real `yyyy-mm-dd`.

### Reason
A stray `z` typed into one Date cell took the entire site down. The root cause
was subtler than the typo: `asDate` split on `-` and accepted any three parts,
so `"not-a-date"` produced an *Invalid Date* object, which is truthy and
slipped past every downstream guard.

### Alternatives Considered
Fixing only the spreadsheet cell — leaves the page one keystroke from breaking
again. Fixing only `build.py` — leaves the page trusting its input.

### Impact
An unusable date now groups under "Date not set" and the rest of the page
carries on. `build.py` prints the offending row at sync time, so a typo is
caught by the committee rather than by staff.

### Status
ACTIVE

---

## DECISION-010

### Date
2026-09-05

### Decision
Remove the "still needs a proctor" summary banner from the page.

### Reason
Requested. The committee did not want the page announcing staffing gaps to
staff.

### Alternatives Considered
Keeping it for coordinators only — no way to tell them apart on a public page.

### Impact
Gaps remain visible where they are actionable — a red dot on that day in the
strip, and a red "Nobody assigned" row against the empty room — but there is
no longer a page-level summary.

### Status
ACTIVE
