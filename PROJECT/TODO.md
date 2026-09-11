# TODO

Remaining work only. Current state lives in `PROJECT_STATUS.md`.

Nothing here is committed work — the project is functionally complete and in
routine use. These are open items and offers the user has not taken up.

---

# High Priority

- [ ] **T-01 — Add the three missing coordinator addresses.**
  `Dr. Lutfi`, `Rami Hammad`, `Razan Fayez Mahmoud Shatnawi`. They are not in
  the committee's `emails.xlsx`, so they cannot be found automatically.
  *Blocked on: the user supplying the addresses.*
  *Effect while open: reminders for their exams reach the proctors but not them.*
  *Where: `Emails` sheet, column C.*

- [ ] **T-02 — Restore the freshness stamp on the page.**
  The old design showed "updated 2 Sep 09:50"; the redesign dropped it and
  nothing replaced it. A reader cannot tell whether the page is current.
  *This is a regression, not a feature request.*
  *Where: `template.html` masthead; `DATA.generatedAt` already carries the value.*

# Medium Priority

- [ ] **T-03 — Remember the viewer between visits.** (F-12)
  One `localStorage` write, so the page opens on the person's own duties
  instead of asking every time. Largest ease-of-use gain for the effort.
  *Depends on: nothing.*

- [ ] **T-04 — Show "your next duty is in N minutes".** (F-13)
  The commonest question deserves the most prominent answer.
  *Depends on: T-05 is natural to do at the same time.*

- [ ] **T-05 — Make the page time-aware within the day.**
  Exams currently dim only once the whole date has passed, so at 15:00 a
  finished morning exam still looks pending.

- [ ] **T-06 — Verify the print stylesheet.** (U-07)
  Written but never exercised. Committees print rosters.

# Low Priority

- [ ] **T-07 — Make room numbers tappable** — "who else is in N-301?".
- [ ] **T-08 — Signal the building** — `W-202` and `N-301` are a real walk apart
  and nothing on the page says so.
- [ ] **T-09 — Include the page link when copying duties.** (I-03)
  Forwarded text currently has no route back to the live version.
- [ ] **T-10 — Sort zero-duty people below a divider** in the name picker.
  Five of 26 have no duties and sit in the middle of the list.
- [ ] **T-11 — Compress the masthead for a returning viewer.**
- [ ] **T-12 — Reorder tabs so "My duties" comes first**, matching actual use.
- [ ] **T-13 — One email per proctor, personally addressed.** (R-12)
  Offered 2026-09-03; the shared-list version was preferred at the time.

# Bugs

*None open.* The last defect — a bad Date cell blanking the whole site — was
fixed and hardened on 2026-09-05 (DECISION-009).

- [x] One bad date cell took the entire page down — fixed 2026-09-05

# Improvements

- [ ] **T-14 — Make `WINDOW_START` less hardcoded** if the site outlives this
  exam period. Currently `2026-09-03`, overridable by environment variable.
- [ ] **T-15 — Add a check that rooms and proctors are listed in the same
  order.** The positional pairing (DECISION-005) is a convention nothing
  enforces; if the committee ever reorders one list, every view is silently
  wrong.
- [ ] **T-16 — Verify a real reminder send.** Only the preview path has been
  exercised, deliberately, since a test send would email real colleagues.

# User Decisions Required

- [ ] **Q-01 — What happens after 21 September?** Keep the site running for the
  next exam period, or retire it?
- [ ] **Q-02 — Should the repo move to an HTU-owned GitHub organisation?**
  It currently sits under `SETcommittee`, created from a personal account.
- [ ] **Q-04 — Are any of T-03 to T-12 wanted** before the exam period ends?
- [ ] **Should the red gap markers stay?** After the summary banner was removed
  (DECISION-010), the red day-strip dot and "Nobody assigned" row are the only
  remaining signs of a staffing gap. Deliberate, but worth confirming.
