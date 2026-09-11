# Requirements

What the project must do. Derived from what was actually asked for and built —
nothing here is speculative. Anything genuinely undecided is marked
`[TBD — USER DECISION REQUIRED]`.

Statuses: CONFIRMED · IN PROGRESS · COMPLETED · TBD · REJECTED

---

## Functional — website

| ID | Requirement | Priority | Status | Notes |
|---|---|---|---|---|
| F-01 | A proctor can find their own duties by name | High | COMPLETED | Search-and-pick; list hidden until typing |
| F-02 | Show the full schedule of every exam in the period | High | COMPLETED | "Full schedule" tab |
| F-03 | A coordinator can see who is proctoring their course | High | COMPLETED | "By course" tab |
| F-04 | Show each proctor the room they are covering | High | COMPLETED | Positional pairing, DECISION-005 |
| F-05 | Search by course, room, coordinator, department or proctor | Medium | COMPLETED | Full schedule only |
| F-06 | Filter to one day | Medium | COMPLETED | Day strip; dots show exams per day |
| F-07 | Filter by status and department | Medium | COMPLETED | Still to come / Everything / Finished |
| F-08 | Per-person shareable link | Medium | COMPLETED | `#p=<name>`, opens on My duties |
| F-09 | Add a duty to a personal calendar | Low | COMPLETED | Google Calendar link, carries the room |
| F-10 | Copy one's duties as text | Low | COMPLETED | Clipboard; no link included (see I-03) |
| F-11 | Show exams short of proctors | Medium | COMPLETED | Per-exam only; banner removed, DECISION-010 |
| F-12 | Remember the viewer between visits | Medium | TBD | Suggested 2026-09-05, not commissioned |
| F-13 | Show "your next duty is in N minutes" | Medium | TBD | Suggested 2026-09-05, not commissioned |

## Functional — workbook and reminders

| ID | Requirement | Priority | Status | Notes |
|---|---|---|---|---|
| R-01 | Hold every person's email address in the workbook | High | COMPLETED | `Emails` sheet; 40 of 43 filled |
| R-02 | Send a reminder per exam to its proctors | High | COMPLETED | `Reminders` sheet, per-row send |
| R-03 | Cc the coordinator, not address them | High | COMPLETED | Not named in the body either |
| R-04 | Preview the email before sending | High | COMPLETED | Live preview pane + Outlook preview |
| R-05 | Send from the committee mailbox | High | COMPLETED | `set.exam.committee@htu.edu.jo` |
| R-06 | Always Cc a fixed set of people | Medium | COMPLETED | Cell E3; de-duplicated against To |
| R-07 | Edit the email wording without touching code | Medium | COMPLETED | `Message` sheet with placeholders |
| R-08 | Name each proctor's room in the email | High | COMPLETED | `{proctorrooms}` |
| R-09 | One row per proctor with their room | Medium | COMPLETED | `Room Plan` sheet, rebuildable |
| R-10 | Flag anyone with no address | High | COMPLETED | Red column on `Reminders` |
| R-11 | Send all upcoming reminders at once | Low | COMPLETED | [NEEDS VERIFICATION] never run in bulk |
| R-12 | One email per proctor, personally addressed | Low | TBD | Offered 2026-09-03, not taken up |

## Data

| ID | Requirement | Priority | Status | Notes |
|---|---|---|---|---|
| D-01 | The workbook is the single source of exam data | High | CONFIRMED | Not in the repo |
| D-02 | Only exams from the window start appear | High | COMPLETED | `WINDOW_START` = 2026-09-03 |
| D-03 | Coverage computed from assignment cells, not formulas | High | COMPLETED | DECISION-002 |
| D-04 | A bad cell must not break the site | High | COMPLETED | DECISION-009 |
| D-05 | Read the workbook while it is open in Excel | High | COMPLETED | Copies to temp first |
| D-06 | Report bad data to the committee at build time | Medium | COMPLETED | `BAD DATES` in build output |

## User interface

| ID | Requirement | Priority | Status | Notes |
|---|---|---|---|---|
| U-01 | Usable on a phone | High | COMPLETED | Verified at 375px, no sideways scroll |
| U-02 | Separate tabs, not one long page | High | COMPLETED | Requested explicitly |
| U-03 | Light theme only | High | COMPLETED | DECISION-008 |
| U-04 | Do not list everyone by default | High | COMPLETED | Search-and-pick with "See everyone" |
| U-05 | A distinctive, non-templated design | Medium | COMPLETED | DECISION-007 |
| U-06 | Every date names its weekday | Low | COMPLETED | Requested 2026-09-05 |
| U-07 | Readable when printed | Low | [NEEDS VERIFICATION] | Print CSS written, never tested |

## Security and privacy

| ID | Requirement | Priority | Status | Notes |
|---|---|---|---|---|
| S-01 | Keep the page out of search engines | High | COMPLETED | `robots.txt` + `noindex` |
| S-02 | Page must not be publicly discoverable | High | REJECTED | Impossible on free Pages; DECISION-003 |
| S-03 | Do not expose committee-internal notes to staff | Medium | COMPLETED | Coverage banner and build notes removed |

## Deployment

| ID | Requirement | Priority | Status | Notes |
|---|---|---|---|---|
| P-01 | One action to publish a schedule change | High | COMPLETED | `sync.cmd` |
| P-02 | A failed build must publish nothing | High | COMPLETED | `sync.cmd` stops on error |
| P-03 | The link must never change | High | COMPLETED | Stable since the org move |
| P-04 | Publish automatically on a timer | Low | REJECTED | Offered twice; user prefers manual control |

## Documentation

| ID | Requirement | Priority | Status | Notes |
|---|---|---|---|---|
| N-01 | Instructions the committee can follow alone | High | COMPLETED | `README.md` |
| N-02 | VBA kept in version control | Medium | COMPLETED | `vba/*.bas` |
| N-03 | Project memory for resuming after a gap | High | COMPLETED | This system, 2026-09-12 |

## Open questions

| ID | Question | Status |
|---|---|---|
| Q-01 | Should the site keep working after 21 September, or be retired? | [TBD — USER DECISION REQUIRED] |
| Q-02 | Should the repo move to an HTU-owned GitHub org? | [TBD — USER DECISION REQUIRED] |
| Q-03 | Do the three missing coordinator addresses exist anywhere? | [TBD — USER DECISION REQUIRED] |
| Q-04 | Are any of the 2026-09-05 usability suggestions wanted? | [TBD — USER DECISION REQUIRED] |
