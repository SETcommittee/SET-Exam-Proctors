# Workflow

```
USER REQUEST
      ↓
UNDERSTAND REQUEST
      ↓
LOAD PROJECT CONTEXT          CLAUDE.md → STATUS → TODO → recent CHANGELOG
      ↓
IDENTIFY REQUIREMENT          which requirement, or is this new?
      ↓
IDENTIFY RELEVANT COMPONENTS  use the file map in ARCHITECTURE.md
      ↓
INSPECT ONLY RELEVANT FILES
      ↓
CHECK DECISIONS / DEPENDENCIES
      ↓
PLAN CHANGE
      ↓
IMPLEMENT
      ↓
VERIFY                        see testing.md
      ↓
DOCUMENT
      ↓
UPDATE STATUS
      ↓
REPORT RESULT
```

---

## Scaling the workflow

**Do not run a full project audit for a small task.** Trimming a label, fixing
a typo, adjusting a value: open the file, change it, check it, done. Skip the
requirement lookup and the documentation pass — there is nothing useful to
record.

**Widen the inspection when the change touches:**

- architecture or a shared contract
- a dependency other components rely on
- security, privacy, or who can see what
- data integrity or a data shape
- more than one major component
- anything published to real users

In those cases: read the relevant decisions first, list every consumer of what
you are changing, and confirm the approach before writing code.

## Step notes

**Understand request.** If two readings would lead to materially different
work, ask. If one reading is clearly more likely, say which you took and
proceed.

**Load context.** Stop as soon as you can name the files you need. Loading more
"to be safe" is the failure this system exists to prevent.

**Identify requirement.** If the request matches nothing on record, it is new
work — say so, and consider whether it belongs in `REQUIREMENTS.md`.

**Check decisions.** If a recorded decision rules out the obvious approach, do
not quietly route around it. Either follow it, or raise that it should be
superseded.

**Plan.** For anything beyond a small change, say what you intend to do before
doing it. Cheaper to correct a plan than an implementation.

**Verify.** Not optional. See `testing.md`.

**Document.** Update only what is now untrue. Do not restate unchanged
information in three places.

**Report.** Say what changed, what was verified, and what was not. If something
was left undone, say so explicitly rather than letting it pass unmentioned.

## Session end

When a meaningful session ends, leave the memory correct:

1. `PROJECT_STATUS.md` — current state.
2. `TODO.md` — tasks opened or closed.
3. `CHANGELOG.md` — a concise entry.
4. `ARCHITECTURE.md` — only if the architecture actually moved.
5. `DECISIONS.md` — only if a lasting decision was taken.

The test: could a fresh session, with no access to this conversation,
understand where the project stands and what to do next?
