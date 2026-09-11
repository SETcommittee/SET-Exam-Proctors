# Project Rules

Ten rules. They apply to any project using this system.

---

## Rule 1 — Source of Truth

The actual project files are authoritative for how things are implemented.
Memory documents describe the project; they are not the project. Where a
document and the files disagree, investigate rather than picking a side.

## Rule 2 — Minimal Context

Do not read the whole project unless the task requires it. Load `CLAUDE.md`,
the status, the TODO and recent history first, then only the files that matter.

## Rule 3 — Relevant Files First

Identify which files implement the behaviour in question before opening
anything else. Use the architecture document's file map. Reading unrelated
files is not thoroughness; it is noise.

## Rule 4 — Minimal Changes

Make the smallest change that correctly solves the task. Do not reformat, do
not rename, do not "improve" code you were not asked to touch.

## Rule 5 — Preserve Existing Behaviour

Anything working today must keep working tomorrow unless a change was
requested. Before altering a shared component, establish who depends on it.

## Rule 6 — No Guessing

If something cannot be verified, mark it `[TBD]` or go and find out. Never
write a plausible-sounding fact into a memory document. An honest gap is more
useful than a confident invention.

## Rule 7 — Track Decisions

A decision that will still constrain the project in a month belongs in
`DECISIONS.md`, with its reasoning and the alternatives weighed. When one is
overturned, mark the old one SUPERSEDED and write a new one. Never delete.

## Rule 8 — Track Progress

Meaningful progress belongs in `PROJECT_STATUS.md`. It must describe the
project as it is now, not as it was planned. Keep it short enough to stay true.

## Rule 9 — Track History

Meaningful changes belong in `CHANGELOG.md`: what was done, what was touched,
what was verified, what remains. Never paste source code into it. The test is
whether a future session could understand the change without the conversation.

## Rule 10 — Verify

Never report work as complete on the strength of the code looking correct. Run
it, load it, check it. Distinguish explicitly between "implemented" and
"implemented and verified", and say plainly what is still untested.

---

## Two failure modes to watch for

**Documentation drift.** A status document that says the project is somewhere
it is not is worse than no document, because it is trusted. If you notice drift
while working, correct it then.

**Memory sprawl.** These files are an index and a state record, not a copy of
the project. If a document is growing toward the size of the thing it
describes, it has stopped doing its job.
