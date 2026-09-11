---
name: project-maintenance
description: Maintain and extend an existing project without reloading all of it. Use when modifying functionality, adding features, fixing bugs, refactoring, changing architecture, reviewing an implementation, debugging, testing, or updating project documentation.
---

# Project Maintenance

## Purpose

Work on an existing project while loading as little of it as possible, and
leave the project memory accurate enough that the next session can pick up
without the previous conversation.

## When to Use

- modifying existing functionality
- adding a feature
- fixing a bug
- refactoring
- changing architecture
- reviewing an implementation
- debugging
- testing
- updating project documentation

## Context Strategy

Load progressively. Stop as soon as you have enough.

1. `CLAUDE.md` — permanent rules.
2. `PROJECT/PROJECT_STATUS.md` — where the project actually is.
3. `PROJECT/TODO.md` — what is outstanding.
4. The most recent `CHANGELOG/CHANGELOG.md` entries — what just happened.
5. Only the relevant requirements and decisions.
6. Only the files the current task needs — find them in the **Relevant File
   Map** in `PROJECT/ARCHITECTURE.md`.

Do not load the whole project unless the task genuinely spans it.

Two signals that you have loaded too much: you are reading files you cannot
connect to the request, or you are re-deriving something the status document
already states.

## Before Editing

Answer these before changing anything:

1. What does the user actually want?
2. Which requirement does it relate to — and is there one?
3. Which component implements that behaviour?
4. Which files will the change touch?
5. Does a recorded decision constrain the approach?
6. What else depends on what you are about to change?

If a file is generated, edit its source instead. If a decision forbids the
obvious approach, say so rather than quietly working around it.

## After Editing

1. Verify the change actually works — not that the code looks right.
2. Look for side effects on anything that shares the component.
3. Update the documentation that is now wrong.
4. Update `PROJECT_STATUS.md` if the state moved.
5. Add a `CHANGELOG` entry if the work was meaningful.
6. Update `TODO.md` if tasks opened or closed.

State clearly whether you implemented, or implemented **and verified**.

## Scope Discipline

Match the depth of inspection to the risk of the change.

| Change | Inspection |
|---|---|
| Copy, styling, a single value | The one file. No audit. |
| A feature inside one component | That component and its immediate callers. |
| Shared component, data shape, build | Every consumer, plus decisions and requirements. |
| Architecture, security, data integrity | Full review, and confirm the approach first. |

A small task does not earn a project audit. A change to a shared contract does.

## Supporting Files

- `project-rules.md` — the non-negotiable rules.
- `workflow.md` — the step order, and when to widen scope.
- `testing.md` — what counts as verified.
