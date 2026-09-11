# Testing and Verification

The point of this file: **"I wrote it" and "I checked it" are different claims.**
Say which one you are making.

---

## Unit / component testing

Exercise the thing you changed, at the smallest level that proves it works.
Where the project has no test framework, this means running the code and
inspecting the result — a build that completes, a function that returns the
expected value, a page that renders.

## Integration testing

Check the components that consume what you changed. A change to a shared data
shape, a build output, or a contract between parts is not verified until every
consumer has been exercised.

## Regression testing

Confirm that important behaviour which was working still works. Prioritise
whatever the user relies on most, and anything that has broken before — past
breakages are the best predictor of future ones.

## Edge cases

Consider, and test where plausible:

- empty input
- invalid or malformed input
- missing or absent data
- boundary values
- duplicate entries
- unexpected or partial states
- failure of an external dependency

Data that arrives from a human — a spreadsheet cell, a form, a filename — should
be assumed malformed until proven otherwise. One bad cell must never take down
a whole system.

## Verification status

Classify every meaningful change as one of:

| Status | Meaning |
|---|---|
| **PASS** | Exercised and behaves correctly |
| **FAIL** | Exercised and does not |
| **PARTIALLY TESTED** | Some paths exercised, others not — say which |
| **NOT TESTED** | Written but not exercised — say why |
| **NOT APPLICABLE** | Nothing meaningful to test |

Record the untested parts in `PROJECT_STATUS.md` under **Verification Status**,
with `[NEEDS VERIFICATION]`. An acknowledged gap is manageable; a silent one is
a future incident.

## Things that are not verification

- The code compiles.
- The diff looks correct.
- It resembles something that worked before.
- No error appeared in the part you happened to look at.

## When you cannot test

Some things should not be exercised: anything that sends real messages, charges
real money, writes to live data, or is visible to other people.

In those cases:

1. Test everything up to the irreversible step — build the message, do not send
   it; compute the write, do not commit it.
2. Substitute safe values where possible (a throwaway address, a copy of the
   file) and say that you did.
3. Record the untested step explicitly. Do not let it pass unmentioned.

## Reporting

State plainly:

- what you verified, and how
- what you did not verify, and why
- anything you found that you were not looking for

If a check failed and you fixed it, say that too. A verification pass that
found and fixed a problem is more trustworthy than one that reports everything
was fine first time.
