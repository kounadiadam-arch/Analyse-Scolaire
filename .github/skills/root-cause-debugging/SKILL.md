---
name: root-cause-debugging
description: 'Debug failing code, reproduce bugs, identify root cause, apply the smallest fix, and verify results with fresh evidence. Use for runtime errors, broken tests, unexpected behavior, build failures, or regression checks.'
---

# Root-Cause Debugging

## When to Use
- A bug, test failure, or runtime error appears
- Behavior is unexpected or inconsistent
- A regression must be diagnosed and fixed
- A change should be validated before completion

## Procedure

1. Reproduce the issue
   - Run the smallest command or test that shows the problem
   - Capture the exact error, stack trace, failing assertion, or wrong output
   - Confirm the problem is real before changing code

2. Gather evidence from the actual failure
   - Read the full message and identify the failing layer: app code, config, dependency, or environment
   - Check the input state, recent changes, and relevant logs or traces
   - Prefer fresh evidence over assumptions

3. Isolate the root cause
   - Trace the data flow to the point where the wrong value or condition appears
   - Compare against a known-good pattern or expected behavior
   - Narrow the issue to the smallest possible component or function

4. Form one focused hypothesis
   - Decide on the single most likely cause
   - Avoid stacking unrelated fixes
   - If the cause is unclear, gather one missing fact instead of changing code blindly

5. Add or update a failing proof
   - Prefer a regression test or minimal repro that demonstrates the bug
   - Make the test fail before the fix if possible

6. Apply the smallest fix
   - Change only the code tied to the root cause
   - Keep the fix scoped and understandable
   - Do not broaden the change to unrelated refactors

7. Verify with fresh output
   - Run the relevant test(s), command(s), or check(s)
   - Confirm the original issue is resolved
   - Check that nearby behavior still passes

8. Close with evidence
   - Report what failed, what caused it, what changed, and how it was verified
   - State the exact validation command and the result

## Decision Points

- If the issue cannot be reproduced, inspect recent edits, environment differences, and configuration before changing code.
- If multiple causes are possible, validate the most likely cause with the smallest probe or trace.
- If the fix affects user-visible behavior, add a regression test or explicit validation step.
- If evidence is incomplete, stop and gather more data rather than guessing.

## Quality Bar

A task is complete only when all of the following are true:
- The root cause is identified, not just the symptom
- The fix is minimal and directly tied to that cause
- Relevant verification has been run successfully with fresh output
- The change does not introduce obvious regressions in the touched area
- The final report includes evidence, not a claim without proof

## Example Prompts
- Diagnose this failing Django test and fix the root cause
- Investigate the runtime error in this endpoint and verify the fix
- Trace why this calculation is wrong and add a regression check
- Reproduce this bug, patch it minimally, and confirm with fresh test output
