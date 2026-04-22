# Unit Tests & Development Skill (Desktop / Excel-DNA)

## Core Principles

- **No code without tests:** Every new function, class, or logic change in the "Clean World" (`Domain` or `Application`) must be accompanied by unit tests.
- **Atomic Iterations:** Break the task into the smallest possible sub-tasks. Run the full Code → Test → Fix cycle on **each** sub-task before moving on.
- **Zero Tolerance for Red:** Do not advance to the next sub-task while any test (new or existing) is failing.
- **Clean World Isolation:** Unit tests must run entirely in memory. If a test requires Excel to be open, a real SQLite database, or WinForms initialization, it is an Integration Test or the architecture boundaries are being violated.

---

## Development Loop (repeat for every sub-task)

### 1. Analyze & Plan
- Identify the exact piece of logic to implement.
- Verify the target layer: Is this a Domain entity? An Application Service/Processor/Mapper? 
- Write down the **expected behavior**, **edge cases**, and **error paths** *before* writing any code.
- Confirm which existing tests could be affected by this change.

### 2. Test Creation (TDD)
- Create or update unit tests in `tests/ConversorPDF.Application.Tests/` or `tests/ConversorPDF.Domain.Tests/` following `.agents/rules/test-convention.txt`.
- Required coverage per sub-task:
  - ✅ Happy path(s)
  - ✅ Edge cases identified in Step 1
  - ✅ Error / failure paths (using `Result` pattern failures, not exceptions)
- The new tests **must fail initially** since the implementation logic isn't written yet.

### 3. Implementation
- Write the **minimum** code required to satisfy the failing tests — no speculative features.
- **Boundary Check**: If you are writing code in `Application`, ensure you are injecting interfaces (e.g., `IPdfGenerator`) and not concrete infrastructure classes (e.g., `PdfSharpGenerator`).
- Include XML comments (`///`) on all new public classes, properties, and methods to automatically generate API documentation with DocFX.
- Follow conventions in:
  - `.agents/rules/desktop-code.txt`
  - `.agents/rules/architecture-rules.txt`
  - `.agents/rules/excel-dna-review.txt`

### 4. Execution & Verification
- Run the **full** test suite from the **Visual Studio Test Explorer** (Do not use CLI `dotnet test` commands for .NET Framework 4.8 Excel-DNA projects to avoid loader context issues).
- Alternatively, ensure the solution builds cleanly with **"Treat warnings as errors"** enabled in VS Build properties.
- All tests must be **green** before proceeding.

### 5. Self-Healing Loop

| Outcome | Decision rule | Action |
|---|---|---|
| Test fails & the assertion is wrong | The test was written incorrectly | Fix the test, document why, re-run Step 4 |
| Test fails & the code is wrong | Implementation bug | Fix the code, re-run Step 4 |
| Build/Warnings fail | Formatting/convention violation | Fix the warning, re-run Step 4 |
| Test & requirements are ambiguous | Spec gap | Clarify requirements, update plan, re-run from Step 1 |
| All checks pass | ✅ | Proceed to next sub-task |

> **Rule:** Never fix a failing test by weakening the assertion (e.g., removing edge-case coverage) unless the requirement itself changed.

---

## Finalization

Once **all** sub-tasks are complete:

1. **Final regression run** — execute all tests from Test Explorer one last time and confirm 100% pass rate.
2. **Write a change summary document** at `docs/articles/changes/<YYYY-MM-DD>-<short-description>.md` with the following sections:
   - **Purpose** — what problem this solves.
   - **Changes** — which files were modified in `Domain` or `Application` (and corresponding `XLL` implementations if interfaces changed).
   - **How to use** — minimal usage example or steps to verify the new behavior via the Excel Ribbon.
   - **Test coverage** — summary of tests added or updated.
3. **No open TODOs** — resolve or create tracked issues for anything left unfinished.

---

## Quick Reference Checklist

- [ ] Sub-task is clearly scoped and targets `Domain` or `Application`
- [ ] Expected behavior & edge cases documented before coding
- [ ] Code follows Desktop and Architecture Boundary rules
- [ ] XML comments (`///`) added or updated for public members
- [ ] Unit tests cover happy path, edge cases, and error paths
- [ ] Tests run successfully without Excel or SQLite running
- [ ] Visual Studio Test Explorer passes with **zero** failures
- [ ] No regression in existing tests
- [ ] Change summary document written