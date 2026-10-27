# Unit Tests & Development Skill (Desktop / Excel-DNA)

## Core Principles

- **No code without tests:** Every new function, class, or logic change in the "Clean World" (`Domain` or `Application`) must be accompanied by unit tests.
- **Atomic Iterations:** Break the task into the smallest possible sub-tasks. Run the full Code → Test → Fix cycle on **each** sub-task before moving on.
- **Zero Tolerance for Red:** Do not advance to the next sub-task while any test (new or existing) is failing.
- **Clean World Isolation:** Unit tests must run entirely in memory. If a test requires Excel to be open, a real SQLite database, or WinForms initialization, it is an Integration Test or the architecture boundaries are being violated.

---

## Project Structure Context

The solution follows a **Modular Monolith** with **Vertical Slices** for self-contained business modules.

```
ConversorPDF.XLL/
├── Academic/                        ← Vertical Slice (fully self-contained)
│   ├── Domain/
│   │   ├── Models/                  (Sesion, Vinculacion, etc.)
│   │   └── Interfaces/              (IAcademicRepository, IAcademicDateSyncService)
│   ├── Application/
│   │   ├── Interfaces/
│   │   ├── Services/                (AcademicDateSyncService, GeneradorSesionesRecurrentes)
│   │   ├── Mappers/                 (GridToSesionMapper)
│   │   └── DTOs/                    (GridRowDto)
│   ├── Infrastructure/
│   │   ├── Persistence/             (AcademicRepository, AcademicUnitOfWork)
│   │   ├── Context/                 (EncryptedSQLiteContext, SQLiteEngineBootstrap)
│   │   └── Sql/                     (AcademicSchema.sql, AcademicSeedData.sql)
│   └── UI/
│       └── Forms/                   (FormSesionesAcademicas)
├── Core/                            ← Legacy shared logic (stable, not being migrated)
│   ├── Models/                      (Viaticos, Terceros, etc.)
│   ├── Processors/
│   ├── Repositories/
│   └── Helpers/                     (ExcelOptimizeHelper, LoggerHelper)
├── UI/                              ← Shared WinForms controls
│   └── (ProgressDialog, FormInfo, DayHoursDropdown, etc.)
└── Ribbon/
    ├── Ribbon.xml
    └── RibbonController.cs          ← Single entry point for all slices
```

Tests must mirror the slice structure:

```
ConversorPDF.Tests/
├── Academic/
│   ├── Application/
│   └── Domain/
├── Core/
├── Integration/
├── UI/
└── Unit/
```

---

## Development Loop (repeat for every sub-task)

### 1. Analyze & Plan
- Identify the exact piece of logic to implement.
- Verify the target layer: Is this a Domain entity? An Application Service/Processor/Mapper?
- If modifying a **Vertical Slice** (e.g., `Academic`), ensure the change stays inside the slice boundaries.
- Write down the **expected behavior**, **edge cases**, and **error paths** *before* writing any code.
- Confirm which existing tests could be affected by this change.

### 2. Test Creation (TDD)
- Create or update unit tests in `ConversorPDF.Tests/` under the corresponding slice/folder following `test-convention.md`.
- Required coverage per sub-task:
  - ✅ Happy path(s)
  - ✅ Edge cases identified in Step 1
  - ✅ Error / failure paths (using `Result` pattern failures, not exceptions)
- The new tests **must fail initially** since the implementation logic isn't written yet.

### 3. Implementation
- Write the **minimum** code required to satisfy the failing tests — no speculative features.
- **Boundary Check**: If you are writing code in `Application`, ensure you are injecting interfaces (e.g., `IPdfGenerator`) and not concrete infrastructure classes (e.g., `PdfSharpGenerator`).
- **Vertical Slice Check**: If adding to `Academic`, place files in `Academic\{Layer}\{Folder}`. Do not create new shared folders in `Core`.
- Include XML comments (`///`) on all new public classes, properties, and methods to automatically generate API documentation with DocFX.
- Follow conventions in:
  - `desktop-code.md`
  - `architecture-rules.md`
  - `excel-dna-review.md`

### 4. Execution & Verification
- Ejecutar `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~<Clase>"` durante el ciclo de implementación.
- Ejecutar `-Mode Full` al cerrar cada sub-tarea. La evidencia generada (`test.trx` + `summary.md`) se vincula directamente al plan.
- Visual Studio Test Explorer queda disponible únicamente como verificación humana opcional.
- No se toleran regresiones respecto al baseline (`tests/baseline/test_baseline.txt`).

### 5. Self-Healing Loop

| Outcome | Decision rule | Action |
|---------|---------------|--------|
| Test fails & the assertion is wrong | The test was written incorrectly | Fix the test, document why, re-run Step 4 |
| Test fails & the code is wrong | Implementation bug | Fix the code, re-run Step 4 |
| Build/Warnings fail | Formatting/convention violation | Fix the warning, re-run Step 4 |
| Test & requirements are ambiguous | Spec gap | Clarify requirements, update plan, re-run from Step 1 |
| All checks pass | ✅ | Proceed to next sub-task |

> **Rule:** Never fix a failing test by weakening the assertion (e.g., removing edge-case coverage) unless the requirement itself changed.

---

## Finalization

Once **all** sub-tasks are complete:

1. **Final regression run** — execute `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full` one last time and confirm 0 regressions.
2. **Write a change summary document** at `docs/articles/changes/<YYYY-MM-DD>-<short-description>.md` with the following sections:
   - **Purpose** — what problem this solves.
   - **Changes** — which files were modified in `Domain` or `Application` (and corresponding `XLL` implementations if interfaces changed).
   - **How to use** — minimal usage example or steps to verify the new behavior via the Excel Ribbon.
   - **Test coverage** — summary of tests added or updated.
3. **No open TODOs** — resolve or create tracked issues for anything left unfinished.

---

## Quick Reference Checklist

- [ ] Sub-task is clearly scoped and targets `Domain` or `Application`
- [ ] If modifying a Vertical Slice, changes remain inside the slice boundary
- [ ] Expected behavior & edge cases documented before coding
- [ ] Code follows Desktop and Architecture Boundary rules
- [ ] XML comments (`///`) added or updated for public members
- [ ] Unit tests cover happy path, edge cases, and error paths
- [ ] Tests run successfully without Excel or SQLite running
- [ ] `scripts/verify.ps1 -Mode Full` passes with 0 regressions vs baseline
- [ ] No regression in existing tests
- [ ] Change summary document written
