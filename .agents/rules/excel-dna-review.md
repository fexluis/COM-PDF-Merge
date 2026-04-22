---
trigger: always_on
---

# Code Review Checklist: Excel-DNA & Desktop Architecture

Before approving any PR or code generation, verify these points:

## 1. Architecture & Boundaries
- [ ] Are there any Office Interop types (`Range`, `Worksheet`) in `ConversorPDF.Application`? (Strictly Forbidden).
- [ ] Are there any WinForms types (`DataGridView`, `Form`) in `ConversorPDF.Application`? (Strictly Forbidden).
- [ ] Did they put SQLite/Dapper logic in `Application`? (Should be in `XLL/Infrastructure`).

## 2. Testability & DI
- [ ] Does the new Service/Processor depend on concretions instead of interfaces?
- [ ] Are Mappers pure static/instance logic without touching UI controls?
- [ ] Is `RibbonController` the ONLY place using `ServiceProvider.GetRequiredService` directly? (Service Locator anti-pattern is only allowed here).

## 3. Excel-DNA & COM Specifics
- [ ] **STA Thread Safety**: If async/await was used, is it properly marshaling back to the Excel UI thread? (Excel hates cross-thread COM calls).
- [ ] **COM Object Cleanup**: Are `Range`, `Worksheet`, and `Application` COM objects being released properly? (Missing releases cause "ghost" Excel.exe processes).
- [ ] **Excel Optimization**: Are heavy Excel read/write loops wrapped in `ExcelOptimizeHelper.Optimize()`?

## 4. SQLite & Shadow Copy
- [ ] Did they bypass the `EncryptedSQLiteContext` and open a raw `SQLiteConnection`? (Risk of shadow-copy `.config` crash).
- [ ] Is the `EncryptedSQLiteContext` lifecycle correct? (Singleton if stateless factory, Transient if holding decryption keys per form).

## 5. Desktop Code Style
- [ ] Did they use C# 12+ syntax like `[]` for collections? (Fails on .NET 4.8).
- [ ] Are exceptions being used for business rule validation instead of `Result` pattern?