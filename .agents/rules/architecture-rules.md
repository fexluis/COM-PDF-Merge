---
trigger: always_on
---

# Engineering Principles: Excel-DNA Modular Monolith (.NET Framework 4.8)

> [!WARNING]
> **CRITICAL CONTEXT**: This is a .NET Framework 4.8 Excel-DNA Add-In using WinForms. 
> **STRICTLY PROHIBITED**: ASP.NET Core concepts (MediatR, EF Core, JWT, Controllers, Middleware, IOptions, Swagger, C# 12+ syntax like `[]` collections).

## 1. The 3-Layer Physical Boundary
The solution is strictly divided into "Clean World" (Testable) and "Dirty World" (Infrastructure).
* **ConversorPDF.Domain**: ZERO external NuGet dependencies. Only pure C# entities, enums, exceptions, and repository interfaces.
* **ConversorPDF.Application**: Depends ONLY on Domain. Contains DTOs, Mappers, Services, Processors, and infrastructure interfaces (e.g., `IPdfGenerator`). NO references to System.Data, System.Drawing, PdfSharp, or Office Interop.
* **ConversorPDF.XLL**: The "Dirty World". Depends on Domain + Application. Contains Excel-DNA, Ribbon, WinForms, SQLite/Dapper, PDFSharp, and COM Interop.

**Enforcement Rule**: If adding a NuGet to `Application` requires an infrastructure library (e.g., MailKit, System.Data.SQLite), **STOP**. Create an interface in `Application` and implement it in `XLL`.

## 2. Logic & Error Handling
* **Result Pattern**: Use `CSharpFunctionalExtensions.Result<T>` for domain and application logic failures. Do not throw exceptions for business rule violations.
* **No Flow-Control Exceptions**: Reserve `try-catch` only for unexpected infrastructure failures (e.g., SQLite lock, COM exceptions in XLL layer). Application layer must wrap all errors into `Result<T>`.
* **NO MediatR**: Do not use CQRS with MediatR. Use standard Application Services and Processors injected via DI.

## 3. Excel Integration Guardrails
* **IExcelDataSource Shield**: The `IExcelDataSource` interface in Application must NEVER return Office Interop types (no `Range`, `Worksheet`). It must return DTOs or primitives. The `ExcelInteropService` in XLL does the dirty translation.
* **Grid Mapper Pattern**: Business logic MUST NEVER accept `DataGridView` or `DataGridViewRow`. Use DTOs and the designated Mapper pattern (see `grid-mapper-pattern.md`).

## 4. Prohibitions (Zero Tolerance)
* ❌ No `.Result` or `.Wait()` on Tasks (prevents thread pool starvation), unless absolutely required by COM STA threads (use extreme caution).
* ❌ No business logic inside WinForms code-behind or RibbonController.
* ❌ No direct use of `DateTime.Now` in Domain/Application; use `IDateTimeProvider`.
* ❌ No C# 12+ syntax (e.g., collection expressions `[]`, primary constructors for DI classes).