---
trigger: always_on
---

# Code Conventions (.NET Framework 4.8 Desktop)

> [!NOTE]
> Tests conventions are listed separately: [Tests conventions](test-convention.md)

## General Conventions 💅

### Access Modifiers
- Use **most strict** access modifiers: `private`/`internal` instead of `public`.
- Use `public` for interfaces and `internal` for their implementations.

### Excel Optimization
- **Rule**: Always use `ExcelOptimizeHelper.Optimize(app)` within a `using` block for operations that modify Excel application state.

### Constructors
- **No Primary Constructors** for classes resolved from DI. Use explicit constructors with `private readonly` fields (e.g., `private readonly IFooService _foo;`).
- **Null Checks**: Do **not** check injected services for null in the constructor. Trust the DI container.

### Async & CancellationToken (Excel-DNA Exception)
- **Infrastructure Methods** (e.g., Repositories, PDF generators): `CancellationToken` is mandatory as the last parameter. No default values.
- **Application Services / Processors**: Because Excel UI runs on a Single-Threaded Apartment (STA), do not force async/await blindly. Use synchronous execution if the Service orchestrates fast UI updates, or ensure proper `Task.Run` marshaling if heavy I/O is needed. Avoid `async void` at all costs.

## Result Pattern
- **Domain**: Methods return `Result` or `Result<T>`. Never throw for business rules.
- **Application**: Always check `if (result.IsFailure) return result;`. Return `Result<T>` to the UI layer.
- **UI (WinForms/XLL)**: Check `result.IsFailure` and show `MessageBox.Show(result.Error)`.

## Data Structures & Collections
- Use **read-only** interfaces (`IReadOnlyCollection<T>`) for returns.
- **Nullability**: Do not use nullable collections (e.g., `List<int>?`). Return empty collections. Avoid `bool?`; use Enums for tri-state.

## Configuration
- **NO IOptions**: This is a Desktop app. Inject a typed `Settings` POCO class directly via DI, or read configuration once at the Composition Root.

## Logging
- **Template Parameters**: Use `UpperCamelCase` (e.g., `{StoreId}`).
- **Null Checks**: Never check logger for null.

## Code Style
- **Braces**: Use braces `{}` for all `if` statements.
- **Else**: Avoid `else`. Use early returns (Guard Clauses).
- **Naming**: No context prefixes (e.g., `class Product { public int Id }`, not `ProductId`).
- **Chain Methods**: Place every chained method on a new line starting with a dot.