# Rule: SQLite — Excel-DNA Shadow Copy & COM Gotchas (.NET Framework 4.8)

## Context
Running `System.Data.SQLite` inside an Excel-DNA Add-in on .NET Framework 4.8 involves a unique combination of COM interop, AppDomain isolation, and CLR assembly shadow-copying that breaks standard singleton and path-resolution patterns.

This rule documents the root causes, their mitigations, and the guardrails for future maintainers.

---

## Known Limitation 1: The "path1 null" ArgumentNullException

### Root Cause — Shadow Copy
Excel-DNA copies DLL assemblies to a CLR-managed temporary path before loading them:

```
%TEMP%\{guid}\assembly\dl3\{hash}\System.Data.SQLite.dll
                                   └─ System.Data.SQLite.dll.config  ← MUST EXIST HERE
```

`SQLiteConnection..ctor` internally calls `Path.Combine(Assembly.Location, ".config")`. 
When the shadow-copy directory is brand new (e.g., after Excel restart), that `.config` file does not yet exist. The CLR's config-reader path resolution returns `null`, causing an `ArgumentNullException("path1")`.

### Mitigation
1. **Config Dummy**: `EncryptedSQLiteContext.EnsureSQLiteConfigurationFile()` dynamically resolves the shadow-copy path via `typeof(SQLiteConnection).Assembly.Location` and creates a dummy `.config` file.
2. **Retry with 50ms delay**: `CreateConnectionWithRetry` retries the constructor once. The delay allows the CLR to update its internal configuration path cache.

> [!IMPORTANT]
> **What if the retry fails?** If the retry also fails after 2 attempts, the exception is propagated. This indicates a fatal system-level IO or security failure (e.g., antivirus blocking dynamic file creation). The user must restart Excel.

---

## Known Limitation 2: Multi-AppDomain Redundant Initialization

### Root Cause — Multiple Load Contexts
Excel-DNA loads the Add-in assembly multiple times (Ribbon, UDFs, COM). Each AppDomain gets its own copy of static fields, so standard `static bool` flags are not shared.

### Mitigation — Process-Level Idempotency
`SQLiteEngineBootstrap.EnsureInitialized()` uses process-wide mechanisms:
1. **`Process.GetCurrentProcess().Modules`**: Detects if `sqlite.interop.dll` is already mapped in `excel.exe`.
2. **Named Mutex** (`ConversorPDF_SQLiteInit_{User}_{Machine}`): Serializes first-time initialization across all concurrent AppDomains.

> [!CAUTION]
> **Timeout Behavior**: If the mutex times out (15s), the system MUST verify via `Process.GetCurrentProcess().Modules` that `SQLite.Interop.dll` is actually loaded before proceeding. Blind continuation risks inconsistency. If it is NOT loaded after 15s, an `InvalidOperationException` is thrown (Fail-Fast).

---

## Known Limitation 3: EntryPointNotFoundException (Native Mismatch)

### Root Cause
Occurs when `System.Data.SQLite.dll` (Managed) cannot find a specific function (like `sqlite3_initialize`) in `SQLite.Interop.dll` (Native). 
- **Cause A**: Version mismatch (Interop version < Managed version).
- **Cause B**: Architecture mismatch (trying to load x86 DLL in x64 process).

### Mitigation
- **NativeLibraryPreloader**: We use a custom Win32 `LoadLibrary` wrapper to force-load the architecture-specific `SQLite.Interop.dll` from the correct path before any managed SQLite interaction.
- **Unified Dependencies**: All projects use the same `Stub.System.Data.SQLite.Core.NetFramework` version (1.0.119.0).

---

## Diagnostic Cheat Sheet

| Symptom | Probable Cause | First Check |
|---------|---------------|--------------|
| `ArgumentNullException("path1")` | Shadow copy without `.config` | Check if `.config` exists in `%TEMP%` path. |
| `EntryPointNotFoundException` | Version/Architecture mismatch | Verify `SQLite.Interop.dll` is in `x64` folder. |
| `DllNotFoundException` | Native DLL not found in search path | Ensure `NativeLibraryPreloader` ran successfully. |
| `TimeoutException` (30s) | Resource contention / Deadlock | Check for orphaned Excel processes in Task Manager. |

---

## Technical Guardrail: AppDomain Isolation
> [!WARNING]
> **SemaphoreSlim Limitation**: `EncryptedSQLiteContext._lock` is a `static SemaphoreSlim`. In an Excel-DNA environment, this lock is **AppDomain-specific**. It does NOT prevent different Excel AppDomains from accessing the same database concurrently. Shared database access must be managed at the process level or avoided via the **Shared Connection** pattern.

---

## Long-Term Migration Note
If the project migrates to **.NET 6+**, evaluate `Microsoft.Data.Sqlite` + `SQLitePCLRaw`. For the current **.NET Framework 4.8 + Excel-DNA** stack, estas mitigaciones son definitivas.
