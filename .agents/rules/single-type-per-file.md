---
trigger: always_on
---

# Rule: Single Type Per File

## Definition
Every top-level type must have its own physical file to ensure Git history cleanliness.

## Constraints
* **Type Isolation:** Each `class`, `interface`, `record`, and `enum` MUST be in its own `.cs` file.
* **Filename Parity:** Filename must match the type exactly (e.g., `ISettingsRepository.cs`).
* **No Grouping:** No "Models.cs" or "Interfaces.cs" files.
* **Namespaces**: Use standard block namespaces `namespace X { ... }` (File-scoped namespaces are not idiomatic for .NET Framework 4.8).

## Agent Instructions
If existing code has multiple types in one block, propose a refactor that splits them into individual files.