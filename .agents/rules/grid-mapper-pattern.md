# Rule: Grid Mapper Pattern for UI Decoupling

## Definition
To ensure 100% unit test coverage and high data integrity, all business logic associated with DataGridView data extraction must be decoupled from WinForms controls using the **Mapper Pattern**.

## Constraints
* **Zero UI Dependency**: Extraction logic MUST NOT accept `DataGridView`, `DataGridViewRow`, or any WinForms control as a parameter.
* **Abstract DTOs**: Use a lightweight DTO (e.g., `GridRowDto`) to represent the visual state of the grid before passing it to the logic layer.
* **Static Mapper**: Create a static class named `{Feature}To{Entity}Mapper` to perform the transformation from DTOs to Domain Models.
* **Integrity Guards**: Mappers MUST verify that the number of successfully extracted entities matches the intended "visual count" (rows checked/marked by the user). Throw `InvalidOperationException` on mismatch.
* **Mandatory Tests**: Every new Mapper class MUST have 100% unit test coverage in the Test project.

## Implementation Pattern

### 1. The DTO
Create a simple DTO to bridge the WinForms row to the logic.
```csharp
public class GridRowDto
{
    public bool IsChecked { get; set; }
    public object Tag { get; set; }
    // Other primitive properties (string, int, etc.)
}
```

### 2. The Mapper
A pure logic class that can be unit tested in seconds.
```csharp
public static class MyMapper
{
    public static List<MyEntity> Extract(IEnumerable<GridRowDto> rows)
    {
        var results = new List<MyEntity>();
        foreach (var row in rows.Where(r => r.IsChecked))
        {
            // ... logic and integrity checks ...
            results.Add(entity);
        }

        // Integrity Guard
        if (results.Count != rows.Count(r => r.IsChecked))
        {
            throw new InvalidOperationException("Falló la integridad de extracción.");
        }

        return results;
    }
}
```

### 3. The UI Call
The WinForms layer is responsible ONLY for flushing the grid and building the DTO list.
```csharp
// Inside the Form
dgv.CurrentCell = null; // Flush current cell
dgv.CommitEdit(DataGridViewDataErrorContexts.Commit);

var dtos = dgv.Rows
    .Cast<DataGridViewRow>()
    .Select(r => new GridRowDto { ... })
    .ToList();

var result = MyMapper.Extract(dtos);
```

## Agent Instructions
1. **Refactoring**: If you encounter legacy code extracting data directly from a `DataGridView` in a button click or private method, propose a refactor to this pattern.
2. **Validation**: Every time a new grid feature is added, ensure a corresponding Mapper and Unit Test file are created.
