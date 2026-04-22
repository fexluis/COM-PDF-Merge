---
trigger: always_on
---

# Test Conventions

## Frameworks & Tools
- **Test Framework**: `xUnit`.
- **Mocking**: `NSubstitute`.
- **Assertions**: `FluentAssertions` (Version 7+).

## Naming & Organization
- **Test Project**: `[ProjectUnderTest].Tests.(Unit|Integration)` (e.g., `ConversorPDF.Application.Tests.Unit`)
- **Test File**: `{SystemUnderTest}Tests.cs` (e.g., `ViaticoMapperTests.cs`)
- **Test Method**: `{TestAction}_When_{Condition(s)}_Should_{ExpectedBehavior}`

## Variable Naming
- **System Under Test**: `_sut` (field).
- **Result**: `result`, `actual`.
- **Inputs**: Descriptive names like `dto`, `expectedEntity`.

## Anatomy of a Test (AAA)
Divide tests into 3 distinct blocks:
1.  **Arrange**: Setup mocks, create DTOs/entities.
2.  **Act**: Single line calling the method.
3.  **Assert**: Verify the result.

## Mocking & Stubbing Guidelines
- **Assert the subject** (SUT), not the internals of collaborators.
- **`Arg.Is<T>(predicate)`**: Preferred to constrain inputs.
- **`Arg.Any<T>()`**: Use when value is irrelevant.
- **`Arg.Do`**: **AVOID**. Do not capture parameters for later assertion.

## Assertions
- **Object Comparison**: **ALWAYS** use `BeEquivalentTo` for DTOs and Entities. Do not assert fields individually.

```csharp
[Fact]
public void Map_FromGridDto_WhenValidData_ShouldReturnViaticoEntity()
{
    // Arrange
    var dto = new ViaticoGridDto { Nombre = "Juan", Monto = 100 };
    
    // Act
    var result = _sut.Map(dto);

    // Assert
    result.Should().BeEquivalentTo(new Viatico { Nombre = "Juan", Monto = 100 });
}