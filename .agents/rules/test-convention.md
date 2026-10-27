---
trigger: always_on
---

# Test Conventions

> [!NOTE]
> Esta regla refleja el estado real y verificado del repositorio (.NET Framework 4.8); no migrar frameworks de testing ni bibliotecas de mocking dentro de tareas de feature o refactor.

## Frameworks & Tools
- **Test Framework**: `xUnit 2.7` (principal). `MSTest 3.3.1` únicamente para tests de UI en hilo STA (carpeta `UI/`).
- **Mocking**: `Moq 4.20.70` (NSubstitute NO está instalado; no introducirlo sin un ADR formal).
- **Assertions**: `FluentAssertions 6.12.0` (la sintaxis de FluentAssertions v7+ no está disponible en este proyecto).

## Naming & Organization
- **Test Project**: Proyecto único `ConversorPDF.Tests` (net48, SDK-style) con carpetas espejo del XLL (`Academic/Application`, `Academic/Domain`, `Core`, `Integration`, `UI`, `Unit`).
- **Test File**: `{SystemUnderTest}Tests.cs` (ej. `ViaticoMapperTests.cs`).
- **Test Method**: `{TestAction}_When_{Condition(s)}_Should_{ExpectedBehavior}`.

## Variable Naming
- **System Under Test**: `_sut` (campo privado de solo lectura).
- **Result**: `result`, `actual`.
- **Inputs**: Nombres descriptivos como `dto`, `expectedEntity`.

## Anatomy of a Test (AAA)
Divide tests into 3 distinct blocks:
1. **Arrange**: Configurar mocks (`Mock<T>`), preparar datos de prueba o DTOs.
2. **Act**: Llamada concisa al método bajo prueba.
3. **Assert**: Verificar el resultado.

## Mocking Guidelines (Moq)
- **Assert the subject** (SUT), no los detalles internos triviales de los colaboradores.
- Usar `It.Is<T>(predicate)` para restringir argumentos.
- Usar `It.IsAny<T>()` cuando el valor exacto del parámetro sea irrelevante para el caso de prueba.
- Verificar invocaciones con `_mock.Verify(x => x.Metodo(...), Times.Once)` solo cuando la interacción sea el contrato esencial.

## Assertions (FluentAssertions 6.12)
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
```
