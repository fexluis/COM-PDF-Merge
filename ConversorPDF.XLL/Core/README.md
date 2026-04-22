# Core Architecture

Este directorio contiene la lógica principal de la aplicación y la definición de dominios, separando responsabilidades (Separation of Concerns). Todo el código está amparado por los principios estandarizados y convenciones del proyecto.

## Estructura de Directorios

La estructura base del módulo `Core` agrupa código por rol funcional en lugar de características (feature-folders), promoviendo que los distintos tipos de responsabilidades no se mezclen.

```mermaid
graph TD
    Core[ConversorPDF.XLL/Core]
    Core --> Enums[Enums/]
    Core --> Interfaces[Interfaces/]
    Core --> Models[Models/]
    Core --> Processors[Processors/]
    Core --> Repositories[Repositories/]
    Core --> Services[Services/]
    Core --> Helpers[Helpers/]

    Enums -.-> |Enumeraciones y Constantes tipadas| Enums
    Interfaces -.-> |Todos los contratos e interfaces del sistema| Interfaces
    Models -.-> |Objetos de transferencia, entidades y configuraciones| Models
    Processors -.-> |Lógica técnica particular de transformación| Processors
    Repositories -.-> |Acceso a datos base y recuperación de información| Repositories
    Services -.-> |Orquestación de alto nivel de negocio| Services
    Helpers -.-> |Utilidades estáticas compartidas (*Helper)| Helpers
```

## Convenciones de Ingeniería y Código

Para garantizar la mantenibilidad y calidad en este módulo, es OBLIGATORIO seguir las siguientes reglas al escribir código dentro de `/Core`:

### 1. Integridad Estructural (Single Type Per File)
- **Regla Estricta**: Cada `class`, `interface`, `record` o `enum` **DEBE** ir en su propio archivo independiente. 
- **Namespaces Locales**: `ConversorPDF.XLL.Core.[NombreCarpeta]` (con modelo *file-scoped* preferiblemente: `namespace ConversorPDF.XLL.Core.Services;`).
- Los modificadores de acceso orgánicos deben tratar de ser `internal` para acoplar sólo lo necesario y exponer únicamente en interfaces.

### 2. Control de Flujo y Errores
- **Result Pattern**: Procurar retornar un objeto `Result<T>` en vez de usar excepciones para control de flujo de la lógica de negocio (como validación fallida o entidad no encontrada).
- **Excepciones**: No usar bloques `try-catch` para validar lógica; dejarlos únicamente para fallas no controladas de infraestructura.
- **Guard Clauses**: Evitar la palabra clave `else`. Retornar tempranamente (early-return) si no se cumplen las condiciones `if (!valid) { return error; }`.

### 3. Operaciones Asíncronas
- **CancellationToken**: Es **obligatorio** agregarlo como el último parámetro en todo método `async`.
- **Prohibición**: Nunca usar `Task.Wait()` ni `Task.Result`. Esto causa *thread starvation*. Todo I/O debe ser `await`.

### 4. Colecciones y Estructuras de Datos
- **Retornos y Parámetros**: Usar tipos inmutables como `IReadOnlyCollection<T>`. 
- **Nulabilidad**: Las colecciones **no** deben ser nulas. Devolver una colección vacía (`[]`) en lugar de `null`.
- **Variables**: Utilizar `var` siempre que el tipo esté implícito.

### 5. Constructores e Inyección de Dependencias
- No comprobar explícitamente nulos (`ArgumentNullException`) en los parámetros del constructor asignados por Inyección de Dependencias.
- No utilizar Primary Constructors (ej. `class Service(IDep dep)`) en clases de la capa `Services` e inyectables; conservar la inicialización clásica para proteger la mutabilidad de los contenedores y el scope local.

### 6. Tests y Calidad (Convención AAA)
Cualquier servicio, procesador o modelo agregado al `Core` debe poderse aislar en una prueba unitaria (en el proyecto `.Tests`) utilizando *xUnit*, *FluentAssertions* y *NSubstitute*, bajo los principios:
1. **Arrange** Setup de datos de prueba
2. **Act** Ejecución (usualmente 1 sola línea)
3. **Assert** Validación de estado mediante `.Should().BeEquivalentTo(...)` u outputs de comportamiento.
