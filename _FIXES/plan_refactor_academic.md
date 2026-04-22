# Plan de Refactorización: Módulo Académico
**Proyecto:** `ConversorPDF.XLL` — Vertical Slice Modular Monolith  
**Rama sugerida:** `refactor/academic-modularization`  
**Estrategia:** Mover → Actualizar namespaces → Limpiar Core → Verificar

---

## Decisión de Arquitectura

Se adopta **`ConversorPDF.XLL\Academic\`** (Modular Monolith) sobre un proyecto físico separado.

| Criterio | `XLL\Academic\` ✅ | `ConversorPDF.Academic\` ❌ |
|---|---|---|
| Impacto en `.csproj` | Ninguno — sigue siendo un solo proyecto | Nuevo `.csproj`, referencias cruzadas |
| RibbonController | Solo cambia el `using`, no la referencia de proyecto | Requiere `<ProjectReference>` y posibles dependencias circulares |
| WinForms / Resources | `.resx` y designer files permanecen en el mismo output | Complicaciones con Forms en proyectos externos al XLL |
| Riesgo | Bajo — solo carpetas y namespaces | Alto — cambio estructural de solución |
| Reversibilidad | Fácil de promover a multiproyecto después | Difícil de revertir |

> El salto a `ConversorPDF.Academic` (proyecto físico separado) es el paso siguiente si en el futuro se quiere reutilizar el módulo en otro host. En un XLL Excel-DNA, no es necesario hoy.

---

## Estado Actual: Inventario de Archivos Académicos

Toda la lógica académica está dispersa en `ConversorPDF.XLL`, mezclada con Viáticos y utilidades globales:

| # | Ubicación Actual | Tipo | Función |
|---|---|---|---|
| 1 | `Core\Models\Academic\*.cs` | Domain | Entidades de dominio (Sesión, Programa, Tercero, etc.) |
| 2 | `Core\Interfaces\IAcademicRepository.cs` | Domain | Contrato de persistencia académica |
| 3 | `Core\Services\Academic\*.cs` | Application | Lógica de negocio (SyncService, GeneradorSesiones, etc.) |
| 4 | `Core\Helpers\GridToSesionMapper.cs` | Application | Mapper de filas del grid a entidad Sesión |
| 5 | `Core\Helpers\GridRowDto.cs` | Application | DTO de fila del grid |
| 6 | `Core\Repositories\AcademicRepository.cs` | Infrastructure | Implementación SQLite + Dapper |
| 7 | `Core\Data\EncryptedSQLiteContext.cs` | Infrastructure | Motor de Vault / encriptación de la DB académica |
| 8 | `Core\Data\SQLiteEngineBootstrap.cs` | Infrastructure | Inicialización del motor nativo SQLite |
| 9 | `Core\Data\AcademicUnitOfWork.cs` | Infrastructure | Gestión de transacciones |
| 10 | `Core\Data\AcademicSchema.sql` | Infrastructure | DDL — Definición de tablas SQLite |
| 11 | `Core\Data\AcademicSeedData.sql` | Infrastructure | Datos iniciales (si aplica) |
| 12 | `FormAcademic\FormSesionesAcademicas.cs` | UI | Formulario WinForms principal |
| 13 | `FormAcademic\FormSesionesAcademicas.Designer.cs` | UI | Código generado por el diseñador |
| 14 | `FormAcademic\FormSesionesAcademicas.resx` | UI | Recursos del formulario |

> **Punto de entrada externo (no se mueve):**  
> `Ribbon\RibbonController.cs` — solo se actualiza el `using` de una línea.

---

## Estructura Destino

```
ConversorPDF.XLL\
└── Academic\
    ├── Domain\
    │   ├── Models\             ← Entidades puras (Sesion, Programa, Tercero, etc.)
    │   └── Interfaces\         ← IAcademicRepository, IAcademicService
    ├── Application\
    │   ├── Services\           ← SyncService, GeneradorSesiones, etc.
    │   ├── Mappers\            ← GridToSesionMapper
    │   └── DTOs\               ← GridRowDto
    ├── Infrastructure\
    │   └── Persistence\
    │       ├── Repositories\   ← AcademicRepository
    │       ├── Context\        ← EncryptedSQLiteContext, SQLiteEngineBootstrap
    │       ├── AcademicUnitOfWork.cs
    │       └── Sql\            ← AcademicSchema.sql, AcademicSeedData.sql
    └── UI\
        └── Forms\              ← FormSesionesAcademicas (.cs, .Designer.cs, .resx)
```

---

## Fases de Implementación

### FASE 0 — Snapshot de seguridad

```bash
git checkout -b refactor/academic-modularization
git add -A && git commit -m "chore: snapshot before academic modularization"
```

No tocar ningún archivo antes de este commit.

---

### FASE 1 — Crear carpetas destino

Crear físicamente la estructura vacía en el proyecto antes de mover cualquier archivo. En Visual Studio: clic derecho sobre el proyecto → *Add → New Folder*.

```
Academic\Domain\Models\
Academic\Domain\Interfaces\
Academic\Application\Services\
Academic\Application\Mappers\
Academic\Application\DTOs\
Academic\Infrastructure\Persistence\Repositories\
Academic\Infrastructure\Persistence\Context\
Academic\Infrastructure\Persistence\Sql\
Academic\UI\Forms\
```

---

### FASE 2 — Mover archivos y renombrar namespaces

Mover los archivos **uno a uno** (no en bloque), actualizando el namespace de cada archivo inmediatamente después de moverlo.

#### 2.1 Domain — Modelos e Interfaces

| # | Origen | Destino |
|---|---|---|
| 1 | `Core\Models\Academic\*.cs` | `Academic\Domain\Models\` |
| 2 | `Core\Interfaces\IAcademicRepository.cs` | `Academic\Domain\Interfaces\` |

**Cambio de namespace por archivo:**

```csharp
// ANTES
namespace ConversorPDF.XLL.Core.Models.Academic
namespace ConversorPDF.XLL.Core.Interfaces

// DESPUÉS
namespace ConversorPDF.XLL.Academic.Domain.Models
namespace ConversorPDF.XLL.Academic.Domain.Interfaces
```

> ⚠️ Si `IAcademicRepository.cs` convive en un archivo con interfaces de Viáticos, **extraer solo la interfaz académica** a un nuevo archivo, sin tocar el original.

---

#### 2.2 Application — Servicios, Mappers y DTOs

| # | Origen | Destino |
|---|---|---|
| 3 | `Core\Services\Academic\*.cs` | `Academic\Application\Services\` |
| 4 | `Core\Helpers\GridToSesionMapper.cs` | `Academic\Application\Mappers\` |
| 5 | `Core\Helpers\GridRowDto.cs` | `Academic\Application\DTOs\` |

**Cambio de namespace:**

```csharp
// ANTES
namespace ConversorPDF.XLL.Core.Services.Academic
namespace ConversorPDF.XLL.Core.Helpers

// DESPUÉS
namespace ConversorPDF.XLL.Academic.Application.Services
namespace ConversorPDF.XLL.Academic.Application.Mappers
namespace ConversorPDF.XLL.Academic.Application.DTOs
```

> ⚠️ Antes de mover `GridToSesionMapper` y `GridRowDto`, verificar con *Find All References* que **no sean usados por el módulo de Viáticos**. Si lo son, dejar el original en `Core\Helpers\` y crear una copia en `Academic\`, o extraer una interfaz compartida.

---

#### 2.3 Infrastructure — Repositorio, Contexto y SQL

| # | Origen | Destino |
|---|---|---|
| 6 | `Core\Repositories\AcademicRepository.cs` | `Academic\Infrastructure\Persistence\Repositories\` |
| 7 | `Core\Data\EncryptedSQLiteContext.cs` | `Academic\Infrastructure\Persistence\Context\` |
| 8 | `Core\Data\SQLiteEngineBootstrap.cs` | `Academic\Infrastructure\Persistence\Context\` |
| 9 | `Core\Data\AcademicUnitOfWork.cs` | `Academic\Infrastructure\Persistence\` |
| 10 | `Core\Data\AcademicSchema.sql` | `Academic\Infrastructure\Persistence\Sql\` |
| 11 | `Core\Data\AcademicSeedData.sql` | `Academic\Infrastructure\Persistence\Sql\` |

**Cambio de namespace:**

```csharp
// ANTES
namespace ConversorPDF.XLL.Core.Repositories
namespace ConversorPDF.XLL.Core.Data

// DESPUÉS
namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context
namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence
```

> ⚠️ **Archivos `.sql`:** Verificar en el `.csproj` el `Build Action` asignado (`EmbeddedResource` o `Content`). Si se leen con `Assembly.GetManifestResourceStream(...)`, actualizar la ruta del recurso al nuevo namespace de carpeta.

```csharp
// Ejemplo — actualizar ruta del recurso embebido:
// ANTES
var stream = assembly.GetManifestResourceStream("ConversorPDF.XLL.Core.Data.AcademicSchema.sql");

// DESPUÉS
var stream = assembly.GetManifestResourceStream("ConversorPDF.XLL.Academic.Infrastructure.Persistence.Sql.AcademicSchema.sql");
```

---

#### 2.4 UI — Formulario WinForms

| # | Origen | Destino |
|---|---|---|
| 12 | `FormAcademic\FormSesionesAcademicas.cs` | `Academic\UI\Forms\` |
| 13 | `FormAcademic\FormSesionesAcademicas.Designer.cs` | `Academic\UI\Forms\` |
| 14 | `FormAcademic\FormSesionesAcademicas.resx` | `Academic\UI\Forms\` |

**Cambio de namespace:**

```csharp
// ANTES
namespace ConversorPDF.XLL.FormAcademic

// DESPUÉS
namespace ConversorPDF.XLL.Academic.UI.Forms
```

> ⚠️ El `.Designer.cs` y el `.resx` deben moverse **junto con** el `.cs` o el diseñador de WinForms romperá el vínculo.  
> En Visual Studio: mover manualmente en el explorador de archivos → clic derecho en el proyecto → *Add → Existing Item* para re-incluirlos.

---

### FASE 3 — Actualizar `using` internos del módulo Academic

Hacer *Find & Replace* **con scope limitado a la carpeta `Academic\`** para no tocar Viáticos ni Core:

| Namespace antiguo | Namespace nuevo |
|---|---|
| `ConversorPDF.XLL.Core.Models.Academic` | `ConversorPDF.XLL.Academic.Domain.Models` |
| `ConversorPDF.XLL.Core.Interfaces` | `ConversorPDF.XLL.Academic.Domain.Interfaces` |
| `ConversorPDF.XLL.Core.Services.Academic` | `ConversorPDF.XLL.Academic.Application.Services` |
| `ConversorPDF.XLL.Core.Helpers` | `ConversorPDF.XLL.Academic.Application.Mappers` / `.DTOs` |
| `ConversorPDF.XLL.Core.Repositories` | `ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories` |
| `ConversorPDF.XLL.Core.Data` | `ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context` |
| `ConversorPDF.XLL.FormAcademic` | `ConversorPDF.XLL.Academic.UI.Forms` |

---

### FASE 4 — Actualizar el único punto de entrada externo

Solo se modifica **una línea** en `RibbonController.cs`. El resto del código del Ribbon no se toca.

```csharp
// RibbonController.cs — ANTES
using ConversorPDF.XLL.FormAcademic;

// RibbonController.cs — DESPUÉS
using ConversorPDF.XLL.Academic.UI.Forms;
```

La instanciación del form y la llamada a `.Show()` permanecen idénticas.

---

### FASE 5 — Primera compilación de control

Compilar **antes de limpiar** para detectar referencias rotas con los archivos origen todavía presentes.

```
Build → Revisar lista de errores
```

Corregir cualquier `using` faltante o namespace incorrecto antes de continuar. En este punto los archivos origen siguen en `Core\` — no hay pérdida de datos.

---

### FASE 6 — Limpieza de Core (eliminar orígenes)

Solo ejecutar después de que la compilación sea exitosa (0 errores).

**Eliminar archivos y carpetas:**

```
Core\Models\Academic\          ← carpeta completa
Core\Services\Academic\        ← carpeta completa
Core\Interfaces\IAcademicRepository.cs
Core\Repositories\AcademicRepository.cs
Core\Helpers\GridToSesionMapper.cs   ← solo si NO es usado por Viáticos
Core\Helpers\GridRowDto.cs           ← solo si NO es usado por Viáticos
Core\Data\EncryptedSQLiteContext.cs
Core\Data\SQLiteEngineBootstrap.cs
Core\Data\AcademicUnitOfWork.cs
Core\Data\AcademicSchema.sql
Core\Data\AcademicSeedData.sql
FormAcademic\                  ← carpeta completa si quedó vacía
```

**Actualizar `ConversorPDF.XLL.csproj`:**

Remover las entradas `<Compile>`, `<EmbeddedResource>` o `<Content>` que apuntan a los archivos eliminados. Agregar las nuevas entradas si Visual Studio no las generó automáticamente.

---

### FASE 7 — Verificación final

#### Compilación

```
Build Solution → 0 errores, 0 warnings nuevos respecto al baseline
```

#### Pruebas manuales

| # | Escenario | Resultado esperado |
|---|---|---|
| 1 | Abrir Excel y hacer clic en el botón "Sesiones Académicas" del Ribbon | El formulario `FormSesionesAcademicas` se abre correctamente |
| 2 | Inicialización de SQLite | El `ProgressManager` aparece; la DB se desencripta desde la nueva ubicación |
| 3 | CRUD Académico | Guardar una sesión y verificar persistencia en la base de datos |
| 4 | Smoke test Viáticos | Toda la lógica de Viáticos funciona sin cambios |
| 5 | Re-apertura en frío | Cerrar Excel, reabrir, repetir escenarios 1–3 |

#### Commit final

```bash
git add -A
git commit -m "refactor: academic module isolated under Academic\ vertical slice"
```

---

## Resumen de Impacto

| Capa | Archivos movidos | Archivos no tocados |
|---|---|---|
| Domain | `*.cs` de Models + `IAcademicRepository` | Todo `Core\` de Viáticos |
| Application | Services, Mappers, DTOs | Helpers compartidos (si los hay) |
| Infrastructure | Repository, Context, UnitOfWork, SQL | Configuración global del proyecto |
| UI | Form + Designer + Resources | Resto de formularios y Ribbon |
| Entry Point | Solo cambia 1 `using` en `RibbonController` | Lógica del Ribbon intacta |

---

## Lo que NO se toca

- Todo `Core\` relacionado con Viáticos
- `RibbonController.cs` más allá del `using` de una línea
- El `.csproj` más allá de remover entradas de archivos eliminados
- Cualquier configuración de Excel-DNA / VSTO
- Paquetes NuGet, referencias a Dapper, SQLite.Interop, etc.
