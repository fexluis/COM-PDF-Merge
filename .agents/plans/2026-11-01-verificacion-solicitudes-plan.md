# Plan R-005: Verificación de Presencia Física de Solicitudes de Viáticos en Disco
**Generado**: 2026-11-01 · **Autor**: Swarm-Planner / Antigravity · **Estado**: approved

## Execution envelope
- branch: refactor/academic-modularization
- workspace: C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace . -Plan .agents/plans/2026-11-01-verificacion-solicitudes-plan.md
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: 952)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado (un unico build/tests a la vez por workspace)
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0:
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Core/Helpers/App.cs
  - ConversorPDF.XLL/Ribbon/RibbonController.cs
- stop_conditions:
  - Acción destructiva no reversible sin respaldo
  - Violación de seguridad o datos
  - Intento de push, merge o publicación a remoto
  - Plan en estado irrecuperable (más de 3 rondas de fix fallidas)
- ledger_owner: swarm-parallel-task
- no_push: true
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: R-005
  - ronda_base: 273
  - fase_base: 1462
  - waves: 3
  - tasks: 6

---

## 1. Resumen Ejecutivo y Decisiones Rectoras

Auditoría de presencia de solicitudes físicas en disco (PDFs) contra los registros de la hoja `CONSULTAS`, marcando la columna H (`INFORME`) como `SIN SOLICITUD` cuando no exista archivo correspondiente `(B, J)`, copiando estados terminales (`REVOCADA`, `RECHAZADA`, `ANULADA`), respetando celdas protegidas (`VERIFICADO`, `REINTEGRO`) sin tocar disco ni modificar celda, y ejecutando guardado condicional del libro de trabajo (`Workbook.Save`) únicamente cuando se hayan producido escrituras efectivas (`CeldasEscritas > 0`).

| # | Decisión | Detalle Técnico |
|---|---|---|
| **D-1** | Auditoría dentro del rango del formulario | El prefill inicial calcula los límites globales del universo en `CONSULTAS` (ej. 126 a 41326); la auditoría se acota estrictamente al rango `(inicio, fin)` filtrado. |
| **D-2** | Búsqueda dual acotada | Sólo `folder` (carpeta del libro de Excel obtenida vía `GetOneDriveFolderPath()`) y `folder\Solicitudes`, vía `CarpetasDeSolicitudes(folder)`. Prohibido recorrer el árbol de OneDrive. |
| **D-3** | Guardado condicional del libro (`Workbook.Save`) | Si `CeldasEscritas > 0`, se invoca `IExcelRepository.SaveActiveWorkbook()`. Si no hubo escrituras, no se toca el archivo. `LibroGuardado` y `ErrorGuardado` se reportan en el DTO de resultado. |
| **D-4** | Optimización de pantalla | Escritura celda a celda protegida por `_excelRepository.ScreenUpdating(false)` en bloque `try/finally`. |
| **D-5** | Protección extendida de columna H | Si `H ∈ {"VERIFICADO", "REINTEGRO"}`: no se consulta disco ni se escribe en celda; se contabiliza en `OmitidasProtegidas`. Si `Estado ∈ {"REVOCADA", "RECHAZADA", "ANULADA"}`: no se consulta disco y se copia el estado a H. |
| **D-6** | Hoja ilegible o ausente | Retorna error estructurado, cero escrituras y cero guardado. |

---

## 2. Parallel execution groups

| Wave | Ronda | Tasks | Can start when |
|---|---|---|---|
| 1 | 273 | T1 | Inicio de plan |
| 2 | 274 | T2, T3 | Wave 1 aprobada (T1 green) |
| 3 | 275 | T4, T5, T6 | Wave 2 aprobada (T2, T3 green) |

---

## 3. Especificación Atómica de Tareas

### T1: Modelos, Constantes y Costura de Guardado en Repositorio
- **id**: T1
- **fase**: 1462
- **depends_on**: []
- **location**: ConversorPDF.XLL/Core
- **files_lock**:
  - ConversorPDF.XLL/Core/Models/EstadoComision.cs
  - ConversorPDF.XLL/Core/Models/ResultadoVerificacionSolicitudes.cs
  - ConversorPDF.XLL/Core/Interfaces/IExcelRepository.cs
  - ConversorPDF.XLL/Core/Repositories/ExcelRepository.cs
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
- **description**: Crear clases de dominio/modelo para `EstadoComision` (incluyendo `Reintegro`) y `ResultadoVerificacionSolicitudes` (métricas y banderas de guardado). Declarar `SaveActiveWorkbook()` en `IExcelRepository` e implementarlo en `ExcelRepository` invocando `GetActiveWorkbook().Save()`. Registrar archivos en `.csproj`.
- **validation**: Interfaz y modelos compilan limpiamente; `SaveActiveWorkbook` invoca `Save()` en COM workbook dinámico.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ExcelRepository"
- **evidence**: _evidence/T1/summary.md
- **status**: pending
- **log**: -
- **files edited/created**: []

---

### T2: Algoritmo de Auditoría y Guardado en ViaticosGestionProcessor
- **id**: T2
- **fase**: 1463
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs
- **files_lock**:
  - ConversorPDF.XLL/Core/Interfaces/IViaticosGestionProcessor.cs
  - ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs
- **description**: Implementar `ObtenerRangoSolicitudes` para determinar límites B en `CONSULTAS`. Implementar `VerificarSolicitudesPresentes(string folder, int inicio, int fin)` con búsqueda acotada a `folder` y subcarpeta `Solicitudes` vía `CarpetasDeSolicitudes(folder)`. Lectura en bloque de columnas B, C, H y J vía `ConstantsHelper`. Aplicar regla D-5 de protección extendida (`VERIFICADO`, `REINTEGRO`) sin lectura de disco, copia de estados terminales (`REVOCADA`, `RECHAZADA`, `ANULADA`), marcación `SIN SOLICITUD` ante ausencia de archivo, idempotencia ante valores ya idénticos (`YaMarcadas`), optimización con `ScreenUpdating` y guardado condicional `SaveActiveWorkbook()` si `CeldasEscritas > 0`.
- **validation**: Filas con VERIFICADO/REINTEGRO intactas sin tocar disco; columna H actualizada; guardado invocado sólo si hubo escrituras.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ViaticosGestionProcessor"
- **evidence**: _evidence/T2/summary.md
- **status**: pending
- **log**: -
- **files edited/created**: []

---

### T3: Integración en Servicio de Viáticos y Formulario UI
- **id**: T3
- **fase**: 1464
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/FormProgramador/FormGestionDocumental.cs
- **files_lock**:
  - ConversorPDF.XLL/Core/Interfaces/IViaticosService.cs
  - ConversorPDF.XLL/Core/Services/ViaticosService.cs
  - ConversorPDF.XLL/FormProgramador/FormGestionDocumental.cs
- **description**: Exponer `ObtenerRangoSolicitudes` y `VerificarSolicitudesPresentes` en `IViaticosService` y `ViaticosService`. En el constructor de `FormGestionDocumental`, tras `InitializeComponent()`, suscribir `this.Load += FormGestionDocumental_Load`. En dicho manejador, prellenar los controles reales `txtViaticoInicial` y `txtViaticoFinal` con el rango obtenido. En `btnRenombrar_Click`, tras renombrar, ejecutar la auditoría de solicitudes y presentar mensaje enriquecido con total de marcas escritas, celdas omitidas y confirmación de guardado.
- **validation**: Al abrir el formulario `txtViaticoInicial` y `txtViaticoFinal` se inicializan automáticamente con el rango de `CONSULTAS`; tras renombrar se ejecuta la verificación y se retroalimenta al usuario.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ViaticosService"
- **evidence**: _evidence/T3/summary.md
- **status**: pending
- **log**: -
- **files edited/created**: []

---

### T4: Suite de Pruebas Unitarias Automatizadas
- **id**: T4
- **fase**: 1465
- **depends_on**: [T2, T3]
- **location**: ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs
- **files_lock**:
  - ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs
- **description**: Implementar 10 pruebas unitarias con Moq y FluentAssertions: (1) Prefill de rango con encabezados y vacíos; (2) Fila con VERIFICADO → sin disco ni escritura; (3) Fila con REINTEGRO → sin disco ni escritura; (4) Estado REVOCADA → copia a H sin consultar disco; (5) Estados ANULADA/RECHAZADA → copia a H; (6) PDF en `Solicitudes` matching (B, J) → sin escritura; (7) B coincide pero cédula no coincide → escribe SIN SOLICITUD; (8) Nombre sin guión → descartado del índice; (9) Fila sin cédula → marca SIN SOLICITUD; (10) `SaveActiveWorkbook` invocado exactamente una vez si `CeldasEscritas > 0` y cero veces si `CeldasEscritas == 0`.
- **validation**: 10 pruebas pasan al 100 % (10 passed, 0 failed).
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ViaticosGestionProcessorTests"
- **evidence**: _evidence/T4/summary.md
- **status**: pending
- **log**: -
- **files edited/created**: []

---

### T5: Documentación Técnica, Changelog y Esquemas DocFX
- **id**: T5
- **fase**: 1466
- **depends_on**: [T2]
- **location**: docs/
- **files_lock**:
  - docs/articles/gestion-documental.md
  - docs/articles/changes/2026-11-01-verificacion-solicitudes.md
  - docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosGestionProcessor.yml
  - docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosService.yml
  - docs/api/ConversorPDF.XLL.Core.Models.ResultadoVerificacionSolicitudes.yml
  - docs/api/ConversorPDF.XLL.Core.Models.EstadoComision.yml
  - docs/api/ConversorPDF.XLL.Core.Models.yml
- **description**: Actualizar documentación técnica de gestión documental describiendo el flujo de verificación y guardado condicional. Crear bitácora en `docs/articles/changes/2026-11-01-verificacion-solicitudes.md`. Generar y sincronizar esquemas DocFX para los nuevos miembros y tipos de API.
- **validation**: Documentación coherente y esquemas DocFX válidos y referenciados en `Models.yml`.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~Documentation"
- **evidence**: _evidence/T5/summary.md
- **status**: pending
- **log**: -
- **files edited/created**: []

---

### T6: Certificación con Gate Full y Cierre Formal del Registro
- **id**: T6
- **fase**: 1467
- **depends_on**: [T4, T5]
- **location**: scripts/verify.ps1
- **files_lock**:
  - .agents/plans/REGISTRY.md
- **description**: Ejecutar `scripts/verify.ps1 -Mode Full`. Verificar la deuda heredada de Fases 1457–1460 registrándolas formalmente en R-003 como aplicadas. Si el gate resulta verde (exit code 0, TOTAL >= 952), sellar formalmente las filas `R-003`, `R-004` y `R-005` a estado `closed` con el commit SHA de cierre. Si el gate falla (exit code != 0), las filas permanecen en `executing` y se abre de inmediato una ronda de fix.
- **validation**: `scripts/verify.ps1 -Mode Full` exit code 0; `REGISTRY.md` sincronizado.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- **evidence**: _evidence/T6/summary.md
- **status**: pending
- **log**: -
- **files edited/created**: []
