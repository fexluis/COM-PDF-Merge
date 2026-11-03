---
numbering:
  registry: .agents/plans/REGISTRY.md
  registry_row: R-006
  ronda_base: 279
  fase_base: 1483
  waves: 3
  tasks: 5
---

# Plan R-006: Duplicados en `_Duplicados` y auditoría de INFORME (columna H)

**Generado**: 2026-10-31 · **Estado**: executing · **Sustituye**: R-005 (`abandoned`)

## Execution envelope
- branch: refactor/academic-modularization
- workspace: `c:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`
- preflight: `powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace . -Plan .agents/plans/2026-10-31-duplicados-y-auditoria-informe-plan.md`
- baseline: `tests/baseline/test_baseline.txt` (TOTAL 952)
- verify: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- build_gate: serializado
- max_parallel_writers: 2
- hot_paths_L0: `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `Core/Helpers/App.cs`, `Ribbon/RibbonController.cs`
- stop_conditions: acción destructiva sin respaldo · push/merge · renombrar datos reales fuera de copia
- no_push: true

## Alcance
1. Duplicados exactos → `Solicitudes\_Duplicados`.
2. Auditoría de presencia física en H: proteger `VERIFICADO`/`REINTEGRO`, copiar estados terminales, limpiar marcas erróneas, marcar `SIN SOLICITUD`, preservar otras marcas y guardar el libro.
3. Remapeo tolerante de prefijos de 6 dígitos.
4. Prefill del rango (columna B) en el formulario y mensaje único con desglose + estado de guardado.

## Tareas

### T1 · Fase 1483 — Duplicados
- **id**: T1 · **depends_on**: [] · **location**: `Core/Processors/ViaticosGestionProcessor.cs`
- **files_lock**: `Core/Processors/ViaticosGestionProcessor.cs`, `Core/Models/ResultadoRenombrado.cs`, `Core/Models/MotivoRenombrado.cs`, `Core/Helpers/App.cs`
- **description**: destino existente + contenido idéntico ⇒ mover a `_Duplicados`; contenido distinto ⇒ `COLISION`.
- **validation**: 14 conflictos pasan a `Duplicados`; `Colisiones` = 0.
- **verification**: `scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ViaticosGestionProcessorTests"`
- **evidence**: `_evidence/<stamp>/summary.md` · **status**: done · **log**: duplicados + MismoContenido

### T2 · Fase 1484 — Auditoría H
- **id**: T2 · **depends_on**: [T1] · **location**: `Core/Processors/ViaticosGestionProcessor.cs`
- **files_lock**: `Core/Processors/ViaticosGestionProcessor.cs`, `Core/Interfaces/IExcelRepository.cs`, `Core/Repositories/ExcelRepository.cs`, `Core/Interfaces/IViaticosGestionProcessor.cs`, `Core/Models/EstadoComision.cs`, `Core/Models/ResultadoVerificacionSolicitudes.cs`, `Core/Models/RangoSolicitudes.cs`, `ConversorPDF.XLL.csproj`
- **description**: `VerificarSolicitudesPresentes` + `ObtenerRangoSolicitudes` + `SaveActiveWorkbook`.
- **validation**: 18 filas sin archivo marcadas; 2 marcas erróneas limpiadas; `VERIFICADO`/`REINTEGRO` intactas; libro guardado.
- **verification**: `scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ViaticosGestionProcessorTests"`
- **evidence**: `_evidence/<stamp>/test.trx` · **status**: done · **log**: `[VERIFICAR_SOLICITUDES]`

### T3 · Fase 1485 — Normalización y remapeo
- **id**: T3 · **depends_on**: [T2] · **location**: `Core/Processors/ViaticosGestionProcessor.cs`
- **files_lock**: `Core/Processors/ViaticosGestionProcessor.cs`
- **description**: `Norm`, `EsMarcaFaltante` (incluye `FALTA SOLICITUD`), `SinAcentos`/`CoincideNombre`, remapeo de prefijos de 6 dígitos.
- **validation**: `202626` → `20626-16796878-GermanLuisRoldanNo.pdf`; prefijo sin coincidencia ⇒ `SIN_COMISION`.
- **verification**: idem T2 · **evidence**: `_evidence/<stamp>/test.trx` · **status**: done

### T4 · Fase 1486 — Pruebas
- **id**: T4 · **depends_on**: [T2, T3] · **location**: `ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs`
- **files_lock**: `ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs`
- **description**: 12 casos (duplicado, colisión, remapeo, prefijo sin coincidencia, rango, marcado+guardado, limpieza, protegidos, estados terminales, otras marcas, idempotencia).
- **validation**: 21 pruebas de la clase en verde; `TOTAL` ≥ baseline.
- **verification**: `scripts/verify.ps1 -Mode Full` · **evidence**: `_evidence/<stamp>/test.trx` · **status**: done

### T5 · Fase 1487 — Documentación y cierre
- **id**: T5 · **depends_on**: [T4] · **location**: `docs/articles/gestion-documental.md`
- **files_lock**: `docs/articles/gestion-documental.md`, `docs/articles/changes/2026-10-31-duplicados-auditoria-informe.md`, `.agents/plans/REGISTRY.md`
- **description**: documentar duplicados + auditoría; sellar R-006 y contadores.
- **validation**: docs actualizados; `REGISTRY` con `R-006 closed` y §2 en `281 / 1487`.
- **verification**: `scripts/verify.ps1 -Mode Mode Full` · **evidence**: `_evidence/<stamp>/summary.md` · **status**: done

## Criterios de aceptación
1. `Duplicados = 14`, `Colisiones = 0` sobre el dataset del 2026-10-31.
2. `CeldasLimpiadas = 2` y `CeldasEscritas ≈ 18`; `LibroGuardado = true`.
3. Segunda corrida: `CeldasEscritas = 0`, `CeldasLimpiadas = 0`, `LibroGuardado = false`.
4. Gate `-Mode Full` exit 0 con `TOTAL` ≥ 952.
