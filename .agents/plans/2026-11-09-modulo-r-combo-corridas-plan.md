---
numbering:
  registry: .agents/plans/REGISTRY.md
  registry_row: R-012
  ronda_base: 289
  fase_base: 1518
  waves: 1
  tasks: 4
---

# Plan R-011: selector de corridas en Módulo R y reparación de la promoción a producción

**Generado**: 2026-11-09 · **Estado**: executing · **Depende de**: R-007 (Fases 1488–1495, 1516–1517)

## Execution envelope
- branch: `refactor/academic-modularization`
- workspace: `c:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`
- verify: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- build_gate: serializado
- no_push: true
- stop_conditions: promoción sin transacción · DELETE en tablas CRDT de producción · editar `APP_VERSION`

| Wave | Ronda | Tasks |
|---|---|---|
| 1 | 289 | T11, T12, T13, T14 |

## Tareas

### T11 · Fase 1518 — Designer: combo de corridas y botón Aprobar
- **id**: T11 · **depends_on**: [] · **fase**: 1518
- **files_lock**: `Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **description**: eliminar `btnCargarUltimaOptimizacion` (instanciación, `Controls.Add`, propiedades, campo); añadir `cmbCargarCorridas` en `(535,10) 195×32` (`DropDownList`, `DropDownWidth 430`); añadir `btnAprobarCorrida` en `pnlAgModo (510,4) 150×28`; mover `lblAgModo` a `(670,10)` con `AutoSize=false`, `Width=276`, `AutoEllipsis`.
- **validation**: `-Mode Build` 0 errores; sin solapes en el header (535–730 / 740–940) ni en `pnlAgModo` (10–500 / 510–660 / 670–946).
- **verification**: `scripts/verify.ps1 -Mode Build` · **status**: done

### T12 · Fase 1519 — Lógica del selector de corridas
- **id**: T12 · **depends_on**: [T11] · **fase**: 1519
- **files_lock**: `Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **description**: eliminar `CargarUltimaOptimizacionAsync` y su handler; `RefrescarComboCorridasAsync()` con `VinculacionQueryRepository(AG).GetCorridasStagingAsync(período, programa, subunidad)` + `AgCorridaSelector`; `CargarCorridaDesdeComboAsync()` con guard `_actualizandoComboCorridas`; refresco en `DropDown` y en `onSeleccion`.
- **validation**: con 2027‑01 + ADMÓN. PÚBLICA TERRITORIAL el combo lista #44 (244 ses) y #46 (sin sesiones) · Error; seleccionar #44 carga grid y métricas.
- **verification**: `scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgCorridaSelector"` · **status**: done

### T13 · Fase 1520 — Aprobación explícita y reparación del promotor
- **id**: T13 · **depends_on**: [T12] · **fase**: 1520
- **files_lock**: `Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `Academic/Application/Services/Genetica/AgPromotorVinculaciones.cs`
- **description**: (G1) `btnAprobarCorrida` → `AprobarCorridaYAulaAsync()` con confirmación; (G2) `AgPromotorVinculaciones` omite el gate de aula si `IdAula <= 0`; (G3) valida `ind.Id > 0`; (G4) `btnPromoverProduccion.Enabled` se restaura en `finally`.
- **validation**: #44 (estado 2, id_aula 0) → aprobar → estado 5 → promover → éxito; rechazo de compuerta deja el botón habilitado.
- **verification**: `scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgPromotorVinculaciones"` · **status**: done

### T14 · Fase 1521 — Pruebas, documentación y sello
- **id**: T14 · **depends_on**: [T13] · **fase**: 1521
- **files_lock**: `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPromotorVinculacionesTests.cs`, `docs/articles/changes/2026-11-09-modulo-r-combo-corridas.md`, `.agents/plans/REGISTRY.md`, `.agents/plans/2026-11-09-modulo-r-combo-corridas-plan.md`
- **description**: casos `WhenCorridaHasNoAula_SkipsAulaGate` y `WhenCorridaHasAulaNotApproved_ShouldReturnFailure`; bitácora del cambio; cierre en el registro.
- **validation**: Fast del promotor 5/5; `-Mode Full` exit 0 sin regresiones.
- **verification**: `scripts/verify.ps1 -Mode Full` · **status**: done

## Criterios de aceptación
1. No existe «Cargar última optimización»; el selector lista las corridas del período/programa con conteos.
2. Seleccionar una corrida rehidrata grid, aptitud, violaciones y vacantes.
3. «Aprobar corrida» deja `ag_corrida.estado = 5` y `ag_aula.estado = 4`.
4. «Promover a Producción» funciona con `id_aula = 0` y sin individuo inválido.
5. Un rechazo de compuerta no deja el botón bloqueado.
6. `-Mode Full` exit 0.
