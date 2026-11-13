---
numbering:
  registry: .agents/plans/REGISTRY.md
  registry_row: R-066
  ronda_base: 358
  fase_base: 1734
  waves: 8
  tasks: 8
---

# Plan: escalado automático de tres casos y planificación balanceada por ranuras
**Generado**: 2026-11-12 · **Autor**: Antigravity/Swarm-Planner · **Estado**: executing

## Execution envelope
- branch: refactor/academic-modularization
- workspace: c:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace "c:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan .agents/plans/2026-11-12-escalado-automatico-tres-casos-plan.md -AllowOneDrive
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: 1223)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado (un único build/tests a la vez; Excel-DNA bloquea bin/)
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0:
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Core/Helpers/App.cs
  - ConversorPDF.XLL/**/Migrations/MigrationV*.cs
  - ConversorPDF.XLL/**/Persistence/SchemaMigrator.cs
  - ConversorPDF.XLL/Ribbon/RibbonController.cs
  - docs/toc.yml
- stop_conditions:
  - Acción destructiva no reversible sin respaldo
  - git clean / git stash / cambio de rama con archivos sin seguimiento
  - Intento de push, merge o publicación a remoto
  - Plan irrecuperable tras más de 3 rondas de fix fallidas
- ledger_owner: swarm-parallel-task
- no_push: true
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: R-066
  - ronda_base: 358          # Wave 1 = Ronda 358
  - fase_base: 1734          # T1 = Fase 1734
  - waves: 8                 # Ronda 358–365
  - tasks: 8                 # Fase 1734–1741

---

## 1. Grafo de dependencias

```text
T1 (RED: Contrato aceptación y reproducción)
 └─► T2 (Origen de patrones y migración V80)
      └─► T3 (Planificación por ranuras y demanda)
           └─► T4 (Evaluador compartido y decisión de escalado)
                └─► T5 (Aplicación transaccional e invalidación)
                     └─► T6 (Integración UI y pre-vuelo en AgFormGeneradorHorarios)
                          └─► T7 (Matriz integral de validación tres casos)
                               └─► T8 (Gates, documentación y handoff)
```

Dependencias declaradas:
- `T1: []`
- `T2: [T1]`
- `T3: [T2]`
- `T4: [T3]`
- `T5: [T4]`
- `T6: [T5]`
- `T7: [T6]`
- `T8: [T7]`

---

## 2. Olas (Parallel execution groups)

| Wave | Ronda | Tasks | Can start when | Leases |
|---|---|---|---|---|
| 1 | 358 | T1 | Inmediato | tests contrato aceptación |
| 2 | 359 | T2 (L0) | T1 done | Origen patron enum, AgAulaAsignatura, MigrationV80, SqliteHelper |
| 3 | 360 | T3 | T2 done | AgParticionCarriles, AgPlanificadorOleadas |
| 4 | 361 | T4 | T3 done | AgEvaluadorDemandaVentana, AgEscaladorVentanaAula |
| 5 | 362 | T5 | T4 done | Persistencia AgEscaladorVentanaAula |
| 6 | 363 | T6 | T5 done | AgFormGeneradorHorarios |
| 7 | 364 | T7 | T6 done | Tests matriz de casos completa |
| 8 | 365 | T8 | T7 done | Handoff, doc cambios, gates |

---

## 3. Especificación de Tareas

### T1: Contrato de aceptación y reproducción de cuellos de botella
- **id**: T1
- **fase**: 1734
- **depends_on**: []
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEscaladoTresCasosTests.cs
- **description**: Crear pruebas de caracterización y reproducción que demuestren las fallas de cuello de botella: 64 sesiones (12,12,8,8,12,12) en LUN-VIE sobrecargando viernes (20 sesiones) y en VIE-DOM sobrecargando domingo (32 sesiones cuando solo hay 17 domingos).
- **validation**: Pruebas unitarias de caracterización escritas y fallando por la causa esperada (RED).
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgEscaladoTresCasosTests"
- **evidence**: _evidence/20261112-190903/summary.md
- **status**: done
- **log**: Verificación Fast ejecutada con éxito (3/3 tests PASS). Se demostró la sobrecarga sistemática de la siembra ingenua en viernes (20 > 17) y domingo (32 > 17).
- **files edited/created**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEscaladoTresCasosTests.cs

### T2: Origen y representación de patrones en AgAulaAsignatura
- **id**: T2
- **fase**: 1735
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic
- **files_lock**:
  - ConversorPDF.XLL/Academic/Domain/Enums/AgOrigenPatron.cs
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgAulaAsignatura.cs
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV80_AddOrigenPatronAgAulaAsignatura.cs
  - ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.Tests/Academic/Domain/AgOrigenPatronTests.cs
- **description**: Incorporar enum AgOrigenPatron (AUTOMATICO, FIJO, LEGADO) en dominio, añadir propiedad OrigenPatron en AgAulaAsignatura, implementar MigrationV80 para agregar columna origen_patron en ag_aula_asignatura con fallback seguro, y registrar en SqliteHelper.
- **validation**: AgAulaAsignatura y DTOs leen/escriben OrigenPatron correctamente. Tests de migración pasan.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgOrigenPatronTests"
- **evidence**: _evidence/20261112-191039/summary.md
- **status**: done
- **log**: Verificación Fast ejecutada con éxito (3/3 tests PASS). AgOrigenPatron enum, AgAulaAsignatura.OrigenPatron, MigrationV80 y registro en SqliteHelper/csproj implementados.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Domain/Enums/AgOrigenPatron.cs
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgAulaAsignatura.cs
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV80_AddOrigenPatronAgAulaAsignatura.cs
  - ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.Tests/Academic/Domain/AgOrigenPatronTests.cs

### T3: Planificación por ranuras y demanda en AgParticionCarriles y AgPlanificadorOleadas
- **id**: T3
- **fase**: 1736
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgParticionCarriles.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgParticionCarrilesTests.cs
- **description**: Sustituir el reparto ingenuo de materias iguales por carril. Asignar materias automáticas según sesiones requeridas y capacidad efectiva de cada ranura/carril. Respetar estrictamente restricciones fijas y legadas sin mutarlas.
- **validation**: Partición de carriles balancea la carga según capacidad y resuelve cuellos de botella sin desbordar carriles pequeños (como domingo en VIE-DOM).
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgParticionCarriles"
- **evidence**: _evidence/20261112-191247/summary.md
- **status**: done
- **log**: Verificación Fast ejecutada con éxito (14/14 tests PASS). AgParticionCarriles evalúa alternativas para patrones automáticos y permite optimización temporal (makespan) cuando el número de infactibles es igual.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgParticionCarriles.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgParticionCarrilesTests.cs

### T4: Evaluador compartido y decisión de escalado (AgEvaluadorDemandaVentana)
- **id**: T4
- **fase**: 1737
- **depends_on**: [T3]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentanaTests.cs
  - ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgEscaladorVentanaAulaTests.cs
- **description**: Crear AgEvaluadorDemandaVentana que evalúe tanto la capacidad global como la factibilidad por carril. Eliminar la salida anticipada errónea de AgEscaladorVentanaAula (demanda <= capacidadTotal) cuando hay saturación de carril. Evaluar ganancia real hacia AgEscaleraVentanas.Siguiente.
- **validation**: AgEscaladorVentanaAula detecta saturación en VIE-DOM o LUN-JUE aunque la suma global de horas fuera menor que el total de ranuras.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgEvaluadorDemandaVentana|AgEscaladorVentanaAulaTests"
- **evidence**: pendiente
- **status**: pending
- **log**: pendiente
- **files edited/created**: []

### T5: Aplicación transaccional e invalidación
- **id**: T5
- **fase**: 1738
- **depends_on**: [T4]
- **location**: ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories
- **files_lock**:
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs
  - ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgEscaladorVentanaAulaPersistenciaTests.cs
- **description**: Asegurar que AplicarAsync actualice ag_aula, ag_arbol_nodo y actualice patrones automáticos redistribuidos en una única transacción atómica. Invalidar caché y verificar idempotencia en segunda corrida.
- **validation**: Transacción atómica confirmada, sin mutaciones parciales ante errores, segunda corrida sobre estado factible no produce propuestas.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgEscaladorVentanaAulaPersistenciaTests"
- **evidence**: pendiente
- **status**: pending
- **log**: pendiente
- **files edited/created**: []

### T6: Integración UI y pre-vuelo en AgFormGeneradorHorarios
- **id**: T6
- **fase**: 1739
- **depends_on**: [T5]
- **location**: ConversorPDF.XLL/Academic/UI/Forms
- **files_lock**:
  - ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs
- **description**: Unificar pre-vuelo entre Diagnosticar y Optimizar. Modal informativo detallado (aula, ruta, días añadidos, fechas originales fijas, sesiones antes/después, demanda pendiente). Bloqueo de acciones y recarga segura.
- **validation**: Mismo diagnóstico desde Diagnosticar y Optimizar; confirmación modal respetada; segunda corrida estable.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
- **evidence**: pendiente
- **status**: pending
- **log**: pendiente
- **files edited/created**: []

### T7: Matriz integral de validación de los tres casos
- **id**: T7
- **fase**: 1740
- **depends_on**: [T6]
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEscaladoTresCasosMatrizTests.cs
- **description**: Implementar la suite completa de pruebas para los casos de la matriz: LUN-JUE 5W, MAR-VIE 5W, VIE-DOM 5W (a JUE-DOM), VIE-DOM 6W (paso 1 JUE-DOM, paso 2 MIE-DOM), MIE-DOM/LUN-VIE terminales, sábado 4h (2 ranuras) y 8h (1 ranura).
- **validation**: Todos los casos de la matriz verificados y en verde (PASS).
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgEscaladoTresCasosMatrizTests"
- **evidence**: pendiente
- **status**: pending
- **log**: pendiente
- **files edited/created**: []

### T8: Gates, documentación y handoff
- **id**: T8
- **fase**: 1741
- **depends_on**: [T7]
- **location**: .agents/plans
- **files_lock**:
  - .agents/plans/2026-11-12-escalado-automatico-tres-casos-handoff.md
  - docs/articles/changes/2026-11-12-escalado-automatico-tres-casos.md
- **description**: Ejecutar suite completa verify.ps1 -Mode Full, actualizar test_baseline.txt si creció la suite sin regresiones, redactar handoff y artículo de cambios técnicos.
- **validation**: Full verify PASS con 0 regresiones.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- **evidence**: pendiente
- **status**: pending
- **log**: pendiente
- **files edited/created**: []
