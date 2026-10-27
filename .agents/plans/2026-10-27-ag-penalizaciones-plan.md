# Plan: Extraer la penalizacion F2 de AgEvaluadorAptitud a un calculador inyectable

**Generado**: 2026-10-27 · **Autor**: Cline/Swarm-Planner · **Estado**: draft
**Prerrequisito bloqueante**: R1 (workspace con el arbol real) y R3 (baseline medido). Sin ellos este plan NO se ejecuta.

## Execution envelope
- branch: refactor/academic-modularization
- workspace: <R1: copia del ARBOL REAL, fuera de OneDrive. El clon `C:\dev\ConversorPDF` actual es un snapshot pre-Academic y esta prohibido>
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <workspace> -Plan .agents/plans/2026-10-27-ag-penalizaciones-plan.md
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: pendiente R3 — medido, nunca estimado)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado (un unico build-verifier a la vez)
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0:
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Core/Helpers/App.cs
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV*.cs
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs
  - ConversorPDF.XLL/Ribbon/RibbonController.cs
  - Composition Root / IoC Container
  - docs/toc.yml
- stop_conditions:
  - Accion destructiva no reversible sin respaldo
  - Violacion de seguridad o datos
  - Intento de push, merge o publicacion a remoto
  - Plan irrecuperable tras 3 rondas de fix
- ledger_owner: swarm-parallel-task
- no_push: true

---

## Overview

`AgEvaluadorAptitud.EvaluarAptitud` calcula `F = F1 - F2 + 2*F3 + F4 + 2*F5`. El bloque F2 (penalizaciones, lineas 32-53 de `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorAptitud.cs`) mezcla tres fuentes: `individuo.SoftViolations` (blandas ya ponderadas), `AgPesosFitness.PorRegla(regla)` por cada violacion dura registrada, y `PenalidadHardBase` por las duras no registradas (`HardViolations - durasRegistradas`), con fallback `contexto.Pesos ?? AgPesosFitness.PorDefecto()`.

Objetivo: extraer F2 a `IAgPenalizacionCalculator` / `AgPenalizacionCalculator` **preservando el valor numerico exacto** (refactor de comportamiento, no de resultados). Es el cambio de menor riesgo del slice y ejercita todo el pipeline (esqueleto -> RED -> GREEN -> integracion -> verificacion).

Contratos reales involucrados:
- `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgEvaluadorAptitud.cs`
- `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPesosFitness.cs` (`PorRegla(AgReglaViolacion)`, `PorDefecto()`, `PenalidadHardBase`)
- `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgIndividuo.cs` (`SoftViolations`, `Violaciones`, `HardViolations`)
- Unico call-site de construccion: `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:84` (`new AgEvaluadorAptitud(validadorReglas)`)

## Dependency graph

```
T1 (integrator: esqueleto + csproj)
 ├── T2 (RED: tests del calculador)  ─┐
 └── T3 (RED: caracterizacion F2)    ─┴── T4 (implementer: calculador) ── T5 (implementer: refactor + call-site) ── T6 (gate + review)
```

---

## Tasks

### T1: Esqueleto de IAgPenalizacionCalculator + AgPenalizacionCalculator (integrator-first)
- **depends_on**: []
- **location**: `ConversorPDF.XLL/Academic/Application/Interfaces/IAgPenalizacionCalculator.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPenalizacionCalculator.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Interfaces/IAgPenalizacionCalculator.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPenalizacionCalculator.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **description**: Declarar `double Calcular(AgIndividuo individuo, AgContexto contexto)` en la interfaz y una implementacion que lance `NotImplementedException`. Registrar ambos archivos como `<Compile Include>` sin duplicar entradas (evitar CS2002). Un archivo por tipo (`single-type-per-file.md`).
- **validation**: `verify.ps1 -Mode Build` exit 0 y el nuevo tipo visible en el ensamblado.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

### T2: Tests RED del calculador (tdd-red-writer)
- **depends_on**: [T1]
- **location**: `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPenalizacionCalculatorTests.cs`
- **files_lock**:
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPenalizacionCalculatorTests.cs`
- **description**: Cubrir el contrato de F2 con xUnit + FluentAssertions 6.12: (a) solo blandas => `SoftViolations`; (b) cada violacion dura registrada suma `pesos.PorRegla(regla)`; (c) `HardViolations > durasRegistradas` suma la diferencia por `PenalidadHardBase`; (d) sin `Violaciones` suma `HardViolations * PenalidadHardBase`; (e) `contexto.Pesos == null` usa `AgPesosFitness.PorDefecto()`. Nombres `{Accion}_When_{Condicion}_Should_{Resultado}`.
- **validation**: Compila y falla con `NotImplementedException` (RED real, no error de compilacion).
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgPenalizacionCalculatorTests"`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

### T3: Caracterizacion del F2 actual en AgEvaluadorAptitud (tdd-red-writer)
- **depends_on**: [T1]
- **location**: `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEvaluadorAptitudTests.cs`
- **files_lock**:
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEvaluadorAptitudTests.cs`
- **description**: Agregar casos que fijen el fitness actual con violaciones duras y blandas combinadas (incluida la multiplicidad F1026/F940) para que el refactor posterior sea verificable: mismo individuo + mismo contexto => mismo `Fitness` antes y despues.
- **validation**: Compilan y pasan contra la implementacion actual (linea base de comportamiento).
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgEvaluadorAptitudTests"`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

### T4: Implementar AgPenalizacionCalculator (implementer)
- **depends_on**: [T2]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPenalizacionCalculator.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPenalizacionCalculator.cs`
- **description**: Implementar `Calcular` replicando exactamente la semantica F2 documentada en T2. Sin tocar `AgEvaluadorAptitud`, sin tocar L0, sin `DateTime.Now`, sin Interop/WinForms.
- **validation**: Test de T2 en GREEN y `Full` sin regresiones ni suite reducida.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgPenalizacionCalculatorTests"`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

### T5: Refactor AgEvaluadorAptitud + call-site (implementer)
- **depends_on**: [T3, T4]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorAptitud.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorAptitud.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **description**: Inyectar `IAgPenalizacionCalculator` por constructor (`private readonly`), reemplazar el bloque F2 por `_penalizaciones.Calcular(individuo, contexto)` y actualizar el unico call-site (`AgFormGeneradorHorarios.cs:84`) a `new AgEvaluadorAptitud(validadorReglas, new AgPenalizacionCalculator())`. No cambiar la formula F1/F3/F4/F5 ni el orden de evaluacion.
- **validation**: T3 sigue verde (mismo fitness) y `Full` sin regresiones.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgEvaluadorAptitudTests|FullyQualifiedName~AgPenalizacion"`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

### T6: Gate final y revision (build-verifier + task-reviewer, sin escritura de codigo)
- **depends_on**: [T5]
- **location**: `_evidence/` (solo lectura del codigo)
- **files_lock**: []
- **description**: Ejecutar la suite completa (serial, un solo gate) y revisar spec + calidad: cero Interop/WinForms en Application, cero C# 12, un tipo por archivo, sin `.Result`/`.Wait()`, `files edited/created` coherente con `files_lock`.
- **validation**: `verify.ps1 -Mode Full` exit 0, 0 regresiones, suite NO reducida vs baseline; veredicto `OK` del task-reviewer.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

---

## Parallel execution groups

| Wave | Tasks | Can start when |
|---|---|---|
| 1 | T1 (integrator, L0) | preflight exit 0 + baseline medido |
| 2 | T2, T3 (tdd-red-writer, leases disjuntos) | T1 completado |
| 3 | T4 (implementer) | T2 en RED |
| 4 | T5 (implementer) | T3 en RED/pass y T4 en GREEN |
| 5 | T6 (build-verifier + task-reviewer) | T5 integrado |

Regla: ningun par de tareas de la misma ola comparte `files_lock`. T2 y T3 escriben archivos de test distintos; T4 y T5 son seriales por dependencia real.

---

## Testing Strategy

- `Fast` por clase durante el ciclo; `Full` al cerrar cada tarea y al cerrar la ola 5.
- La caracterizacion de T3 es el contrato del refactor: si cambia el valor de fitness con el mismo individuo/contexto, T5 esta mal.
- Prohibido debilitar asserts, subir timeouts o excluir tests con `--filter` fuera del SUT.
- La evidencia (`_evidence/<timestamp>/summary.md`) se enlaza en el campo `evidence` de cada tarea; no se commitea.

---

## Risks & Mitigations

| Riesgo | Mitigacion |
|---|---|
| Baseline inexistente o medido sobre un clon obsoleto | Bloqueo R1+R3: `preflight.ps1` y `Full` fallan antes de empezar |
| El refactor altera el fitness por redondeo u orden | T3 fija el valor actual; T5 compara identico |
| Un segundo constructor/call-site no detectado | T1 incluye `grep_search 'new AgEvaluadorAptitud'`; hoy solo `AgFormGeneradorHorarios.cs:84` |
| Dos gates concurrentes bloquean `bin/*.xll` | `build_gate: serializado` en el envelope |
| `csproj` editado por dos tareas | Solo T1 (integrator) lo toca; L0 exclusivo |
