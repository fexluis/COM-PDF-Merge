# Plan: Reparación automática de R5 con inventario de ranuras (segunda ronda)

**Generado**: 2026-11-22 · **Autor**: Swarm-Planner · **Estado**: closed

## Execution envelope
- branch: reparacion-automatica-r5
- workspace: C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace "C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan ".agents/plans/2026-11-22-reparacion-automatica-r5-plan.md" -AllowOneDrive
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt
- verify: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj
- build_gate: serializado
- max_parallel_writers: 2
- stop_conditions:
  - Acción destructiva no reversible sin respaldo
  - Violación de seguridad o datos
  - Plan en estado irrecuperable
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: R-197
  - ronda_base: 525          # Wave 525 = Ronda 525
  - fase_base: 2322          # T1 = Fase 2322
  - waves: 1                 # Ronda 525
  - tasks: 10                # Fase 2322..2331

---

## Mapeo de Olas y Rondas

| Wave | Ronda | Tasks |
|---|---|---|
| Wave 525 | Ronda 525 | T1..T10 (Fase 2322–2331) |

---

## Tasks

### T1: Rama de protección en GitHub
- **id**: T1
- **fase**: 2322
- **depends_on**: []
- **location**: git / origin
- **files_lock**:
  - `.agents/plans/REGISTRY.md`
- **descripción**: Crear rama `reparacion-automatica-r5` desde `rendimiento-app` y publicarla en origin.

### T2: Reserva de numeración y plan swarm
- **id**: T2
- **fase**: 2323
- **depends_on**: [T1]
- **location**: .agents/plans/
- **files_lock**:
  - `.agents/plans/REGISTRY.md`
  - `.agents/plans/2026-11-22-reparacion-automatica-r5-plan.md`
- **descripción**: Reservar contadores en REGISTRY.md §2 y fila R-197 en §3, y crear plan de ejecución formal.

### T3: AgInventarioRanuras
- **id**: T3
- **fase**: 2324
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgInventarioRanuras.cs`
- **descripción**: Construir modelo de inventario de ranuras libres, válidas, liberables y ocupadas por aula y cohorte.

### T4: AgReparadorR5 núcleo y scoring
- **id**: T4
- **fase**: 2325
- **depends_on**: [T3]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR5.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **descripción**: Extraer helper canónico de carriles y armar estructura de scoring y presupuesto de reparación.

### T5: Candidatos C1 (desplazamiento) y C2 (compresión)
- **id**: T5
- **fase**: 2326
- **depends_on**: [T4]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR5.cs`
- **descripción**: Generar operadores de forward shifting (C1) y compresión sobre ranuras reales (C2).

### T6: Candidatos C3/C4 y validación V1–V6
- **id**: T6
- **fase**: 2327
- **depends_on**: [T5]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR5.cs`
- **descripción**: Pre-validación V1–V6 contra colisiones de docente, aula, carril R6 y reglas duras con rollback garantizado.

### T7: Hook en AgGeneradorHorarios
- **id**: T7
- **fase**: 2328
- **depends_on**: [T6]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`
- **descripción**: Integrar hook de llamada a `AgReparadorR5.Reparar` tras `RealinearOleadas` con control de rollback.

### T8: Telemetría
- **id**: T8
- **fase**: 2329
- **depends_on**: [T7]
- **location**: ConversorPDF.XLL/Academic/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgResultadoCorrida.cs`
- **descripción**: Exponer métricas de reparación R5 (intentos, hard pre/post, desperdicio pre/post) y serializar en parámetros JSON.

### T9: Pruebas unitarias
- **id**: T9
- **fase**: 2330
- **depends_on**: [T8]
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgInventarioRanurasTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgReparadorR5Tests.cs`
- **descripción**: Cobertura completa de casos para inventario de ranuras, operadores de reparación y reglas de rollback.

### T10: Auditoría, gates y cierre
- **id**: T10
- **fase**: 2331
- **depends_on**: [T9]
- **location**: scripts/ / REGISTRY.md
- **files_lock**:
  - `.agents/plans/REGISTRY.md`
  - `implementation_plan.md`
  - `walkthrough.md`
- **descripción**: Certificación por scripts de auditoría, preflight y verificación de tests, y cierre documental de R-197.
