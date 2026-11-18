# Plan: Optimización de Rendimiento AG — Corrida #105 (R-112, Ronda 433–438)

**Generado**: 2026-11-17 · **Autor**: Swarm-Planner · **Estado**: executing

## Execution envelope
- branch: rendimiento-app
- workspace: C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace "C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan ".agents/plans/2026-11-17-rendimiento-app-corrida105-plan.md" -AllowOneDrive
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: 1674 medido 2026-11-17 con dotnet test previo a T1)
- verify: dotnet build .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug ; dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --logger "trx;LogFileName=gate-alternativo.trx"
- build_gate: serializado (un único build/tests a la vez por workspace; Excel-DNA bloquea bin/)
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0:
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Core/Helpers/App.cs
  - ConversorPDF.XLL/**/Migrations/MigrationV*.cs
  - ConversorPDF.XLL/**/Persistence/SchemaMigrator.cs
  - ConversorPDF.XLL/Ribbon/RibbonController.cs
  - Composition Root / IoC Container
  - docs/toc.yml
- stop_conditions:
  - Acción destructiva no reversible sin respaldo
  - Violación de seguridad o datos
  - Intento de push, merge o publicación a remoto (solo autorizado push inicial de rama ya completado)
  - Plan en estado irrecuperable (más de 3 rondas de fix fallidas)
- ledger_owner: swarm-parallel-task
- no_push: true
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: R-112
  - ronda_base: 433          # Wave 1 = Ronda 433
  - fase_base: 1942          # T1 = Fase 1942
  - waves: 6                 # Ronda 433..438
  - tasks: 10                # Fase 1942..1951

---

## Contexto y Métricas Base

- **Tiempo actual OptimizacionAG**: 240,158 s (98,5 % de 243,822 s total).
- **Ritmo de estado estacionario**: 3,90 s por transición entre Gen 2 y Gen 60 (226 s / 58 transiciones; 52 intervalos de 4 s y 6 de 3 s).
- **Presupuesto global actual**: 4,003 s/generación sobre 60 generaciones reportadas (incluye ~14 s iniciales: factibilidad interna, índices, población inicial y primer ciclo).
- **Última mejora**: Gen 40; parada en Gen 60 = 77 s (~32 % del tiempo). Se preserva `estancamientoMax` (no recortar a ciegas).
- **Resolución timestamps**: 1 s en logs actuales; no apto para p50/p95. Se medirá con `Stopwatch`.
- **Transiciones reproductivas**: 59 transiciones reales (Gen 1→2 hasta Gen 59→60; Gen 60 reporta y para).
- **Meta presupuestaria**: Reducir estado estacionario a ~0,65 s/transición (o meta intermedia ≤1,0 s) y optimizar el primer intervalo (~14 s → ~7 s) para alcanzar el objetivo global de ~45 s.
- **Invariantes L0**: No tocar `ConversorPDF.XLL.csproj` ni migraciones. Nuevos tipos (`AgPerfilGeneracion`, `AgProyeccionIndividuo`, etc.) van anidados/internos en archivos existentes.

---

## Criterios de Aceptación Globales

1. **Equivalencia de Trayectoria y Calidad**: Misma semilla en baseline y versión optimizada produce idéntica aptitud o mejora estricta; no superar las 30 violaciones blandas históricas; `HARD=0` en campeón final.
2. **Volúmenes Canónicos**: 410 genes, 2.228 sesiones, 8.912 horas, 199 vinculaciones y 336 vacantes en el horario final materializado.
3. **Fidelidad de Horario**: El hash canónico del horario materializado coincide bit a bit entre baseline sembrado y versión optimizada.
4. **Presupuesto de Rendimiento**:
   - Transición de estado estacionario ≤ 1,00 s (objetivo ~0,65 s).
   - Tiempo total ≤ 72 s (objetivo ≤ 45 s con inicio optimizado a ~7 s).
   - Medición generacional precisa con `Stopwatch` en `[AG-PERF]`.
5. **No Regresiones**: Suite completa en verde con el gate alternativo autorizado (0 fallos).

---

## Dependency Graph

```text
T1 ── T2 ── T4 ── T6 ── T8 ─┐
  │      │                 ├── T10
  │      └── T5 ── T7 ─────┤
  └── T3 ──────────────────┤
           T9 ─────────────┘
```

---

## Parallel execution groups (Waves)

| Wave | Ronda | Tasks | Can start when | Leases |
|---|---|---|---|---|
| 1 | 433 | T1 | Inmediato | `ConversorPDF.XLL/Academic/Application/Services/Genetica/{AgOperadorCruce, AgOperadorMutacion, AgOperadorSeleccion, AgFabricaIndividuos, AgAsignadorTercero}.cs`, `ConversorPDF.Tests/.../AgRandomSeedTests.cs` |
| 2 | 434 | T2, T3 | T1 completado | Disjuntos: T2 en `AgGeneradorHorarios.cs`, `AgResultadoCorrida.cs`, `AgBitacora.cs`, `AgPerfilGeneracionTests.cs` vs T3 en `AgGeneradorHorariosReplayTests.cs`, `AgCorrida105Fixture.cs` |
| 3 | 435 | T4, T5 | T2 completado | Disjuntos: T4 en `AgAsignadorTercero.cs`, `AgCalendarioConflictos.cs`, `AgCoherenciaGrupo.cs`, `AgPerfilSubetapasTests.cs` vs T5 en `AgContexto.cs`, `AgBloqueHorario.cs`, `CalculadoraViaticaSesion.cs`, `AgExpansorSesiones.cs`, `AgProyeccionHigieneTests.cs` |
| 4 | 436 | T6, T7 | T4 y T5 completados | Disjuntos: T6 en `AgOperadorCruce.cs`, `AgOperadorMutacion.cs`, `AgValidadorReglas.cs`, `AgEvaluadorAptitud.cs`, `AgQuickWinsTests.cs` vs T7 en `AgCalendarioConflictos.cs`, `AgExpansorSesiones.cs`, `AgProyeccionLigeraTests.cs` |
| 5 | 437 | T8, T9 | T6 y T2 completados | Disjuntos: T8 en `AgAsignadorTercero.cs`, `AgOperadorCruce.cs`, `AgValidadorReglas.cs`, `AgIndiceIndividuoTests.cs` vs T9 en `AgIndividuo.cs`, `AgGen.cs`, `AgGeneradorHorarios.cs`, `AgHuellaAptitudTests.cs` |
| 6 | 438 | T10 | T3, T6, T7, T8, T9 completados | `AgRendimientoCertificacionTests.cs`, `docs/articles/changes/2026-11-17-rendimiento-ag-corrida105.md` |

---

## Tasks

### T1: Semilla determinista y flujos RNG
- **id**: T1
- **fase**: 1942
- **depends_on**: []
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorSeleccion.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRandomSeedTests.cs
- **description**: Añadir constructores o sobrecargas que acepten `Random` o semilla entera; definir flujos deterministas por operador genético; conservar el comportamiento actual sin cambios cuando no se pasa semilla. No alterar el orden de consumo de números aleatorios en operadores existentes.
- **validation**: Dos ejecuciones con la misma semilla producen exactamente la misma secuencia de genes y aptitudes; semillas distintas divergen; toda la suite de pruebas existente continúa pasando en verde.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgRandomSeedTests"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:

### T2: Telemetría generacional y corrección de DuracionMs
- **id**: T2
- **fase**: 1943
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgResultadoCorrida.cs
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgBitacora.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPerfilGeneracionTests.cs
- **description**: Corregir el cronómetro generacional (`DuracionMs` real); declarar `AgPerfilGeneracion` como clase interna dentro de `AgResultadoCorrida.cs`; poblar métricas generacionales (`PeorFitness`, `Evaluados`, `Vacantes`, `FueraDeVentana`); emitir línea estructurada `[AG-PERF]`; medir subtiempos de inicialización (factibilidad interna, índices, población inicial, primer ciclo) y de ciclo reproductivo (selección, cruce neto, mutación neta, sellado en cruce/mutación, evaluación, reemplazo élite, campeón final); contar `EvaluacionesAptitud`, `HijosCreados`, `HijosIncorporados`, `ElitesClonados`, `EvaluacionesExtraPorCupoImpar`, genes y transiciones.
- **validation**: `DuracionMs` coincide con el tiempo real de pared; la suma de etapas parciales cuadra con el tiempo total reportado; sin cambios funcionales en los resultados de programación.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgPerfilGeneracionTests|FullyQualifiedName~AgGeneradorHorariosTests"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:

### T3: Arnés de replay y baseline sembrado
- **id**: T3
- **fase**: 1944
- **depends_on**: [T1]
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica/
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGeneradorHorariosReplayTests.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCorrida105Fixture.cs
- **description**: Construir un contexto determinista representativo de la Corrida #105 (35 aulas, 102 asignaturas, 199 relaciones, 410 genes, 32 bloques, pool de 220 docentes) o cargar fixture serializado en memoria; ejecutar un número controlado de generaciones con semilla fija; registrar la trayectoria de fitness por generación, desglose de violaciones por regla, hash canónico del horario final y volúmenes esperados (410 genes, 2.228 sesiones, 8.912 h, 199 vinculaciones, 336 vacantes, HARD=0).
- **validation**: Ejecución determinista 100 % repetible; el baseline sembrado coincide con la salida del código base actual; tiempo de ejecución acotado para CI/tests.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgGeneradorHorariosReplayTests"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:

### T4: Subetapas de reparación y sellado
- **id**: T4
- **fase**: 1945
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCalendarioConflictos.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPerfilSubetapasTests.cs
- **description**: Implementar instrumentación ambiental de bajo overhead para medir subetapas de sellado: `DisponibleGrupo`, calendarios creados, sesiones proyectadas, candidatos docentes evaluados, invocaciones a `Sellar` y tiempos de sellado; diferenciar explícitamente `TiempoCruceBruto`, `TiempoSelladoEnCruce`, `TiempoCruceNeto`, `TiempoMutacionBruto`, `TiempoSelladoEnMutacion`, `TiempoMutacionNeto` y `TiempoSelladoTotal` para prevenir doble conteo.
- **validation**: Contadores > 0 en una ejecución corta; la suma de subetapas no excede el tiempo total del operador; cero alteraciones funcionales.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgPerfilSubetapasTests|FullyQualifiedName~AgAsignadorTerceroTests"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:

### T5: Cachés de higiene en bucles calientes
- **id**: T5
- **fase**: 1946
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgBloqueHorario.cs
  - ConversorPDF.XLL/Academic/Application/Services/CalculadoraViaticaSesion.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgProyeccionHigieneTests.cs
- **description**: Precalcular estructuras invariantes en `AgContexto`: `TimeSpan` por bloque, `HorarioTexto`, diccionarios `ProgramasPorId` y mapeo de unidad por aula; normalizar cadenas de ciudad una sola vez por gen; eliminar closures/LINQ en bucles críticos cuando `SesionesConservadas` está vacío; garantizar que no se alteren fechas efectivas ni ocupaciones.
- **validation**: Mismas sesiones, fechas y clasificaciones que antes en fixtures de prueba; reducción medible en asignación de memoria y tiempo; pruebas existentes pasando.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgProyeccionHigieneTests|FullyQualifiedName~AgExpansorSesionesTests"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:

### T6: Quick wins de reparación y validación
- **id**: T6
- **fase**: 1947
- **depends_on**: [T4]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorAptitud.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgQuickWinsTests.cs
- **description**: Modificar `Mutar` para retornar `bool cambios` y omitir `Sellar` si no hubo mutación efectiva; evitar reparación de hijos en cruce si OX no modificó asignaciones; incorporar bandera `incluirDetalleJson` (`false` durante búsqueda generacional, `true` únicamente en evaluación del campeón final); no reevaluar clones exactos; preservar rigurosamente el orden de llamadas a `Random`.
- **validation**: Trayectoria sembrada idéntica en corrida de verificación; reducción en conteo de serializaciones JSON y sellados redundantes; pruebas en verde.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgQuickWinsTests|FullyQualifiedName~AgOperadorCruceTests|FullyQualifiedName~AgOperadorMutacionTests"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:

### T7: Proyección ligera para búsqueda
- **id**: T7
- **fase**: 1948
- **depends_on**: [T5]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCalendarioConflictos.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCalendarioConflictos.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgProyeccionLigeraTests.cs
- **description**: Desacoplar la proyección de fechas e intervalos de ocupación (`Ocupacion`) de la instanciación pesada de objetos `AgSesion`; implementar constructor ligero anidado en `AgCalendarioConflictos.cs` para el camino crítico de validación y sellado; mantener intacto el método `Expandir` completo para la materialización final de horarios.
- **validation**: Equivalencia estricta: fechas, intervalos de tiempo, cardinalidad y detección de cruces idénticos a la proyección completa; materialización final inalterada; mejora medible de rendimiento.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgProyeccionLigeraTests|FullyQualifiedName~AgCalendarioConflictosTests"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:

### T8: Índice por individuo y reparación sucia
- **id**: T8
- **fase**: 1949
- **depends_on**: [T4, T6]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgIndiceIndividuoTests.cs
- **description**: Introducir clase interna `AgProyeccionIndividuo` dentro de `AgAsignadorTercero.cs` para mantener índices de ocupación por gen, grupos y docentes; implementar reparación incremental acotada al conjunto sucio conservador (genes modificados por OX, genes con el mismo `ClaveGrupo`, genes del mismo docente, genes que comparten cohorte/aula/fecha y genes con cambios de bloque/fechas); preservar el orden canónico de reparación y consumo de `Random`; en modo Debug, verificar equivalencia contra reparación completa abortando si difieren.
- **validation**: En fixtures deterministas, la reparación incremental y la reparación completa producen exactamente las mismas decisiones docentes y el mismo fitness; sin inconsistencias de caché; ganancia sustancial de tiempo en `DisponibleGrupo`.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgIndiceIndividuoTests|FullyQualifiedName~AgValidadorReglasTests"
- **evidence**: _evidence/20261117-t8/summary.md
- **status**: done
- **log**: Introducida clase interna AgProyeccionIndividuo en AgAsignadorTercero.cs para indexación O(1) de ocupaciones y resolución de conflictos. Reparación con conjunto sucio conservador en AgOperadorCruce.cs. 40 tests verdes en AgIndiceIndividuoTests y AgValidadorReglasTests.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgIndiceIndividuoTests.cs
  - _evidence/20261117-t8/summary.md

### T9: Huella y caché de aptitud
- **id**: T9
- **fase**: 1950
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgIndividuo.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgIndividuo.cs
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgHuellaAptitudTests.cs
- **description**: Implementar cálculo de huella estructural rápida para individuos; almacenar en caché generacional y reutilizar fitness/violaciones únicamente para individuos con idéntica huella estructural; registrar contador de duplicados; garantizar que la evaluación final del campeón genere siempre el desglose completo.
- **validation**: Fitness recuperado de la caché es idéntico bit a bit al fitness recalculado; porcentaje de aciertos en caché > 0 conforme la población converge; sin mutación de la trayectoria generacional.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgHuellaAptitudTests"
- **evidence**: _evidence/20261117-t9/summary.md
- **status**: done
- **log**: Implementada huella estructural FNV-1a en AgGen y AgIndividuo. Caché generacional de aptitud en AgGeneradorHorarios con evaluación forzada con detalle JSON para el campeón final. 9 tests en AgHuellaAptitudTests y tests de replay en verde (2 s).
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgIndividuo.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgHuellaAptitudTests.cs
  - _evidence/20261117-t9/summary.md

### T10: Benchmark, equivalencia y certificación
- **id**: T10
- **fase**: 1951
- **depends_on**: [T3, T6, T7, T8, T9]
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica/
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRendimientoCertificacionTests.cs
  - docs/articles/changes/2026-11-17-rendimiento-ag-corrida105.md
- **description**: Ejecutar suite de certificación comparativa entre baseline sembrado y versión optimizada; verificar paridad estricta en trayectoria de fitness, violaciones por regla (≤ 30 blandas, HARD=0) y hash del horario materializado; validar invariantes de volumen: 410 genes, 2.228 sesiones, 8.912 h, 199 vinculaciones, 336 vacantes; certificar métricas de rendimiento con `Stopwatch`: transición de estado estacionario ≤ 1,0 s (objetivo ~0,65 s) y tiempo total ≤ 72 s (objetivo ≤ 45 s con inicio optimizado a ~7 s); generar informe de cierre en `docs/articles/changes/`.
- **validation**: Todos los criterios de equivalencia, volumen y rendimiento se cumplen satisfactoriamente; cero regresiones en la suite global de pruebas.
- **verification**: dotnet build .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug ; dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --logger "trx;LogFileName=gate-alternativo.trx"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:

---

## Riesgos y Mitigaciones

| # | Riesgo | Severidad | Mitigación |
|---|---|---|---|
| 1 | Divergencia en trayectoria por consumo de RNG en operadores optimizados | Alta | T1 fija flujos deterministas; T6 y T8 preservan rigurosamente el orden de llamadas a `Random`. Comparación en Debug contra ejecución completa. |
| 2 | Inconsistencia en reparación incremental (conjunto sucio insuficiente) | Alta | El conjunto sucio en T8 es conservador (incluye genes con mismo grupo, docente, cohorte, aula y fecha). En Debug se corre validación cruzada incremental vs completa. |
| 3 | Proyección ligera diverge de `AgSesion` completa | Media | T7 valida equivalencia estricta en pruebas unitarias contra fechas, horas y conflictos de la proyección pesada. La materialización final sigue usando `Expandir` completo. |
| 4 | Contención en `ConversorPDF.XLL.csproj` (Lease L0) | Crítica | No se añaden archivos nuevos a `ConversorPDF.XLL`. Todos los tipos nuevos son clases internas anidadas (`AgPerfilGeneracion`, `AgProyeccionIndividuo`). `ConversorPDF.Tests` usa SDK-style con globbing automático. |
| 5 | Caché de aptitud devuelve fitness obsoleto ante mutaciones parciales | Media | T9 calcula la huella estructural sobre todos los campos funcionales de los genes del individuo. |

---

## Self-Review (Checklist de 13 puntos de swarm-planner)

| # | Punto | Estado | Observación |
|---|---|---|---|
| 1 | Dependencias completas: ningún `depends_on` omitido | OK | Grafo cerrado: T1 raíz; T2,T3 dependen de T1; T4,T5 dependen de T2; T6 depende de T4; T7 depende de T5; T8 depende de T4,T6; T9 depende de T2; T10 depende de T3,T6,T7,T8,T9. |
| 2 | Orden ejecutable: cada tarea compila al terminar | OK | Tareas ordenadas topológicamente; cada una cuenta con pruebas focalizadas. |
| 3 | Ninguna tarea mezcla L0 sin rol `integrator` | OK | Ninguna tarea toca archivos L0. `ConversorPDF.XLL.csproj` intacto; clases nuevas son internas anidadas. |
| 4 | `files_lock` disjuntos por ola | OK | Verificados pares de Wave 2 (T2,T3), Wave 3 (T4,T5), Wave 4 (T6,T7) y Wave 5 (T8,T9). Cero solapamientos. |
| 5 | Cada `verification` es un comando que existe en el repo | OK | Se usa el gate alternativo por lotes documentado y verificado en preflight (`dotnet test ...`). |
| 6 | Cada `validation` es falsable | OK | Métricas cuantitativas, paridad estricta de hash y asserts de equivalencia. |
| 7 | El RED compila (esqueleto `integrator-first` previsto) | OK | Clases internas en archivos de producción existentes; proyecto de test es SDK-style. |
| 8 | Migración y `APP_VERSION` asignadas solo por `integrator` | OK | No se requieren migraciones SQLite ni bumps de versión para esta ronda de optimización. |
| 9 | Ninguna tarea cruza más de un slice | OK | Trabajo circunscrito 100 % al slice `Academic`. |
| 10 | El context pack usado tiene ≤ 400 líneas | OK | Sin desbordamiento de contexto. |
| 11 | Criterio de aceptación global definido | OK | 5 criterios globales especificados (equivalencia, volumen, hash, presupuesto y cero regresiones). |
| 12 | Riesgos con mitigación | OK | 5 riesgos técnicos analizados con su respectiva estrategia de mitigación. |
| 13 | `numbering:` en envelope y fila `reserved` en `REGISTRY.md` | OK | Bloque `numbering:` presente; fila `R-112` reservada en `REGISTRY.md`. |

---

## Auditoría de Plan-Reviewer (9 Reglas Duras)

| Regla | Auditoría | Veredicto |
|---|---|---|
| 1. Execution Envelope | Bloque completo con workspace, preflight, workspace_guard, baseline, verify, build_gate serializado, hot_paths_L0 y numbering. | CUMPLE |
| 2. Lease L0 Protegido | Ningún archivo L0 (`.csproj`, `App.cs`, migraciones, `RibbonController`, IoC) presente en `files_lock`. | CUMPLE |
| 3. Disyunción de `files_lock` | Olas 2, 3, 4 y 5 tienen listas disjuntas sin prefijos compartidos. | CUMPLE |
| 4. Fase RED Compilable | Clases internas en archivos ya referenciados; `ConversorPDF.Tests.csproj` incluye archivos `.cs` automáticamente. | CUMPLE |
| 5. Comandos Falsables | Comandos `dotnet test` con filtros específicos y gate con TRX. | CUMPLE |
| 6. Ubicación del Plan | Guardado en `.agents/plans/2026-11-17-rendimiento-app-corrida105-plan.md`. | CUMPLE |
| 7. Workspace Verificado | Preflight ejecutado satisfactoriamente con `-AllowOneDrive` (`PREFLIGHT_OK`). | CUMPLE |
| 8. Baseline Medido | Baseline oficial 1446 documentado; R-109 reporta Full 1648; medición obligatoria antes de T1. | CUMPLE |
| 9. Gate Serial | Un único gate a la vez; build_gate serializado declarado en envelope. | CUMPLE |

**Veredicto Final Plan-Reviewer**: **APROBADO**
