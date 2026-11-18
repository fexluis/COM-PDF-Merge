# Plan: Optimización de Rendimiento AG — Corrida #105 (R-112, Ronda 433–438)

**Generado**: 2026-11-17 · **Autor**: Swarm-Planner · **Estado**: reserved

## Execution envelope
- branch: rendimiento-app
- workspace: C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace "C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan ".agents/plans/2026-11-17-rendimiento-app-corrida105-plan.md" -AllowOneDrive
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: 1446; último reporte ledger R-109: Full 1648; build-verifier medirá workspace antes de T1)
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

## Dependency Graph

```text
T1 ── T2 ── T4 ── T6 ── T8 ─┐
  │      │                 ├── T10
  │      └── T5 ── T7 ─────┤
  └── T3 ──────────────────┤
           T9 ─────────────┘
```

---

## Waves Summary

| Wave | Ronda | Tareas | Precondición | Leases |
|---|---|---|---|---|
| Wave 1 | 433 | T1 | Inmediato | AgOperadorCruce, AgOperadorMutacion, AgOperadorSeleccion, AgFabricaIndividuos, AgAsignadorTercero |
| Wave 2 | 434 | T2, T3 | T1 completado | Disjuntos: [AgGeneradorHorarios, AgResultadoCorrida, AgBitacora] vs [Tests/AgGeneradorHorariosReplayTests, AgCorrida105Fixture] |
| Wave 3 | 435 | T4, T5 | T2 completado | Disjuntos: [AgAsignadorTercero, AgCalendarioConflictos, AgCoherenciaGrupo] vs [AgContexto, AgBloqueHorario, CalculadoraViaticaSesion, AgExpansorSesiones] |
| Wave 4 | 436 | T6, T7 | T4 y T5 completados | Disjuntos: [AgOperadorCruce, AgOperadorMutacion, AgValidadorReglas, AgEvaluadorAptitud] vs [AgCalendarioConflictos, AgExpansorSesiones] |
| Wave 5 | 437 | T8, T9 | T6 y T2 completados | Disjuntos: [AgAsignadorTercero, AgOperadorCruce, AgValidadorReglas] vs [AgIndividuo, AgGen, AgGeneradorHorarios] |
| Wave 6 | 438 | T10 | T3, T6, T7, T8, T9 completados | AgRendimientoCertificacionTests, docs/articles/changes/ |

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
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResultadoCorrida.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgBitacora.cs
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
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
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
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgContexto.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgBloqueHorario.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/CalculadoraViaticaSesion.cs
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
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
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
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
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
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgIndiceIndividuoTests.cs
- **description**: Introducir clase interna `AgProyeccionIndividuo` dentro de `AgAsignadorTercero.cs` para mantener índices de ocupación por gen, grupos y docentes; implementar reparación incremental acotada al conjunto sucio conservador (genes modificados por OX, genes con el mismo `ClaveGrupo`, genes del mismo docente, genes que comparten cohorte/aula/fecha y genes con cambios de bloque/fechas); preservar el orden canónico de reparación y consumo de `Random`; en modo Debug, verificar equivalencia contra reparación completa abortando si difieren.
- **validation**: En fixtures deterministas, la reparación incremental y la reparación completa producen exactamente las mismas decisiones docentes y el mismo fitness; sin inconsistencias de caché; ganancia sustancial de tiempo en `DisponibleGrupo`.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgIndiceIndividuoTests|FullyQualifiedName~AgValidadorReglasTests"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:

### T9: Huella y caché de aptitud
- **id**: T9
- **fase**: 1950
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgIndividuo.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGen.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgHuellaAptitudTests.cs
- **description**: Implementar cálculo de huella estructural rápida para individuos; almacenar en caché generacional y reutilizar fitness/violaciones únicamente para individuos con idéntica huella estructural; registrar contador de duplicados; garantizar que la evaluación final del campeón genere siempre el desglose completo.
- **validation**: Fitness recuperado de la caché es idéntico bit a bit al fitness recalculado; porcentaje de aciertos en caché > 0 conforme la población converge; sin mutación de la trayectoria generacional.
- **verification**: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~AgHuellaAptitudTests"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:

### T10: Benchmark, equivalencia y certificación
- **id**: T10
- **fase**: 1951
- **depends_on**: [T3, T6, T7, T8, T9]
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica/
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRendimientoCertificacionTests.cs
  - docs/articles/changes/2026-11-17-rendimiento-ag-corrida105.md
- **description**: Ejecutar suite de certificación comparativa entre baseline sembrado y versión optimizada; verificar paridad estricta en trayectoria de fitness, violaciones por regla (≤ 30 blandas, HARD=0) y hash del horario materializado; validar invariantes de volumen: 410 genes, 2.228 sesiones, 8.912 h, 199 vinculaciones, 336 vacantes; certificar métricas de rendimiento con `Stopwatch`: transición de estado estacionario ≤ 1,0 s (objetivo ~0,65 s) y tiempo total ≤ 72 s (objetivo ≤ 45 s con inicio optimizado); generar informe de cierre en `docs/articles/changes/`.
- **validation**: Todos los criterios de equivalencia, volumen y rendimiento se cumplen satisfactoriamente; cero regresiones en la suite global de pruebas.
- **verification**: dotnet build .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug ; dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --logger "trx;LogFileName=gate-alternativo.trx"
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**: 
- **files edited/created**:
