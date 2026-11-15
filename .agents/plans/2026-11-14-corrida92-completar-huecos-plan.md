# Plan: Corrida 92 — Completar Huecos y Unificar Evaluación y Generación
**Generado**: 2026-11-14 · **Autor**: Antigravity/Swarm-Planner · **Estado**: executing

## Execution envelope
- branch: codex/huellas-escalado-20261114
- workspace: C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -AllowOneDrive -Workspace "C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan ".agents/plans/2026-11-14-corrida92-completar-huecos-plan.md"
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: 1359)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado
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
  - Intento de push, merge o publicación a remoto
  - Plan en estado irrecuperable (más de 3 rondas de fix fallidas)
- ledger_owner: swarm-parallel-task
- no_push: true
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: R-078
  - ronda_base: 384
  - fase_base: 1798
  - waves: 8
  - tasks: 8

## Dependency graph
```text
Wave 1: T1 (Fase 1798) [Regression Fixtures & Tests RED en suite dedicada]
          │
          ▼
Wave 2: T2 (Fase 1799) [Validador Canónico ValidarFechado + Tests Unitarios]
          │
          ▼
Wave 3: T3 (Fase 1800) [CompletarReferencia & Desacople LimiteMayo + Tests]
          │
          ▼
Wave 4: T4 (Fase 1801) [Selección Común y Unificada ConstruirInterno -> GREEN T1]
          │
          ▼
Wave 5: T5 (Fase 1802) [Consumo de Plan Fechado en Fabrica & Coherencia]
          │
          ▼
Wave 6: T6 (Fase 1803) [Persistencia Staging & Trazabilidad Reproducible]
          │
          ▼
Wave 7: T7 (Fase 1804) [Pruebas de Aceptación E2E - 12 Criterios]
          │
          ▼
Wave 8: T8 (Fase 1805) [Certificación Gate Full & Calibración Baseline]
```

## Parallel execution groups
| Wave | Ronda | Tasks | Can start when |
|---|---|---|---|
| Wave 1 | Ronda 384 | T1 | Preflight exitoso y repositorio limpio |
| Wave 2 | Ronda 385 | T2 | T1 completada |
| Wave 3 | Ronda 386 | T3 | T2 completada |
| Wave 4 | Ronda 387 | T4 | T3 completada |
| Wave 5 | Ronda 388 | T5 | T4 completada |
| Wave 6 | Ronda 389 | T6 | T5 completada |
| Wave 7 | Ronda 390 | T7 | T6 completada |
| Wave 8 | Ronda 391 | T8 | T7 completada |

## Tasks

### T1: Fixtures y Pruebas de Regresión Cali III, IV y Cartago X-A (Fase RED)
- **id**: T1
- **fase**: 1798
- **depends_on**: []
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasRegresionTests.cs
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasRegresionTests.cs
- **description**: Crear la suite aislada `AgPlanificadorRanurasRegresionTests.cs` para evitar contaminar la suite base de `AgPlanificadorRanurasTests`. Implementar fixtures y pruebas de regresión pura para Cali III (32, 48, 48, 48, 32, 48 h; patrones 1,2 / 3,4 / 5 / 1,2 / 3,4; última AUTOMATICO), Cali IV (32, 48, 48, 48, 48, 48 h) y Cartago X-A (32, 32, 64, 112 h; ventana JUE-DOM) en el período 05/02/2027 a 30/05/2027 con festivos canónicos de Semana Santa (22–28/03/2027) y 10/05/2027. Exigir que las sesiones se completen antes de la fecha límite (13/05 para III, 19/05 para IV) y verificar que `Construir` y `ConstruirEvaluacion` arrojen idéntico resultado sin ampliar el plazo.
- **validation**: La prueba `Mixta_ConservaRestriccionesYCompletaAntesDelLimite` compila y falla (RED) por forzamiento actual de la referencia en `AgPlanificadorOleadas` cuando hay patrones.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "ConversorPDF.Tests.Academic.Application.Services.Genetica.AgPlanificadorRanurasRegresionTests"
- **evidence**: _evidence/20261114-211320/summary.md
- **status**: done
- **log**: "RED verificado: 4 tests implementados y fallando exactamente por forzamiento de referencia con patrones en AgPlanificadorOleadas. Evidencia _evidence/20261114-211320/summary.md."
- **files edited/created**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasRegresionTests.cs

### T2: Validador Canónico de Planes Fechados ValidarFechado y Tests Unitarios
- **id**: T2
- **fase**: 1799
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorRanuras.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorRanuras.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasTests.cs
- **description**: Añadir el método `internal static Result ValidarFechado(...)` en `AgPlanificadorRanuras.cs`. Valida de forma exhaustiva: plan no nulo (`I5_PLAN_NULO`), demandas requeridas vs sesiones ubicadas (`I5_PLAN_DEMANDA_INCONSISTENTE`), validez de patrones y días restringidos (`I1_PATRON_INVALIDO`), coincidencia de horas de ranura, límites de sub-ventana por materia (`I5_PLAN_SESION_INVALIDA`), ausencia de colisiones y existencia en calendario (`I5_PLAN_FECHA_O_COLISION`), y consistencia exacta entre materias con déficit y la lista de `Infactibles` (`I5_PLAN_PENDIENTES_INCONSISTENTES`). El validador admite planes parciales legítimos (`Sesiones.Count <= requerida`) siempre que toda materia incompleta figure en `Infactibles`, distinguiendo déficit admisible de inconsistencia estructural. Escribir pruebas unitarias en `AgPlanificadorRanurasTests.cs` verificando casos aprobados y cada código de fallo.
- **validation**: `ValidarFechado` expone `Result`, compila limpiamente y todas las pruebas unitarias nuevas de validación en `AgPlanificadorRanurasTests.cs` pasan en GREEN sin interferir con T1.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "ConversorPDF.Tests.Academic.Application.Services.Genetica.AgPlanificadorRanurasTests"
- **evidence**: _evidence/20261114-212158/summary.md
- **status**: done
- **log**: "ValidarFechado implementado con validación canónica de consistencia, colisiones, patrones y soporte de déficit estructurado. 24/24 pruebas unitarias en GREEN. Evidencia _evidence/20261114-212158/summary.md. Aprobado por task-reviewer."
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorRanuras.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasTests.cs

### T3: Ocupación de Huecos por Sesión CompletarReferencia y Desacople de LimiteMayo
- **id**: T3
- **fase**: 1800
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorRanuras.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorRanuras.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasTests.cs
- **description**: Implementar `internal static Result<AgPlanOleadasResultado> CompletarReferencia(...)` en `AgPlanificadorRanuras.cs`. Conserva exactamente las sesiones de materias restringidas completas de la referencia validada con `ValidarFechado`. Para las materias restantes, aplica una heurística Most-Constrained-First (ordenando primero aquellas con restricciones de días `diasPatron.Count > 0` por menor grado de libertad, y al final las `AUTOMATICO`), ubicando sesión por sesión en los huecos libres del calendario común sin segregación por carril. En la búsqueda acotada de `AgPlanificadorRanuras.Construir`, eliminar la constante fija `limiteMayo = 30/05/2027` garantizando que el bucle de particiones continúe buscando compactaciones con menor `mejorFin` sin corte prematuro en la primera partición completa. Añadir tests unitarios específicos para `CompletarReferencia` en `AgPlanificadorRanurasTests.cs`.
- **validation**: `CompletarReferencia` rescata y completa materias aprovechando huecos entre bandas; el test existente `Construir_WhenReferenciaCabeEnVentanaAmpliada_ShouldBuscarFinMasTemprano` sigue pasando sin depender de `limiteMayo`.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "ConversorPDF.Tests.Academic.Application.Services.Genetica.AgPlanificadorRanurasTests"
- **evidence**: _evidence/20261114-212655/summary.md
- **status**: done
- **log**: "CompletarReferencia implementado con heurística Most-Constrained-First y validación canónica. Desacoplado limiteMayo de AgPlanificadorRanuras.Construir sin truncar compactación. 27/27 tests en GREEN. Evidencia _evidence/20261114-212655/summary.md. Aprobado por task-reviewer."
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorRanuras.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasTests.cs

### T4: Selección Común y Unificada de Planes en AgPlanificadorOleadas
- **id**: T4
- **fase**: 1801
- **depends_on**: [T3]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasRegresionTests.cs
- **description**: Reemplazar la selección en `AgPlanificadorOleadas.ConstruirInterno`. Eliminar la regla que forzaba referencia en presencia de patrones (`generacion_con_patrones`) y el requisito de 7 días de diferencia. Ensamblar lista de alternativas (`referenciaValida`, `completado`, y búsqueda acotada como rescate si `completado` es parcial), filtrar aquellas que preserven materias previamente completas, priorizar planes 100% completos, ordenar determinísticamente por mayor cantidad de sesiones ubicadas, menor número de infactibles y menor fecha fin máxima. Asegurar que evaluación y generación elijan exactamente el mismo plan completo ante entradas idénticas.
- **validation**: `Construir` y `ConstruirEvaluacion` seleccionan el mismo plan; la prueba `Mixta_ConservaRestriccionesYCompletaAntesDelLimite` en `AgPlanificadorRanurasRegresionTests.cs` (T1) pasa de RED a GREEN sin ampliar fechas.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "ConversorPDF.Tests.Academic.Application.Services.Genetica.AgPlanificadorRanurasRegresionTests"
- **evidence**: _evidence/20261114-213030/summary.md
- **status**: done
- **log**: "Selección unificada implementada en ConstruirInterno integrando referenciaValida, completado y búsqueda acotada de rescate con ValidarFechado. Eliminado forzamiento de patrones y umbral de 7 días. 4/4 pruebas de regresión en GREEN. Evidencia _evidence/20261114-213030/summary.md. Aprobado por task-reviewer."
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs

### T5: Consumo Fiel del Plan Fechado en Fabrica y Coherencia de Grupo
- **id**: T5
- **fase**: 1802
- **depends_on**: [T4]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs
- **description**: Ajustar el consumo de planes en `AgFabricaIndividuos` y `AgCoherenciaGrupo`. En la fábrica, asegurar coincidencia de bloque horario (`Activo == 1 && DiaSemanaNumero == ranura.Dia && Math.Abs(DuracionHoras - ranura.Horas) < 0.001 && Franja == ranura.Franja`), manteniendo fallback defensivo si el catálogo no dispone de la franja exacta, y fijar `FechasGen` y `HorasGen = FechasGen.Count * horasEfectivas` por ranura sin truncamiento. En `AgCoherenciaGrupo.Sellar`, respetar la distribución multi-día del plan cuando la suma total de horas cubre el esperado (`Math.Abs(suma - esperado) < 0.001`), evitando redistribuciones uniformes (`12/N`) incluso si algún gen tiene 0 horas por concentración en ranuras efectivas.
- **validation**: Los genes del individuo conservan las fechas exactas del planificador y sus horas reflejan la cantidad real de sesiones; no se introducen solapes ni mutaciones espurias en `Sellar`.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "ConversorPDF.Tests.Academic.Application.Services.Genetica"
- **evidence**: _evidence/20261114-214422/summary.md
- **status**: done
- **log**: "Fábrica ajustada con coincidencia exacta y activa de bloques por franja. Coherencia de grupo Sellar blindada para preservar FechasGen y HorasGen de planes fechados sin redistribución forzada. 50/50 tests en GREEN. Evidencia _evidence/20261114-214422/summary.md. Aprobado por task-reviewer."
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs

### T6: Persistencia en Staging y Trazabilidad Reproducible
- **id**: T6
- **fase**: 1803
- **depends_on**: [T5]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs
- **description**: Enriquecer `AgExpansorSesiones` y `AgMaterializadorSesiones` para garantizar que la expansión y materialización en las tablas de staging (`ag_sesiones`) utilice fielmente `FechasGen` sin recurrir a generadores legacy cuando el plan es factible. Enriquecer los logs con métricas de decisión (`ELEGIR`, candidatos evaluados, fechas límite y pendientes) y asegurar que los metadatos de corrida permitan reproducir el plan de slots exacto.
- **validation**: Las sesiones en staging coinciden una a una con el plan fechado; la trazabilidad en log permite reconstruir la alternativa seleccionada.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "ConversorPDF.Tests.Academic.Application.Services.Genetica"
- **evidence**: _evidence/20261114-215342/summary.md
- **status**: done
- **log**: "AgExpansorSesiones enriquecido para adaptar fInicio..fFin a FechasGen sin truncamiento. AgMaterializadorSesiones sincroniza fechas extremas en ag_vinculaciones y enriquece logs con métricas de decisión. 63/63 tests en GREEN. Evidencia _evidence/20261114-215342/summary.md. Aprobado por task-reviewer."
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs

### T7: Suite de Aceptación End-to-End (12 Criterios de Aceptación)
- **id**: T7
- **fase**: 1804
- **depends_on**: [T6]
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasAceptacionTests.cs
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasAceptacionTests.cs
- **description**: Implementar y ejecutar las 12 comprobaciones exhaustivas de aceptación descritas en T7: 1) Cali III y IV sin sesiones post-30/05 y horas exactas; 2) Paridad estricta entre generación y evaluación en mezclas LEGADO/AUTOMATICO; 3) Asignaturas que aprovechan múltiples días/bandas sin exclusividad artificial; 4) Compartición de calendario sin colisiones; 5) Invariantes de días permitidos para materias no reclasificadas; 6) Selección de plan más temprano aunque la diferencia sea < 7 días; 7) Preservación de materias previamente completas en planes parciales; 8) Rechazo controlado de anomalías de calendario/sesión; 9) Soporte para semanas incompletas, sábado M/T y unidades de 8h; 10) Independencia de años o constantes rígidas; 11) Fidelidad de fechas entre fábrica, coherencia y staging; 12) Ausencia de regresiones en escalado de ventanas 4->5 y 4->5->6.
- **validation**: 12/12 pruebas de aceptación pasan en GREEN con 100% de éxito.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "ConversorPDF.Tests.Academic.Application.Services.Genetica.AgPlanificadorRanurasAceptacionTests"
- **evidence**: _evidence/pending/summary.md
- **status**: pending
- **log**: ""
- **files edited/created**: []

### T8: Certificación Gate Full y Calibración de Baseline
- **id**: T8
- **fase**: 1805
- **depends_on**: [T7]
- **location**: tests/baseline/test_baseline.txt
- **files_lock**:
  - tests/baseline/test_baseline.txt
- **description**: Ejecutar la certificación global en modo Full con `scripts/verify.ps1 -Mode Full`, validar que no exista ninguna regresión en los 1359+ tests de la suite y calibrar el nuevo baseline con los tests adicionales creados.
- **validation**: `scripts/verify.ps1 -Mode Full` devuelve código de salida 0; 0 fallos, 0 errores, total ejecutado >= 1359.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- **evidence**: _evidence/pending/summary.md
- **status**: pending
- **log**: ""
- **files edited/created**: []

## Context pack
- Slice: Academic / Genetica (Planificador de Ranuras, Oleadas, Fábrica y Materialización).
- Clases principales: `AgPlanificadorRanuras`, `AgPlanificadorOleadas`, `AgFabricaIndividuos`, `AgCoherenciaGrupo`, `AgExpansorSesiones`, `AgMaterializadorSesiones`.
- Convenciones: .NET Framework 4.8, xUnit 2.7, FluentAssertions 6.12, Result pattern (`CSharpFunctionalExtensions`).
- Prohibido: C# 12+ (colecciones `[]`), MediatR, llamadas bloqueantes async void.

## Self-Review: Checklist de 13 Puntos
| # | Punto | Cumple | Justificación |
|---|---|---|---|
| 1 | Dependencias completas: ningún `depends_on` omitido | SÍ | Cadena lineal T1->T2->T3->T4->T5->T6->T7->T8 verificada |
| 2 | Orden ejecutable: cada tarea compila al terminar | SÍ | Cada tarea deja el árbol de compilación en estado verde |
| 3 | Ninguna tarea mezcla L0 sin rol `integrator` | SÍ | Ninguna tarea toca archivos de Lease L0 |
| 4 | `files_lock` disjuntos por ola | SÍ | Leases estrictamente delimitados y separados por olas secuenciales |
| 5 | Cada `verification` es un comando que existe en el repo | SÍ | `scripts/verify.ps1` con filtros exactos de xUnit a nivel de clase |
| 6 | Cada `validation` es falsable | SÍ | Criterios cuantitativos y observables |
| 7 | El RED compila (esqueleto `integrator-first` previsto) | SÍ | Tipos existentes, métodos agregados de forma compatible |
| 8 | Migración y `APP_VERSION` asignadas solo por `integrator` | SÍ | No aplica (no hay migraciones de base de datos) |
| 9 | Ninguna tarea cruza más de un slice | SÍ | Todas acotadas estrictamente a `Academic/Genetica` |
| 10 | El context pack usado tiene ≤ 400 líneas | SÍ | Módulo de genética acotado a planificadores y expansores |
| 11 | Criterio de aceptación global definido | SÍ | 12 criterios de aceptación de T7 y Gate Full en T8 |
| 12 | Riesgos con mitigación | SÍ | Documentados en sección de riesgos |
| 13 | `numbering:` en el envelope y fila `reserved` añadida a `REGISTRY.md` | SÍ | Fila R-078 reservada: Rondas 384–391, Fases 1798–1805 |

## Riesgos y Mitigaciones
1. **Riesgo**: Que `CompletarReferencia` genere colisiones horarias entre materias que comparten la misma franja horaria.
   - **Mitigación**: `ValidarFechado` valida de forma estricta la ocupación sesión por sesión en el calendario discreto con índice único antes de aceptar cualquier plan.
2. **Riesgo**: Que `AgCoherenciaGrupo.Sellar` sobreescriba `FechasGen` con fechas recalculadas por reparto uniforme de horas cuando la suma ya es igual al esperado.
   - **Mitigación**: T5 preserva `FechasGen` y `HorasGen` si `Math.Abs(suma - esperado) < 0.001` sin importar si algún gen individual tiene 0 horas, evitando redistribución forzada.
3. **Riesgo**: Que la eliminación de `limiteMayo` interrumpa prematuramente la compactación de horizontes en ventanas extendidas.
   - **Mitigación**: T3 estipula que la condición de parada evalúe compactación continua sin corte anticipado en la primera partición factible.
4. **Riesgo**: Regresiones en la suite global de 1359 pruebas.
   - **Mitigación**: T8 ejecuta `scripts/verify.ps1 -Mode Full` garantizando que ninguna funcionalidad existente se degrade.
