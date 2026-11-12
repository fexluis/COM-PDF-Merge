# Plan: aprovechamiento dinámico de ranuras semanales (4 h → 8 h)
**Generado**: 2026-11-12 · **Autor**: Cline · **Estado**: draft (sin reserva de numeración)

> Análisis, hallazgos del libro y restricciones verificadas: `docs/2026-11-12-plan-ranuras-dinamicas-4-a-8-horas.md`.

## Execution envelope

- branch: `refactor/academic-modularization`
- workspace: copia local verificada fuera de OneDrive
- preflight: `powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ws> -Plan .agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: `tests/baseline/test_baseline.txt` (TOTAL remedido en el workspace tras el commit de R-047; nunca inventado)
- verify: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- build_gate: serializado (Excel-DNA bloquea `bin/`)
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0:
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/**/Migrations/MigrationV*.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs`
- stop_conditions:
  - Acción destructiva no reversible sin respaldo
  - `git clean`/`git stash`/cambio de rama con archivos sin seguimiento
  - Intento de push, merge o publicación a remoto
  - Plan irrecuperable tras más de 3 rondas de fix fallidas
- no_push: true
- numbering:
  - registry: `.agents/plans/REGISTRY.md`
  - registry_row: **PENDIENTE** (Fase A)
  - ronda_base: **PENDIENTE**
  - fase_base: **PENDIENTE**
  - waves: 10
  - tasks: 10

> **Incidencia de numeración (declarada):** este plan no tiene fila en el registro. Las tareas se
> referencian como `T1..T10`; `- **fase**` se asigna al reservar. El registro quedó en
> `ronda 340 / fase 1694` tras las revisiones manuales R-048..R-056. No reutilizar rangos.

## Precondición verificada

HEAD auditado: `c0d986d`. Árbol de trabajo: limpio. `AgParticionCarriles.cs` y tests asociados están formalmente versionados y rastreados.
Build/Fast/Full: pendientes de ejecución en copia local fuera de OneDrive.

## Prohibiciones

- No fijar `HorasSemanales = 8` en un bloque de 4 h.
- No ampliar patrones estrictos ni añadir días a materias con patrón explícito.
- No eliminar `I5` ni cambiar horas del pensum ni acortar fechas de exportación.
- No ampliar la fecha límite por decisión del motor.
- No reconstruir semanalmente un plan explícito que falló al cargar ni degradar a fallback legado.
- No declarar mejora sin cobertura exacta y cero violaciones duras nuevas.

## Parallel execution groups

| Wave | Ronda | Tasks | Can start when |
|---|---|---|---|
| 1 | pendiente | T1 | Aprobación de preflight en copia local |
| 2 | pendiente | T2 | T1 completado |
| 3 | pendiente | T3 | T2 completado |
| 4 | pendiente | T4 | T3 completado |
| 5 | pendiente | T5 | T4 completado |
| 6 | pendiente | T6 | T5 completado |
| 7 | pendiente | T7 | T6 completado |
| 8 | pendiente | T8 | T7 completado |
| 9 | pendiente | T9 | T8 completado |
| 10 | pendiente | T10 | T9 completado |

Regla de oro: `ConversorPDF.XLL.csproj`, `SqliteHelper.cs` y `AgPlanificadorOleadas.cs` son L0; cualquier tarea que los toque se serializa.

## Tasks

### T1: Reproducibilidad de #62 y fixtures anónimos
- **id**: T1
- **fase**: pendiente de reserva
- **depends_on**: []
- **location**: `ConversorPDF.Tests/Academic/Domain/`, `_evidence/<stamp>/`
- **files_lock**: `ConversorPDF.Tests/Academic/Domain/AgRanurasDinamicasFixtures.cs` (nuevo), `docs/articles/changes/2026-11-12-ranuras-dinamicas-diagnostico.md` (nuevo)
- **description**: copiar `Libro2.xlsx` a `_evidence/<stamp>/`, registrar SHA256, extraer de copia de solo lectura de #62 período/programa/territorial, aula/cohorte, demanda por materia, ventana original y ampliada, bloques con horas exactas, festivos, patrón y procedencia, genes, sesiones, docentes, fijaciones y restricciones. Fixtures anónimos mínimos. Informe oferta/demanda por aula/cohorte y por subconjunto de ranuras compatibles. Identificar el aula de la corrida #62 e I5 de 62 ranuras.
- **validation**: hash del libro registrado; informe con déficit clasificado (físico / compatibilidad / fragmentación); fixtures sin datos personales; filas 8, 33 y 36 vinculadas por identificador, no por nombre.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgRanurasDinamicas"`
- **evidence**: `_evidence/<stamp>/summary.md`
- **status**: pending
- **log**: —
- **files edited/created**: pendiente de ejecución

### T2: Contrato de dominio y esquema del plan de sesiones
- **id**: T2
- **fase**: pendiente de reserva
- **depends_on**: [T1]
- **location**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/`, `.../Infrastructure/Persistence/Migrations/`
- **files_lock**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPlanSesiones.cs` (nuevo), `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgSesionPlan.cs` (nuevo), `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV80_AddFechasGenAgGen.cs` (nuevo), `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `ConversorPDF.Tests/Academic/Domain/AgPlanSesionesTests.cs` (nuevo), `ConversorPDF.Tests/Academic/Domain/AgGenTests.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/Migrations/MigrationV80Tests.cs` (nuevo)
- **description**: plan inmutable de sesiones `(asignatura, cohorte, fecha, bloque, inicio, fin, duración, clave de reserva)`; tipos `internal sealed`; `AgGen.PlanExplicito` (prohibido fallback a fechas semanales); `Clonar` copia el flag; constructor `AgGen(bool planExplicito = false)`; migración no aditiva (`IMigration` sin `IMigrationTolerante`), `Version => 80`, `Up(IDbConnection, IDbTransaction)` con guardas `PRAGMA table_info`; registro en `SqliteHelper.cs:380`.
- **validation**: migración aplicada en vault de prueba; reintento sin error de columna duplicada; fallo forzado aborta arranque; `PlanExplicito` sobrevive `Clonar`; test reflectivo de propiedades en `Clonar`; `ALTER TABLE` no toca CRDT (`ag_gen` no está en `CrdtTables`).
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~MigrationV80|FullyQualifiedName~AgGenTests"`
- **evidence**: `_evidence/<stamp>/summary.md`
- **status**: pending
- **log**: —
- **files edited/created**: pendiente de ejecución

### T3: Calendario único de ranuras fechadas
- **id**: T3
- **fase**: pendiente de reserva
- **depends_on**: [T2]
- **location**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/`
- **files_lock**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgRanuras.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgBandasDias.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgDiasVentana.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `ConversorPDF.Tests/Academic/Domain/AgRanurasTests.cs`, `ConversorPDF.Tests/Academic/Domain/AgBandasDiasTests.cs`, `ConversorPDF.Tests/Academic/Domain/AgDiasVentanaTests.cs`
- **description**: enumeración de ranuras fechadas con identidad por `(recurso, fecha, intervalo físico)`; festivos, límites inclusivos y deduplicación de catálogo; mañana ≠ tarde; conservar el ancla de semana (`fechaInicio`) y usar ISO solo para identidad/reportes; `TryParseCarril` como único parser de patrón (rechaza `|`, rangos, texto libre).
- **validation**: bloques distintos solapados colisionan; sin doble conteo por duplicados; `TryParseCarril("5")=true`, `TryParseCarril("VIERNES|4|08:00 A 12:00")=false`, `TryParseCarril("JUEVES–VIERNES")=false`; slices idénticos a los actuales para `fechaInicio` no lunes.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgRanuras|FullyQualifiedName~AgBandasDias|FullyQualifiedName~AgDiasVentana"`
- **evidence**: `_evidence/<stamp>/summary.md`
- **status**: pending
- **log**: —
- **files edited/created**: pendiente de ejecución

### T4: Factibilidad y diagnóstico I5 por subconjuntos compatibles
- **id**: T4
- **fase**: pendiente de reserva
- **depends_on**: [T3]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/`
- **files_lock**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgParticionCarriles.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgParticionCarrilesTests.cs`
- **description**: conservar exactamente los códigos vigentes (`I5_INFACTIBLE`, `I5_INFACTIBLE_OLEADA`, `I5_CAPACIDAD_CARRIL_PATRON`, `I5_SIN_CAPACIDAD`, `I5_SIN_SEMANAS_HABILES`, `I5_HOLGURA_BAJA`, `I5_HORAS_DESCONOCIDAS`); déficit en horas y sesiones; capacidad por subconjunto de ranuras accesibles; `hayPatronValido`, `ContarPatronesNoAlojables` y `ObtenerIndicesAlojables` usan `TryParseCarril`.
- **validation**: un patrón no-carril nunca cuenta como alojable ni activa búsqueda de partición; `64 > 62` reporta déficit 2 ranuras/8 h; patrón saturado, festivo, viernes inicial y semana incompleta con mensajes esperados; búsqueda sin solución ≠ infactibilidad matemática.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgValidadorVentanaTests|FullyQualifiedName~AgParticionCarrilesTests"`
- **evidence**: `_evidence/<stamp>/summary.md`
- **status**: pending
- **log**: —
- **files edited/created**: pendiente de ejecución

### T5: Planificador por reservas con restricciones duras
- **id**: T5
- **fase**: pendiente de reserva
- **depends_on**: [T4]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/`, `.../Domain/Models/Genetica/`
- **files_lock**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorReservas.cs` (nuevo), `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgParticionCarriles.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCohorte.cs` (nuevo), `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorReservasTests.cs` (nuevo), `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCohorteTests.cs` (nuevo)
- **description**: reservas por `(recurso, fecha, bloque)`; guardas R1, R2, R3, R4 (unicidad cohorte/día/franja + tope diario 1, sábado 2), R5, R6, R7, R8, R9, R10; `AgCohorte.De` como clave única (reutilizada por `AgValidadorReglas`); orden determinista por restricciones, orden curricular e Id; presupuesto y cancelación; sin solver externo; legado intacto hasta aceptación; `EsParticionFiel` estricto con `PatronesExplicitos` y `TryParseCarril` (superconjunto prohibido).
- **validation**: 4→8 cuando la segunda ranura es compatible; V estricto produce 6 semanas de 4 h; J ocupado semanas 1–2 y libre desde la 3 produce V,V,J+V,J+V para 24 h; ranura liberada reutilizada solo en día/franja distinto; dos bloques el mismo día hábil bloqueados; docente único por `ClaveGrupo`; candidato aceptado solo sin pérdida de cobertura ni violaciones nuevas.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgPlanificadorReservasTests|FullyQualifiedName~AgGateG1ParticionCarrilesTests"`
- **evidence**: `_evidence/<stamp>/summary.md`
- **status**: pending
- **log**: —
- **files edited/created**: pendiente de ejecución

### T6: Integración con genes y procedencia
- **id**: T6
- **fase**: pendiente de reserva
- **depends_on**: [T5]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/`
- **files_lock**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGeneradorHorariosTests.cs`
- **description**: un gen por `(asignatura, día, bloque)`; `FechasGen` exactas; `HorasGen = sesiones × unidad`; `PlanExplicito` solo en la ruta canónica (`new AgGen` en `:473` con fechas en `:524`), `false` en `:558` y `:700`; todos los genes de igual `ClaveGrupo` comparten `IdTercero`; clave de `_cachePlanes` incluye versión de estrategia; re-sellado de hijos antes de evaluar (`AgGeneradorHorarios:174-180`).
- **validation**: ninguna materia con 8 h en un bloque de 4 h; sin genes con `PlanExplicito` y `FechasGen` vacío al salir de la fábrica; hijos no evaluados con `HorasGen = 0`; caché no sirve planes de otra estrategia.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgFabricaIndividuosTests|FullyQualifiedName~AgGeneradorHorariosTests"`
- **evidence**: `_evidence/<stamp>/summary.md`
- **status**: pending
- **log**: —
- **files edited/created**: pendiente de ejecución

### T7: `Sellar`, cruce y replanificación atómica sin degradación
- **id**: T7
- **fase**: pendiente de reserva
- **depends_on**: [T6]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/`
- **files_lock**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparacionCorridaService.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorCruceTests.cs`
- **description**: guarda quirúrgica en el bloque de reparto de horas (`:173-315`) cuando `GrupoConPlanExplicitoValido`; I6, vacante y sábado siguen ejecutándose siempre; ante invalidación de fechas en planes explícitos se ejecuta replanificación atómica del grupo (`ReplanificarGrupo`) y si falla se rechaza el candidato (prohibido `DemoverPlanExplicito` o fallback legado a `AddDays(7)`); clonación y reparación conservan plan y flag intactos.
- **validation**: `Sellar_PlanExplicitoConsistente_NoModificaFechasNiHoras`; `Sellar_NormalizaSobreDuracionAunConPlanExplicito`; `Cruce_ReplanificaGrupoORechazaSinFallbackLegado`; `Reparacion_NoDuplicaSesiones`; ningún gen explícito degradado a fechas heurísticas.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgCoherenciaGrupoTests|FullyQualifiedName~AgOperadorCruceTests"`
- **evidence**: `_evidence/<stamp>/summary.md`
- **status**: pending
- **log**: —
- **files edited/created**: pendiente de ejecución

### T8: Persistencia, roundtrip y cobertura exacta
- **id**: T8
- **fase**: pendiente de reserva
- **depends_on**: [T7]
- **location**: `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/`, `.../Application/Services/Genetica/`
- **files_lock**: `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepositoryTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgExpansorSesionesTests.cs`
- **description**: INSERT (`:457`) y ambos SELECT (`:549`, `:1145`) incluyen `fechas_gen` y `plan_explicito`; roundtrip exacto fecha/bloque/duración; `Expansor` exige cobertura estricta (`FechasGen.Count * unidad == HorasGen`), rechaza duplicados, fechas fuera de ventana o días no concordantes con el bloque, y emite `Result.Failure` ante déficit por festivos o fin de ventana (los festivos se descuentan en la fase de reservas del planificador); #62 histórica intacta.
- **validation**: `GuardarYCargarGen_ConservaFechasGenYPlanExplicito`; `Expansor_DeficitPorFestivo_Falla`; `Expansor_PlanExplicitoIncompleto_Falla`; reintento de materialización sin duplicar sesiones.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgScheduleRepository|FullyQualifiedName~AgExpansorSesionesTests"`
- **evidence**: `_evidence/<stamp>/summary.md`
- **status**: pending
- **log**: —
- **files edited/created**: pendiente de ejecución

### T9: Validación por reglas reales y presentación
- **id**: T9
- **fase**: pendiente de reserva
- **depends_on**: [T8]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/`, `.../UI/Forms/`
- **files_lock**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPesosFitness.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEvaluadorAptitudTests.cs`
- **description**: auditar R1, R2, R3, R4 (dos ramas), R5, R6, R7, R8 (dos ramas), R9 (dos ramas), R10, R14 (dos ramas); R11/R12/R13 reservadas sin prueba nueva; intensidad como penalización secundaria opt-in, solo por oportunidades factibles omitidas; K indica días/horas usadas; exportaciones existentes conservadas.
- **validation**: `R4_DosFranjasMismoDia_EsViolacionDura`; reglas sin doble conteo en solape real; penalización no altera cobertura; exportaciones idénticas al formato previo cuando la estrategia está desactivada.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgValidadorReglas|FullyQualifiedName~AgEvaluadorAptitudTests"`
- **evidence**: `_evidence/<stamp>/summary.md`
- **status**: pending
- **log**: —
- **files edited/created**: pendiente de ejecución

### T10: Comparación con #62, activación y entrega
- **id**: T10
- **fase**: pendiente de reserva
- **depends_on**: [T9]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/`, `docs/articles/changes/`
- **files_lock**: `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`, `docs/articles/changes/2026-11-12-ranuras-dinamicas-4-a-8-horas.md` (nuevo), `implementation_plan.md`, `walkthrough.md`
- **description**: reejecutar sobre copia de #62 con cortes 30/05 y 23/07, mismo calendario y semilla; comparar por cohorte (demanda, capacidad, horas cubiertas, fecha final, huecos usables, semanas de 4 h evitables, violaciones duras); activar estrategia por opción versionada con reversión; no regenerar ni promover #62.
- **validation**: cobertura exacta; fecha límite de mayo no ampliada por el motor; tabla comparativa con evidencia; modo legado reproducible; documentación con resultados medidos.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- **evidence**: `_evidence/<stamp>/summary.md`
- **status**: pending
- **log**: —
- **files edited/created**: pendiente de ejecución

## Verificación y evidencia

```powershell
powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ws> -Plan .agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~<suite>"
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full      # sin -Filter; 0 = OK, exit 3 = regresión, SUITE_SHRUNK o baseline ausente (consultar summary.md)
```

- Un único gate a la vez. Evidencia en `_evidence/<stamp>/{build.log,test.trx,summary.md}`; no se commitea.
- `-Mode Full` nunca con `-Filter`: el filtro reduce la suite y devuelve `3`.
- Toda cita `archivo:línea` se resuelve contra el commit sellado.

## Criterio final

Cobertura exacta, cero violaciones duras nuevas, transición 4→8 probada con segunda ranura autorizada, fechas persistidas idénticas al plan y uso de huecos sin ampliar la fecha límite. El cierre de cohortes al 30/05 solo se declara si el diagnóstico y la ejecución lo demuestran.
