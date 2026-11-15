# Plan: Correcciones Libro 8 / Corrida 94 — Aulas Aprobadas, Autoridad de Conflictos Fechados y Consistencia Staging/UI

**Generado**: 2026-11-15 · **Autor**: Swarm-Planner · **Estado**: executing

## Execution envelope
- branch: codex/huellas-escalado-20261114
- workspace: C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace "C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan ".agents/plans/2026-11-15-libro8-corrida94-plan.md" -AllowOneDrive
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: 1387)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado (un unico build/tests a la vez por workspace; Excel-DNA bloquea bin/)
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
  - registry_row: R-079
  - ronda_base: 392          # Wave 1 = Ronda 392
  - fase_base: 1806          # T1 = Fase 1806
  - waves: 6                 # Ronda 392..397
  - tasks: 9                 # Fase 1806..1814

## Dependency graph
```text
Wave 1 (Ronda 392) [L0]:       T1 (Esqueleto e inclusión csproj de AgCalendarioConflictos)
                                ├─────────────────────────────┐
Wave 2 (Ronda 393):            T2 (Implementación Calendario) T3 (A1: ValidarAprobadas en AgGeneradorHorarios)
                                ├─────────────────┐           │
Wave 3 (Ronda 394):            T4 (A2: Reglas)   T5 (A3: Terceros)
                                │                             │
Wave 4 (Ronda 395):            T6 (A4/A6: AgScheduleRepository)
                                ├─────────────────────────────┤
Wave 5 (Ronda 396):            T7 (A6: FormVerVinculaciones)  T8 (A4/A5: AgFormGeneradorHorarios)
                                └─────────────────┬───────────┘
Wave 6 (Ronda 397) [Gate]:     T9 (Full Gate e Integración Global)
```

## Tasks

### T1: Esqueleto de AgCalendarioConflictos e inclusión en csproj
- **id**: T1
- **fase**: 1806
- **depends_on**: []
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCalendarioConflictos.cs
- **files_lock**:
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCalendarioConflictos.cs
- **description**: Rol Integrator (Lease L0). Crea el esqueleto inicial compilable de `AgCalendarioConflictos.cs` con su clase anidada `Ocupacion` y firmas públicas/internas con `NotImplementedException`, y lo registra formalmente en `ConversorPDF.XLL.csproj` dentro del ItemGroup de compilación correspondiente.
- **validation**: La solución compila en Debug/x64 y el tipo `AgCalendarioConflictos` es accesible desde el proyecto de pruebas.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "Genetica"
- **evidence**: _evidence/20261115-135905/summary.md
- **status**: done
- **log**: Integrator creó el esqueleto compilable de AgCalendarioConflictos con clase anidada Ocupacion y firmas con NotImplementedException. Registrado en ConversorPDF.XLL.csproj. Compilación verify.ps1 -Mode Build exitosa (código 0). Auditado y aprobado por task-reviewer.
- **files edited/created**:
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCalendarioConflictos.cs

### T2: Autoridad común de ocupación fechada (AgCalendarioConflictos)
- **id**: T2
- **fase**: 1807
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCalendarioConflictos.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCalendarioConflictos.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCalendarioConflictosTests.cs
- **description**: Rol Implementer. Implementa la lógica completa de `AgCalendarioConflictos`:
  - `Ocupacion`: campos `Gen`, `Fecha`, `Inicio`, `Fin`, `Franja`.
  - `Crear(IEnumerable<AgGen> genes, AgContexto contexto)`: usa `AgExpansorSesiones` para proyectar fechas efectivas de cada gen con `HorasGen > 0`, validando bloques activos e intervalos coherentes. Si la expansión falla o produce sesiones sin fecha, retorna fallo estructurado (`I6_BLOQUE_INVALIDO` o `I5_SESION_SIN_FECHA`) sin asumir agenda libre.
  - `SeCruzan(Ocupacion a, Ocupacion b)`: intersección real por fecha e intervalos horarios `a.Fecha == b.Fecha && a.Inicio < b.Fin && b.Inicio < a.Fin`.
  - `Cohorte(AgGen g)`: normalización canónica de cohorte `${g.IdSede}|${g.IdPrograma}|${semestre}|${g.Grupo}` respetando fallback de semestre romano.
  - `DocenteReal(AgGen g, AgContexto c)`: valida `g.EsVacante == 0 && g.IdTercero > 0 && g.IdTercero != c.IdTerceroVacante`.
  Escribe pruebas unitarias xUnit en `AgCalendarioConflictosTests.cs`.
- **validation**: Pruebas unitarias pasan validando proyección correcta de fechas intercaladas, detección de cruces reales y descarte de falsos positivos temporales.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgCalendarioConflictosTests"
- **evidence**: _evidence/20261115-141028/summary.md
- **status**: done
- **log**: Implementada lógica completa en AgCalendarioConflictos (Ocupacion, Crear con AgExpansorSesiones, SeCruzan, Cohorte, DocenteReal). 27 pruebas unitarias en AgCalendarioConflictosTests ejecutadas y superadas al 100% (código 0). Auditado y aprobado por task-reviewer.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCalendarioConflictos.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCalendarioConflictosTests.cs

### T3: A1 — Bloqueo específico de aula aprobada con déficit
- **id**: T3
- **fase**: 1808
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGeneradorHorariosAprobadasTests.cs
- **description**: Rol Implementer. Implementa `ValidarAprobadas(AgContexto contexto)` en `AgGeneradorHorarios` invocado antes de `ValidarFactibilidadAsync` en `EjecutarOptimizacionAsync`. Si un aula con `Estado == AgEstadoAula.Aprobado` no cabe en plazo bajo la ventana efectiva recibida, retorna error explícito: `I5_AULA_APROBADA_CON_DEFICIT: {aula.AulaRef}; requiere revisión explícita de aprobación para replanificar. El escalador no modifica aulas aprobadas.`. Protege el estado aprobado sin convertirlo automáticamente a Disponible ni omitir aulas silenciosamente. Añade pruebas unitarias xUnit en `AgGeneradorHorariosAprobadasTests.cs`.
- **validation**: Si un aula aprobada presenta déficit respecto a la ventana efectiva, la optimización falla de inmediato con el código de error `I5_AULA_APROBADA_CON_DEFICIT` y el mensaje prescriptivo.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgGeneradorHorariosAprobadasTests"
- **evidence**: _evidence/20261115-141050/summary.md
- **status**: done
- **log**: Implementado ValidarAprobadas en AgGeneradorHorarios antes del pre-vuelo. 6 pruebas unitarias en AgGeneradorHorariosAprobadasTests más 3 en AgGeneradorHorariosTests ejecutadas y aprobadas (9/9 PASS). Auditado y aprobado por task-reviewer.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGeneradorHorariosAprobadasTests.cs
- **log**: (lo llena el orquestador)
- **files edited/created**: (lo llena el orquestador)

### T4: A2 — Sustitución de comprobaciones temporales en AgValidadorReglas por proyección fechada y calibración de fixtures
- **id**: T4
- **fase**: 1809
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs
- **description**: Rol Implementer.
  1. En `AgValidadorReglas.Evaluar`: Reemplaza las comprobaciones temporales basadas en envolventes de `FechaInicioEstimada`..`FechaFinEstimada` (R1, R2, R3, R4 y R7) por la autoridad común `AgCalendarioConflictos.Crear`:
     - R1: Superposición real de docente en misma fecha e intervalo horario.
     - R2: Superposición real de aula en misma fecha e intervalo horario.
     - R3: Superposición real de grupo en misma fecha e intervalo horario.
     - R4: Unicidad por cohorte, fecha concreta y franja, y tope de franjas diarias (2 en sábado, 1 de lunes a viernes).
     - R7: Tercero en otro grupo o con ocupación externa validado en cada fecha proyectada real (`contexto.TerceroOcupado`).
     - Serializa `detalleJson` con `indicesGenes`, `fecha` e `intervalos` para resolución posterior de IDs. Si la proyección falla, marca `R14_BloqueHorarioNoAsignado` e infactible sin asumir agenda libre.
  2. **Calibración obligatoria de fixtures preexistentes en `AgValidadorReglasTests.cs`**:
     - Varios tests unitarios existentes instancian `AgContexto` asignando `IdBloqueHorario` sin poblar `contexto.BloquesPorId` con bloques activos e intervalos válidos.
     - El implementador calibrará los fixtures de prueba de R1-R4 y R7 asegurando que registren bloques activos con horas de inicio y fin válidas en `contexto.BloquesPorId` (mediante helper de fixture común en el archivo de tests), evitando que fallen con `I6_BLOQUE_INVALIDO` / R14 espurio.
  3. Añade regresión específica demostrando que asignaturas con fechas alternadas sin cruce real (como Estadística y Teorías en CALI III-A) arrojan 0 HARD, mientras cruces efectivos por día/hora generan la violación correspondiente.
- **validation**: Toda la suite existente y nueva de `AgValidadorReglasTests` pasa en verde; elimina los falsos positivos por solapamiento de rangos envolventes y preserva la detección de cruces reales.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgValidadorReglasTests"
- **evidence**: _evidence/20261115-143048/summary.md
- **status**: done
- **log**: Sustituidos bloques temporales por proyección fechada AgCalendarioConflictos en AgValidadorReglas (R1, R2, R3, R4, R7). Calibrados fixtures preexistentes con bloques activos en contexto.BloquesPorId. Añadida regresión Estadística vs Teorías con 0 HARD. 32/32 tests superados en AgValidadorReglasTests. Auditado y aprobado por task-reviewer.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs

### T5: A3 — Disponibilidad para el grupo completo en AgAsignadorTercero
- **id**: T5
- **fase**: 1810
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAsignadorTerceroTests.cs
- **description**: Rol Implementer. En `AgAsignadorTercero`:
  - Añade helper `MismoGrupo(AgGen a, AgGen b)` para agrupar genes de la misma asignatura/aula/programa/sede/grupo.
  - Implementa `DisponibleGrupo(int docente, AgGen gen, AgIndividuo individuo, AgContexto contexto)` con null-safety defensiva (`individuo?.Genes != null`), validando que el docente esté habilitado para todo el grupo, no choque con genes `Fijado` de otro docente, y que todas las fechas proyectadas del grupo no choquen con otras asignaciones del docente en el individuo ni con ocupación externa (`contexto.TerceroOcupado`).
  - Actualiza `DocenteDisponible` para delegar en `DisponibleGrupo`.
  - En `BuscarAlterno(AgGen gen, AgContexto contexto, AgIndividuo individuo)`: además de evaluar `DisponibleGrupo`, encapsula directamente la actualización sincronizada de los miembros no fijados del grupo cuando se encuentra un candidato alterno apto, o provee el método público de conveniencia `BuscarYAsignarAlternoGrupo(AgGen gen, AgContexto contexto, AgIndividuo individuo)` dentro de `AgAsignadorTercero.cs` para mantener la consistencia sin forzar modificaciones fuera del `files_lock`.
  - Añade pruebas unitarias xUnit en `AgAsignadorTerceroTests.cs` cubriendo asignación multi-día, consistencia de grupo y prevención de choques posteriores.
- **validation**: `AgAsignadorTerceroTests` pasa verificando que un docente no es asignado si tiene cruces en fechas intermedias o posteriores de la asignatura y que el grupo hermano se mantiene sincronizado.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgAsignadorTerceroTests"
- **evidence**: _evidence/20261115-142233/summary.md
- **status**: done
- **log**: Implementado MismoGrupo, DisponibleGrupo, actualización sincronizada de grupo hermano en BuscarAlterno y helper BuscarYAsignarAlternoGrupo en AgAsignadorTercero. 8/8 pruebas superadas en AgAsignadorTerceroTests. Auditado y aprobado por task-reviewer.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAsignadorTerceroTests.cs
- **log**: (lo llena el orquestador)
- **files edited/created**: (lo llena el orquestador)

### T6: A4/A6 — Persistencia de vacantes y resolución de IDs de genes en AgScheduleRepository
- **id**: T6
- **fase**: 1811
- **depends_on**: [T4]
- **location**: ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs
  - ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgScheduleRepositoryCorridaVacantesTests.cs
- **description**: Rol Implementer. En `AgScheduleRepository`:
  - En `UpdateCorridaAsync`: añade `vacantes = @Vacantes` en la cláusula `SET` del SQL y en el objeto anónimo de parámetros, asegurando que el recuento de vacantes persista en `ag_corrida`.
  - En `SaveIndividuoAsync`: dentro de la transacción existente, después de asignar `g.Id` a todos los genes y antes de insertar en `ag_violacion`, deserializa `DetalleJson`, resuelve `indicesGenes` hacia los IDs generados (`idGenes`), asigna `violacion.IdGen = ids[0]` y re-serializa el detalle.
  - Añade pruebas unitarias xUnit en `AgScheduleRepositoryCorridaVacantesTests.cs` verificando la persistencia de vacantes y el mapeo de `id_gen` en las violaciones.
- **validation**: Las pruebas de persistencia confirman que `vacantes` se actualiza en base de datos y que `ag_violacion` queda poblada con `id_gen` válido cuando el detalle contiene `indicesGenes`.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgScheduleRepositoryCorridaVacantesTests"
- **evidence**: _evidence/20261115-144154/summary.md
- **status**: done
- **log**: Incorporada persistencia de vacantes en UpdateCorridaAsync y resolución transaccional de indicesGenes hacia idGenes/id_gen en SaveIndividuoAsync. 4/4 pruebas unitarias superadas en AgScheduleRepositoryCorridaVacantesTests. Auditado y aprobado por task-reviewer.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs
  - ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgScheduleRepositoryCorridaVacantesTests.cs
- **log**: (lo llena el orquestador)
- **files edited/created**: (lo llena el orquestador)

### T7: A6 — Consulta y visualización de violaciones globales y enlazadas en FormVerVinculaciones
- **id**: T7
- **fase**: 1812
- **depends_on**: [T4, T6]
- **location**: ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs
  - ConversorPDF.Tests/Academic/UI/Forms/FormVerVinculacionesViolacionesTests.cs
- **description**: Rol Implementer. En `FormVerVinculaciones.cs`:
  - Modifica `CargarViolacionesAsync` para no requerir `INNER JOIN ag_sesiones`: consulta todas las violaciones del individuo óptimo de la corrida vinculada con filtro tolerante de subunidad `WHERE c.id IN @IdCorridas AND (@Subunidad = '' OR COALESCE(c.subunidad, '') IN ('', @Subunidad))`.
  - Las violaciones con `id_gen` y fecha resoluble se asocian a las sesiones correspondientes.
  - Las violaciones globales (sin `id_gen` o sin fecha enlazada) se muestran bajo un nodo raíz dedicado en el árbol (`TreeNode nodoGlobales = new TreeNode("⚠️ Violaciones Generales de la Corrida")`) para evitar que queden invisibles en la UI.
  - **Convención de testing (regla `test-convention.md`)**: El archivo de pruebas `ConversorPDF.Tests/Academic/UI/Forms/FormVerVinculacionesViolacionesTests.cs` debe implementarse con **MSTest 3.3.1** bajo hilo STA (`[TestClass]`, `[STATestMethod]`), evitando xUnit en la carpeta de UI.
- **validation**: Las pruebas MSTest en STA validan que las violaciones de la corrida se cargan y son visibles en el visor aun cuando no tengan sesión específica enlazada (`id_gen IS NULL`).
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FormVerVinculacionesViolacionesTests"
- **evidence**: _evidence/20261115-150433/summary.md
- **status**: done
- **log**: Modificado CargarViolacionesAsync sin INNER JOIN ag_sesiones, nodo raíz dedicado para violaciones globales en TreeView, y filtro tolerante de subunidad. 6/6 tests MSTest STA superados en FormVerVinculacionesViolacionesTests. Auditado y aprobado por task-reviewer.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs
  - ConversorPDF.Tests/Academic/UI/Forms/FormVerVinculacionesViolacionesTests.cs

### T8: A4/A5 — Conteos/puntuaciones separados y compuerta estricta de factibilidad/promoción en AgFormGeneradorHorarios
- **id**: T8
- **fase**: 1813
- **depends_on**: [T3, T4, T6]
- **location**: ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs
  - ConversorPDF.Tests/Academic/UI/Forms/AgFormGeneradorHorariosResultadosTests.cs
- **description**: Rol Implementer. En `AgFormGeneradorHorarios.cs`:
  - En `CargarModuloResultadosAsync`: separa conteos de registros (`registrosHard`, `registrosSoft`) de infracciones e importe de penalización (`ind.HardViolations`, `ind.SoftViolations`). Corrige la advertencia de log para no sumar penalización a recuento de filas.
  - En la finalización de optimización / post-materialización:
    - Lee las sesiones de staging con `GetSesionesStagingAsync`.
    - Calcula `sesionesVacantes` y `horasVacantes`.
    - Valida cobertura exacta comparando multiconjuntos de sesiones proyectadas vs materializadas (`gen|fecha|horas`).
    - Define `factible = stagingOk && cobertura && mejor.HardViolations == 0 && mejor.Vacantes == 0 && sesionesVacantes == 0`.
    - Asigna `corrida.Vacantes = mejor.Vacantes;` antes de `UpdateCorridaAsync`.
    - Habilita `btnPromoverProduccion.Enabled = factible;` y actualiza `lblStatusAG.Text` con detalle claro de faltantes si es parcial.
  - **Unificación de compuerta en combo selector (`CargarCorridaDesdeComboAsync`)**:
    - Actualiza la línea 3321 para que la rehidratación desde el combo use la misma compuerta estricta: `btnPromoverProduccion.Enabled = ind != null && ind.HardViolations == 0 && corrida.Vacantes == 0 && sesionesVacantes == 0;`, impidiendo el bypass al recargar una corrida con vacantes.
  - **Convención de testing (regla `test-convention.md`)**: El archivo de pruebas `ConversorPDF.Tests/Academic/UI/Forms/AgFormGeneradorHorariosResultadosTests.cs` debe implementarse con **MSTest 3.3.1** bajo hilo STA (`[TestClass]`, `[STATestMethod]`).
- **validation**: Pruebas MSTest STA confirman que el módulo de resultados muestra separadamente registros y puntuaciones sin generar warnings espurios, y que la promoción a producción queda bloqueada tanto tras optimización como tras recargar desde el combo si existen HARD > 0, vacantes > 0 o falta de cobertura.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgFormGeneradorHorariosResultadosTests"
- **evidence**: _evidence/20261115-145842/summary.md
- **status**: done
- **log**: Implementada separación de conteos y penalizaciones en CargarModuloResultadosAsync, compuerta estricta de promoción/factibilidad con cobertura multiconjunto en post-materialización y en CargarCorridaDesdeComboAsync. 5/5 pruebas MSTest STA superadas en AgFormGeneradorHorariosResultadosTests. Auditado y aprobado por task-reviewer.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs
  - ConversorPDF.Tests/Academic/UI/Forms/AgFormGeneradorHorariosResultadosTests.cs
- **log**: (lo llena el orquestador)
- **files edited/created**: (lo llena el orquestador)

### T9: Full Gate e Integración Global de Corrida 94
- **id**: T9
- **fase**: 1814
- **depends_on**: [T5, T7, T8]
- **location**: ConversorPDF.Tests
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasAceptacionTests.cs
  - tests/baseline/test_baseline.txt
- **description**: Rol Integrator / Build-verifier.
  - Incorpora prueba de aceptación `Criterio14_Corrida94_CalendarioConflictos_Aprobadas_Y_ConsistenciaStaging` en `AgPlanificadorRanurasAceptacionTests.cs` validando el ciclo integral: aula aprobada con déficit bloqueada pre-vuelo, fechas intercaladas sin falsos HARD, asignación de grupo sin choques docentes, vacantes persistidas en `ag_corrida`, y compuerta estricta de promoción.
  - Ejecuta la suite de verificación completa `scripts/verify.ps1 -Mode Full`.
  - Actualiza `tests/baseline/test_baseline.txt` con la nueva cardinalidad neta de pruebas aprobadas sin regresiones.
- **validation**: Suite completa en verde (Full Mode), 0 fallos, 0 omitidos, `TOTAL >= 1391`.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- **evidence**: (la llena build-verifier)
- **status**: pending
- **log**: (lo llena el orquestador)
- **files edited/created**: (lo llena el orquestador)

## Parallel execution groups
| Wave | Ronda | Tasks | Can start when |
|---|---|---|---|
| 1 | 392 | T1 (L0) | Inmediato (Inicio serializado Integrator L0) |
| 2 | 393 | T2, T3 | T1 completado (leases disjuntos entre Calendario y GeneradorHorarios) |
| 3 | 394 | T4, T5 | T2 completado (leases disjuntos entre ValidadorReglas y AsignadorTercero) |
| 4 | 395 | T6 | T4 completado (Persistencia en AgScheduleRepository) |
| 5 | 396 | T7, T8 | T6 completado (leases disjuntos entre FormVerVinculaciones y AgFormGeneradorHorarios) |
| 6 | 397 | T9 (Gate) | T5, T7 y T8 completados (Integración Full Gate) |

## Testing strategy
- Fast mode tras cada tarea atómica utilizando filtros específicos sobre las clases creadas y modificadas.
- Pruebas unitarias de aislamiento con xUnit 2.7 para servicios y dominio (`AgCalendarioConflictosTests`, `AgValidadorReglasTests`, `AgAsignadorTerceroTests`, `AgGeneradorHorariosAprobadasTests`, `AgScheduleRepositoryCorridaVacantesTests`).
- Pruebas unitarias en hilo STA con MSTest 3.3.1 para formularios y controles UI bajo `ConversorPDF.Tests/Academic/UI/Forms/` (`FormVerVinculacionesViolacionesTests`, `AgFormGeneradorHorariosResultadosTests`), siguiendo `test-convention.md`.
- Pruebas de integración sobre SQLite in-memory para el repositorio y staging.
- Verificación Full serializada en Wave 6 con `scripts/verify.ps1 -Mode Full` comprobando que no existan regresiones en los 1387+ tests de baseline.

## Risks & mitigations
- **Riesgo 1**: Costo computacional de proyectar fechas completas con `AgExpansorSesiones` durante la evaluación genética de poblaciones grandes.
  - *Mitigación*: Indexar proyecciones por recurso (aula, docente, cohorte) y evaluar cruces únicamente sobre sesiones concurrentes en la misma fecha.
- **Riesgo 2**: Bloqueo de aulas aprobadas impida optimizaciones legítimas.
  - *Mitigación*: `ValidarAprobadas` emite mensaje claro prescriptivo indicando que el aula requiere reapertura administrativa formal; no altera datos silenciosamente.
- **Riesgo 3**: Pérdida de retrocompatibilidad en violaciones históricas sin `id_gen`.
  - *Mitigación*: `FormVerVinculaciones` tolera `id_gen == null` y visualiza violaciones bajo un nodo raíz dedicado `⚠️ Violaciones Generales de la Corrida`.
- **Riesgo 4**: Colisión en compilación de archivos del proyecto XLL.
  - *Mitigación*: `ConversorPDF.XLL.csproj` se modifica exclusivamente en T1 por el rol `integrator`.
- **Riesgo 5**: Regresiones en tests existentes de `AgValidadorReglasTests.cs` por falta de bloques activos en fixtures.
  - *Mitigación*: T4 incluye explícitamente la calibración de fixtures para registrar bloques activos en `contexto.BloquesPorId`.
- **Riesgo 6**: Excepciones de Cross-Thread / STA en test runners al probar componentes WinForms.
  - *Mitigación*: T7 y T8 emplean MSTest 3.3.1 con `[STATestMethod]` en `UI/Forms/`.

## Self-Review: checklist de 13 puntos
| # | Punto | Cumplimiento | Detalle |
|---|---|:---:|---|
| 1 | Dependencias completas: ningún `depends_on` omitido | SÍ | Grafo acíclico y explícito de T1 a T9. |
| 2 | Orden ejecutable: cada tarea compila al terminar | SÍ | Esqueleto L0 primero (T1) antes de uso o prueba. |
| 3 | Ninguna tarea mezcla L0 sin rol `integrator` | SÍ | Solo T1 y T9 (Gate) asignados a `integrator`. |
| 4 | `files_lock` disjuntos por ola | SÍ | En Wave 2 (T2, T3), Wave 3 (T4, T5) y Wave 5 (T7, T8) los locks son disjuntos. |
| 5 | Cada `verification` es un comando que existe en el repo | SÍ | Usa `scripts/verify.ps1` con `-Mode Fast` o `-Mode Full`. |
| 6 | Cada `validation` es falsable | SÍ | Criterios verificables con aserciones concretas. |
| 7 | El RED compila (esqueleto `integrator-first` previsto) | SÍ | T1 crea el esqueleto compilable de `AgCalendarioConflictos.cs`. |
| 8 | Migración y `APP_VERSION` asignadas solo por `integrator` | SÍ | No se añaden migraciones de esquema de base de datos; la tabla `ag_corrida` ya tiene columna `vacantes` y `ag_violacion` ya tiene `id_gen`. |
| 9 | Ninguna tarea cruza más de un slice | SÍ | Todo el trabajo pertenece al slice `Academic`. |
| 10 | El context pack usado tiene ≤ 400 líneas | SÍ | Revisión estricta y lectura quirúrgica de contexto. |
| 11 | Criterio de aceptación global definido | SÍ | Resoluciones A1 a A6 del diagnóstico Libro 8 verificadas en T9. |
| 12 | Riesgos con mitigación | SÍ | 6 riesgos técnicos identificados con sus mitigaciones. |
| 13 | `numbering:` en el envelope y fila `reserved` añadida a `REGISTRY.md` | SÍ | Envelope tiene `registry_row: R-079`, `ronda_base: 392`, `fase_base: 1806`, y fila en `REGISTRY.md`. |
