# Plan: Reclasificación Exclusiva de Materias con Sesiones Pendientes dentro del Plazo
**Generado**: 2026-11-14 · **Autor**: Cline/Swarm-Planner · **Estado**: closed

## Execution envelope
- branch: refactor/reclasificacion-solo-pendientes
- workspace: C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -AllowOneDrive -Workspace "C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan ".agents/plans/2026-11-14-reclasificacion-solo-pendientes-plan.md"
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: 1259)
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
  - registry_row: R-077
  - ronda_base: 376
  - fase_base: 1788
  - waves: 8
  - tasks: 10

## Dependency graph
```text
Wave 1: T1 (Fase 1788) [Model]               T2 (Fase 1789) [Evaluador & Tests]
          │                                       │
          ▼                                       ▼
Wave 2: T3 (Fase 1790) [Calibr Fixtures Tests]   T4 (Fase 1791) [UI Diálogo]
          │                                       │
          ▼                                       │
Wave 3: T5 (Fase 1792) [Escalador Eleg/Candidatas/SQL]
          │                                       │
          ▼                                       │
Wave 4: T6 (Fase 1793) [Validación Tipada Escalador]
          │                                       │
          ▼                                       │
Wave 5: T7 (Fase 1794) [Tests Rechazos Tipados]   │
          │                                       │
          └───────────────────┬───────────────────┘
                              ▼
Wave 6:                 T8 (Fase 1795) [UI Orquestación Flujo]
                              │
                              ▼
Wave 7:                 T9 (Fase 1796) [Tests Aceptación E2E]
                              │
                              ▼
Wave 8:                 T10 (Fase 1797) [Gate Full Baseline]
```

## Tasks

### T1: Enriquecimiento de Modelo AgReclasificacionPatron
- **id**: T1
- **fase**: 1788
- **depends_on**: []
- **location**: ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgReclasificacionPatron.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgReclasificacionPatron.cs
- **description**: Extender AgReclasificacionPatron para portar métricas de déficit y metadatos legibles (SesionesPendientes, HorasPendientes, FechaLimite, Sede, Programa, SemestreGrupo) garantizando su copia profunda en Clone() y asegurando que las validaciones de transición sigan funcionando.
- **validation**: AgReclasificacionPatron expone los campos de déficit y metadatos legibles; Clone() produce copias independientes de todos los nuevos campos.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
- **evidence**: (la llena build-verifier)
- **status**: done
- **log**: Implementada extensión de AgReclasificacionPatron con métricas de déficit y metadatos legibles (SesionesPendientes, HorasPendientes, FechaLimite, Sede, Programa, SemestreGrupo) y Clone() profunda. Compilación x64 limpia en ConversorPDF.XLL.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgReclasificacionPatron.cs

### T2: Desglose de Déficit por Asignatura en Evaluador y Cobertura Unitaria
- **id**: T2
- **fase**: 1789
- **depends_on**: []
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResultadoEvaluacionVentana.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentanaTests.cs
- **description**: Exponer en AgResultadoEvaluacionVentana el diccionario o desglose de sesiones pendientes por asignatura (DemandaRequerida - SesionesUbicadas) y escribir pruebas unitarias dedicadas en AgEvaluadorDemandaVentanaTests para verificar su cálculo exacto en casos factibles, saturados y con búsqueda agotada.
- **validation**: AgEvaluadorDemandaVentanaTests verifica que AgResultadoEvaluacionVentana discrimina correctamente las sesiones pendientes por IdAsignatura.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "ConversorPDF.Tests.Academic.Application.Services.Genetica.AgEvaluadorDemandaVentanaTests"
- **evidence**: _evidence/20261114-183431/summary.md
- **status**: done
- **log**: Implementado desglose de déficit por asignatura (SesionesPendientesPorAsignatura y HorasPendientesPorAsignatura) en AgResultadoEvaluacionVentana y AgEvaluadorDemandaVentana. 5/5 tests unitarios PASS.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResultadoEvaluacionVentana.cs
  - ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentanaTests.cs

### T3: Calibración Preventiva de Fixtures en Pruebas de Redistribución
- **id**: T3
- **fase**: 1790
- **depends_on**: [T1, T2]
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionCorreccionesTests.cs
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionCorreccionesTests.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionRanurasAceptacionTests.cs
- **description**: Calibrar preventivamente los fixtures en los tests preexistentes de redistribución para que aquellos tests que esperan candidatos de reclasificación configuren ventanas con saturación real (déficit > 0) y no asuman que materias que caben en plazo son candidatas, previniendo regresiones anticipadas al implementar la política estricta en T5.
- **validation**: Los fixtures existentes ejecutan limpiamente sin fallos de Sequence contains no elements ni aserciones prematuras.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgRedistribucion"
- **evidence**: _evidence/20261114-184000/summary.md
- **status**: done
- **log**: Calibración preventiva de fixtures en AgRedistribucionCorreccionesTests y AgRedistribucionRanurasAceptacionTests con déficit real. 40/40 tests PASS.
- **files edited/created**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionCorreccionesTests.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionRanurasAceptacionTests.cs

### T4: Diálogo UI de Autorización con Columnas Legibles y Déficit
- **id**: T4
- **fase**: 1791
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic/UI/Forms/FormAutorizarReclasificacion.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/UI/Forms/FormAutorizarReclasificacion.cs
- **description**: Enriquecer la presentación visual en FormAutorizarReclasificacion: mostrar sede, programa, semestre/grupo, nombre de materia, patrón anterior y sesiones/horas pendientes. Mantener selección individual desmarcada por defecto, y respetar Cancelar/Cerrar sin alterar selecciones.
- **validation**: Compilación limpia en Debug/x64 y CheckedListBox renderiza el texto compuesto con metadatos y déficit por materia.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
- **evidence**: _evidence/20261114-183727/summary.md
- **status**: done
- **log**: Enriquecida presentación de PoblarLista() con metadatos compuestos ([Sede - Programa], [SemestreGrupo], déficit sesiones/horas). Compilación x64 limpia en verify.ps1 -Mode Build.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/UI/Forms/FormAutorizarReclasificacion.cs

### T5: Centralización de Elegibilidad, Candidatas Fechadas y Aislamiento SQL
- **id**: T5
- **fase**: 1792
- **depends_on**: [T3]
- **location**: ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorRanuraAula.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorRanuraAula.cs
- **description**: Centralizar el predicado canónico EsAulaElegible(aula): Id > 0, AulaRef no vacía y Estado == AgEstadoAula.Disponible (1). Aulas con Estado Aprobado (4), Inactiva, Mantenimiento o Reservada quedan excluidas de generar candidatas y de autorizaciones. Reformar la generación de candidatas para basarse en el resultado de evaluación fechado: solo materias LEGADO con sesiones pendientes > 0 en la evaluación fechada. Si CabeEnPlazo == true, retornar lista vacía. Mantener sobrecarga de compatibilidad IdentificarPatronesRedistribuibles(contexto, aula, evaluacion = null). Enriquecer la clase Decision con IdAula y asegurar aislamiento SQL estricto 'WHERE id_aula = @IdAula' en AplicarDetalladoAsync.
- **validation**: IdentificarPatronesRedistribuibles solo genera candidatos para aulas Disponibles y materias con déficit positivo; materias completas nunca son ofertadas.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgRedistribucion"
- **evidence**: _evidence/20261114-185016/summary.md
- **status**: done
- **log**: Centralizado predicado canónico EsAulaElegible, candidatos acotados a déficit positivo desde evaluación fechada, sobrecarga retrocompatible, IdAula en Decision y aislamiento SQL WHERE id_aula = @IdAula. 40/40 tests PASS.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorRanuraAula.cs

### T6: Validación Tipada de Selección, Manejo Controlado de Rechazos y Huellas
- **id**: T6
- **fase**: 1793
- **depends_on**: [T5]
- **location**: ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorRanuraAula.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorRanuraAula.cs
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgHuellaEscaladoPersistido.cs
- **description**: Proveer validación tipada de selecciones en AgEscaladorRanuraAula que retorne un resultado estructurado en lugar de lanzar ArgumentException indiscriminada. Diferenciar explícitamente motivos: aula no elegible (ej. Estado != Disponible), materia ajena a candidatas evaluadas, configuración modificada o huella cambiada. Registrar advertencias WARN con IDs de aula/asignatura y detalle del rechazo. Asegurar rollback y cero escrituras ante rechazo.
- **validation**: La solución compila limpiamente en Debug/x64 y los métodos de validación estructurada discriminan con precisión la causa de rechazo.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
- **evidence**: _evidence/20261114-185823/summary.md
- **status**: done
- **log**: Implementada clase AgResultadoValidacionSeleccion y método ValidarSelecciones(...) con discriminación precisa de causas de rechazo (AULA_NO_ELEGIBLE, CONFIGURACION_MODIFICADA, etc.), logs WARN y verificación de estado en huella. Compilación limpia y 40/40 tests PASS.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorRanuraAula.cs
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgHuellaEscaladoPersistido.cs

### T7: Cobertura de Pruebas Unitarias de Elegibilidad y Rechazos Controlados
- **id**: T7
- **fase**: 1794
- **depends_on**: [T6]
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionCorreccionesTests.cs
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionCorreccionesTests.cs
- **description**: Escribir pruebas unitarias exhaustivas en AgRedistribucionCorreccionesTests adaptadas a la nueva validación tipada de T6: 1) Aula elegible (Disponible) genera candidatos solo para materias con déficit. 2) Aula no elegible (Estado=4 Aprobado, caso aula 12) nunca genera candidatos y su autorización forzada se rechaza con motivo controlado 'aula no elegible'. 3) Dos aulas con misma materia y distinto estado. 4) Dos programas homónimos en misma sede mantienen identidades aisladas. 5) Cero escrituras ante rechazo o cancelación.
- **validation**: Todos los tests de AgRedistribucionCorreccionesTests pasan satisfactoriamente verificando el aislamiento y la política de rechazo.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgRedistribucionCorreccionesTests"
- **evidence**: _evidence/20261114-190300/summary.md
- **status**: done
- **log**: Pruebas unitarias completadas para aulas no elegibles (Aprobado/Inactiva/Mantenimiento/Reservada), motivo AULA_NO_ELEGIBLE, aislamiento de aulas Disponible vs Aprobado, programas homónimos y cero escrituras. 28/28 tests PASS.
- **files edited/created**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionCorreccionesTests.cs

### T8: Orquestación del Flujo UI en Generador de Horarios
- **id**: T8
- **fase**: 1795
- **depends_on**: [T4, T6]
- **location**: ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs
- **files_lock**:
  - ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs
- **description**: Invertir y ordenar el flujo en AgFormGeneradorHorarios: 1) Evaluar plan fechado con el contexto actual. 2) Extraer candidatas exclusivamente de las materias con déficit en esa evaluación reutilizando el resultado fechado. 3) Si no existen candidatas, no abrir FormAutorizarReclasificacion y continuar con propuesta regular. 4) Si hay candidatas, abrir diálogo modal. Si el usuario cancela, terminar sin guardar. 5) Si el usuario confirma, simular en copia, evaluar misma ventana y avanzar a escalera solo si persiste déficit y aporta mejora. 6) Gestionar errores de validación de forma controlada sin propagar excepciones no capturadas.
- **validation**: Compilación correcta en Debug/x64; AgFormGeneradorHorarios no invoca FormAutorizarReclasificacion si la evaluación inicial no arroja materias pendientes elegibles; captura de errores muestra aviso informativo al operador.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
- **evidence**: _evidence/20261114-191000/summary.md
- **status**: done
- **log**: Reordenado flujo en AgFormGeneradorHorarios.cs (evaluar primero fechado, extraer candidatas con déficit, modal condicional, validación tipada con aviso explicativo y try/catch de excepciones de negocio sin ExcepcionNoControlada). Compilación limpia y 48/48 PASS.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs

### T9: Cobertura de Pruebas de Aceptación End-to-End
- **id**: T9
- **fase**: 1796
- **depends_on**: [T7, T8]
- **location**: ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionRanurasAceptacionTests.cs
- **files_lock**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionRanurasAceptacionTests.cs
- **description**: Actualizar y ampliar las pruebas de aceptación en AgRedistribucionRanurasAceptacionTests: 1) Caso plan completo (CabeEnPlazo=true) produce 0 candidatos y no abre diálogo. 2) Aula con una materia completa y una pendiente ofrece únicamente la materia pendiente. 3) LEGADO pendiente vs FIJO pendiente: solo LEGADO es seleccionable; FIJO permanece intacto. 4) AUTOMATICO pendiente evalúa escalera sin diálogo. 5) Reclasificación en la misma ventana resuelve déficit sin ampliar días. 6) Idempotencia: segunda aplicación produce SinCambios y cero escrituras.
- **validation**: Todos los tests de AgRedistribucionRanurasAceptacionTests pasan satisfactoriamente.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "AgRedistribucionRanurasAceptacionTests"
- **evidence**: _evidence/20261114-191600/summary.md
- **status**: done
- **log**: Pruebas de aceptación end-to-end completadas (materias que caben con 0 candidatos, exclusión de materia completa en aula mixta, LEGADO vs FIJO, resolución en misma ventana, idempotencia). 20/20 PASS.
- **files edited/created**:
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRedistribucionRanurasAceptacionTests.cs

### T10: Certificación de la Suite Completa y Gate Full
- **id**: T10
- **fase**: 1797
- **depends_on**: [T9]
- **location**: tests/baseline/test_baseline.txt
- **files_lock**:
  - tests/baseline/test_baseline.txt
- **description**: Ejecutar la suite completa de pruebas mediante verify.ps1 -Mode Full. Validar 0 errores, 0 regresiones contra el baseline (TOTAL >= 1259) y certificar que la integridad global del proyecto se preserva.
- **validation**: scripts/verify.ps1 -Mode Full devuelve exit code 0 con TOTAL >= 1259.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- **evidence**: _evidence/20261114-192003/summary.md
- **status**: done
- **log**: Suite completa ejecutada con verify.ps1 -Mode Full. 1356 tests ejecutados, 1356 PASS, 0 fallos, 0 regresiones. Baseline actualizado canónicamente a 1356. GREEN_VALIDATED.
- **files edited/created**:
  - tests/baseline/test_baseline.txt

## Parallel execution groups
| Wave | Ronda | Tasks | Can start when |
|---|---|---|---|
| 1 | 376 | T1, T2 | Inmediato (leases disjuntos: AgReclasificacionPatron vs AgEvaluadorDemandaVentana/Tests) |
| 2 | 377 | T3, T4 | Wave 1 completada (leases disjuntos: AgRedistribucionCorreccionesTests/Ranuras vs FormAutorizarReclasificacion) |
| 3 | 378 | T5 | Wave 2 completada (lease exclusivo: AgEscaladorRanuraAula) |
| 4 | 379 | T6 | Wave 3 completada (lease exclusivo: AgEscaladorRanuraAula/Huella) |
| 5 | 380 | T7 | Wave 4 completada (lease exclusivo: AgRedistribucionCorreccionesTests) |
| 6 | 381 | T8 | Wave 5 completada y T4 lista (lease exclusivo: AgFormGeneradorHorarios) |
| 7 | 382 | T9 | Wave 6 completada (lease exclusivo: AgRedistribucionRanurasAceptacionTests) |
| 8 | 383 | T10 | Wave 7 completada (Gate Full suite completa) |

## Testing strategy
- Pruebas unitarias focalizadas por componente usando verify.ps1 -Mode Fast -Filter.
- Verificación de no-regresión en la suite completa (1259 tests) usando verify.ps1 -Mode Full.
- Casos específicos para exclusión de aulas no elegibles (Aprobado/Inactiva/Mantenimiento/Reservada).
- Validación de que materias completas nunca son ofertadas para reclasificación.

## Risks & mitigations
- **Riesgo**: Tests preexistentes que asumían que cualquier LEGADO generaba candidato fallen.
  - **Mitigación**: T3 calibra preventivamente los fixtures en la Ola 2 antes de la implementación de T5 en la Ola 3.
- **Riesgo**: Colisiones de concurrencia al compilar en Excel-DNA.
  - **Mitigación**: build_gate serializado y límite estricto de max_parallel_writers: 2 con leases disjuntos.
- **Riesgo**: Bloqueo de UI por excepciones en llamadas STA.
  - **Mitigación**: T6 y T8 convierten ArgumentException en validaciones tipadas y manejo controlado con avisos al usuario.

## Self-Review (13 Puntos)
| # | Punto | Cumplimiento |
|---|---|---|
| 1 | Dependencias completas: ningún depends_on omitido | OK (Grafo acíclico cerrado en 8 olas) |
| 2 | Orden ejecutable: cada tarea compila al terminar | OK (Model/Evaluador -> Fixtures -> Escalador -> Validación -> Tests -> UI -> Aceptación -> Gate) |
| 3 | Ninguna tarea mezcla L0 sin rol integrator | OK (No se tocan archivos L0; todos los tipos ya existen) |
| 4 | files_lock disjuntos por ola | OK (Auditado rigurosamente en las 8 olas) |
| 5 | Cada verification es un comando que existe en el repo | OK (verify.ps1 con -Mode Fast, -Mode Build y -Mode Full) |
| 6 | Cada validation es falsable | OK |
| 7 | El RED compila (esqueleto integrator-first previsto) | OK (Todos los tipos ya existen, no hay clases nuevas) |
| 8 | Migración y APP_VERSION asignadas solo por integrator | OK (No aplica; no hay migraciones DB) |
| 9 | Ninguna tarea cruza más de un slice | OK (Solo slice Academic) |
| 10 | El context pack usado tiene <= 400 líneas | OK |
| 11 | Criterio de aceptación global definido | OK |
| 12 | Riesgos con mitigación | OK |
| 13 | numbering: en el envelope y fila reserved añadida a REGISTRY.md | OK (R-077, Ronda 376-383, Fase 1788-1797) |
