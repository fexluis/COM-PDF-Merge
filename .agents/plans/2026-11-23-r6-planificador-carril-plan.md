# Plan: Planificador determinista de carril y re-ranuración R6 (Ronda 534–538)

**Generado**: 2026-11-24 · **Estado**: closed

## Execution envelope
- branch: reparacion-automatica-r5
- workspace: C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace . -Plan ".agents/plans/2026-11-23-r6-planificador-carril-plan.md" -AllowOneDrive
- baseline: tests/baseline/test_baseline.txt
- verify: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj
- build_gate: serializado
- max_parallel_writers: 1
- no_push: true
- stop_conditions:
  - Acción destructiva no reversible sin respaldo
  - Violación de seguridad o datos
  - Plan en estado irrecuperable
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: R-206
  - ronda_base: 534
  - fase_base: 2375
  - waves: 5
  - tasks: 16

---

## Mapeo de Olas y Rondas

| Wave | Ronda | Tasks |
|---|---|---|
| Wave 534 | Ronda 534 | T1..T5 (Fase 2375–2379) |
| Wave 535 | Ronda 535 | T6..T9 (Fase 2380–2383) |
| Wave 536 | Ronda 536 | T10..T12 (Fase 2384–2386) |
| Wave 537 | Ronda 537 | T13..T15 (Fase 2387–2389) |
| Wave 538 | Ronda 538 | T16 (Fase 2390) |

---

## Tasks

### T1: Planificación multi-gen independiente por día
- **id**: T1
- **fase**: 2375
- **depends_on**: []
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorCarril.cs`
- **descripción**: Secuenciar a nivel de cada gen individual respetando su propio `DiaSemana` para evitar que asignaturas multi-gen asignen fechas de un único día.

### T2: Re-ranuración validada por programa, bloque y patrón
- **id**: T2
- **fase**: 2376
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
- **descripción**: Restringir re-ranuración a aulas del mismo `IdPrograma`, requerir `TryGetBloque` válido, deducir `DiasPatron` real y descartar aulas con vínculo `LEGADO/FIJO`.

### T3: Orden curricular y anclas fijas completas
- **id**: T3
- **fase**: 2377
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorCarril.cs`
- **descripción**: Respetar `OrdenCurricular` de `AulaAsignatura` antes de permutaciones y calcular `IniFijo/FinFijo` con fallback a `FechasGen`.

### T4: Limpieza de telemetría y conteo de intentos
- **id**: T4
- **fase**: 2378
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
- **descripción**: Resetear `resultado.Motivo` al inicio de cada iteración, registrar delta real de hard y sumar `intentos` al evaluar el planificador.

### T5: Pruebas unitarias robustas y cierre de registro
- **id**: T5
- **fase**: 2379
- **depends_on**: [T3, T4]
- **location**: ConversorPDF.Tests/Academic/ + .agents/plans/
- **files_lock**:
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgReparadorR6Tests.cs`
  - `.agents/plans/REGISTRY.md`
- **descripción**: Aserciones por invariantes relacionales (no-solape temporal y HARD) en vez de fechas mágicas, verificación de aula destino y fixture multi-día.

### T6: Validación estricta de periodo y vínculo destino en re-ranuración
- **id**: T6
- **fase**: 2380
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
- **descripción**: Exigir igualdad de `IdPeriodo` y existencia de `relDestino` válido en catálogo antes de re-ranurar.

### T7: Presupuesto del planificador y orden curricular O(1)
- **id**: T7
- **fase**: 2381
- **depends_on**: [T3, T4]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorCarril.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
- **descripción**: Limitar evaluaciones por `intentos >= MaxIntentos`, indexar orden curricular con diccionario O(1) y proteger orden curricular frente a permutaciones.

### T8: Certificación de compuertas y aserción exacta
- **id**: T8
- **fase**: 2382
- **depends_on**: [T5]
- **location**: ConversorPDF.Tests/Academic/ + tests/baseline/
- **files_lock**:
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgReparadorR6Tests.cs`
  - `tests/baseline/test_baseline.txt`
- **descripción**: Aserción exacta `SIN_FECHAS`, verificación de compuertas `verify Full` (2024 tests) y actualización de baseline.

### T9: Capacidad real sobre fechas libres
- **id**: T9
- **fase**: 2383
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorCarril.cs`
- **descripción**: Evaluar capacidad y semanas disponibles deduciendo previamente las colisiones externas en el aula/cohorte.

### T10: Filtro de aulas programables por estado y subunidad
- **id**: T10
- **fase**: 2384
- **depends_on**: [T6]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
- **descripción**: Excluir aulas inactivas o en mantenimiento requiriendo Estado Disponible o Aprobado, y verificar consistencia de Subunidad con el aula víctima.

### T11: Actualización integral del gen re-ranurado
- **id**: T11
- **fase**: 2385
- **depends_on**: [T10]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
- **descripción**: Sincronizar ClaveGrupo con la nueva aula alternativa y resetear PatronSemanal e IndiceDiaPatron para evitar claves obsoletas.

### T12: Telemetría de diagnóstico y pruebas de re-ranuración
- **id**: T12
- **fase**: 2386
- **depends_on**: [T10, T11]
- **location**: ConversorPDF.Tests/Academic/ + ConversorPDF.XLL/Academic/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResultadoReparacionR6.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgReparadorR6Tests.cs`
- **descripción**: Emitir MotivoReranuracion = SIN_VINCULO cuando ninguna aula candidata pase filtros y validar en pruebas unitarias el rechazo de inactivas, legadas y actualización de clave.

### T13: Sufijo efectivo en ClaveGrupo tras re-ranurar
- **id**: T13
- **fase**: 2387
- **depends_on**: [T11]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
- **descripción**: Conservar el sufijo de grupo efectivo existente en ClaveGrupo (o fallback a Grupo/"A") al re-ranurar hacia un aula alternativa.

### T14: Validación de patrón y subunidad en vínculo destino
- **id**: T14
- **fase**: 2388
- **depends_on**: [T10]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
- **descripción**: Rechazar aulas alternativas cuyo vínculo catálogo tenga DiasPatron incompatible con la víctima o Subunidad divergente.

### T15: Diagnóstico verídico de SIN_FECHAS en asignación de cadena y cierre
- **id**: T15
- **fase**: 2389
- **depends_on**: [T13, T14]
- **location**: ConversorPDF.XLL/Academic/ + .agents/plans/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorCarril.cs`
  - `.agents/plans/REGISTRY.md`
- **descripción**: Retornar el motivo real SIN_FECHAS registrado durante AsignarCadena en lugar de ocultarlo bajo SIN_ORDEN cuando todos los órdenes fallan por falta de fechas.

### T16: Validación de subunidad contra aula destino
- **id**: T16
- **fase**: 2390
- **depends_on**: [T14]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgReparadorR6Tests.cs`
- **descripción**: Comparar Subunidad de relDestino contra aulaAlt.Subunidad en vez de aulaV.Subunidad, evitando rechazos espurios cuando el aula de origen carece de subunidad.
