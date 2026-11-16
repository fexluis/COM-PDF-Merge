# Prioridad por horas en la planificación (implementación parcial)

## Purpose

Asignar las materias de cada aula y semestre por horas efectivas descendentes, con desempate por orden curricular normalizado e identificador, sin modificar las horas almacenadas.

## Changes

- `AgOrdenAsignaturas` centraliza la prioridad para oleadas, ranuras y evaluación de demanda.
- La reserva de carriles recorre las demandas una sola vez por prioridad, conservando restricciones y diagnósticos.
- El completado mantiene la protección de las materias restringidas completas de la referencia válida.
- El completado exige capacidad en el sufijo secuencial propio y capacidad residual en cualquier índice para las otras materias pendientes tras cada reserva hipotética. Conserva el fallback al primer hueco elegible; se retiró la preferencia por terminación temprana restringida.

## How to use

No hay opciones nuevas. La prioridad se aplica en los planificadores existentes. **No entregar todavía:** la verificación Full sigue fallando.

## Test coverage

Se corrigió, con autorización del padre, una sobreespecificación de `AgPrioridadHorasTests`: se retiró únicamente la aserción de orden por `FechaInicio`. La prioridad gobierna el orden de asignación (`Oleada`, cuya aserción permanece), no la cronología global de carriles o días independientes. Es una corrección del exceso del test, no una relajación del requisito. Se conservan cuotas, días restringidos, determinismo por permutación, ausencia de colisiones y ventana. Las suites externas de aceptación y regresión permanecen sin modificaciones.

## Observed results

- RED inicial: 11 aprobadas, 6 fallidas; `_evidence/20261115-200944/test.trx`.
- Con capacidad residual solamente: 15 aprobadas, 2 fallidas; CALI III termina 27/05 frente a 13/05 (14 días tarde), IV 25/05 frente a 19/05 (6 días tarde). `_evidence/20261115-201517/test.trx`.
- Con preferencia de terminación restringida: 14 aprobadas, 3 fallidas; ambos límites CALI pasan. `_evidence/20261115-201618/test.trx`.
- Filtro combinado: 70 aprobadas, 4 fallidas; `_evidence/20261115-201708/test.trx`.
- Full: 1462 aprobadas, 4 fallidas, 1466 ejecutadas; `_evidence/20261115-201739/test.trx`.

### Última corrección autorizada: sufijo propio y capacidad residual ajena

- Prioridad: 15 aprobadas, 2 fallidas; salida 3; `_evidence/20261115-202109/test.trx`.
- Filtro combinado: 68 aprobadas, 6 fallidas; salida 3; `_evidence/20261115-202133/test.trx`.
- Full: 1460 aprobadas, 6 fallidas, 1466 ejecutadas, ninguna omitida; salida 3; `_evidence/20261115-202202/test.trx`.
- Pasan las comprobaciones de orden de asignación y `Criterio7_PreservacionMaterias_PreviamenteCompletasEnParciales` (antes fallaba en línea 410).

## Remaining blockers

- Solo fallan los límites CALI, duplicados en tres suites: `AgPrioridadHorasTests.cs:229`, `AgPlanificadorRanurasAceptacionTests.cs:90` y `AgPlanificadorRanurasRegresionTests.cs:106`.
- III: 64 sesiones / 256 horas completas, déficit total 0 sesiones / 0 horas, pero termina el 27/05 frente al límite 13/05 (14 días tarde).
- IV: 68 sesiones / 272 horas completas, déficit total 0 sesiones / 0 horas, pero termina el 25/05 frente al límite 19/05 (6 días tarde).
- El TRX no cuantifica las sesiones posteriores al límite; no se atribuye un déficit horario específico al corte sin esa evidencia.
- La única corrección de producción autorizada ya se consumió. No se modificaron los límites externos ni la validación.

### Opciones acotadas propuestas, pendientes de autorización

1. Comparar candidatos por el máximo entre la terminación factible propia y los q-ésimos huecos ajenos, manteniendo ambos chequeos de capacidad y el fallback. Evitaría proteger fechas ajenas a costa de un déficit propio.
2. Buscar un horizonte mínimo factible mediante prefijos finitos del calendario y conteos de capacidad residual; asignar por horas dentro del primer prefijo que complete la demanda. Mantener las sesiones protegidas de la referencia.
3. Reparación local limitada a intercambios entre sesiones no protegidas y huecos compatibles, aceptando solo intercambios que reduzcan la fecha final y mantengan cuotas y restricciones.

Estas opciones son propuestas, no implementaciones ni garantías de éxito. Continuación sugerida: escritor, alcance `AgPlanificadorRanuras.cs`, una corrección como máximo y los mismos tres comandos de prioridad, filtro combinado y Full, sin modificar aserciones externas.
