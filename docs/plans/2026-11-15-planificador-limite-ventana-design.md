# Planificador fuera de ventana — diseño de corrección

Estado: IMPLEMENTADO (empaquetado con flujo máximo de horizonte mínimo + respaldo al relleno).
Citas por símbolo; los números de línea del plan v2 están obsoletos.
Sustituye a los tres intentos revertidos (`odd/tasks/cobertura-aprobadas-i5.md`, L45).

Objetivo: que las sesiones no excedan la fecha límite esperada en los escenarios mixtos Cali III/IV, **sin romper** los invariantes que fijan otras pruebas del planificador.

Objetivo medible: `AgPlanificadorRanurasAceptacionTests.Criterio1_*` y `AgPlanificadorRanurasRegresionTests.Mixta_*` (cuarto False/True) en verde, con la familia `AgPlanificadorRanuras|AgPrioridadHorasTests` completa (64 casos, incluye `AgPlanificadorRanurasFallbackTests`) y la suite completa sin fallas.

## Correcciones respecto de la v1

| Afirmación v1 | Verificación | Corrección v2 |
|---|---|---|
| «Cae a `CompletarReferencia` cuando la referencia tiene infactibles» | **Invertida** | `CompletarReferencia` se invoca **siempre** que alguna asignatura tiene patrón (`AgPlanificadorOleadas.cs:107-111`). El test de infactibles (`:114`) condiciona el **DFS `Construir`** (`:116-124`), no el completado |
| Carril único de la automática, citando `:435-479` | **Confirmado**, cita excedida | reparto por carril en `Construir` (cursores por banda); `CompletarReferencia` propaga `carrilPorId` (`:472-479`) |
| Paridad `Construir`↔`ConstruirEvaluacion` la fija `Criterio1` | **Atribución errónea** | La fijan `AgPlanificadorRanurasRegresionTests:92-100` y `Criterio2:168-178`; `Criterio1` sólo llama `Construir` (`:71-78`) |
| Orden por horas y factibilidad sólo por conteo | Confirmado con matiz | El bucle de relleno de respaldo, `AgPlanificadorRanuras.cs:433-439` exige `j > i`, así que el orden relativo sí se respeta; sólo se ignoran las fechas absolutas |

## Invariantes que deben cumplirse a la vez

| Invariante | Quién lo fija |
|---|---|
| Fin compacto: ≤ 13/05 (cuarto=False) y ≤ 19/05 (cuarto=True) | `RegresionTests:105-106`, `AceptacionTests:90` |
| Cantidad exacta de sesiones (64/68) y patrones por día | `RegresionTests:88-89,109-113`, `AceptacionTests:83` |
| Fechas exactas de las 5 asignaturas restringidas | `AceptacionTests:120` |
| Paridad `Construir`↔`ConstruirEvaluacion` | `RegresionTests:92-100`, `Criterio2:168-178` |
| Orden de inicio y `Oleada` = orden por horas | `AgPrioridadHorasTests:64-65` |
| Firma estable ante entradas permutadas | `AgPrioridadHorasTests:72-73,76` |
| **`Carril` y `Oleada` exactos de la automática y su paralela** | `AgPrioridadHorasTests:108-118` (`Construir_WhenPatronMenorCompiteConAutomatica`), con `BeEquivalentTo(evaluacion)` |
| Referencia preservada con carga pendiente mayor | `AgPrioridadHorasTests:156-167` |
| Unicidad de sesiones y contención en ventana | `AgPrioridadHorasTests:68-73` |
| Uso de ≥3 días incluyendo {1,3,5} por la automática | `Criterio3` |
| Casos `JUE-DOM` y criterios 2–7 | `AceptacionTests`, `RegresionTests:116-307` |
| Respaldo sin mutación: día de patrón sin hueco ⇒ `false` y relleno previo completo | `AgPlanificadorRanurasFallbackTests:20-56` |
| Día de patrón presente en ranuras pero ausente en ventana ⇒ relleno cubre lo disponible | `AgPlanificadorRanurasFallbackTests:59-83` |

El invariante de `Carril`/`Oleada` es el más restrictivo y **no** estaba en la v1.

## Intentos descartados (no repetir)

| Intento | Efecto observado |
|---|---|
| Reordenar `restantes` con restringidas primero | Arregla las 4 de fecha; rompe 4 de `AgPrioridadHorasTests` |
| Reservar los primeros huecos para las restringidas | Igual |
| Dos pasadas evitando días restringidos | Fechas siguen fallando; rompe `Criterio7` y `CompletarReferencia_*` |

Conclusión: cambiar el orden de llenado rompía el contrato; el empaquetado lo resuelve sin reordenar el inicio.

## Opciones de diseño (prioridad revisada)

### B. Factibilidad consciente de fecha dentro del greedy — **preferida**

Mantener el orden por horas y el orden de inicio, y sustituir relleno por `EmpaquetarPendientes` (`:160-245`) para exigir que cada asignatura pendiente conserve los huecos necesarios de su conjunto de días **dentro del horizonte compacto**, no sólo en cualquier punto del calendario.

- Ventaja: cambio local; no toca partición de carriles, `Carril`, `Oleada` ni el DFS.
- Restricción dura: **no debe adelantar el inicio** de la restringida por delante de la automática, ni desplazar las fechas exactas que fija `CompletarReferencia_*` (`:162-165`).
- Descarte si: cambia la firma estable, el orden de inicio o las fechas exactas.

### A. Repartir la automática entre carriles — **riesgosa, requiere rediseño**

Hacer que una asignatura sin patrón ocupe varios carriles.

- Coste real: reparto por carril en `Construir` (`:115-137`), planes (`:482-496`), `AgPlanOleada.Carril` como `int` único (`AgPlanOleada.cs:13`) y la reserva de `EvaluarParticion`. Las particiones se limitan a `kMin..kActual` bandas (`AgParticionCarriles.cs:111-129`): sería una **nueva clase de partición**, no un parámetro.
- Riesgo principal: hacer factible la referencia **cambia el plan ganador** del greedy al DFS, con huella de una banda de 2 ranuras por asignatura (`ValidarFechado` (`:531-605`); huella 2 ranuras = `AgCarriles.cs:13`), lo que puede alterar `Carril`/`Oleada` y el `{1,3,5}` de `Criterio3`.

### C. Compactación posterior

Descartada salvo que A y B fallen: puede violar la firma estable y la paridad.

## Orden de trabajo

1. Prueba RED propia que fije el fin compacto de forma aislada.
2. Implementar **B**; sólo evaluar **A** si B no satisface todos los invariantes.
3. En cada paso: familia `AgPlanificadorRanuras|AgPrioridadHorasTests`; al cerrar, suite completa y build.
4. Si tras dos ciclos no se cumplen todos los invariantes:
   - no usar `git checkout --` del archivo;
   - revertir sólo los símbolos nuevos (`EmpaquetarPendientes`, `RedFlujo`, `asignado`) y borrar `AgPlanificadorRanurasFallbackTests.cs`;
   - restaurar la expectativa anterior de `AgPrioridadHorasTests` antes de re-ejecutar;
   - registrar el bloqueo en `odd/tasks/cobertura-aprobadas-i5.md`.

## Criterio de cierre

- [x] Aserciones intactas salvo la excepción registrada: `CompletarReferencia_WhenMayorCargaPendiente` (fechas de asigs. 2 y 3), autorizada por decisión humana; quedan intactos cuotas, días permitidos, sesión protegida y fin compacto.
- [x] Familia `AgPlanificadorRanuras|AgPrioridadHorasTests` verde (64 casos esperados; incluye `AgPlanificadorRanurasFallbackTests`).
      Evidencia: `_evidence/20261116-gate-alt/gate-alternativo.trx` (suite completa 1584/1584, 0 fallas, 2026-11-16T17:23:32-05:00).
- [x] Suite completa sin fallas y build Debug correcto.
      Evidencia: mismo TRX (`outcome=Completed`, `failed=0`, `error=0`).

## Límites de evidencia

No afirmar sin ejecución: «61/61 verde», «4 fallas reproducidas en HEAD», «1581/1581», ni que `Criterio3`/`JUE-DOM` no cambian. El conteo 1581 proviene de un log de sesión, no de una corrida actual.

## Superficies previstas

Superficies: `AgPlanificadorRanuras.cs`, `AgPlanificadorOleadas.cs`, `AgOrdenAsignaturas.cs`, `AgPlanificadorRanurasFallbackTests.cs` y las tres pruebas del planificador.
