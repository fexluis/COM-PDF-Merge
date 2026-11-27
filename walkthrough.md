# Indice de Historial - walkthrough.md

Escision: 2026-11-21 - commit `a790167` - trozo objetivo: 1000 lineas

Contenido historico congelado en `docs/plans/history/walkthrough/`. Orden append-only: no se reordena ni se edita.

| Archivo | Fases | Rondas | Lineas |
|---|---|---|---|
| [walkthrough_F0620_F0710.md](docs/plans/history/walkthrough/walkthrough_F0620_F0710.md) | 620-710 | 102..117 | 1239 |
| [walkthrough_F0707_F0799.md](docs/plans/history/walkthrough/walkthrough_F0707_F0799.md) | 707-799 | 145, 146, 147, 148, 149, 150 | 1233 |
| [walkthrough_F0800_F0885.md](docs/plans/history/walkthrough/walkthrough_F0800_F0885.md) | 800-885 | 151..169 | 1246 |
| [walkthrough_F0883_F0990.md](docs/plans/history/walkthrough/walkthrough_F0883_F0990.md) | 883-990 | 170..191 | 1248 |
| [walkthrough_F0990_F1107.md](docs/plans/history/walkthrough/walkthrough_F0990_F1107.md) | 990-1107 | 192..205 | 1246 |
| [walkthrough_F1103_F1347.md](docs/plans/history/walkthrough/walkthrough_F1103_F1347.md) | 1103-1347 | 232 | 1232 |
| [walkthrough_F1342_F1509.md](docs/plans/history/walkthrough/walkthrough_F1342_F1509.md) | 1342-1509 | 264..286 | 1243 |
| [walkthrough_F1510_F1629.md](docs/plans/history/walkthrough/walkthrough_F1510_F1629.md) | 1510-1629 | 287..328 | 1250 |
| [walkthrough_F1613_F1819.md](docs/plans/history/walkthrough/walkthrough_F1613_F1819.md) | 1613-1819 | 329..398 | 590 |
| [walkthrough_F1820_F1982.md](docs/plans/history/walkthrough/walkthrough_F1820_F1982.md) | 1820-1982 | 399..449 | 1247 |
| [walkthrough_F1983_F2020.md](docs/plans/history/walkthrough/walkthrough_F1983_F2020.md) | 1983-2020 | 450..456 | 137 |
| [walkthrough_F2021_F2267.md](docs/plans/history/walkthrough/walkthrough_F2021_F2267.md) | 2021-2267 | 457..507 | 995 |
| [`walkthrough.md`](walkthrough.md) (vivo) | posteriores al corte | ultimas 5 | 107 |

## Busqueda por fase
Toda `Fase N` se localiza en el primer archivo cuyo rango la contenga; ante solape de rangos, usar la columna Rondas.

## Citas historicas
Las citas `walkthrough.md:NNNN` anteriores al corte apuntan al archivo monolitico; su ubicacion actual esta en [LINE_MAP.md](docs/plans/history/LINE_MAP.md).

---

## Ronda 508 — Criterio de partición de origen (re-emisión) (Fases 2268–2271)

### Resumen de la Ronda
- **Fase 2268:** Re-emisión de `AgCriterioParticion` y guardas `NoEmpeora`/`NoEmpeoraSinMakespan` con control de infactibles, sesiones ubicadas, makespan y materias no alojables.
- **Fase 2269:** Extensión de signaturas en `AgPlanificadorOleadas.cs` y `AgEvaluadorDemandaVentana.cs` para propagar `AgCriterioParticion`.
- **Fase 2270:** Paridad de criterio entre `AgEscaladorRanuraAula.cs` y `AgValidadorVentana.cs` (`PreferirOrigen` con opción, `Estricto` en referencia).
- **Fase 2271:** Trazabilidad en log de oleadas (`criterio=`) y matriz de casos de prueba unitaria.

### Verificación de Principios (Ronda 508)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2268 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:85-147` |
| F2269 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:30-59, 236, 357-364`, `AgEvaluadorDemandaVentana.cs:98, 131-133` |
| F2270 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:502-586`, `AgValidadorVentana.cs:353-355, 428-429` |
| F2271 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:377-391`, `AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs` |

---

## Ronda 509 — Cierre de 2258 y diagnóstico de aplicación (Fases 2272–2274)

### Resumen de la Ronda
- **Fase 2272:** Ajuste de `destinoEsperado` en postcondición same-window (`Desde == Hacia`), omitiendo comprobación de sustitución de nodo cuando no hay cambio de ventana.
- **Fase 2273:** Emisión anticipada de log `APLICAR` previa a escrituras SQL y reporte detallado de `APLICAR_FALLO` en caso de excepción.
- **Fase 2274:** Limpieza de flags y sellos residuales de partición en `AgEscaladorRanuraAula.Decision` y `HuellaAccion`.

### Verificación de Principios (Ronda 509)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2272 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:898-1001`, `AgHuellaEscaladoPersistido.cs:100-145` |
| F2273 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:925, 996` |
| F2274 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:53-56, 1145-1150` |

---

## Ronda 510 — Reclasificación en misma ventana rompe el hash esperado (Fases 2275–2279)

### Resumen de la Ronda
- **Fase 2275:** Desacople de `destino != null` en `Calcular` para reflejar mutaciones de `ag_aula_asignatura` en el hash esperado sin depender de cambio de ventana.
- **Fase 2276:** Aceptación de `destino == null` en `EstadoCompleto` y propagación de `destinoEsperado` en reaplicación idempotente y postcondición de `AgEscaladorRanuraAula.cs`.
- **Fase 2277:** Definición de `SoloParticion` para emitir advertencia de partición vigente únicamente ante decisiones intra-ventana sin reclasificaciones.
- **Fase 2278:** Registro anticipado de `APLICAR` antes de escrituras SQL y reporte de `APLICAR_FALLO` ante excepciones.
- **Fase 2279:** Supresión de tokens `pm`/`pdr` en `HuellaAccion` sin eliminar las propiedades de `Decision`.

### Verificación de Principios (Ronda 510)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2275 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgHuellaEscaladoPersistido.cs:33, 39` (Implementado) |
| F2276 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgHuellaEscaladoPersistido.cs:100-112`, `AgEscaladorRanuraAula.cs:919-943, 1010` (Implementado) |
| F2277 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:57, 857` (Implementado) |
| F2278 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:943, 1029` (Implementado) |
| F2279 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:1169` (Implementado) |

---

## Ronda 511 — `UPDATE ag_aula` se ejecuta con guarda distinta a `destinoEsperado` (Fases 2280–2283)

### Resumen de la Ronda
- **Fase 2280:** Uso de `cambiaVentana` en lugar de comparación directa en `AgEscaladorRanuraAula.cs:948` y flag `huboCambioVentana` (línea 1044).
- **Fase 2281:** Validación estricta simétrica previa de autorizaciones y consistencia de aula/subunidad en reclasificaciones (`AgEscaladorRanuraAula.cs:980-987`).
- **Fase 2282:** Actualización del prefijo de versión a `propuesta-escalado-v5` en `HuellaAccion` (`AgEscaladorRanuraAula.cs:1164`).
- **Fase 2283:** Ampliación de `SoloParticion` para admitir `ParticionDifiereDeReferencia` y unificación del texto de advertencia a `"Sin cambios persistidos; la partición se recalcula en cada corrida."`.

### Verificación de Principios (Ronda 511)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2280 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:948, 1044` (Implementado) |
| F2281 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:980-987` (Implementado) |
| F2282 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:1164`, `AgEscaladorRanuraAulaTests.cs:340` (Implementado) |
| F2283 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:57-61, 859, 1048`, `AgEscaladorRanuraAulaPersistenciaTests.cs:771, 850` (Implementado) |

---

## Ronda 512 — `Aplica` omite liberación de cuello de botella y el escalador pierde cambios persistentes (Fases 2284–2288)

### Resumen de la Ronda
- **Fase 2284:** Incorporación de condición de cuello de botella (`!ConservaPatronOrigen && IdAsignaturasBottleneck != null && IdAsignaturasBottleneck.Count > 0`) en `Decision.Aplica` (`AgEscaladorRanuraAula.cs:68-75`).
- **Fase 2285:** Suma del conteo de filas afectadas por la liberación de cuellos de botella (`afectadasBottleneck`) a `modificadas` y salvaguarda en post-loop (`AgEscaladorRanuraAula.cs:985-995, 1070-1080`).
- **Fase 2286:** Reconciliación de `dias_ventana` en `ag_arbol_nodo` mediante consulta de `nodoActual` previo al UPDATE (`AgEscaladorRanuraAula.cs:965-985`).
- **Fase 2287:** Preservación estricta de patrones con `ConservaPatronOrigen == true`, desvinculando la liberación del flag `cambiaVentana` (`AgEscaladorRanuraAula.cs:930, 985-990`).
- **Fase 2288:** Desacoplamiento de `HuellaAccionPartes` para verificar contractualmente el token inicial `propuesta-escalado-v5` (`AgEscaladorRanuraAula.cs:1184-1215`, `AgEscaladorRanuraAulaTests.cs:342-365`).

### Verificación de Principios (Ronda 512)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2284 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:68-75`, `AgEscaladorRanuraAulaTests.cs:342-358` (Implementado) |
| F2285 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:990-995, 1075`, `AgEscaladorRanuraAulaPersistenciaTests.cs:854-915` (Implementado) |
| F2286 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:965-985`, `AgEscaladorRanuraAulaPersistenciaTests.cs:917-980` (Implementado) |
| F2287 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:930, 985-990`, `AgEscaladorRanuraAulaPersistenciaTests.cs:982-1045` (Implementado) |
| F2288 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:1184-1215`, `AgEscaladorRanuraAulaTests.cs:360-380` (Implementado) |

---

## Ronda 513 — Nulabilidad incompleta, decisiones no aplicables silenciosas y huella sensible a esquema (Fases 2289–2292)

### Resumen de la Ronda
- **Fase 2289:** Null-check `d.IdAsignaturasBottleneck != null` en `AgEscaladorRanuraAula.cs:991`.
- **Fase 2290:** Advertencia descriptiva ante decisiones no aplicables que no son solo partición (`AgEscaladorRanuraAula.cs:867`).
- **Fase 2291:** Remoción de `CapacidadDespues > CapacidadAntes` en `Decision.Aplica`, habilitando persistencia de mejoras con idéntica capacidad (`AgEscaladorRanuraAula.cs:69-76`).
- **Fase 2292:** Proyección selectiva de columnas relevantes en `AgHuellaEscaladoPersistido.cs:58-65, 151-158`, independizando el hash de futuras alteraciones en el esquema.

### Verificación de Principios (Ronda 513)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2289 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:991` (Implementado) |
| F2290 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:867`, `AgEscaladorRanuraAulaPersistenciaTests.cs:1068-1110` (Implementado) |
| F2291 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:69-76`, `AgEscaladorRanuraAulaTests.cs:360-375` (Implementado) |
| F2292 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgHuellaEscaladoPersistido.cs:58-65, 151-158`, `AgEscaladorRanuraAulaPersistenciaTests.cs:1112-1138` (Implementado) |

---

## Ronda 514 — Nodo vacío/duplicado, proyección incompleta y métrica de modificadas (Fases 2293–2296)

### Resumen de la Ronda
- **Fase 2293:** Cardinalidad estricta `Count == 1` y tolerancia a nodo vacío en `AgEscaladorRanuraAula.cs:971-992`.
- **Fase 2294:** Inclusión de atributos de dominio en `AgHuellaEscaladoPersistido.cs:58-85, 115-135, 185-195` (`sqlRel`, `sqlNodo`, `sqlAula`, `CoincideOrigen`).
- **Fase 2295:** Registro métrico `modificadas += nodosAfectados` al actualizar `ag_arbol_nodo` (`AgEscaladorRanuraAula.cs:991`).
- **Fase 2296:** Proyección explícita selectiva en tablas de catálogo estáticas (`AgHuellaEscaladoPersistido.cs:75-95`).

### Verificación de Principios (Ronda 514)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2293 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:971-992`, `AgEscaladorRanuraAulaPersistenciaTests.cs:1145-1215` (Implementado) |
| F2294 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgHuellaEscaladoPersistido.cs:58-85, 115-135, 185-195` (Implementado) |
| F2295 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:991`, `AgEscaladorRanuraAulaPersistenciaTests.cs:1217-1270` (Implementado) |
| F2296 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgHuellaEscaladoPersistido.cs:75-95`, `AgEscaladorRanuraAulaPersistenciaTests.cs:1272-1300` (Implementado) |

---

## Ronda 515 — Allowlists de catálogo incompletas y sondeo de esquema repetido (Fases 2297–2298)

### Resumen de la Ronda
- **Fase 2297:** Cobertura completa de columnas de dominio existentes en `catalogoCols` (`asignaturas`, `ag_bloque_horario`, `ag_arbol`, `ag_arbol_nodo`), garantizando detección de cambios en demanda, ventanas y perfiles.
- **Fase 2298:** Cache concurrente `_cacheColumnas` por `(ConnectionString, PRAGMA schema_version, table)` en `AgHuellaEscaladoPersistido.cs`, evitando overhead de sondeo repetitivo en bucles de evaluación/persistencia.

### Verificación de Principios (Ronda 515)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2297 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgHuellaEscaladoPersistido.cs:58-96`, `AgHuellaEscaladoTests.cs:293-330` (Implementado) |
| F2298 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgHuellaEscaladoPersistido.cs:15-35`, `AgHuellaEscaladoTests.cs:293-330` (Implementado) |

---

## Ronda 516 — Allowlists de catálogo usan nombres que no existen en el esquema real (Fases 2299–2300)

### Resumen de la Ronda
- **Fase 2299:** Sustitución de nombres sintéticos en allowlists por los identificadores reales de esquema de producción (`asignaturas.idprograma`, `int_horaria`, etc.; `ag_bloque_horario.dia_semana`; `ag_arbol.manual_ventana`, etc.) en `AgHuellaEscaladoPersistido.cs:97-106`.
- **Fase 2300:** Alineación de tablas sintéticas de test con el esquema real en `AgEscaladorRanuraAulaPersistenciaTests.cs:131-155` e implementación de pruebas para verificar alteración de huella al mutar campos de catálogo.

### Verificación de Principios (Ronda 516)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2299 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgHuellaEscaladoPersistido.cs:97-106` (Implementado) |
| F2300 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAulaPersistenciaTests.cs:131-155, 1392-1447` (Implementado) |

---

## Ronda 517 — Cobertura explícita de idprograma en pruebas de huella (Fase 2301)

### Resumen de la Ronda
- **Fase 2301:** Incorporación de la prueba unitaria `Huella_CambioIdProgramaAsignatura_AlteraHash` en `AgEscaladorRanuraAulaPersistenciaTests.cs:1420-1447`, garantizando cobertura específica para cambios en `idprograma`.

### Verificación de Principios (Ronda 517)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2301 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAulaPersistenciaTests.cs:1420-1447` (Implementado) |

---

## Ronda 518 — Selección autoritativa de ventana y escalado R5-aware (Fases 2302–2306)

### Resumen de la Ronda
- **Fase 2302:** Creación de pruebas unitarias RED en `AgDimensionesPerfilAulaTests.cs` y `AgConcurrenciaCohorteTests.cs`, y tests de escalado ante infactibilidad R5 en `AgEscaladorRanuraAulaTests.cs`.
- **Fase 2303:** Implementación de `AgDimensionesPerfilResolver` (`Aplicar` y `Detectar`) y resolución autoritativa de dimensiones del perfil sobre `ctx.Aulas` en `AgFormGeneradorHorarios.cs`.
- **Fase 2304:** Implementación de modelo único `AgConcurrenciaCohorte` y `AgResultadoConcurrencia`, unificación de evaluación R5 en `AgEvaluadorDemandaVentana` (`PasoMejora` y `Evaluar`) y refactor de R5 en `AgValidadorReglas`.
- **Fase 2305:** Habilitación de escalado R5-aware en `AgEscaladorRanuraAula`, guarda de omisión considerando `R5Insatisfacible`, trazabilidad `R5_SIN_ESCALON` y condición de mejora R5 en `Decision.Mejora`.
- **Fase 2306:** Implementación de `SincronizarLoteAsync` en `AgAulaDimensionesSync` y preflight de reconciliación en `AgFormGeneradorHorarios.cs`.

### Verificación de Principios (Ronda 518)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2302 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDimensionesPerfilAulaTests.cs`, `AgConcurrenciaCohorteTests.cs`, `AgEscaladorRanuraAulaTests.cs:398-418` (Implementado) |
| F2303 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDimensionesPerfilResolver.cs`, `AgContextoSeleccionDto.cs:62-72`, `AgFormGeneradorHorarios.cs:1128-1138` (Implementado) |
| F2304 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConcurrenciaCohorte.cs`, `AgResultadoConcurrencia.cs`, `AgResultadoEvaluacionVentana.cs:39-44`, `AgEvaluadorDemandaVentana.cs:81-105, 195-207`, `AgValidadorReglas.cs:266-278` (Implementado) |
| F2305 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:556, 574, 651, 679-684, 712` (Implementado) |
| F2306 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaDimensionesSync.cs:160-198`, `AgFormGeneradorHorarios.cs:1115-1127` (Implementado) |

---

## Ronda 519 — R5 no evaluado tratado como factible, lote no atómico y dimensiones inválidas (Fases 2307–2310)

### Resumen de la Ronda
- **Fase 2307:** Exigencia estricta de `despues.R5Evaluado && despues.R5Factible` en `AgEvaluadorDemandaVentana.PasoMejora` y `AgEscaladorRanuraAula.Decision.Mejora`.
- **Fase 2308:** Sincronización batch atómica en `AgAulaDimensionesSync.SincronizarLoteAsync` bajo una única transacción y conexión con rollback total ante error.
- **Fase 2309:** Validación estricta de dimensiones de perfil (`DiasVentana`, `Jornada`, `HorasDia == 4 || 8`) en `AgDimensionesPerfilResolver` (`Aplicar` y `Detectar`).
- **Fase 2310:** Validación unívoca de cohorte en `AgConcurrenciaCohorte.Calcular` evitando comparaciones erróneas de múltiples cohortes contra un solo `maxCarriles`.

### Verificación de Principios (Ronda 519)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2307 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs:88-89`, `AgEscaladorRanuraAula.cs:680, 716-717`, `AgEscaladorRanuraAulaTests.cs:410-417` (Implementado) |
| F2308 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaDimensionesSync.cs:163-228` (Implementado) |
| F2309 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDimensionesPerfilResolver.cs:31-38, 65-72`, `AgDimensionesPerfilAulaTests.cs:53-68` (Implementado) |
| F2310 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConcurrenciaCohorte.cs:28-33`, `AgConcurrenciaCohorteTests.cs:76-90` (Implementado) |

---

## Ronda 520 — Fixture de resolución inválido, guard de cohorte unívoco y cobertura batch (Fases 2311–2313)

### Resumen de la Ronda
- **Fase 2311:** Actualización del fixture de prueba en `AgDimensionesPerfilAulaTests.cs` con dimensiones válidas (`Jornada="TARDE"`, `HorasDia=4`) satisfaciendo el validador `perfilValido`.
- **Fase 2312:** Normalización exhaustiva de cohorte con `StringComparer.Ordinal` en `AgConcurrenciaCohorte.Calcular` y adición de test de mezcla de cohortes vacías/no vacías.
- **Fase 2313:** Validación de existencia previa en `AgAulaDimensionesSync.SincronizarLoteAsync` con rollback ante aulas inexistentes y suite completa de pruebas transaccionales en `AgAulaDimensionesSyncTests.cs`.

### Verificación de Principios (Ronda 520)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2311 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDimensionesPerfilAulaTests.cs:14-36` (Implementado) |
| F2312 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConcurrenciaCohorte.cs:28-40`, `AgConcurrenciaCohorteTests.cs:88-102` (Implementado) |
| F2313 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaDimensionesSync.cs:195-230`, `AgAulaDimensionesSyncTests.cs:181-290` (Implementado) |

---

## Ronda 521 — Tolerancia post-commit en MarkDirty y validación de inversión de fechas en concurrencia (Fases 2314–2315)

### Resumen de la Ronda
- **Fase 2314:** Blindaje post-commit de `MarkDirty()` en `AgAulaDimensionesSync.SincronizarLoteAsync` con logging de advertencia.
- **Fase 2315:** Detección de intervalos temporales invertidos en `AgConcurrenciaCohorte.Calcular` y adición de test unitario.

### Verificación de Principios (Ronda 521)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2314 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaDimensionesSync.cs:236-249` (Implementado) |
| F2315 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConcurrenciaCohorte.cs:38-48`, `AgConcurrenciaCohorteTests.cs:90-102` (Implementado) |

---

## Ronda 522 — Consistencia post-commit en métodos unitarios de sincronización (Fase 2316)

### Resumen de la Ronda
- **Fase 2316:** Blindaje post-commit de `MarkDirty()` en `AgAulaDimensionesSync.SincronizarAsync` y `SincronizarConfirmadoAsync`.

### Verificación de Principios (Ronda 522)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2316 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaDimensionesSync.cs:69-78, 152-161` (Implementado) |

---

## Ronda 523 — Consolidar R6 como regla canónica de carril y retirar R13 hard (Fases 2317–2319)

### Resumen de la Ronda
- **Fase 2317:** Implementación de suite de pruebas para semántica de carril en `AgValidadorReglasOleadasTests.cs`, verificando tolerancia a frontera compartida, detección de solape real por R6, detección de solape no adyacente en orden curricular all-pairs, no emisión de R13 y sincronización del test del realineador.
- **Fase 2318:** Sustitución de comparación adyacente en R6 por all-pairs estricto en `AgValidadorReglas.cs`, permitiendo fronteras coincidentes sin solape. Retiro completo de la emisión de `R13_OleadasDispersas` y actualización documental en `AgReglaViolacion.cs` y `AgPesosFitness.cs`.
- **Fase 2319:** Cierre documental e integración en `REGISTRY.md` (fila R-195) y actualización de archivos vivos con alineación a contratos canónicos.

### Verificación de Principios (Ronda 523)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2317 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasOleadasTests.cs:348-435` (Implementado) |
| F2318 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:314-345, 389-391`, `AgReglaViolacion.cs:37`, `AgPesosFitness.cs:27` (Implementado) |
| F2319 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `.agents/plans/REGISTRY.md:27-28, 232`, `implementation_plan.md`, `walkthrough.md` (Implementado) |

---

## Ronda 524 — Alineación de HaySolape con R6 estricto y aislamiento de fixture de frontera compartida (Fases 2320–2321)

### Resumen de la Ronda
- **Fase 2320:** Reemplazo de la evaluación con `<=` y bucle adyacente en `AgPlanificadorOleadas.HaySolape` por comparación all-pairs con límites estrictos (`Ini < Fin && Fin > Ini`), armonizando el realineador con la regla canónica R6 de carril. Pruebas de no alteración ante frontera compartida integradas en suites de oleadas y planificador.
- **Fase 2321:** Aislamiento del fixture `R6_FronteraCompartidaMismoCarril_NoEsViolacionDura` en `AgValidadorReglasOleadasTests.cs`, unificando docente y diferenciando bloques para prevenir infracciones de R10 o R2/R14 en la fecha común, asegurando `HardViolations == 0`.

### Verificación de Principios (Ronda 524)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2320 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:1147-1165`, `AgPlanificadorOleadasTests.cs:847-876`, `AgValidadorReglasOleadasTests.cs:597-613` (Implementado) |
| F2321 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasOleadasTests.cs:351-404`, `.agents/plans/REGISTRY.md:27-28, 233` (Implementado) |

---

## Ronda 525 — Reparación automática de R5 con inventario de ranuras (Fases 2322–2331)

### Resumen de la Ronda
- **Fases 2322–2323:** Creación y publicación de la rama de trabajo `reparacion-automatica-r5`, reserva de fila R-197 en `REGISTRY.md` y creación del plan formal `.agents/plans/2026-11-22-reparacion-automatica-r5-plan.md`.
- **Fases 2324–2325:** Modelado de ranuras libres/liberables/ocupadas en `AgInventarioRanuras.cs`, creación de `AgReparadorR5.cs` y tipos de soporte desacoplados (`AgSobrecargaR5.cs`, `AgResultadoReparacionR5.cs`), y extracción de helper canónico `AgValidadorReglas.MaxCarrilesCohorte`.
- **Fases 2326–2327:** Implementación de operadores C1 (desplazamiento temporal) y C2 (compresión en huecos reales) junto con V1–V6 pre-validaciones contra colisiones R6/docente/aula y reversión transaccional garantizada.
- **Fase 2328:** Inyección de hook de reparación post-oleadas en `AgGeneradorHorarios.cs`.
- **Fase 2329:** Integración de métricas de telemetría de reparación R5 en `AgResultadoCorrida.cs` y persistencia en JSON en `AgFormGeneradorHorarios.cs`.
- **Fases 2330–2331:** Cobertura de 11 tests unitarios en `AgInventarioRanurasTests.cs` y `AgReparadorR5Tests.cs`, auditoría de tareas con `audit-tasks.ps1` limpia, y certificación completa con `verify.ps1 -Mode Full` (2010 tests passed).

### Verificación de Principios (Ronda 525)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2322 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `git branch`: `reparacion-automatica-r5` (Implementado) |
| F2323 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:27-28, 234`, plan R-197 (Implementado) |
| F2324 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgInventarioRanuras.cs:1-125` (Implementado) |
| F2325 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs`, `AgSobrecargaR5.cs`, `AgResultadoReparacionR5.cs`, `AgValidadorReglas.cs:259-278` (Implementado) |
| F2326 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:125-245` (Implementado) |
| F2327 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:246-340` (Implementado) |
| F2328 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs:33, 331-366`, `AgModoReparacionR5.cs` (Implementado) |
| F2329 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgResultadoCorrida.cs:18-22`, `AgFormGeneradorHorarios.cs:3422-3430` (Implementado) |
| F2330 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgInventarioRanurasTests.cs:1-76`, `AgReparadorR5Tests.cs:1-260` (Implementado) |
| F2331 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `audit-tasks.ps1` (0 findings), `verify.ps1 -Mode Full` (2010 passed) |

---

## Ronda 526 — Reparación R5 en producción y pruebas calibradas (Fases 2332–2342)

### Resumen de la Ronda
- **Fase 2332:** Reevaluación forzada de candidatos reseteando `Estado = NoEvaluado` e invocando `EvaluarAptitud(cand, contexto, forzar: true, incluirDetalleJson: true)` con sobrecarga explícita en `IAgEvaluadorAptitud`.
- **Fases 2333–2334:** Desbloqueo del desplazamiento para asignaturas `LEGADO` manteniendo estricto respeto a `EsFijo`, y filtrado de genes a nivel de cohorte excluyendo vacantes.
- **Fases 2335–2336:** Paridad de identidad posicional por índice en clones de individuo para C3/C4 y V4; soporte de unidad de bloque dinámica, atributos de patrón e intercambio C4.
- **Fases 2337–2338:** Compresión C2 acotada a huecos reales y validación integral V1–V6 incorporando colisión de aula (R2), calendario docente global (`TerceroOcupado`), sesiones conservadas y tope R4 diario, junto con guarda post-`Sellar` de vacantes.
- **Fases 2339–2340:** Resiliencia ante excepciones con try/catch no abortivo, acumulación de mejoras en ciclo iterativo, métrica de desperdicio disjunta e inventario multi-aula.
- **Fases 2341–2342:** Sustitución de mocks en tests por evaluador real validando fixture 132 con HARD 1→0 y movimiento a 2027-04-02; persistencia de telemetría de reparación R5 previa a compuertas de factibilidad y sincronización de baseline y candados.

### Verificación de Principios (Ronda 526)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2332 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:115-119, 162-166, 212-216`, `IAgEvaluadorAptitud.cs:11` (Implementado) |
| F2333 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:304-315, 535-550` (Implementado) |
| F2334 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:330-335, 410-415, 465-470` (Implementado) |
| F2335 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:340-345, 520-535` (Implementado) |
| F2336 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:475-555` (Implementado) |
| F2337 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:385-450` (Implementado) |
| F2338 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:105-110, 520-630` (Implementado) |
| F2339 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:65-85, 230-245` (Implementado) |
| F2340 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:645-655`, `AgInventarioRanuras.cs:45-75, 125-130` (Implementado) |
| F2341 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5Tests.cs:15-180` (Implementado) |
| F2342 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:3267-3295`, `test_baseline.txt:2`, `REGISTRY.md:27-28, 235` (Implementado) |

---

## Ronda 527 — Corrección de tipos, forward shifting con ranuras libres y suite fiel 132 (Fases 2343–2352)

### Resumen de la Ronda
- **Fases 2343–2346:** Corrección de contratos y tipados en `AgReparadorR5.cs` (semántica R4 por franjas, llamada a `TerceroOcupado` con `FranjaHoraria` y solape temporal de `SesionesConservadas`) y tipado de `Festivos` como tuplas `(DateTime Inicio, DateTime Fin)` en tests.
- **Fases 2347–2349:** Forward shifting C1 guiado por ranuras libres reales (`libres` excluyendo ocupadas ajenas e integrando `Liberables`), inversión de orden ejecutando `Sellar` antes de `Validar`, y compresión C2 sobre huecos intermedios reales.
- **Fase 2350:** Fixture 132 fiel con desplazamiento a `2027-04-02` (HARD 1→0), test R6 aislado sin choque R2 de aula, y aserción estricta de secuencias en determinismo.
- **Fases 2351–2352:** Trazabilidad fidedigna en logs, contadores desacoplados, consumo de `Liberables` y sincronización de registro R-199 y plan swarm.

### Verificación de Principios (Ronda 527)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2343 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:835-855` (Implementado) |
| F2344 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:748-757` (Implementado) |
| F2345 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:805-832` (Implementado) |
| F2346 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5Tests.cs:70-75` (Implementado) |
| F2347 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:390-445` (Implementado) |
| F2348 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:115-125, 175-185, 225-235` (Implementado) |
| F2349 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:450-540` (Implementado) |
| F2350 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5Tests.cs:86-175, 335-360, 395-405` (Implementado) |
| F2351 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:140-155, 190-205, 240-255` (Implementado) |
| F2352 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:405-415`, `REGISTRY.md:27-28, 236` (Implementado) |

---

## Ronda 528 — Reparación R5 con víctima 132 calibrada y cursor multi-gen (Fases 2353–2357)

### Resumen de la Ronda
- **Fase 2353:** Calibración temporal del fixture 132 en `AgReparadorR5Tests.cs` (149 arranca 02-03, 152 arranca 02-02, 150 arranca 03-19) asegurando que el orden descendente por fecha de inicio seleccione la Asig 150 como víctima trasladada a 2027-04-02.
- **Fase 2354:** Priorización de ranuras libres reales (`inv.LibresReales`) sobre liberables y ventana ampliada a 8 inicios viables.
- **Fase 2355:** Soporte multi-gen en `GenerarDesplazamientos` con cursor temporal por gen ordenado por `DiaSemana`.
- **Fase 2356:** Cálculo consistente de desperdicio antes y después (`invPre`/`invPost`) sobre la misma cohorte reparada.
- **Fase 2357:** Registro de numeración global (Ronda 528, Fase 2353–2357, fila R-200) en `REGISTRY.md`.

### Verificación de Principios (Ronda 528)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2353 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5Tests.cs:88-160` (Implementado) |
| F2354 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:411-426` (Implementado) |
| F2355 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:432-460` (Implementado) |
| F2356 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:268-274` (Implementado) |
| F2357 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:27-28, 237` (Implementado) |

---

## Ronda 529 — Fechas propias en opciones, cursor multi-gen sin solape y desperdicio por cohorte (Fases 2358–2361)

### Resumen de la Ronda
- **Fase 2358:** Concatenación de `fechasVictima` en `opciones` con validación sobre `inv.FechasValidas` permitiendo reubicar a la víctima en las fechas que ella misma desocupa (desplazamiento a 2027-04-02 en fixture 132).
- **Fase 2359:** Control multi-gen evitando solapes de fechas entre genes de la misma asignatura y día mediante el conjunto `usadas`.
- **Fase 2360:** Métrica de desperdicio y log referenciados a la cohorte específica de la sobrecarga (`despPreCohorte`).
- **Fase 2361:** Registro de numeración global (Ronda 529, Fases 2358–2361, fila R-201) en `REGISTRY.md`.

### Verificación de Principios (Ronda 529)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2358 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:420-436`, `AgReparadorR5Tests.cs:188-191` (Implementado) |
| F2359 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:449-478` (Implementado) |
| F2360 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:110-114, 156, 277-283` (Implementado) |
| F2361 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:27-28, 238` (Implementado) |

---

## Ronda 530 — Resolución analítica de corrida 132 y preparación de commit/gates (Fase 2362)

### Resumen de la Ronda
- **Fase 2362:** Validación del flujo de resolución de `IdCorrida=132` sobre el fixture de tests (desplazamiento a `2027-04-02` con reducción de HARD 1→0) y formalización de la compuerta de cierre documental y de proceso en `REGISTRY.md` (Ronda 530, Fase 2362, fila R-202).

### Verificación de Principios (Ronda 530)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2362 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:27-28, 239` (Implementado) |























---

## Ronda 531 — Orden R6→R5, validación incremental y ripple por carril (Fases 2363–2367)

### Resumen de la Ronda
- **Fases 2363–2364:** Reparación de carriles multipasada con aceptación por peso (R6=850 &gt; R5=800) y `ClavesDuras`/`Validar` que sólo rechaza duras nuevas; guardas de vacantes contra `actual`.
- **Fase 2365:** Diagnóstico de duras no-R1/R7 con cohorte/tramo/carril en `Describir` y `ResumenCorto`.
- **Fase 2366:** Inventario por carril `AgInventarioCarril` y `AgInventarioRanuras.Carriles`.
- **Fase 2367:** Operador A (`GenerarDesplazamientosRipple`) sobre la cadena del carril, telemetría `rippleIntentos/rippleSinCola` y pruebas del fixture #135.

### Verificación de Principios (Ronda 531)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2363 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs:41-89, 393-452` (Implementado) |
| F2364 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:781-785, 905-1002` (Implementado) |
| F2365 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDiagnosticoInfactibilidad.cs:166-175, 328-369` (Implementado) |
| F2366 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgInventarioCarril.cs:1-30`, `AgInventarioRanuras.cs:24-44, 239-285` (Implementado) |
| F2367 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR5.cs:263-307, 641-774`, `AgReparadorR5Tests.cs:454-651` (Implementado) |

### Verificación
- `verify.ps1 -Mode Fast` (filtro Genética): 21/21 PASS.
- `verify.ps1 -Mode Full`: 2016/2016 PASS, 0 regresiones.
- `audit-tasks.ps1 T1–T4`: AUDIT_PASS. `T5`: 25 bloqueantes heredados en `AgFormGeneradorHorarios.cs` (preexistentes; 0 introducidos).


---

## Ronda 532 — Reparación R6 dedicada y orden Sellar→plan (Fases 2368–2370)

### Resumen de la Ronda
- **Fase 2368:** Se elimina el `Sellar` posterior a `RealinearOleadas`; era el que borraba `FechasGen` y anulaba el plan del carril.
- **Fase 2369:** `AgReparadorR6` dedicado: mueve la asignatura corta del solape después del bloque largo dentro del carril, con el inventario real.
- **Fase 2370:** Hook R6 antes de R5, telemetría `ReparacionR6*` y pruebas del carril largo + cortas.

### Verificación de Principios (Ronda 532)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2368 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs:413-416` (Implementado) |
| F2369 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:1-295`, `AgResultadoReparacionR6.cs:1-20` (Implementado) |
| F2370 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs:455-469`, `AgResultadoCorrida.cs:42-45`, `AgReparadorR6Tests.cs:1-144` (Implementado) |

### Verificación
- `verify.ps1 -Mode Fast` (R6+R5+Generador): 19/19 PASS.
- `verify.ps1 -Mode Full`: 2018/2018 PASS, 0 regresiones.
- `audit-tasks.ps1 T1–T3`: AUDIT_PASS.

---

## Ronda 533 — Planificador determinista de carril y re-ranuración R6 (Fases 2371–2374)

### Resumen de la Ronda
- **Fase 2371:** Implementación de `AgPlanificadorCarril.cs` y `AgResultadoPlanCarril.cs`: cálculo determinista de semanas requeridas vs disponibles, agrupación de cupo por día real de sesión, mutación de clones por índices posicionales, anclaje estricto de genes con `EsFijo == 1` y filtro de colisiones externas `ocupadasFuera`.
- **Fase 2372:** En `AgReparadorR6.cs`, evaluación jerárquica: 1) planificador por carril único con reinicio de detección vía `break`, 2) `GenerarReRanuracion` hacia aulas candidatas con cupo en la sede, 3) fallback de desplazamiento al final del bloque.
- **Fase 2373:** Telemetría de carril enriquecida en `AgResultadoReparacionR6.cs` (`Motivo`, `SemanasRequeridas`, `SemanasDisponibles`) y log único consolidado por carril.
- **Fase 2374:** Pruebas unitarias en `AgReparadorR6Tests.cs` cubriendo cadena #137, capacidad insuficiente, fechas ocupadas, anclaje de fijos y re-ranuración a aula alternativa.

### Verificación de Principios (Ronda 533)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2371 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorCarril.cs:1-268`, `AgResultadoPlanCarril.cs:1-17`, `ConversorPDF.XLL.csproj:798-799` (Implementado) |
| F2372 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:55-155, 235-325` (Implementado) |
| F2373 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgResultadoReparacionR6.cs:20-27`, `AgReparadorR6.cs:85-95` (Implementado) |
| F2374 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6Tests.cs:146-285`, `REGISTRY.md:27-28, 242` (Implementado) |

---

## Ronda 534 — Planificador multi-gen, orden curricular y re-ranuración restringida (Fases 2375–2379)

### Resumen de la Ronda
- **Fase 2375:** Asignación secuencial de multi-gen por `DiaSemana` en `AgPlanificadorCarril.cs`.
- **Fase 2376:** Blindaje de re-ranuración en `AgReparadorR6.cs`: coincidencia de `IdPrograma`, validación estricta de bloque y exclusión de vínculos `LEGADO/FIJO`.
- **Fase 2377:** Orden curricular prioritario y anclaje de asignaturas fijas con soporte de fechas en `FechasGen`.
- **Fase 2378:** Telemetría depurada en `AgReparadorR6.cs`: reseteo de motivo stale, conteo de intentos de planificador y deltas de HARD verídicos.
- **Fase 2379:** Formalización de plan swarm `2026-11-23-r6-planificador-carril-plan.md`, fila `R-206` en `REGISTRY.md`, test multi-día y aserciones por invariantes relacionales.

### Verificación de Principios (Ronda 534)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2375 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorCarril.cs:255-320` (Implementado) |
| F2376 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:285-350` (Implementado) |
| F2377 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorCarril.cs:70-85, 200-220` (Implementado) |
| F2378 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:50-95` (Implementado) |
| F2379 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6Tests.cs:175-320`, `REGISTRY.md:27-28, 243` (Implementado) |

---

## Ronda 535 — Validación de periodo, presupuesto O(1) y capacidad sobre libres (Fases 2380–2383)

### Resumen de la Ronda
- **Fase 2380:** Filtrado estricto por `IdPeriodo` e `IdPrograma` y verificación de `relDestino` en catálogo en `AgReparadorR6.cs`.
- **Fase 2381:** Presupuesto de intentos en planificador y búsqueda O(1) de orden curricular protegiéndolo de permutaciones.
- **Fase 2382:** Aserción estricta `SIN_FECHAS`, certificación de compuertas `verify Full` (2024 tests PASS) y actualización de baseline.
- **Fase 2383:** Cálculo de capacidad y semanas disponibles sobre fechas libres reales tras deducir ocupación externa.

### Verificación de Principios (Ronda 535)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2380 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:285-330` (Implementado) |
| F2381 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorCarril.cs:200-240`, `AgReparadorR6.cs:55-70` (Implementado) |
| F2382 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6Tests.cs:205-215`, `test_baseline.txt:1-5` (Implementado) |
| F2383 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorCarril.cs:100-150` (Implementado) |

### Verificación
- `verify.ps1 -Mode Fast` (R6+R5): 19/19 PASS.
- `verify.ps1 -Mode Full`: 2024/2024 PASS, 0 regresiones.
- `audit-tasks.ps1 T1–T5`: AUDIT_PASS (0 hallazgos).

---

## Ronda 536 — Cierre de re-ranuración y robustez de asignación (Fases 2384–2386)

### Resumen de la Ronda
- **Fase 2384:** Filtrado de aulas candidatas por `Estado` (`Disponible` o `Aprobado`) y coherencia de `Subunidad` en `AgReparadorR6.cs`.
- **Fase 2385:** Actualización completa de `ClaveGrupo` y reseteo de `PatronSemanal` e `IndiceDiaPatron` en el gen re-ranurado en `AgReparadorR6.cs`.
- **Fase 2386:** Diagnóstico `MotivoReranuracion = "SIN_VINCULO"` y batería de 3 pruebas unitarias en `AgReparadorR6Tests.cs`. Registro y envelope sincronizados.

### Verificación de Principios (Ronda 536)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2384 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:305-316` (Implementado) |
| F2385 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:407-417` (Implementado) |
| F2386 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgResultadoReparacionR6.cs:22-24`, `AgReparadorR6.cs:115-119`, `AgReparadorR6Tests.cs:425-500` (Implementado) |

### Verificación
- `verify.ps1 -Mode Fast` (R6+R5): 22/22 PASS.
- `verify.ps1 -Mode Full`: 2027/2027 PASS, 0 regresiones.
- `audit-tasks.ps1 T10–T12`: AUDIT_PASS (0 hallazgos).

---

## Ronda 537 — Cierre real de re-ranuración y diagnóstico de carril (Fases 2387–2389)

### Resumen de la Ronda
- **Fase 2387:** Extracción y preservación del sufijo de grupo efectivo en `ClaveGrupo` (`AgReparadorR6.cs`).
- **Fase 2388:** Validación de compatibilidad de `DiasPatron` y `Subunidad` exigidos por `relDestino` en `AgReparadorR6.cs`.
- **Fase 2389:** Retorno del diagnóstico verídico `SIN_FECHAS` en `AgPlanificadorCarril.cs` ante fallas en asignación secuencial y registro de Wave 537 en `REGISTRY.md`.

### Verificación de Principios (Ronda 537)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2387 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:412-421`, `AgReparadorR6Tests.cs:522-556` (Implementado) |
| F2388 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:340-345`, `AgReparadorR6Tests.cs:558-595` (Implementado) |
| F2389 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorCarril.cs:192-212, 281-348`, `REGISTRY.md:27-28, 245` (Implementado) |

---

## Ronda 538 — Validación de subunidad contra aula destino (Fase 2390)

### Resumen de la Ronda
- **Fase 2390:** Corrección de la validación de `Subunidad` en el vínculo catálogo comparándolo contra `aulaAlt.Subunidad` en `AgReparadorR6.cs` y prueba unitaria en `AgReparadorR6Tests.cs`.

### Verificación de Principios (Ronda 538)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2390 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:344-345`, `AgReparadorR6Tests.cs:597-635` (Implementado) |

---

## Ronda 539 — Día residual 5/1, catálogo de 5 días y fidelidad de partición (Fases 2391–2396)

### Resumen de la Ronda
- **Fase 2391:** Respaldo de opciones de catálogo sin singletons (`{2,3|1,4,5}` y `{1,2,5|3,4}`) para ventanas de 5 días con patrones mixtos en `AgOpcionesParticion.cs` y fallback en `AgPlanificadorOleadas.cs`.
- **Fase 2392:** Partición canónica sin singleton `AgCarriles.ParticionSinSingleton` restringida exclusivamente a 5 días y su sustitución canónica en `AgPlanificadorOleadas.cs`.
- **Fase 2393:** Construcción del patrón y fidelidad evaluada sobre `ranurasEfectivas` en lugar de la banda completa en `AgPlanificadorOleadas.cs`.
- **Fase 2394:** Normalización de patrones residuales (`ForzarPatronResidual`) en días 5 y 1, propagación a `DiasPatron` y traza diagnóstica `[AG-RESIDUAL]`.
- **Fase 2395:** Métrica `BandasSingleton` en `AgEvaluacionParticion`, compuerta de selección que minimiza singletons y prefiere menos bandas (`ThenBy(ev => ev.Bandas)`).
- **Fase 2396:** Batería de pruebas unitarias y de integración para 5 días, desempate y días residuales, adaptando `Caso76` a 2 bandas y certificando Wave 539 en `REGISTRY.md`.

### Verificación de Principios (Ronda 539)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2391 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticion.cs:48-75`, `AgPlanificadorOleadas.cs:360-366` (Implementado) |
| F2392 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCarriles.cs:41-67`, `AgPlanificadorOleadas.cs:353-359` (Implementado) |
| F2393 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:565-625, 680-705` (Implementado) |
| F2394 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:547-555, 706-728` (Implementado) |
| F2395 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:20-38, 130-145, 230-245, 275-305` (Implementado) |
| F2396 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarrilesTests.cs:220-310`, `AgParticionCarrilesIntegracionTests.cs:155-161`, `AgCarrilesTests.cs:105-160`, `AgOpcionesParticionTests.cs:275-315`, `AgPlanificadorOleadasTests.cs:980-1025` (Implementado) |

---

## Ronda 540 — Regresiones en Fases 2391–2396 (día residual, catálogo y tests) (Fases 2397–2402)

### Resumen de la Ronda
- **Fase 2397:** Eliminación de la llave de cierre redundante en `AgParticionCarrilesTests.cs:319`, restaurando la compilación limpia de la suite.
- **Fase 2398:** Restricción de `sinSingleton` en `AgPlanificadorOleadas.cs` a orígenes explícitos `MarVie` o `LunJue`, preservando neutralidad ante `Origen.Ninguno`.
- **Fase 2399:** Ampliación incondicional de `ForzarPatronResidual` en `AgPlanificadorOleadas.cs` para cualquier patrón asignado a bandas `{1,2,5}` o `{1,4,5}`, agregando test de regresión `{1,2} → {1,2,5}`.
- **Fase 2400:** Actualización de la expectativa de bandas en `Caso76_Ausencia` a 2 bandas `{4,5}|{6,7}` en `AgParticionCarrilesIntegracionTests.cs`.
- **Fase 2401:** Eliminación del descarte duro de particiones con singleton en la enumeración contigua de `AgParticionCarriles.cs`, delegando la preferencia al orden de desempate.
- **Fase 2402:** Corrección del cálculo de capacidad en `AgPlanificadorCarril.cs:138-147` comparando sesiones requeridas contra `lane.FechasValidas.Count` en vez de número de semanas. Sincronización de Wave 540 y fila `R-211` en `REGISTRY.md`.

### Verificación de Principios (Ronda 540)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2397 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarrilesTests.cs:315-321` (Implementado) |
| F2398 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:357-366` (Implementado) |
| F2399 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:710-725`, `AgPlanificadorOleadasTests.cs:1020-1045` (Implementado) |
| F2400 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarrilesIntegracionTests.cs:181-188` (Implementado) |
| F2401 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:200-218` (Implementado) |
| F2402 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorCarril.cs:138-147`, `REGISTRY.md:27-28, 248` (Implementado) |

---

## Ronda 541 — Siembra de patrones aún crea carriles de 1 día; R6 pierde índice de patrón (Fases 2403–2406)

### Resumen de la Ronda
- **Fase 2403:** Mapeo de días sueltos a bandas completas sin singleton en 5 días (`BandaSinSingletonParaDia`) en `AgCarriles.cs`, propagado a `AgPlanificadorPatronSemanal.cs` y resguardo en `AgFabricaIndividuos.cs`.
- **Fase 2404:** Política explícita `PermitirSingleton = false` en `AgParticionCarriles.cs` con filtro preferente a 0 singletons y fallback seguro ante necesidad extrema.
- **Fase 2405:** Preservación de `IndiceDiaPatron` en re-ranuración R6 según la posición de `DiaSemana` en `DiasPatron` en `AgReparadorR6.cs`.
- **Fase 2406:** Corrección semántica de telemetría de carril incorporando `SesionesDisponibles` y derivando `SemanasDisponibles` en `AgResultadoPlanCarril.cs` y `AgPlanificadorCarril.cs`. Sincronización de Wave 541 y fila `R-212` en `REGISTRY.md`.

### Verificación de Principios (Ronda 541)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2403 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCarriles.cs:64-80`, `AgPlanificadorPatronSemanal.cs:54-95, 130-150`, `AgFabricaIndividuos.cs:434-442`, `AgCarrilesTests.cs:160-175`, `AgFabricaIndividuosTests.cs:636-665` (Implementado) |
| F2404 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:288-305` (Implementado) |
| F2405 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReparadorR6.cs:415-427`, `AgReparadorR6Tests.cs:517-545` (Implementado) |
| F2406 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgResultadoPlanCarril.cs:12-20`, `AgPlanificadorCarril.cs:138-147`, `REGISTRY.md:27-28, 249` (Implementado) |

---

## Ronda 542 — Mapeo singleton de 5 días no es neutral al origen; DiasKDe y RanurasDe divergen (Fases 2407–2409)

### Resumen de la Ronda
- **Fase 2407:** Parametrización origin-aware en `AgCarriles.BandaSinSingletonParaDia` distinguiendo MarVie (`{1,4,5}` y `{2,3}`) de LunJue (`{1,2,5}` y `{3,4}`), y validación de ventana canónica `{1,2,3,4,5}`. Suite en `AgCarrilesTests.cs`.
- **Fase 2408:** Detección de origen y resolución simétrica y unificada de ranuras a partir de los días calculados en `AgPlanificadorPatronSemanal.cs`.
- **Fase 2409:** Separación de conteo de fechas (`SesionesDisponibles`) y semanas calculadas (`SemanasDisponibles => ceil(SesionesDisponibles / 2.0)`) en `AgResultadoPlanCarril.cs` y `AgPlanificadorCarril.cs`. Sincronización de Wave 542 y fila `R-213` en `REGISTRY.md`.

### Verificación de Principios (Ronda 542)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2407 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCarriles.cs:75-98`, `AgCarrilesTests.cs:165-195` (Implementado) |
| F2408 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:54-95, 120-155` (Implementado) |
| F2409 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgResultadoPlanCarril.cs:12-25`, `AgPlanificadorCarril.cs:128-220`, `REGISTRY.md:27-28, 250` (Implementado) |

---

## Ronda 543 — Origen residual ambiguo, SemanasDisponibles duplica sesiones y singleton escapa por catálogo (Fases 2410–2413)

### Resumen de la Ronda
- **Fase 2410:** Implementación de `DetectarEsMarVie` tri-estado (`bool?`) en `AgPlanificadorPatronSemanal.cs` con heurística para días 1, 4, 5 y advertencia `ORIGEN_RESIDUAL_DESCONOCIDO` ante indeterminación. Heurística default en `AgCarriles.cs:97-102`.
- **Fase 2411:** Eliminación del setter de semanas en `AgResultadoPlanCarril.cs` desacoplando `SesionesDisponibles` de `SemanasEquivalentes`, y asignación coherente en `AgPlanificadorCarril.cs`.
- **Fase 2412:** Filtrado anti-singleton de catálogo (`BandasSingleton == 0`) en `AgParticionCarriles.cs:135-155` bajo política `PermitirSingleton = false`.
- **Fase 2413:** Suite de pruebas para patrones de 1 día y origen en `AgFabricaIndividuosTests.cs` y `AgCarrilesTests.cs`. Sincronización de Wave 543 y fila `R-214` en `REGISTRY.md`.

### Verificación de Principios (Ronda 543)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2410 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:54-190`, `AgCarriles.cs:97-102` (Implementado) |
| F2411 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgResultadoPlanCarril.cs:11-23`, `AgPlanificadorCarril.cs:128-218` (Implementado) |
| F2412 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:135-155` (Implementado) |
| F2413 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuosTests.cs:660-705`, `AgCarrilesTests.cs:184-188`, `REGISTRY.md:27-28, 251` (Implementado) |

---

## Ronda 544 — La siembra automática sin AgAulaAsignatura sigue produciendo {5} (Fases 2414–2416)

### Resumen de la Ronda
- **Fase 2414:** Retorno determinista de LunJue (`false`) para siembra automática sin `aa` o sin patrón declarado en `AgPlanificadorPatronSemanal.cs`, garantizando que carriles residuales en 5 días mapeen a `{1,2,5}`.
- **Fase 2415:** Exclusión de días 2, 3 y 4 de la inferencia heurística sin origen declarado en `DetectarEsMarVie` y `AgCarriles.BandaSinSingletonParaDia`, emitiendo `ORIGEN_DIA_NO_RESIDUAL_DESCONOCIDO`.
- **Fase 2416:** Suite de pruebas en `AgFabricaIndividuosTests.cs` comprobando que asignaturas automáticas tengan `RanurasPatron` multi-día y que patrones no residuales de 1 día conserven su configuración sin alteración. Sincronización de Wave 544 y fila `R-215` en `REGISTRY.md`.

### Verificación de Principios (Ronda 544)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2414 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:170-196`, `AgFabricaIndividuosTests.cs:662-675` (Implementado) |
| F2415 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:54-75, 186-195`, `AgCarriles.cs:97-108` (Implementado) |
| F2416 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuosTests.cs:676-735`, `AgCarrilesTests.cs:184-190`, `REGISTRY.md:27-28, 252` (Implementado) |

---

## Ronda 545 — Patrones de 3 días se normalizan a 2; 2/3/4 sin origen siguen en singleton (Fases 2417–2418)

### Resumen de la Ronda
- **Fase 2417:** Aceptación de patrones residuales explícitos de 3 días (`{1,2,5}` y `{1,4,5}`) en `AgPlanificadorPatronSemanal.DiasKDe` evitando normalización truncada, y erradicación de singletons para 2, 3 y 4 sin origen vía fallback determinista no singleton en `AgCarriles.cs`.
- **Fase 2418:** Actualización de la suite de pruebas en `AgFabricaIndividuosTests.cs` y `AgCarrilesTests.cs` verificando la resolución de 2 y 3 días para patrones no residuales y la conservación íntegra de patrones residuales de 3 días. Sincronización de Wave 545 y fila `R-216` en `REGISTRY.md`.

### Verificación de Principios (Ronda 545)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2417 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:40-145`, `AgCarriles.cs:97-105` (Implementado) |
| F2418 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuosTests.cs:696-745`, `AgCarrilesTests.cs:184-190`, `REGISTRY.md:27-28, 253` (Implementado) |

---

## Ronda 546 — Test de fábrica cuenta genes incorrectamente; semanas de carril siguen asumiendo 2 días (Fases 2419–2420)

### Resumen de la Ronda
- **Fase 2419:** Corrección en `AgFabricaIndividuosTests.cs` esperando 7 genes totales (uno por ranura para las 3 asignaturas con carriles `{1,2}`, `{3,4}` y `{1,2,5}`) y 3 asignaturas distintas sin singletons.
- **Fase 2420:** `SesionesPorSemana` dinámico en `AgResultadoPlanCarril.cs` y `AgPlanificadorCarril.cs` con cálculo proporcional en `SemanasEquivalentes` y `SemanasRequeridasEquivalentes`. Prueba unitaria en `AgReparadorR6Tests.cs`. Sincronización de Wave 546 y fila `R-217` en `REGISTRY.md`.

### Verificación de Principios (Ronda 546)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2419 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuosTests.cs:670-696` (Implementado) |
| F2420 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgResultadoPlanCarril.cs:11-30`, `AgPlanificadorCarril.cs:120-220`, `AgReparadorR6Tests.cs:663-685`, `REGISTRY.md:27-28, 254` (Implementado) |

---

## Ronda 547 — Separación de sesiones y semanas requeridas; trazabilidad de origen inferido (Fases 2421–2422)

### Resumen de la Ronda
- **Fase 2421:** `SesionesRequeridas` como propiedad mutable y `SemanasRequeridas` como propiedad calculada proporcional a `SesionesPorSemana` en `AgResultadoPlanCarril.cs` y `AgPlanificadorCarril.cs`.
- **Fase 2422:** Emisión de advertencia diagnóstica `ORIGEN_DIA_INFERIDO` en `AgPlanificadorPatronSemanal.cs` para patrones de 1 día que adoptan la heurística de fallback MarVie. Sincronización de Wave 547 y fila `R-218` en `REGISTRY.md`.

### Verificación de Principios (Ronda 547)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2421 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgResultadoPlanCarril.cs:11-30`, `AgPlanificadorCarril.cs:128-229`, `AgReparadorR6Tests.cs:662-687` (Implementado) |
| F2422 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:58-67`, `REGISTRY.md:27-28, 255` (Implementado) |

---

## Ronda 548 — Eliminación de setter ambiguo en SemanasRequeridas (Fase 2423)

### Resumen de la Ronda
- **Fase 2423:** Eliminación del setter mutable en `SemanasRequeridas` en `AgResultadoPlanCarril.cs`, forzando el uso exclusivo de `SesionesRequeridas` para asignaciones numéricas y preservando el getter de semanas calculadas. Sincronización de Wave 548 y fila `R-219` en `REGISTRY.md`.

### Verificación de Principios (Ronda 548)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2423 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgResultadoPlanCarril.cs:14-23`, `REGISTRY.md:27-28, 256` (Implementado) |

---

## Ronda 549 — Registro diagnóstico de inferencia de origen para días no residuales (Fase 2424)

### Resumen de la Ronda
- **Fase 2424:** Emisión estandarizada de advertencia diagnóstica `ORIGEN_DIA_NO_RESIDUAL` en `AgPlanificadorPatronSemanal.cs` ante inferencia de MarVie en patrones `{2}`, `{3}`, `{4}` sin origen declarado. Sincronización de Wave 549 y fila `R-220` en `REGISTRY.md`.

### Verificación de Principios (Ronda 549)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2424 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:58-75`, `REGISTRY.md:27-28, 257` (Implementado) |

---

## Ronda 550 — Integración de catálogo y canónicas sin singleton en AgPlanificadorOleadas (Fases 2425–2427)

### Resumen de la Ronda
- **Fase 2425:** Integración en `AgPlanificadorOleadas.cs` de `AgCarriles.ParticionSinSingleton` sobre `canonicas` según `DetectarOrigen`, junto con la validación de aplicabilidad (`AgParticionCarriles.EsOpcionAplicable`) y resolución de opciones directas o de fallback sin singleton (`ResolverFallbackSinSingleton`).
- **Fase 2426:** Implementación de pruebas unitarias en `AgPlanificadorOleadasTests.cs` validando que `Construir` en `LUN-VIE` resuelve a 2 bandas sin singleton y que en `4,5,6,7` aplica catálogo de fin de semana.
- **Fase 2427:** Registro y trazabilidad de Wave 550 (Fases 2425–2427) en `REGISTRY.md` fila `R-221`.

### Verificación de Principios (Ronda 550)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2425 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:352-378` (Implementado) |
| F2426 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs:985-1045` (Implementado) |
| F2427 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:27-28, 258` (Implementado) |

---

## Ronda 551 — Validación de aplicabilidad en opción directa y aserciones de catálogo real (Fases 2428–2429)

### Resumen de la Ronda
- **Fase 2428:** Validación previa de aplicabilidad mediante `AgParticionCarriles.EsOpcionAplicable` sobre `opcionDirecta` en `AgPlanificadorOleadas.cs` antes de asignarla a `opciones`, permitiendo la ejecución limpia de `ResolverFallbackSinSingleton` y emisión de telemetría diagnóstica ante opciones directas inaplicables.
- **Fase 2429:** Refuerzo de pruebas en `AgPlanificadorOleadasTests.cs` con aserción exacta de días para catálogo de 4 días `{5,6}` y `{4,6,7}`, validación de partición de 5 días mixtos (`2.3|1.4.5` o `1.2.5|3.4`) y verificación de fallback activo ante opciones inválidas. Sincronización de Wave 551 y fila `R-222` en `REGISTRY.md`.

### Verificación de Principios (Ronda 551)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2428 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:374-395` (Implementado) |
| F2429 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs:985-1080`, `REGISTRY.md:25-28, 259` (Implementado) |

---

## Ronda 552 — Desempate de catálogo admitiendo bandas equivalentes sin empeoramiento (Fase 2430)

### Resumen de la Ronda
- **Fase 2430:** Modificación en `AgParticionCarriles.cs` de la guarda de selección de catálogo para aceptar opciones con igual número de bandas (`Bandas <= referencia.Bandas`), permitiendo que `5,6M|4,6T,7` gane frente a la canónica de 2 bandas en `JUE-DOM` (`4,5,6,7`) sin empeorar métricas. Sincronización de Wave 552 y fila `R-223` en `REGISTRY.md`.

### Verificación de Principios (Ronda 552)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2430 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:144-157`, `REGISTRY.md:25-28, 260` (Implementado) |

---

## Ronda 553 — Comparación global catálogo vs enumeración y desempate por catálogo (Fase 2431)

### Resumen de la Ronda
- **Fase 2431:** Supresión del cortocircuito en catálogo en `AgParticionCarriles.cs`, ejecutando la enumeración contigua y resolviendo mediante `Comparar` simétrico. En empate exacto se selecciona el catálogo; si la enumeración es superior, se selecciona ésta con `catalogo_fallback`. Se añaden pruebas de empate y mejora contigua en `AgParticionCarrilesOpcionesTests.cs`. Sincronización de Wave 553 y fila `R-224` en `REGISTRY.md`.

### Verificación de Principios (Ronda 553)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2431 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:120-170, 286-370`, `AgParticionCarrilesOpcionesTests.cs:570-665`, `REGISTRY.md:25-28, 261` (Implementado) |

---

## Ronda 554 — Cobertura de empate, métrica conjunta de candidatas y catalogo_fallback (Fase 2432)

### Resumen de la Ronda
- **Fase 2432:** Unificación del cálculo de `CandidatasEvaluadas` en `AgParticionCarriles.cs` como la suma total de opciones y particiones evaluadas (`evaluadasOpciones + evaluadas`), y adición de pruebas en `AgParticionCarrilesOpcionesTests.cs` validando empate exacto con catálogo, victoria de enumeración por desempate de singleton y consistencia métrica. Sincronización de Wave 554 y fila `R-225` en `REGISTRY.md`.

### Verificación de Principios (Ronda 554)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2432 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:338-356`, `AgParticionCarrilesOpcionesTests.cs:570-715`, `REGISTRY.md:25-28, 262` (Implementado) |

---

## Ronda 555 — Política catálogo-only, cobertura sin enumeración, escalado obligatorio y fijación de patrón de 3 días (Fases 2433–2437)

### Resumen de la Ronda
- **Fase 2433:** Configuración de `PermitirEnumeracion = false` en `AgParticionCarriles.cs` y supresión de enumeración contigua en `Seleccionar`. Retorno determinista de `sin_catalogo` cuando el catálogo no satisface condiciones y emisión de `I5_SIN_CATALOGO_AMPLIAR` en `AgPlanificadorOleadas.cs`.
- **Fase 2434:** Cobertura determinista en `AgOpcionesParticion.cs` para `4,5,6,7` (`5,6M|4,6T,7`), detección de patrones de 1 día como señales de origen y ordenamiento por frecuencia en fallback mixto sin singleton en 5 días (`2,3|1,4,5` vs `1,2,5|3,4`).
- **Fase 2435:** Implementación del bucle de escalado obligatorio en `AgPlanificadorOleadas.ConstruirInterno` sobre `AgEscaleraVentanas.Siguiente` hasta `MaxPasos` ante `I5_SIN_CATALOGO_AMPLIAR` (Opción 2A).
- **Fase 2436:** Erradicación de singletons de 1 día en 5 días en `AgPlanificadorPatronSemanal.cs`, `AgPlanificadorOleadas.cs` y `AgFabricaIndividuos.cs`: fijación de 3 días completos en bandas `{1,2,5}` y `{1,4,5}` (Opción 1A).
- **Fase 2437:** Pruebas unitarias de política catálogo-only, catálogo de fin de semana, escalado de `VIE-DOM` y patrones de 3 días en `AgParticionCarrilesOpcionesTests.cs` y `AgPlanificadorOleadasTests.cs`. Registro en `REGISTRY.md` (Wave 555, fila `R-226`).

### Verificación de Principios (Ronda 555)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2433 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:97-185`, `AgPlanificadorOleadas.cs:450-465` (Implementado) |
| F2434 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticion.cs:35-155` (Implementado) |
| F2435 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:91-140` (Implementado) |
| F2436 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:122-130`, `AgPlanificadorOleadas.cs:645-655`, `AgFabricaIndividuos.cs:424-440` (Implementado) |
| F2437 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarrilesOpcionesTests.cs:660, 770, 785-910`, `AgPlanificadorOleadasTests.cs:1115-1165`, `REGISTRY.md:27-28, 263` (Implementado) |

---

## Ronda 556 — Compatibilidad de enumeración por defecto, visibilidad de detector y sincronización de ranuras expandidas (Fases 2438–2440)

### Resumen de la Ronda
- **Fase 2438:** Restitución de `permitirEnumeracion = true` como valor por defecto en `AgParticionCarriles.Seleccionar` para preservar compatibilidad en tests existentes, y pase de `permitirEnumeracion: false` en `AgPlanificadorOleadas.cs` para mantener la política catálogo-only en producción.
- **Fase 2439:** Modificación de visibilidad a `internal static bool? DetectarEsMarVie(AgAulaAsignatura aa)` en `AgPlanificadorPatronSemanal.cs`, corrigiendo error CS0122 en `AgFabricaIndividuos.cs`.
- **Fase 2440:** Sincronización de `ranurasPatron` al expandir `diasK` a banda completa en `AgFabricaIndividuos.cs` evitando que genes de 3 días queden con ranuras de 1 día, y test unitario en `AgFabricaIndividuosTests.cs`. Registro en `REGISTRY.md` (Wave 556, fila `R-227`).

### Verificación de Principios (Ronda 556)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2438 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:103`, `AgPlanificadorOleadas.cs:457` (Implementado) |
| F2439 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:185` (Implementado) |
| F2440 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:421-448`, `AgFabricaIndividuosTests.cs:780-815`, `REGISTRY.md:27-28, 264` (Implementado) |

---

## Ronda 557 — Cobertura de catálogo para MIE-DOM y completitud de escalera (Fase 2441)

### Resumen de la Ronda
- **Fase 2441:** Adición de opción de catálogo determinista sin singleton en `AgOpcionesParticion.cs` para ventana `MIE-DOM` (`3,4,5,6,7`), cubriendo ranuras de sábado (6M/6T) sin dejar ranuras sobrantes ni caer en `I5_SIN_CATALOGO_AGOTADO`. Cobertura de tests en `AgOpcionesParticionTests.cs` y `AgPlanificadorOleadasTests.cs`. Registro en `REGISTRY.md` (Wave 557, fila `R-228`).

### Verificación de Principios (Ronda 557)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2441 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticion.cs:40-52, 195-205`, `AgOpcionesParticionTests.cs:320-345`, `AgPlanificadorOleadasTests.cs:1165-1190`, `REGISTRY.md:27-28, 265` (Implementado) |

---

## Ronda 558 — Aserción exacta de partición MIE-DOM y cobertura de variante sin 6T (Fase 2442)

### Resumen de la Ronda
- **Fase 2442:** Fortalecimiento de aserciones en `AgPlanificadorOleadasTests.cs` validando la clave de partición exacta `5.6|3.4.6.7` y las franjas canónicas 6M y 6T. Incorporación de prueba para `3,4,5,6,7` sin franja tarde (resolviendo `5,6M|3,4,7`) en `AgOpcionesParticionTests.cs`. Registro en `REGISTRY.md` (Wave 558, fila `R-229`).

### Verificación de Principios (Ronda 558)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2442 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs:1185-1196`, `AgOpcionesParticionTests.cs:335-370`, `REGISTRY.md:27-28, 266` (Implementado) |














