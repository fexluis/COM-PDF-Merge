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





















