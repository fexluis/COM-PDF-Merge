# Indice de Historial - implementation_plan.md

Escision: 2026-11-21 - commit `a790167` - trozo objetivo: 1000 lineas

Contenido historico congelado en `docs/plans/history/implementation_plan/`. Orden append-only: no se reordena ni se edita.

| Archivo | Fases | Rondas | Lineas |
|---|---|---|---|
| [implementation_plan_F0151_F0272.md](docs/plans/history/implementation_plan/implementation_plan_F0151_F0272.md) | 151-272 | - | 1228 |
| [implementation_plan_F0270_F0461.md](docs/plans/history/implementation_plan/implementation_plan_F0270_F0461.md) | 270-461 | 42, 43, 44 | 1207 |
| [implementation_plan_F0457_F0633.md](docs/plans/history/implementation_plan/implementation_plan_F0457_F0633.md) | 457-633 | 44 | 1234 |
| [implementation_plan_F0634_F0779.md](docs/plans/history/implementation_plan/implementation_plan_F0634_F0779.md) | 634-779 | 44 | 1250 |
| [implementation_plan_F0777_F0964.md](docs/plans/history/implementation_plan/implementation_plan_F0777_F0964.md) | 777-964 | 44 | 1234 |
| [implementation_plan_F0965_F1364.md](docs/plans/history/implementation_plan/implementation_plan_F0965_F1364.md) | 965-1364 | 232 | 1241 |
| [implementation_plan_F1311_F1602.md](docs/plans/history/implementation_plan/implementation_plan_F1311_F1602.md) | 1311-1602 | 264..317 | 1237 |
| [implementation_plan_F1603_F1819.md](docs/plans/history/implementation_plan/implementation_plan_F1603_F1819.md) | 1603-1819 | 318..398 | 473 |
| [implementation_plan_F1820_F2020.md](docs/plans/history/implementation_plan/implementation_plan_F1820_F2020.md) | 1820-2020 | 399..456 | 1226 |
| [implementation_plan_F2021_F2267.md](docs/plans/history/implementation_plan/implementation_plan_F2021_F2267.md) | 2021-2267 | 457..507 | 994 |
| [`implementation_plan.md`](implementation_plan.md) (vivo) | posteriores al corte | ultimas 5 | 109 |

## Busqueda por fase
Toda `Fase N` se localiza en el primer archivo cuyo rango la contenga; ante solape de rangos, usar la columna Rondas.

## Citas historicas
Las citas `implementation_plan.md:NNNN` anteriores al corte apuntan al archivo monolitico; su ubicacion actual esta en [LINE_MAP.md](docs/plans/history/LINE_MAP.md).

---

## Ronda 508: Criterio de partición de origen (re-emisión) (Fases 2268–2271)

### Resumen de la Ronda
- **Fase 2268:** Re-emisión de la definición de `AgCriterioParticion` (`Estricto`, `PreferirOrigen`, `ForzarOrigen`), cálculo auxiliar `Ubicadas` y predicados de guarda `NoEmpeora`/`NoEmpeoraSinMakespan` con verificación estricta de cobertura (`Ubicadas(ev) >= Ubicadas(r)`) en `AgParticionCarriles.cs:85-147`, permitiendo que la opción de catálogo `2,3|1,4,5` no se descarte frente a referencias canónicas 0/0.
- **Fase 2269:** Propagación del parámetro `AgCriterioParticion criterio = AgCriterioParticion.Estricto` en `AgPlanificadorOleadas.Construir`, `ConstruirEvaluacion`, `ConstruirInterno` (`AgPlanificadorOleadas.cs:30-59, 236, 357-364`) y `AgEvaluadorDemandaVentana.Evaluar` (`AgEvaluadorDemandaVentana.cs:98, 131-133`).
- **Fase 2270:** Paridad y alineación de criterio entre escalador (`AgEscaladorRanuraAula.cs:502-586`) y validador (`AgValidadorVentana.cs:353-355, 428-429`), empleando `PreferirOrigen` en evaluaciones con opción de catálogo y `Estricto` en la evaluación de referencia sin opción.
- **Fase 2271:** Trazabilidad de partición en log de oleadas (`[Oleadas] ... criterio={criterio}; ...`, `AgPlanificadorOleadas.cs:377-391`) y batería de pruebas unitarias (`AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs`) para verificar selección de catálogo de origen, guardas de cobertura y fallback nominal.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2268 | `AgParticionCarriles.cs` | Planificado |
| 2269 | `AgPlanificadorOleadas.cs`, `AgEvaluadorDemandaVentana.cs` | Planificado |
| 2270 | `AgEscaladorRanuraAula.cs`, `AgValidadorVentana.cs` | Planificado |
| 2271 | `AgPlanificadorOleadas.cs`, `AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Planificado |

---

## Ronda 509: Cierre de 2258 y diagnóstico de aplicación (Fases 2272–2274)

### Resumen de la Ronda
- **Fase 2272:** Ajuste de postcondición same-window (`Desde == Hacia`) en `AgEscaladorRanuraAula.cs:898-1001` y `AgHuellaEscaladoPersistido.cs:100-145`. En decisiones solo de reclasificación o intra-ventana sin cambio de ventana, `destino` se pasa como `null` a `Calcular` y `EstadoCompleto`. `EstadoCompleto` omite la comprobación de `dias_ventana` en aula y nodo cuando `destino == null`, evitando fallos espurios de postcondición por divergencias preexistentes entre el árbol y `ag_aula`. Asimismo, `Calcular` no condiciona la sustitución de reclasificaciones ni liberación de cuellos de botella a `destino != null`.
- **Fase 2273:** Adelanto del registro de log `[ESCALADO] APLICAR` previo a la ejecución del primer `UPDATE` en `AgEscaladorRanuraAula.cs` (~línea 925) y emisión diagnóstica `[ESCALADO] APLICAR_FALLO` con aula, desde, hacia, tipo y mensaje de excepción en el bloque `catch` para garantizar trazabilidad ante fallos de escritura SQL.
- **Fase 2274:** Limpieza de propiedades residuales de partición en `AgEscaladorRanuraAula.Decision` (`ParticionMejora`, `ParticionDifiereDeReferencia`) y sus correspondientes tokens (`pm=`, `pdr=`) en `HuellaAccion`, una vez desacoplada la persistencia de decisiones intra-ventana.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2272 | `AgEscaladorRanuraAula.cs`, `AgHuellaEscaladoPersistido.cs` | Planificado |
| 2273 | `AgEscaladorRanuraAula.cs` | Planificado |
| 2274 | `AgEscaladorRanuraAula.cs` | Planificado |

---

## Ronda 510: Reclasificación en misma ventana rompe el hash esperado (Fases 2275–2279)

### Resumen de la Ronda
- **Fase 2275:** Desacople de `destino != null` para la sustitución de `liberar` (cuello de botella) y `reclasificaciones` en `AgHuellaEscaladoPersistido.Calcular` (`AgHuellaEscaladoPersistido.cs:33, 39`). Con `Desde == Hacia` (o `Hacia` vacío) y solo reclasificación, los `UPDATE` mutan `origen_patron` y `dias_patron`; al desacoplar la condición, el hash esperado refleja fielmente los cambios en `ag_aula_asignatura` sin requerir cambio de ventana.
- **Fase 2276:** Adaptación de `AgHuellaEscaladoPersistido.EstadoCompleto` para aceptar `destino == null` (`AgHuellaEscaladoPersistido.cs:100-112`). Si `destino == null`, se verifica la existencia unívoca del aula y se omite la validación de `dias_ventana` en aula y nodo. En `AgEscaladorRanuraAula.cs` (líneas 907-917 y 998) se calcula `destinoEsperado = cambiaVentana ? d.Hacia : null` tanto para la rama idempotente como para la verificación postcondición.
- **Fase 2277:** Corrección de la clasificación de advertencia `"Partición ya vigente; sin cambios persistidos."` en `AgEscaladorRanuraAula.cs:852-855`. Se introduce la propiedad `SoloParticion` (`ParticionMejora && !Aplica && Desde == Hacia && SinReclasificaciones`) evitando atribuir a partición decisiones de escalada rechazadas por falta de capacidad.
- **Fase 2278:** Reubicación de la traza `[ESCALADO] APLICAR` previo al primer `UPDATE` SQL en `AgEscaladorRanuraAula.cs` (~línea 925) e incorporación de log `[ESCALADO] APLICAR_FALLO` con referencia, desde, hacia, conserva, tipo y mensaje de error en el bloque `catch`.
- **Fase 2279:** Exclusión de los flags `pm=` y `pdr=` del cómputo de `HuellaAccion` en `AgEscaladorRanuraAula.cs:1155-1156`, manteniendo las propiedades en `Decision` para el diagnóstico de `SoloParticion` pero sin alterar el sello criptográfico de la acción.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2275 | `AgHuellaEscaladoPersistido.cs` | Implementado |
| 2276 | `AgHuellaEscaladoPersistido.cs`, `AgEscaladorRanuraAula.cs` | Implementado |
| 2277 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2278 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2279 | `AgEscaladorRanuraAula.cs` | Implementado |

---

## Ronda 511: `UPDATE ag_aula` se ejecuta con guarda distinta a `destinoEsperado` (Fases 2280–2283)

### Resumen de la Ronda
- **Fase 2280:** Sustitución de `!string.Equals(d.Desde, d.Hacia)` por `cambiaVentana` en `AgEscaladorRanuraAula.cs:948` y en el flag global `huboCambioVentana` (línea 1044). Si `Hacia` viene vacío o no nulo pero igual a `Desde`, no se ejecuta `UPDATE ag_aula SET dias_ventana` ni se modifica `ag_arbol_nodo`, manteniendo estricta paridad con `destinoEsperado = cambiaVentana ? d.Hacia : null`.
- **Fase 2281:** Validación simétrica estricta de cada reclasificación en `AgEscaladorRanuraAula.cs:980-987`. Se verifica `rec != null && rec.Autorizado && rec.IdAula == idAula && rec.AulaRef == d.AulaRef && rec.Subunidad == d.Subunidad` antes de ejecutar cualquier `UPDATE` en `ag_aula_asignatura`, arrojando `InvalidOperationException` si la propuesta fue manipulada o presenta divergencias.
- **Fase 2282:** Bump de versión del prefijo de acción en `HuellaAccion` de `propuesta-escalado-v4` a `propuesta-escalado-v5` (`AgEscaladorRanuraAula.cs:1164`) para reflejar la eliminación de los tokens `pm=` y `pdr=`.
- **Fase 2283:** Ampliación de la guarda de `SoloParticion` en `AgEscaladorRanuraAula.cs:57-61` (`(ParticionMejora || ParticionDifiereDeReferencia)`) y unificación del texto de advertencia a `"Sin cambios persistidos; la partición se recalcula en cada corrida."` en las líneas 859 y 1048. Actualización de aserciones en tests correspondientes.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2280 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2281 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2282 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaTests.cs` | Implementado |
| 2283 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs`, `AgEscaladorRanuraAulaTests.cs` | Implementado |

---

## Ronda 512: `Aplica` omite liberación de cuello de botella y el escalador pierde cambios persistentes (Fases 2284–2288)

### Resumen de la Ronda
- **Fase 2284:** Incorporación de condición de cuello de botella (`!ConservaPatronOrigen && IdAsignaturasBottleneck != null && IdAsignaturasBottleneck.Count > 0`) en `Decision.Aplica` (`AgEscaladorRanuraAula.cs:68-75`), evitando que decisiones con `Desde == Hacia` que liberan patrones queden erróneamente excluidas en `decisionesAplicables` o clasificadas como `SoloParticion`.
- **Fase 2285:** Suma del conteo de filas afectadas por la liberación de cuellos de botella (`afectadasBottleneck`) a `modificadas` en el bucle de aplicación y consideración de `huboLiberacionBottleneck` en el guard post-loop (`AgEscaladorRanuraAula.cs:985-995, 1070-1080`), previniendo que una corrida con solo liberación de cuello de botella resulte en `Modificadas=0`, `Confirmadas=0` ("Sin cambios persistidos").
- **Fase 2286:** Reconciliación en la actualización de `ag_arbol_nodo` (`AgEscaladorRanuraAula.cs:965-985`). Se consulta `dias_ventana` actual del nodo (`nodoActual`) en lugar de asumir que coincide con `ag_aula.dias_ventana` (`d.Desde`), actualizando con `COALESCE(dias_ventana,'')=@NodoActual` y eliminando rollbacks espurios por representaciones divergentes (ej. formato numérico `1,2,3,4,5` vs `MAR-VIE`).
- **Fase 2287:** Eliminación de la dependencia espuria `|| cambiaVentana` en `liberarEsperado` y en el vaciado de `ag_aula_asignatura` (`AgEscaladorRanuraAula.cs:930, 985-990`). Cuando `ConservaPatronOrigen == true`, los patrones de asignaturas se preservan intactos incluso si la ventana escala.
- **Fase 2288:** Exposición de `HuellaAccionPartes(Decision d)` como lista de preimágenes para verificación directa de la versión del prefijo (`propuesta-escalado-v5`) en pruebas unitarias (`AgEscaladorRanuraAula.cs:1184-1215`, `AgEscaladorRanuraAulaTests.cs:342-365`).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2284 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaTests.cs` | Implementado |
| 2285 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |
| 2286 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |
| 2287 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |
| 2288 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaTests.cs` | Implementado |

---

## Ronda 513: Nulabilidad incompleta, decisiones no aplicables silenciosas y huella sensible a esquema (Fases 2289–2292)

### Resumen de la Ronda
- **Fase 2289:** Blindaje de nulabilidad de `d.IdAsignaturasBottleneck != null` previo a evaluar `.Count > 0` en `AgEscaladorRanuraAula.cs:991`, previniendo excepciones `NullReferenceException` dentro de la transacción de aplicación.
- **Fase 2290:** Diferenciación explícita de advertencia en `AgEscaladorRanuraAula.cs:867` cuando no hay decisiones aplicables (`!d.Aplica`) y no corresponden a `SoloParticion`, informando `"Sin cambios persistidos; hay decisiones no aplicables (revise demanda/capacidad)."` en lugar de devolver advertencia nula silenciosa.
- **Fase 2291:** Corrección de la propiedad `Decision.Aplica` (`AgEscaladorRanuraAula.cs:69-76`) al eliminar la condición redundante y restrictiva `CapacidadDespues > CapacidadAntes`, permitiendo que mejoras válidas de distribución o reducción de demanda con igual capacidad total sean aplicadas persistiendo el cambio de ventana.
- **Fase 2292:** Proyección explícita de columnas en `AgHuellaEscaladoPersistido.cs:58-65, 151-158` para `ag_aula`, `ag_aula_asignatura` y `ag_arbol_nodo`, desensibilizando la huella SQL frente a migraciones futuras o columnas adyacentes no involucradas en el escalado.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2289 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2290 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |
| 2291 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaTests.cs` | Implementado |
| 2292 | `AgHuellaEscaladoPersistido.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |

---

## Ronda 514: Nodo vacío/duplicado, proyección incompleta y métrica de modificadas (Fases 2293–2296)

### Resumen de la Ronda
- **Fase 2293:** En `AgEscaladorRanuraAula.cs:971-992`, reemplazo de `ExecuteScalarAsync` por `QueryAsync<string>`, verificando cardinalidad exacta `nodos.Count == 1` para detectar tanto nodos duplicados como ausentes, tolerando `dias_ventana = ''` sin provocar rollback espurio.
- **Fase 2294:** Proyección explícita defensiva ampliada en `AgHuellaEscaladoPersistido.cs:58-85, 115-135, 185-195` incorporando columnas de dominio (`exclusivo`, `prioridad`, `requerida`, `subunidad`, `habilitado`, `capacidad`, `tipo_aula`, etc.) en huellas y en la validación `CoincideOrigen`.
- **Fase 2295:** Suma de `modificadas += nodosAfectados` tras la actualización de `ag_arbol_nodo` (`AgEscaladorRanuraAula.cs:991`), reflejando fielmente `Modificadas == 2` cuando se escala ventana con perfil activo.
- **Fase 2296:** Proyección selectiva defensiva en catálogos estáticos (`asignaturas`, `programas`, `ag_bloque_horario`, `festivos`, `presenciales`, `periodos`, `ag_arbol`) en `AgHuellaEscaladoPersistido.cs:75-95`, desensibilizando la huella ante columnas añadidas en migraciones de catálogos.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2293 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |
| 2294 | `AgHuellaEscaladoPersistido.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |
| 2295 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |
| 2296 | `AgHuellaEscaladoPersistido.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |

---

## Ronda 515: Allowlists de catálogo incompletas y sondeo de esquema repetido (Fases 2297–2298)

### Resumen de la Ronda
- **Fase 2297:** En `AgHuellaEscaladoPersistido.cs:58-96, 189-195`, ampliación de `catalogoCols` incorporando columnas existentes relevantes:
  - `asignaturas`: `id, id_programa, semestre, nombre, horas, int_horaria, creditos, horas_semanales, id_sede, modalidad, subunidad`.
  - `ag_bloque_horario`: `id, activo, dia_semana_numero, franja, duracion_horas, hora_inicio, hora_fin, subunidad`.
  - `ag_arbol`: `id, idperiodo, nombre, fecha_inicio, fecha_fin, ventana_origen, poblacion, generaciones, semilla, activo`.
  - `ag_arbol_nodo`: ampliación de `colsNodoDeseadas` con `capacidad, tipo_aula, es_virtual, caracteristicas_json, codigo, nombre, idperiodo, idsede, idprograma, semestre, grupo, modalidad`.
- **Fase 2298:** Optimización del sondeo de esquema mediante memoización thread-safe en `AgHuellaEscaladoPersistido.cs` (`_cacheColumnas` con clave `ConnectionString|PRAGMA schema_version|tabla`), eliminando consultas repetidas a `pragma_table_info` en llamadas consecutivas sin perder reactividad ante DDL.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2297 | `AgHuellaEscaladoPersistido.cs`, `AgHuellaEscaladoTests.cs` | Implementado |
| 2298 | `AgHuellaEscaladoPersistido.cs`, `AgHuellaEscaladoTests.cs` | Implementado |

---

## Ronda 516: Allowlists de catálogo usan nombres que no existen en el esquema real (Fases 2299–2300)

### Resumen de la Ronda
- **Fase 2299:** En `AgHuellaEscaladoPersistido.cs:97-106`, corrección y alineación de `catalogoCols` con el esquema canónico de producción:
  - `asignaturas`: reemplazo de nombres sintéticos (`id_programa`, `horas`, etc.) por las columnas reales de producción (`idprograma`, `num_nivel`, `cod_unidad`, `cod_materia`, `uni_teorica`, `metodologia`, `abreviatura`, `nucleo`, `nombre`, `creditos`, `int_horaria`, `tipoasignatura`, `tipocredito`, `prerequisito`, `subunidad`).
  - `ag_bloque_horario`: corrección de `dia_semana_numero` a `dia_semana`.
  - `ag_arbol`: adición de `manual_ventana`, `subunidad`, `tasa_cruce`, `tasa_mutacion`, `elitismo`, `estancamiento_max`.
- **Fase 2300:** En `AgEscaladorRanuraAulaPersistenciaTests.cs:131-155`, actualización de los fixtures de prueba para reproducir el esquema real de `asignaturas` y `ag_bloque_horario` (evitando drift entre allowlist y esquema), y adición de tests de alteración de hash ante mutación de campos reales (`Huella_CambioCampoExistenteAsignatura_AlteraHash`, `Huella_CambioDiaSemanaBloque_AlteraHash`).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2299 | `AgHuellaEscaladoPersistido.cs` | Implementado |
| 2300 | `AgEscaladorRanuraAulaPersistenciaTests.cs`, `AgHuellaEscaladoTests.cs` | Implementado |

---

## Ronda 517: Cobertura explícita de idprograma en pruebas de huella (Fase 2301)

### Resumen de la Ronda
- **Fase 2301:** En `AgEscaladorRanuraAulaPersistenciaTests.cs:1393-1446`, desagregación del test de alteración de hash en asignaturas implementando `Huella_CambioIdProgramaAsignatura_AlteraHash` para verificar explícitamente que cambios en `idprograma` (clave foránea de programa) invalidan la huella y previenen regresiones en `catalogoCols`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2301 | `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |

---

## Ronda 518: Selección autoritativa de ventana y escalado R5-aware (Fases 2302–2306)

### Resumen de la Ronda
- **Fase 2302:** Creación de suites de pruebas unitarias (`AgDimensionesPerfilAulaTests.cs`, `AgConcurrenciaCohorteTests.cs`, y adición en `AgEscaladorRanuraAulaTests.cs`) para verificar que las dimensiones del perfil prevalecen sobre ag_aula, el cálculo canónico de concurrencia R5 cohorte-carril y la propuesta de escalón cuando la demanda cabe pero R5 es infactible.
- **Fase 2303:** Implementación de `AgDimensionesPerfilResolver.cs` y su integración en `AgFormGeneradorHorarios.cs` (en `ConstruirContextoAsync`), asegurando que las dimensiones seleccionadas en el perfil de `ag_arbol_nodo` sobrescriban autoritativamente las dimensiones de `ctx.Aulas` en memoria con trazabilidad de desalineaciones.
- **Fase 2304:** Implementación del modelo canónico `AgConcurrenciaCohorte.cs` y DTO `AgResultadoConcurrencia.cs`. Extensión de `AgResultadoEvaluacionVentana.cs` con propiedades R5 (`R5Evaluado`, `R5Factible`, `MaxConcurrentesR5`, `CarrilesR5`, `R5Insatisfacible`). Integración de la evaluación R5 en `AgEvaluadorDemandaVentana.cs` y `PasoMejora`, y refactorización de la regla R5 en `AgValidadorReglas.cs` para compartir la detección de sobrecargas sin duplicación.
- **Fase 2305:** Adaptación de `AgEscaladorRanuraAula.cs` para considerar `R5Insatisfacible` como condición de escalado, registrar advertencia y traza `R5_SIN_ESCALON` cuando ningún escalón resuelve la infactibilidad, y actualizar `Decision.Mejora` para contemplar transiciones hacia R5 factible.
- **Fase 2306:** Incorporación de método `SincronizarLoteAsync` en `AgAulaDimensionesSync.cs` y llamada preflight en `AgFormGeneradorHorarios.cs` para reconciliar y persistir automáticamente las dimensiones desalineadas detectadas en base de datos.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2302 | `AgDimensionesPerfilAulaTests.cs`, `AgConcurrenciaCohorteTests.cs`, `AgEscaladorRanuraAulaTests.cs` | Implementado |
| 2303 | `AgDimensionesPerfilResolver.cs`, `AgContextoSeleccionDto.cs`, `AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL.csproj` | Implementado |
| 2304 | `AgConcurrenciaCohorte.cs`, `AgResultadoConcurrencia.cs`, `AgResultadoEvaluacionVentana.cs`, `AgEvaluadorDemandaVentana.cs`, `AgValidadorReglas.cs` | Implementado |
| 2305 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2306 | `AgAulaDimensionesSync.cs`, `AgFormGeneradorHorarios.cs` | Implementado |

---

## Ronda 519: R5 no evaluado tratado como factible, lote no atómico y dimensiones inválidas (Fases 2307–2310)

### Resumen de la Ronda
- **Fase 2307:** En `AgEvaluadorDemandaVentana.cs:88` y `AgEscaladorRanuraAula.cs:680, 716`, exigencia estricta de `despues.R5Evaluado && despues.R5Factible` (junto con `antes.R5Evaluado && !antes.R5Factible`) para considerar que un paso mejora R5, evitando que pasos sin evaluación R5 sean aceptados incorrectamente como mejora. Se añade prueba unitaria `Escalador_CuandoPasoNoEvaluadoR5_NoProponePaso`.
- **Fase 2308:** En `AgAulaDimensionesSync.cs:163-228`, reescritura de `SincronizarLoteAsync` utilizando una única conexión y transacción atómica compartida (`conn.BeginTransaction()`), asegurando rollback completo ante fallas y disparo unificado de `SyncContext.Scheduler.MarkDirty()`.
- **Fase 2309:** En `AgDimensionesPerfilResolver.cs:31-38, 65-72`, validación integral de dimensiones del perfil (`DiasVentana` no vacía, `Jornada` no vacía y `HorasDia` igual a 4 u 8) tanto en `Aplicar` como en `Detectar`, impidiendo sobreescritura de valores inválidos en aulas o sincronizaciones espurias. Se añade prueba unitaria `Resolver_PerfilInvalido_NoModificaAulaNiGeneraDiff`.
- **Fase 2310:** En `AgConcurrenciaCohorte.cs:28-33`, verificación estricta de cardinalidad de cohorte en `Calcular`, impidiendo mezclar cohortes con distintos números de carriles en una misma evaluación global. Se añade prueba unitaria `Calcular_ConMultiplesCohortes_LanzaArgumentException`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2307 | `AgEvaluadorDemandaVentana.cs`, `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaTests.cs` | Implementado |
| 2308 | `AgAulaDimensionesSync.cs` | Implementado |
| 2309 | `AgDimensionesPerfilResolver.cs`, `AgDimensionesPerfilAulaTests.cs` | Implementado |
| 2310 | `AgConcurrenciaCohorte.cs`, `AgConcurrenciaCohorteTests.cs` | Implementado |

---

## Ronda 520: Fixture de resolución inválido, guard de cohorte unívoco y cobertura batch (Fases 2311–2313)

### Resumen de la Ronda
- **Fase 2311:** En `AgDimensionesPerfilAulaTests.cs:14-36`, corrección del fixture de prueba `Resolver_PrevalecePerfilSobreAula_WhenDifiere` asignando `Jornada = "TARDE"` y `HorasDia = 4` explícitas, alineándolo con el contrato de validación de `perfilValido` introducido en Fase 2309.
- **Fase 2312:** En `AgConcurrenciaCohorte.cs:28-40`, normalización unívoca de claves de cohorte (`x.Cohorte ?? string.Empty` con `StringComparer.Ordinal`) en la detección de cardinalidad de cohorte y en `GroupBy`, impidiendo mezclar cohortes vacías y no vacías en un mismo cómputo. Se añade prueba `Calcular_ConCohorteVaciaYNoVacia_LanzaArgumentException`.
- **Fase 2313:** En `AgAulaDimensionesSync.cs:195-230`, incorporación de verificación previa de existencia (`sqlExiste`) dentro de la transacción por lote y rollback inmediato ante aulas inexistentes. Se agregan pruebas unitarias en `AgAulaDimensionesSyncTests.cs` para verificar commit atómico de lote, rollback ante aula inexistente y rollback total ante fallo de trigger.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2311 | `AgDimensionesPerfilAulaTests.cs` | Implementado |
| 2312 | `AgConcurrenciaCohorte.cs`, `AgConcurrenciaCohorteTests.cs` | Implementado |
| 2313 | `AgAulaDimensionesSync.cs`, `AgAulaDimensionesSyncTests.cs` | Implementado |

---

## Ronda 521: Tolerancia post-commit en MarkDirty y validación de inversión de fechas en concurrencia (Fases 2314–2315)

### Resumen de la Ronda
- **Fase 2314:** En `AgAulaDimensionesSync.cs:236-249`, protección post-commit con bloque `try-catch` alrededor de `SyncContext.Scheduler.MarkDirty()`, garantizando que si el scheduler falla o no está disponible tras persistir el lote, la operación retorne `Result.Success(totalActualizadas)` sin falsear un fallo en la base de datos.
- **Fase 2315:** En `AgConcurrenciaCohorte.cs:38-48`, validación de consistencia temporal en intervalos (`Fin >= Ini`), retornando `Evaluado=false` y `Factible=false` ante intervalos invertidos para evitar declaraciones falsas de factibilidad. Se añade prueba `Calcular_ConIntervaloInvertido_NoDeclaraFactible`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2314 | `AgAulaDimensionesSync.cs` | Implementado |
| 2315 | `AgConcurrenciaCohorte.cs`, `AgConcurrenciaCohorteTests.cs` | Implementado |

---

## Ronda 522: Consistencia post-commit en métodos unitarios de sincronización (Fase 2316)

### Resumen de la Ronda
- **Fase 2316:** En `AgAulaDimensionesSync.cs:69-78, 152-161`, alineación de `SincronizarAsync` y `SincronizarConfirmadoAsync` con el patrón post-commit tolerante de Fase 2314, protegiendo `SyncContext.Scheduler.MarkDirty()` con `try-catch` para evitar fallos falsos o retornos de error tras confirmarse el `UPDATE` en base de datos.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2316 | `AgAulaDimensionesSync.cs` | Implementado |

---

## Ronda 523: Consolidar R6 como regla canónica de carril y retirar R13 hard (Fases 2317–2319)

### Resumen de la Ronda
- **Fase 2317:** En `AgValidadorReglasOleadasTests.cs`, adición de pruebas unitarias para semántica canónica de carril: frontera compartida sin violación dura (`R6_FronteraCompartidaMismoCarril_NoEsViolacionDura`), solape real con violación dura R6 (`R6_SolapeRealMismoCarril_EsViolacionDura`), solape no adyacente con orden curricular all-pairs (`R6_SolapeNoAdyacenteConOrdenCurricularInvertido_EsViolacionDura`), ausencia de emisión de R13 en validador (`R13_NoSeEmiteEnValidador`) y actualización del test del realineador (`RealinearOleadas_CampeonConSolapesR6_EliminaViolaciones`).
- **Fase 2318:** En `AgValidadorReglas.cs:314-412`, unificación estructural del control de carril en R6 con comparación estricta all-pairs (`Ini < Fin && Fin > Ini`) evaluando `Date` para permitir fronteras compartidas consecutivas, y retiro del bloque de evaluación dura de `R13_OleadasDispersas`. Actualización de comentarios en `AgReglaViolacion.cs:37` y `AgPesosFitness.cs:27,63` documentando R13 como reservada/no emitida y retenida para compatibilidad de deserialización.
- **Fase 2319:** Verificación y sincronización de contratos de validación de carril en tests y documentación viva; mantenimiento de baseline y registro R-195 cerrado en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2317 | `AgValidadorReglasOleadasTests.cs` | Implementado |
| 2318 | `AgValidadorReglas.cs`, `AgReglaViolacion.cs`, `AgPesosFitness.cs` | Implementado |
| 2319 | `AgValidadorReglasOleadasTests.cs`, `AgValidadorReglasTests.cs`, `.agents/plans/REGISTRY.md` | Implementado |

---

## Ronda 524: Alineación de HaySolape con R6 estricto y aislamiento de fixture de frontera compartida (Fases 2320–2321)

### Resumen de la Ronda
- **Fase 2320:** En `AgPlanificadorOleadas.cs:1147-1165`, alineación del método `HaySolape` con la semántica canónica de R6, unificando la verificación en un único bucle all-pairs con comparación estricta de intervalos (`Ini < Fin && Fin > Ini`) y eliminando la evaluación `<=` inclusiva que provocaba reintentos espurios de reparación en fronteras compartidas consecutivas (`Fin == Ini`). Se agregan pruebas `RealinearOleadas_FronteraCompartida_NoCambia` en `AgPlanificadorOleadasTests.cs` y `AgValidadorReglasOleadasTests.cs`.
- **Fase 2321:** En `AgValidadorReglasOleadasTests.cs:351-403`, aislamiento del fixture de prueba `R6_FronteraCompartidaMismoCarril_NoEsViolacionDura` asignando docente idéntico (`IdTercero = 20`) y bloques horarios independientes (`IdBloqueHorario = 24` y `29`) en aula 30 y patrón 4, impidiendo falsas violaciones colaterales de R10 o R2/R14 en la fecha de frontera, e incorporando la aserción de invariante `ind.HardViolations.Should().Be(0)`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2320 | `AgPlanificadorOleadas.cs`, `AgPlanificadorOleadasTests.cs`, `AgValidadorReglasOleadasTests.cs` | Implementado |
| 2321 | `AgValidadorReglasOleadasTests.cs`, `.agents/plans/REGISTRY.md` | Implementado |

---

## Ronda 525: Reparación automática de R5 con inventario de ranuras (Fases 2322–2331)

### Resumen de la Ronda
- **Fase 2322 (T1):** Creación y publicación de rama de protección `reparacion-automatica-r5` en repositorio remoto (`origin/reparacion-automatica-r5`).
- **Fase 2323 (T2):** Reserva de numeración en `REGISTRY.md` (§2 contadores Ronda 526 / Fase 2332, §3 fila R-197) y formalización del plan de ejecución `.agents/plans/2026-11-22-reparacion-automatica-r5-plan.md` (T1..T10, Fases 2322..2331).
- **Fase 2324 (T3):** Creación del modelo `AgInventarioRanuras.cs` en `ConversorPDF.XLL/Academic/Application/Services/Genetica/`, modelando ranuras válidas, ocupadas, libres reales, liberables, cola muerta y balance por día para cálculo de desperdicio y capacidad de alojamiento.
- **Fase 2325 (T4):** Creación del componente de reparación `AgReparadorR5.cs`, modelo de sobrecarga `AgSobrecargaR5.cs`, resultado `AgResultadoReparacionR5.cs` (cumpliendo Single Type Per File) y extracción del helper canónico `AgValidadorReglas.MaxCarrilesCohorte` evitando duplicación o divergencia de cálculo de carriles máximos.
- **Fase 2326 (T5):** Implementación de operadores de forward shifting C1 (`GenerarDesplazamientos`) con ordenamiento decreciente de candidatos para resolver sobrecargas R5 hacia la cola de la ventana, y compresión C2 (`GenerarCompresiones`) restringido a ranuras libres reales en asignaturas automáticas.
- **Fase 2327 (T6):** Implementación de operadores C3/C4 (`GenerarReRanuracionEIntercambio`) y validación estricta V1–V6 (`Validar`) previniendo violaciones de docente, aula, carril R6 o reglas duras con rollback garantizado.
- **Fase 2328 (T7):** Integración del hook de reparación en `AgGeneradorHorarios.cs:33, 331-366` post-`RealinearOleadas` condicionado a `ReparacionR5Activa` y modo `AgModoReparacionR5`, con manejo y logging seguro de excepciones.
- **Fase 2329 (T8):** Exposición de telemetría en `AgResultadoCorrida.cs` (`ReparacionR5Intentos`, `ReparacionR5HardPre`, `ReparacionR5HardPost`, `ReparacionR5DesperdicioPre`, `ReparacionR5DesperdicioPost`) y serialización del bloque `"reparacionR5"` en `ParametrosJson` dentro de `AgFormGeneradorHorarios.cs`.
- **Fase 2330 (T9):** Cobertura de pruebas unitarias exhaustivas en `AgInventarioRanurasTests.cs` (3 pruebas) y `AgReparadorR5Tests.cs` (8 pruebas canónicas: fixture sintético, ausencia de ranuras, respeto a LEGADO/EsFijo, no creación de vacantes, rollback ante R6, minimización de desperdicio, determinismo e integración con evaluador).
- **Fase 2331 (T10):** Auditoría estática con `audit-tasks.ps1` (0 errores, 0 advertencias), validación de preflight con `preflight.ps1` (OK) y certificación de suite completa `verify.ps1 -Mode Full` (2010 superados, 0 fallos, 0 omitidos).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2322 | git (`reparacion-automatica-r5`) | Implementado |
| 2323 | `.agents/plans/REGISTRY.md`, `.agents/plans/2026-11-22-reparacion-automatica-r5-plan.md` | Implementado |
| 2324 | `AgInventarioRanuras.cs`, `ConversorPDF.XLL.csproj` | Implementado |
| 2325 | `AgReparadorR5.cs`, `AgSobrecargaR5.cs`, `AgResultadoReparacionR5.cs`, `AgValidadorReglas.cs` | Implementado |
| 2326 | `AgReparadorR5.cs` | Implementado |
| 2327 | `AgReparadorR5.cs` | Implementado |
| 2328 | `AgGeneradorHorarios.cs`, `AgModoReparacionR5.cs` | Implementado |
| 2329 | `AgResultadoCorrida.cs`, `AgFormGeneradorHorarios.cs` | Implementado |
| 2330 | `AgInventarioRanurasTests.cs`, `AgReparadorR5Tests.cs` | Implementado |
| 2331 | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md` | Implementado |























