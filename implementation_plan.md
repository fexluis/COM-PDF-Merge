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

---

## Ronda 526: Reparación R5 en producción y pruebas calibradas (Fases 2332–2342)

### Resumen de la Ronda
- **Fase 2332 (T1):** En `AgReparadorR5.cs:108-112, 150-154, 188-192`, reinicio explícito de `cand.Estado = AgEstadoIndividuo.NoEvaluado` y llamada forzada a `evaluadorAptitud.EvaluarAptitud(cand, contexto, forzar: true, incluirDetalleJson: true)` para evitar que la evaluación retorne tempranamente por estado clonado sin recalcular violaciones. En `IAgEvaluadorAptitud.cs:11`, exposición de la sobrecarga con parámetros `forzar` e `incluirDetalleJson`.
- **Fase 2333 (T2):** En `AgReparadorR5.cs:304-320, 535-555`, eliminación de la exclusión de asignaturas `LEGADO` en `GenerarDesplazamientos` (permitiendo forward shifting cuando son el candidato idóneo para aliviar sobrecarga) y ajuste en `Validar` para comprobar invariantes de patrón (días y bloque) sin bloquear el desplazamiento de fechas.
- **Fase 2334 (T3):** En `AgReparadorR5.cs`, acotación de la selección de genes a desplazar o re-ranurar filtrando estrictamente por `g.EsVacante == 0`, `g.IdAsignatura == candAsig.IdAsignatura` y concordancia de cohorte con `sob.Cohorte`, impidiendo mover vacantes o genes de otras cohortes/sedes.
- **Fase 2335 (T4):** En `AgReparadorR5.cs`, sustitución de búsquedas por `x.Id == g.Id` por indexación posicional exacta en el clon (`clon.Genes[idx]`), eliminando la pérdida de candidatos e inoperancia de `Validar` provocada por `Clonar()` asignando `Id = 0`.
- **Fase 2336 (T5):** En `AgReparadorR5.cs`, completitud de C3/C4 mediante cálculo de unidad de sesión con `AgUnidadesSemanales.HorasSesionAula`, asignación de `DiasPatron`, `IndiceDiaPatron = 0`, `PatronSemanal = string.Empty`, e implementación del intercambio C4 entre pares de asignaturas automáticas de la cohorte.
- **Fase 2337 (T6):** En `AgReparadorR5.cs`, aseguramiento de compresión C2 condicionada estrictamente a la existencia de huecos intermedios (`HuecosIntermedios > 0`) y ranuras libres reales en el carril sin colisionar con fechas ocupadas.
- **Fase 2338 (T7):** En `AgReparadorR5.cs:Validar`, robustecimiento de V5 con verificación de choques de aula (R2), sesiones conservadas, docente ocupado en calendario global (`TerceroOcupado`) y tope diario de cohorte (R4); en bucle principal, guarda post-`Sellar` descartando candidatos que incrementen vacantes.
- **Fase 2339 (T8):** En `AgReparadorR5.cs:Ejecutar`, contención total de excepciones en bloque `try / catch` con logging de advertencia y retorno seguro de estado pre-reparación sin abortar la optimización; adopción de iteración acumulativa (`actual = mejorClon ?? campeon`) con re-detección de sobrecargas tras cada aceptación.
- **Fase 2340 (T9):** En `AgReparadorR5.cs:CalcularDesperdicio`, eliminación del doble conteo de huecos intermedios calculando sobre `inv.LibresReales.Count` y balance por día; en `AgInventarioRanuras.cs`, consolidación de inventario multi-aula mediante agregación de aulas de la cohorte.
- **Fase 2341 (T10):** En `AgReparadorR5Tests.cs`, sustitución de mocks por instancia real de `AgEvaluadorAptitud(new AgValidadorReglas())`, calibración del fixture 132 con asignaturas 149 y 150 `LEGADO`, 152 `AUTOMATICO`, ventana MAR-VIE y festivos, certificando reducción HARD 1→0 y desplazamiento exacto a `2027-04-02`.
- **Fase 2342 (T11):** En `AgFormGeneradorHorarios.cs:3267`, persistencia de telemetría `dict["reparacionR5"]` en `corrida.ParametrosJson` antes del gate de factibilidad para registrar diagnóstico incluso en soluciones infactibles; actualización de `files_lock` en plan de ejecución y sincronización de baseline a 2010.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2332 | `AgReparadorR5.cs`, `IAgEvaluadorAptitud.cs` | Implementado |
| 2333 | `AgReparadorR5.cs` | Implementado |
| 2334 | `AgReparadorR5.cs` | Implementado |
| 2335 | `AgReparadorR5.cs` | Implementado |
| 2336 | `AgReparadorR5.cs` | Implementado |
| 2337 | `AgReparadorR5.cs` | Implementado |
| 2338 | `AgReparadorR5.cs` | Implementado |
| 2339 | `AgReparadorR5.cs` | Implementado |
| 2340 | `AgReparadorR5.cs`, `AgInventarioRanuras.cs` | Implementado |
| 2341 | `AgReparadorR5Tests.cs` | Implementado |
| 2342 | `AgFormGeneradorHorarios.cs`, `.agents/plans/2026-11-22-reparacion-automatica-r5-plan.md`, `tests/baseline/test_baseline.txt`, `.agents/plans/REGISTRY.md` | Implementado |

---

## Ronda 527: Corrección de tipos, forward shifting con ranuras libres y suite fiel 132 (Fases 2343–2352)

### Resumen de la Ronda
- **Fase 2343 (T1):** En `AgReparadorR5.cs:Validar`, reemplazo de `ctx.MaxHorasDiaCohorte` inexistente por la semántica canónica de R4 (`cf.Select(x => x.IdBloqueHorario).Distinct().Count() > 1`), impidiendo violaciones de franja horaria en la cohorte para un mismo día.
- **Fase 2344 (T2):** En `AgReparadorR5.cs:Validar`, adaptación de la llamada a `ctx.TerceroOcupado(g.IdTercero, f, new FranjaHoraria(bloque.TimeSpanInicio, bloque.TimeSpanFin))` obteniendo el bloque correspondiente desde `ctx.Bloques`.
- **Fase 2345 (T3):** En `AgReparadorR5.cs:Validar`, verificación de solapes frente a `ctx.SesionesConservadas` mediante comparación de intervalos temporales `FranjaHoraria.SeSolapaCon(new FranjaHoraria(sc.Inicio, sc.Fin))`, corrigiendo error de compilación por `sc.IdBloqueHorario`.
- **Fase 2346 (T4):** En `AgReparadorR5Tests.cs:70-74`, tipado estricto de `ctx.Festivos` como `List<(DateTime Inicio, DateTime Fin)>` corrigiendo error CS0029.
- **Fase 2347 (T5):** En `AgReparadorR5.cs:GenerarDesplazamientos`, filtrado de ranuras libres excluyendo las ocupadas por asignaturas ajenas (`ocupadasAjenas`), incorporando `inv.Liberables` y evaluando múltiples inicios disponibles en ranuras libres para cada día de bloque.
- **Fase 2348 (T6):** En `AgReparadorR5.cs:Ejecutar`, reordenamiento del ciclo de evaluación: `AgCoherenciaGrupo.Sellar(cand, contexto)` primero, guarda de vacantes, luego `Validar` y finalmente `EvaluarAptitud(cand, contexto, forzar: true, incluirDetalleJson: true)`.
- **Fase 2349 (T7):** En `AgReparadorR5.cs:GenerarCompresiones`, redefinición de compresión insertando sesiones en huecos libres intermedios reales (`LibresReales` entre `ini` y `fin`) garantizando la compactación efectiva del makespan (`Max() < fin`).
- **Fase 2350 (T8):** En `AgReparadorR5Tests.cs`, calibración de fixture 132 con asignaturas 149 y 150 `LEGADO`, 152 `AUTOMATICO`, ventana MAR-VIE y festivos, certificando desplazamiento a `2027-04-02` y HARD 1→0; aislamiento de test R6 asignando fechas disjuntas para evitar colisión R2 de aula; y verificación de determinismo mediante igualdad de secuencias de fechas.
- **Fase 2351 (T9):** En `AgReparadorR5.cs`, separación de contadores `intentosEvaluados` e `intentosAceptados`, captura previa de `hardAntes` para logging fidedigno y recomputación de desperdicio final sobre la cohorte inicial de diagnóstico.
- **Fase 2352 (T10):** En `AgReparadorR5.cs`, consumo activo de `inv.Liberables` en `GenerarDesplazamientos`; actualización de `files_lock` en plan de ejecución y actualización de contadores y fila R-199 en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2343 | `AgReparadorR5.cs` | Implementado |
| 2344 | `AgReparadorR5.cs` | Implementado |
| 2345 | `AgReparadorR5.cs` | Implementado |
| 2346 | `AgReparadorR5Tests.cs` | Implementado |
| 2347 | `AgReparadorR5.cs` | Implementado |
| 2348 | `AgReparadorR5.cs` | Implementado |
| 2349 | `AgReparadorR5.cs` | Implementado |
| 2350 | `AgReparadorR5Tests.cs` | Implementado |
| 2351 | `AgReparadorR5.cs` | Implementado |
| 2352 | `AgReparadorR5.cs`, `.agents/plans/2026-11-22-reparacion-automatica-r5-plan.md`, `.agents/plans/REGISTRY.md` | Implementado |

---

## Ronda 528: Reparación R5 con víctima 132 calibrada y cursor multi-gen (Fases 2353–2357)

### Resumen de la Ronda
- **Fase 2353 (T1):** Calibración temporal del fixture 132 en `AgReparadorR5Tests.cs:88-160` (Asig 149 miércoles arranca `2027-02-03`, Asig 152 martes arranca `2027-02-02`, y Asig 150 viernes LEGADO arranca `2027-03-19`), asegurando que la ordenación descendente por fecha de inicio seleccione fielmente a Asig 150 como la víctima que debe desplazarse a `2027-04-02`.
- **Fase 2354 (T2):** En `AgReparadorR5.cs:GenerarDesplazamientos`, priorización de `inv.LibresReales` sobre `Liberables` y ampliación de ventana de búsqueda a `Take(8)` inicios viables.
- **Fase 2355 (T3):** En `AgReparadorR5.cs:GenerarDesplazamientos`, manejo de asignaturas multi-gen mediante avance ordenado por `DiaSemana` y cursor temporal por sesión (`cursor = nuevasFechas.Min().AddDays(1)`), impidiendo que múltiples genes colapsen en fechas idénticas.
- **Fase 2356 (T4):** En `AgReparadorR5.cs:Ejecutar`, consistencia en el cálculo de desperdicio pre y post (`invPre` e `invPost`) evaluados sobre la misma cohorte reparada (`cohorteReparada`).
- **Fase 2357 (T5):** Sincronización de numeración global en `.agents/plans/REGISTRY.md` (fila `R-200`, contadores ronda 528 / siguiente 529, fase 2357 / siguiente 2358).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2353 | `AgReparadorR5Tests.cs` | Implementado |
| 2354 | `AgReparadorR5.cs` | Implementado |
| 2355 | `AgReparadorR5.cs` | Implementado |
| 2356 | `AgReparadorR5.cs` | Implementado |
| 2357 | `REGISTRY.md` | Implementado |

---

## Ronda 529: Fechas propias en opciones, cursor multi-gen sin solape y desperdicio por cohorte (Fases 2358–2361)

### Resumen de la Ronda
- **Fase 2358 (T1):** En `AgReparadorR5.cs:420-436`, construcción de `opciones` mediante `inv.LibresReales.Concat(fechasVictima)` filtrado por `inv.FechasValidas`, asegurando que el desplazamiento de la víctima considere sus propias fechas liberadas y permita reubicar la Asignatura 150 a `2027-04-02` resolviendo la sobrecarga R5 (HARD 1→0) en el fixture de la corrida 132.
- **Fase 2359 (T2):** En `AgReparadorR5.cs:449-478`, control estricto de solape multi-gen mediante el conjunto `usadas`, impidiendo que múltiples genes de la misma asignatura programados para el mismo día reciban fechas idénticas.
- **Fase 2360 (T3):** En `AgReparadorR5.cs:110-114, 156, 277-283`, cálculo de `invPre` y `despPreCohorte` local por sobrecarga (`sob.Cohorte`) en lugar de `primerCohorte`, garantizando consistencia en las comparaciones de desperdicio antes y después de la reparación.
- **Fase 2361 (T4):** Actualización de trazabilidad y numeración global en `.agents/plans/REGISTRY.md` (fila `R-201`, contadores ronda 529 / siguiente 530, fase 2361 / siguiente 2362).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2358 | `AgReparadorR5.cs`, `AgReparadorR5Tests.cs` | Implementado |
| 2359 | `AgReparadorR5.cs` | Implementado |
| 2360 | `AgReparadorR5.cs` | Implementado |
| 2361 | `REGISTRY.md` | Implementado |

---

## Ronda 530: Resolución analítica de corrida 132 y preparación de commit/gates (Fase 2362)

### Resumen de la Ronda
- **Fase 2362 (T1):** Confirmación analítica del cierre funcional del defecto en `IdCorrida=132` (la inclusión de fechas propias en `opciones` permite desplazar Asig 150 a `2027-04-02`, reduciendo HARD 1→0 en 149/150/152). Documentación de la ejecución de compuertas (`git commit`, `verify.ps1 -Mode Build/Fast/Full`) y actualización de numeración global en `.agents/plans/REGISTRY.md` (fila `R-202`, contadores ronda 530 / siguiente 531, fase 2362 / siguiente 2363).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2362 | `REGISTRY.md`, `git` | Implementado |

























---

## Ronda 531: Orden R6→R5, validación incremental y ripple por carril (Fases 2363–2367)

### Resumen de la Ronda
- **Fase 2363 (T1):** En `AgGeneradorHorarios.cs`, reparación de carriles multipasada (máx. 3) con aceptación por **peso de reglas** (`PesoDuras`, `MejoraReparacion`) y hash de estado; se acepta el intercambio R6(850)→R5(800) y se elimina el rollback por conteo.
- **Fase 2364 (T2):** En `AgReparadorR5.cs`, `ClavesDuras` (pares R2/R4/R6) y `Validar(..., preexistentes)` que rechaza sólo violaciones **nuevas**; guardas de vacantes contra `actual`.
- **Fase 2365 (T3):** En `AgDiagnosticoInfactibilidad.cs`, emisión de mensajes R5/R6 (cohorte, tramo, carril) en `Describir` y `ResumenCorto`, truncados.
- **Fase 2366 (T4):** Nuevo `AgInventarioCarril.cs` y `AgInventarioRanuras.Carriles` (fechas válidas/ocupadas/libres por `aula|patrón`), con `Iso` y `FechasDeGen`.
- **Fase 2367 (T5):** `GenerarDesplazamientosRipple` (cadena del carril), telemetría `RippleIntentos/RippleSinCola`, props en `AgResultadoCorrida`, persistencia JSON en ambos bloques y 6 tests nuevos (fixture #135, R6 preexistente, sin cola, carriles, pesos).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2363 | `AgGeneradorHorarios.cs` | Implementado |
| 2364 | `AgReparadorR5.cs` | Implementado |
| 2365 | `AgDiagnosticoInfactibilidad.cs` | Implementado |
| 2366 | `AgInventarioCarril.cs`, `AgInventarioRanuras.cs`, `ConversorPDF.XLL.csproj` | Implementado |
| 2367 | `AgReparadorR5.cs`, `AgResultadoReparacionR5.cs`, `AgResultadoCorrida.cs`, `AgFormGeneradorHorarios.cs`, tests | Implementado |

### Verificación
- `verify.ps1 -Mode Fast` (AgReparadorR5|AgInventarioRanuras|AgGeneradorHorarios): 21/21 PASS.
- `verify.ps1 -Mode Full`: **2016/2016, 0 fallos, 0 regresiones**, veredicto PASS.
- `audit-tasks.ps1 T1–T4`: AUDIT_PASS (0 hallazgos). `T5`: 25 bloqueantes heredados en `AgFormGeneradorHorarios.cs` (preexistentes, fuera del diff; 0 introducidos) → deuda para ronda futura.
- Baseline actualizado a 2016.


---

## Ronda 532: Reparación R6 dedicada y orden Sellar→plan (Fases 2368–2370)

### Resumen de la Ronda
- **Fase 2368 (T1):** En `AgGeneradorHorarios.cs` se retira el `AgCoherenciaGrupo.Sellar` posterior a `RealinearOleadas`; el sellado previo ya ocurrió y el `Sellar` post-plan borraba `FechasGen` (`AgCoherenciaGrupo.cs:171-176`), anulando el reordenamiento del carril.
- **Fase 2369 (T2):** Nuevo `AgReparadorR6.cs` + `AgResultadoReparacionR6.cs`: detecta pares R6 por cohorte/carril, elige la asignatura corta no fija y reproyecta su bloque tras el fin del solapado usando `AgInventarioRanuras.Carriles`; acepta sólo si baja el HARD; no sella.
- **Fase 2370 (T3):** Hook `AgReparadorR6` antes del hook R5, telemetría `ReparacionR6Intentos/HardPre/HardPost` en `AgResultadoCorrida` y tests `AgReparadorR6Tests` (carril largo + dos cortos; cortas fijas).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2368 | `AgGeneradorHorarios.cs` | Implementado |
| 2369 | `AgReparadorR6.cs`, `AgResultadoReparacionR6.cs`, `ConversorPDF.XLL.csproj` | Implementado |
| 2370 | `AgGeneradorHorarios.cs`, `AgResultadoCorrida.cs`, `AgReparadorR6Tests.cs` | Implementado |

### Verificación
- `verify.ps1 -Mode Fast` (R6+R5+Generador): 19/19 PASS.
- `verify.ps1 -Mode Full`: **2018/2018, 0 fallos, 0 regresiones**, veredicto PASS.
- `audit-tasks.ps1 T1–T3`: AUDIT_PASS (0 hallazgos).
- Baseline actualizado a 2018.

---

## Ronda 533: Planificador determinista de carril y re-ranuración R6 (Fases 2371–2374)

### Resumen de la Ronda
- **Fase 2371 (T1):** Creación de `AgPlanificadorCarril.cs` y `AgResultadoPlanCarril.cs`: mide capacidad real por día (`porDia`) y semanas disponibles de la ventana, mapea índices posicionales para mutar el clon sin depender de instancias por referencia, respeta asignaturas fijas como anclas inmutables y deduce colisiones externas en el aula/cohorte (`ocupadasFuera`).
- **Fase 2372 (T2):** En `AgReparadorR6.cs`, integración previa de `AgPlanificadorCarril.Planificar` por carril único; si no hay mejora, ejecuta `GenerarReRanuracion` explorando aulas alternativas de la misma sede con cupo libre en los días requeridos; fallback determinista a `GenerarCandidatos`. Control de flujo con `break` para re-evaluación limpia de pares tras cada aceptación.
- **Fase 2373 (T3):** Ampliación de `AgResultadoReparacionR6.cs` con `Motivo` (`SIN_CARRIL`, `SIN_FECHAS`, `CAPACIDAD`, `SIN_ORDEN`, `FIJO`), `SemanasRequeridas` y `SemanasDisponibles`. Log consolidado por carril en `AgReparadorR6.cs`.
- **Fase 2374 (T4):** En `AgReparadorR6Tests.cs`, batería de pruebas con caso #137 (cadena 152/153/151), detección de `SIN_FECHAS` en ventana ocupada, detección de `CAPACIDAD` en ventana insuficiente, anclaje de asignaturas fijas y re-ranuración hacia aula alternativa. Asignación R-205 y contadores actualizados en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2371 | `AgPlanificadorCarril.cs`, `AgResultadoPlanCarril.cs`, `ConversorPDF.XLL.csproj` | Implementado |
| 2372 | `AgReparadorR6.cs` | Implementado |
| 2373 | `AgResultadoReparacionR6.cs`, `AgReparadorR6.cs` | Implementado |
| 2374 | `AgReparadorR6Tests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 534: Planificador multi-gen, orden curricular y re-ranuración restringida (Fases 2375–2379)

### Resumen de la Ronda
- **Fase 2375 (T1):** En `AgPlanificadorCarril.cs`, iteración independiente por cada gen del sujeto (`OrderBy(clon.Genes[i].DiaSemana)`) extrayendo fechas exactas para su propio `DiaSemana` con avance secuencial del cursor temporal, resolviendo la asignación errónea de multi-gen en un solo día.
- **Fase 2376 (T2):** En `AgReparadorR6.cs`, filtro estricto de aulas alternativas por `IdPrograma` y sede compatible; validación obligatoria de bloque con `ctx.TryGetBloque`; proyección de `DiasPatron` deducido de los días reales; y descarte si el aula destino tiene vínculo `LEGADO` o `FIJO`.
- **Fase 2377 (T3):** En `AgPlanificadorCarril.cs`, respeto al `OrdenCurricular` de `AulaAsignatura` como primer criterio de ordenamiento y cómputo de `IniFijo`/`FinFijo` con fallback sobre `FechasGen` ante valores nulos.
- **Fase 2378 (T4):** En `AgReparadorR6.cs`, limpieza de `resultado.Motivo = string.Empty` al inicio de cada iteración, log con deltas reales de HARD y cómputo de `intentos++` por evaluación de planificador.
- **Fase 2379 (T5):** En `AgReparadorR6Tests.cs`, aserciones por invariante de no-solape temporal y HARD, verificación explícita de aula destino `IdAula == 13`, fixture de asignatura multi-día y formalización de plan swarm `.agents/plans/2026-11-23-r6-planificador-carril-plan.md` y fila `R-206` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2375 | `AgPlanificadorCarril.cs` | Implementado |
| 2376 | `AgReparadorR6.cs` | Implementado |
| 2377 | `AgPlanificadorCarril.cs` | Implementado |
| 2378 | `AgReparadorR6.cs` | Implementado |
| 2379 | `AgReparadorR6Tests.cs`, `REGISTRY.md`, `.agents/plans/2026-11-23-r6-planificador-carril-plan.md` | Implementado |

---

## Ronda 535: Validación de periodo, presupuesto O(1) y capacidad sobre libres (Fases 2380–2383)

### Resumen de la Ronda
- **Fase 2380 (T6):** En `AgReparadorR6.cs`, validación estricta de `IdPeriodo` e `IdPrograma` al buscar aulas candidatas de re-ranuración y verificación de existencia de `relDestino` válido en catálogo descartando `LEGADO` y `FIJO`.
- **Fase 2381 (T7):** En `AgReparadorR6.cs`, control estricto de presupuesto con `intentos >= MaxIntentos` antes de planificar y reseteo de `SemanasRequeridas` y `SemanasDisponibles`. En `AgPlanificadorCarril.cs`, consulta de `OrdenCurricular` en O(1) vía `ctx.AulaAsignaturaPorAulaAsig` y supresión de permutaciones cuando existe orden curricular explícito.
- **Fase 2382 (T8):** En `AgReparadorR6Tests.cs`, aserción exacta `SIN_FECHAS`, verificación de compuertas `verify Full` (2024 tests superados, 0 errores), actualización de baseline (`tests/baseline/test_baseline.txt`) a 2024 y sincronización de registro.
- **Fase 2383 (T9):** En `AgPlanificadorCarril.cs`, deducción previa de colisiones externas en aula/cohorte antes de evaluar capacidad, calculando `semanasDisp` directamente sobre fechas libres efectivas por día.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2380 | `AgReparadorR6.cs` | Implementado |
| 2381 | `AgPlanificadorCarril.cs`, `AgReparadorR6.cs` | Implementado |
| 2382 | `AgReparadorR6Tests.cs`, `REGISTRY.md`, `tests/baseline/test_baseline.txt` | Implementado |
| 2383 | `AgPlanificadorCarril.cs` | Implementado |

### Verificación
- `verify.ps1 -Mode Fast` (R6+R5): 19/19 PASS.
- `verify.ps1 -Mode Full`: **2024/2024, 0 fallos, 0 regresiones**, veredicto PASS.
- `audit-tasks.ps1 T1–T5`: AUDIT_PASS (0 hallazgos).
- Baseline actualizado a 2024.

---

## Ronda 536: Cierre de re-ranuración y robustez de asignación (Fases 2384–2386)

### Resumen de la Ronda
- **Fase 2384 (T10):** En `AgReparadorR6.cs`, filtro riguroso de aulas alternativas para re-ranuración garantizando que `a.Estado == AgEstadoAula.Disponible || a.Estado == AgEstadoAula.Aprobado` y consistencia de `Subunidad` (coincidente con la del aula víctima o vacía).
- **Fase 2385 (T11):** En `AgReparadorR6.cs`, actualización integral del gen clonado tras re-ranuración: recálculo de `ClaveGrupo = $"{gClon.IdAsignatura}|{aulaAlt.Id}|{gClon.Grupo}"`, reseteo de `PatronSemanal = string.Empty` e `IndiceDiaPatron = 0`.
- **Fase 2386 (T12):** En `AgResultadoReparacionR6.cs`, propiedad `MotivoReranuracion`. En `AgReparadorR6.cs`, emisión diagnóstica `MotivoReranuracion = "SIN_VINCULO"` cuando ninguna aula candidata pase los filtros. En `AgReparadorR6Tests.cs`, 3 pruebas unitarias dedicadas (`Reparar_ReRanuracion_NoUsaAulaInactiva`, `Reparar_ReRanuracion_RechazaVinculoLegado`, `Reparar_ReRanuracion_ActualizaClaveGrupo`). Formalización de Wave 536 y fila `R-207` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2384 | `AgReparadorR6.cs` | Implementado |
| 2385 | `AgReparadorR6.cs` | Implementado |
| 2386 | `AgResultadoReparacionR6.cs`, `AgReparadorR6.cs`, `AgReparadorR6Tests.cs`, `REGISTRY.md`, `.agents/plans/2026-11-23-r6-planificador-carril-plan.md` | Implementado |

### Verificación
- `verify.ps1 -Mode Fast` (R6+R5): 22/22 PASS.
- `verify.ps1 -Mode Full`: **2027/2027, 0 fallos, 0 regresiones**, veredicto PASS.
- `audit-tasks.ps1 T10–T12`: AUDIT_PASS (0 hallazgos).
- Baseline actualizado a 2027.

---

## Ronda 537: Cierre real de re-ranuración y diagnóstico de carril (Fases 2387–2389)

### Resumen de la Ronda
- **Fase 2387 (T13):** En `AgReparadorR6.cs`, derivación del sufijo de grupo efectivo en `ClaveGrupo` inspeccionando las partes del identificador previo (`partes.Length >= 3 ? partes[2] : (Grupo ?? "A")`), garantizando que genes con `Grupo` vacío conserven su sufijo asignado original.
- **Fase 2388 (T14):** En `AgReparadorR6.cs`, validación de coincidencia de `DiasPatron` y `Subunidad` exigidos por la relación catálogo `relDestino` en el aula destino antes de autorizar la re-ranuración.
- **Fase 2389 (T15):** En `AgPlanificadorCarril.cs`, propagación de `ultimoMotivo` capturado en `AsignarCadena` (devolviendo `SIN_FECHAS` ante agotamiento de días por asignatura en vez de enmascararlo con `SIN_ORDEN`). Pruebas unitarias de grupo y patrón en `AgReparadorR6Tests.cs`. Sincronización de Wave 537 y fila `R-208` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2387 | `AgReparadorR6.cs`, `AgReparadorR6Tests.cs` | Implementado |
| 2388 | `AgReparadorR6.cs`, `AgReparadorR6Tests.cs` | Implementado |
| 2389 | `AgPlanificadorCarril.cs`, `REGISTRY.md`, `.agents/plans/2026-11-23-r6-planificador-carril-plan.md` | Implementado |

---

## Ronda 538: Validación de subunidad contra aula destino (Fase 2390)

### Resumen de la Ronda
- **Fase 2390 (T16):** En `AgReparadorR6.cs:344-345`, comparación de `relDestino.Subunidad` contra `aulaAlt.Subunidad` (aula destino candidata) en lugar de `aulaV.Subunidad` (aula origen víctima), evitando rechazos espurios cuando el aula de origen tiene subunidad no especificada o vacía. Prueba unitaria en `AgReparadorR6Tests.cs`. Sincronización de Wave 538 y fila `R-209` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2390 | `AgReparadorR6.cs`, `AgReparadorR6Tests.cs`, `REGISTRY.md`, `.agents/plans/2026-11-23-r6-planificador-carril-plan.md` | Implementado |

---

## Ronda 539: Día residual 5/1, catálogo de 5 días y fidelidad de partición (Fases 2391–2396)

### Resumen de la Ronda
- **Fase 2391 (T1):** Creación de `AgOpcionesParticion.ResolverFallbackSinSingleton` para ventanas de 5 días (1..5) con patrones mixtos o de 1 día que producen `Origen.Ninguno`, generando opciones binarias de respaldo `{2,3|1,4,5}` y `{1,2,5|3,4}` sin singletons. Integración de fallback en `AgPlanificadorOleadas.cs:355-365`.
- **Fase 2392 (T2):** Implementación de `AgCarriles.ParticionSinSingleton` acotada con guarda estricta a ventanas de exactamente 5 días (1..5), evitando alterar ventanas de 3 o 7 días. Sustitución de canónicas con singleton en `AgPlanificadorOleadas.cs:352-358`.
- **Fase 2393 (T3):** En `AgPlanificadorOleadas.cs:545-620`, limitación de `ranurasEfectivas` y construcción del patrón al subconjunto de ranuras efectivas de la banda en lugar de la banda completa, permitiendo que `EsParticionFiel` no descarte opciones de catálogo válidas.
- **Fase 2394 (T4):** Implementación de `ForzarPatronResidual` en `AgPlanificadorOleadas.cs`, normalizando patrones de asignaturas que tocan días extremos ({5} en banda {1,2,5} pasa a {1,2,5}; {1} en banda {1,4,5} pasa a {1,4,5}), propagando `DiasPatron` y `RanurasPatron` y emitiendo traza diagnóstica `[AG-RESIDUAL]`.
- **Fase 2395 (T5):** Propiedad `BandasSingleton` en `AgEvaluacionParticion`. En `AgParticionCarriles.cs:129-144, 246-274`, compuerta de selección y desempate que prioriza minimizar `BandasSingleton`, invierte el desempate a `.ThenBy(ev => ev.Bandas)` para favorecer menos bandas (2 sobre 3) y descarta particiones con singletons cuando existan alternativas sin singleton.
- **Fase 2396 (T6):** Batería de pruebas de regresión y cobertura para 5 días y días residuales en `AgOpcionesParticionTests.cs`, `AgCarrilesTests.cs`, `AgParticionCarrilesTests.cs`, `AgParticionCarrilesIntegracionTests.cs` (adaptación de `Caso76` a 2 bandas) y `AgPlanificadorOleadasTests.cs`. Sincronización de Wave 539 y fila `R-210` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2391 | `AgOpcionesParticion.cs`, `AgPlanificadorOleadas.cs` | Implementado |
| 2392 | `AgCarriles.cs`, `AgPlanificadorOleadas.cs` | Implementado |
| 2393 | `AgPlanificadorOleadas.cs` | Implementado |
| 2394 | `AgPlanificadorOleadas.cs` | Implementado |
| 2395 | `AgParticionCarriles.cs` | Implementado |
| 2396 | `AgParticionCarrilesTests.cs`, `AgParticionCarrilesIntegracionTests.cs`, `AgCarrilesTests.cs`, `AgOpcionesParticionTests.cs`, `AgPlanificadorOleadasTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 540: Regresiones en Fases 2391–2396 (día residual, catálogo y tests) (Fases 2397–2402)

### Resumen de la Ronda
- **Fase 2397 (T1):** Corrección sintáctica en `AgParticionCarrilesTests.cs:319`, eliminando la llave de cierre redundante en el método `Seleccionar_CincoDiasMarVieConSingleton_ShouldElegirDosTres_UnoCuatroCinco` que expulsaba el test subsecuente fuera de la clase.
- **Fase 2398 (T2):** En `AgPlanificadorOleadas.cs:357-366`, sustitución condicional de la partición canónica sin singleton únicamente cuando `origen == Origen.MarVie` o `origen == Origen.LunJue`, preservando neutralidad ante `Origen.Ninguno`.
- **Fase 2399 (T3):** En `AgPlanificadorOleadas.cs:710-725`, eliminación de la comprobación restrictiva `Contains(5)` / `Contains(1)` en `ForzarPatronResidual`, asegurando que toda asignatura asignada a una banda residual `{1,2,5}` o `{1,4,5}` adopte el patrón ampliado correspondiente (incorporando prueba unitaria).
- **Fase 2400 (T4):** En `AgParticionCarrilesIntegracionTests.cs:181-188`, actualización de la aserción en `Caso76_Ausencia` para esperar 2 bandas (`{4,5}|{6,7}`) en conformidad con la partición canónica y desempate sin singletons.
- **Fase 2401 (T5):** En `AgParticionCarriles.cs:200-218`, supresión del descarte anticipado y duro `haySinSingleton`, permitiendo que particiones con singleton sean evaluadas y únicamente penalizadas en el orden de desempate (`ThenBy(ev => ev.BandasSingleton)`).
- **Fase 2402 (T6):** En `AgPlanificadorCarril.cs:138-147`, cálculo verídico de la capacidad del carril comparando sesiones requeridas totales contra `lane.FechasValidas.Count` en lugar del número de semanas. Sincronización de Wave 540 y fila `R-211` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2397 | `AgParticionCarrilesTests.cs` | Implementado |
| 2398 | `AgPlanificadorOleadas.cs` | Implementado |
| 2399 | `AgPlanificadorOleadas.cs`, `AgPlanificadorOleadasTests.cs` | Implementado |
| 2400 | `AgParticionCarrilesIntegracionTests.cs` | Implementado |
| 2401 | `AgParticionCarriles.cs` | Implementado |
| 2402 | `AgPlanificadorCarril.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 541: Siembra de patrones aún crea carriles de 1 día; R6 pierde índice de patrón (Fases 2403–2406)

### Resumen de la Ronda
- **Fase 2403 (T1):** Implementación de `AgCarriles.BandaSinSingletonParaDia` para mapear cualquier día suelto en ventanas de 5 días a su banda sin singleton respectiva (`{5} → {1,2,5}`, `{1} → {1,4,5}`, `{2,3} → {2,3}`, `{4} → {1,4,5}`). Integrado en `AgPlanificadorPatronSemanal.DiasKDe` y `RanurasDe`, y blindaje en `AgFabricaIndividuos.cs:434-442` para evitar siembra automática de patrones `{5}` o `{1}`.
- **Fase 2404 (T2):** En `AgParticionCarriles.cs:288-305`, formalización de la política de motor `PermitirSingleton = false`, filtrando estrictamente candidatas a `BandasSingleton == 0` cuando exista al menos una alternativa sin singletons, y emitiendo advertencia `SINGLETON_NECESARIO` con fallback seguro sólo si ninguna partición elimina singletons.
- **Fase 2405 (T3):** En `AgReparadorR6.cs:415-427`, preservación determinista de `gClon.IndiceDiaPatron` al mover asignaturas a aulas alternativas calculando `Math.Max(0, diasPatron.IndexOf(gClon.DiaSemana))` en lugar de forzarlo a 0. Aserciones en `AgReparadorR6Tests.cs`.
- **Fase 2406 (T4):** En `AgResultadoPlanCarril.cs` y `AgPlanificadorCarril.cs:138-147`, separación de la semántica de telemetría introduciendo `SesionesDisponibles` (fechas totales del carril) y derivando `SemanasDisponibles = ceil(SesionesDisponibles / 2.0)`. Sincronización de Wave 541 y fila `R-212` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2403 | `AgCarriles.cs`, `AgPlanificadorPatronSemanal.cs`, `AgFabricaIndividuos.cs`, `AgCarrilesTests.cs`, `AgFabricaIndividuosTests.cs` | Implementado |
| 2404 | `AgParticionCarriles.cs` | Implementado |
| 2405 | `AgReparadorR6.cs`, `AgReparadorR6Tests.cs` | Implementado |
| 2406 | `AgResultadoPlanCarril.cs`, `AgPlanificadorCarril.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 542: Mapeo singleton de 5 días no es neutral al origen; DiasKDe y RanurasDe divergen (Fases 2407–2409)

### Resumen de la Ronda
- **Fase 2407 (T1):** En `AgCarriles.cs:75-98`, parametrización con `bool esMarVie` en `AgCarriles.BandaSinSingletonParaDia` y validación estricta de ventana canónica `{1,2,3,4,5}`. Si `esMarVie` es true, mapea 1, 4, 5 a `{1,4,5}` y 2, 3 a `{2,3}`. Si es false (LunJue), mapea 1, 2, 5 a `{1,2,5}` y 3, 4 a `{3,4}`. Si la ventana difiere de `{1,2,3,4,5}` o el día no está en 1..5, devuelve `null`. Pruebas exhaustivas en `AgCarrilesTests.cs`.
- **Fase 2408 (T2):** En `AgPlanificadorPatronSemanal.cs:54-95, 120-155`, detección de origen en `DetectarEsMarVie(AgAulaAsignatura)` y unificación del cálculo de ranuras en `RanurasDe` apoyándose directamente en `DiasKDe`, asegurando estricta paridad entre `DiasPatron` y `RanurasPatron`.
- **Fase 2409 (T3):** En `AgResultadoPlanCarril.cs` y `AgPlanificadorCarril.cs:128-220`, desacople formal de fechas vs semanas: `SesionesDisponibles` refleja las fechas físicas útiles del carril, mientras `SemanasDisponibles` pasa a ser una propiedad calculada sólo lectura `(int)Math.Ceiling(SesionesDisponibles / 2.0)`. Sincronización de Wave 542 y fila `R-213` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2407 | `AgCarriles.cs`, `AgCarrilesTests.cs` | Implementado |
| 2408 | `AgPlanificadorPatronSemanal.cs` | Implementado |
| 2409 | `AgResultadoPlanCarril.cs`, `AgPlanificadorCarril.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 543: Origen residual ambiguo, SemanasDisponibles duplica sesiones y singleton escapa por catálogo (Fases 2410–2413)

### Resumen de la Ronda
- **Fase 2410 (T1):** En `AgPlanificadorPatronSemanal.cs:164-190`, `DetectarEsMarVie` retorna `bool?` distinguiendo orígenes explícitos (`MAR-VIE`, `LUN-JUE`), detección de patrones de múltiples días y heurística para patrones de 1 día (5 → false/LunJue, 1/4 → true/MarVie, 2/3 → false/LunJue; otros → null sin inventar origen emitiendo `ORIGEN_RESIDUAL_DESCONOCIDO`). En `AgCarriles.cs:97-102`, la sobrecarga sin origen adopta la heurística `dia != 5`.
- **Fase 2411 (T2):** En `AgResultadoPlanCarril.cs:11-23`, eliminación del setter confuso en `SemanasDisponibles` (que multiplicaba por 2). Se introduce `SesionesDisponibles` (entero mutable) y `SemanasEquivalentes` (sólo lectura, calculada como `ceil(SesionesDisponibles / 2.0)`), manteniendo `SemanasDisponibles => SemanasEquivalentes` por compatibilidad. En `AgPlanificadorCarril.cs:128-218`, todas las asignaciones usan `SesionesDisponibles = ...`.
- **Fase 2412 (T3):** En `AgParticionCarriles.cs:135-155`, aplicación estricta de la política `PermitirSingleton = false` en la evaluación de opciones de catálogo (`evaluadasCat.Where(x => x.Ev.BandasSingleton == 0)`), impidiendo que una opción de catálogo con singleton sea aceptada cuando se prohíben singletons.
- **Fase 2413 (T4):** En `AgFabricaIndividuosTests.cs:660-705` y `AgCarrilesTests.cs:184-188`, cobertura de patrones de 1 día `{1}` y `{5}` con orígenes `MAR-VIE` y `LUN-JUE` mapeando a `{1,4,5}` y `{1,2,5}`, y verificación de la heurística residual default. Sincronización de Wave 543 y fila `R-214` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2410 | `AgPlanificadorPatronSemanal.cs`, `AgCarriles.cs` | Implementado |
| 2411 | `AgResultadoPlanCarril.cs`, `AgPlanificadorCarril.cs` | Implementado |
| 2412 | `AgParticionCarriles.cs` | Implementado |
| 2413 | `AgFabricaIndividuosTests.cs`, `AgCarrilesTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 544: La siembra automática sin AgAulaAsignatura sigue produciendo {5} (Fases 2414–2416)

### Resumen de la Ronda
- **Fase 2414 (T1):** En `AgPlanificadorPatronSemanal.cs:170-196`, cuando `aa == null` o `aa.DiasPatron` está vacío (ruta automática), `DetectarEsMarVie` retorna deterministamente `false` (LunJue), permitiendo que la siembra automática en ventanas de 5 días resuelva el tercer carril residual mapeándolo a `{1,2,5}` sin retornar singletons `{5}`.
- **Fase 2415 (T2):** En `AgPlanificadorPatronSemanal.cs:186-195` y `AgCarriles.cs:97-108`, delimitación de heurística: únicamente los residuos 5 (`false`/LunJue) y 1 (`true`/MarVie) infieren origen; los días 2, 3 y 4 sin `OrigenPatron` devuelven `null` emitiendo `ORIGEN_DIA_NO_RESIDUAL_DESCONOCIDO` y preservando su asignación canónica sin inventar bandas.
- **Fase 2416 (T3):** En `AgFabricaIndividuosTests.cs:662-730` y `AgCarrilesTests.cs:184-190`, ampliación de cobertura verificando que la siembra automática sin `aa` produzca genes con al menos 2 días en `RanurasPatron`, y que patrones `{2}`, `{3}` y `{4}` sin origen mantengan sus días sin forzar bandas residuales. Sincronización de Wave 544 y fila `R-215` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2414 | `AgPlanificadorPatronSemanal.cs`, `AgFabricaIndividuosTests.cs` | Implementado |
| 2415 | `AgPlanificadorPatronSemanal.cs`, `AgCarriles.cs` | Implementado |
| 2416 | `AgFabricaIndividuosTests.cs`, `AgCarrilesTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 545: Patrones de 3 días se normalizan a 2; 2/3/4 sin origen siguen en singleton (Fases 2417–2418)

### Resumen de la Ronda
- **Fase 2417 (T1):** En `AgPlanificadorPatronSemanal.cs:40-85`, soporte de patrones de 3 días `{1,2,5}` y `{1,4,5}` en ventanas de 5 días exceptuándolos de la normalización por sobrecupo de días (`AgCarriles.DiasPorCarril`) y aceptándolos como carriles residuales válidos. Asimismo, eliminación del singleton para días 2, 3 y 4 sin origen declarado aplicando el fallback determinista `esMarVieOrigen ?? true` en `BandaSinSingletonParaDia` (`2/3 → {2,3}`, `4 → {1,4,5}`). En `AgCarriles.cs:97-105`, la sobrecarga sin origen adopta este mapeo no singleton.
- **Fase 2418 (T2):** En `AgFabricaIndividuosTests.cs:696-745` y `AgCarrilesTests.cs:184-190`, actualización de expectativas en `Fabrica_PatronUnDiaNoResidualSinOrigen_UsaDosTresDias` (`2` y `3` a `"2,3"`, `4` a `"1,4,5"` con `RanurasPatron >= 2`), y prueba de preservación de patrones residuales de 3 días en `DiasKDe_PatronResidualTresDias_SeConserva`. Sincronización de Wave 545 y fila `R-216` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2417 | `AgPlanificadorPatronSemanal.cs`, `AgCarriles.cs` | Implementado |
| 2418 | `AgFabricaIndividuosTests.cs`, `AgCarrilesTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 546: Test de fábrica cuenta genes incorrectamente; semanas de carril siguen asumiendo 2 días (Fases 2419–2420)

### Resumen de la Ronda
- **Fase 2419 (T1):** En `AgFabricaIndividuosTests.cs:670-696`, corrección de la aserción de conteo de genes en `Fabrica_AutomaticaSinAa_NoUsaRanuraSingleton`: un `AgGen` por ranura genera 7 genes para 3 asignaturas (2 en carril 0, 2 en carril 1 y 3 en carril 2 `{1,2,5}`), verificando `Distinct(IdAsignatura).Count() == 3` y `ind.Genes.Count == 7`.
- **Fase 2420 (T2):** En `AgResultadoPlanCarril.cs:11-30` y `AgPlanificadorCarril.cs:120-220`, desacople de la división fija entre 2.0 introduciendo `SesionesPorSemana` (calculado dinámicamente según los días distintos del carril) tanto en `SemanasEquivalentes` como en `SemanasRequeridasEquivalentes`, permitiendo que carriles de 3 días (`{1,2,5}` / `{1,4,5}`) calculen semanas reales fieles. Suite en `AgReparadorR6Tests.cs`. Sincronización de Wave 546 y fila `R-217` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2419 | `AgFabricaIndividuosTests.cs` | Implementado |
| 2420 | `AgResultadoPlanCarril.cs`, `AgPlanificadorCarril.cs`, `AgReparadorR6Tests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 547: Separación de sesiones y semanas requeridas; trazabilidad de origen inferido (Fases 2421–2422)

### Resumen de la Ronda
- **Fase 2421 (T1):** En `AgResultadoPlanCarril.cs:11-30` y `AgPlanificadorCarril.cs:128-229`, separación semántica de sesiones requeridas brutas (`SesionesRequeridas`) frente a semanas requeridas (`SemanasRequeridas => ceil(SesionesRequeridas / SesionesPorSemana)`), evitando ambigüedades dimensionales en logs y métricas de carril.
- **Fase 2422 (T2):** En `AgPlanificadorPatronSemanal.cs:58-67`, registro diagnóstica `ORIGEN_DIA_INFERIDO` cuando un patrón de 1 día no residual (2, 3 o 4) adopta el fallback determinista MarVie por ausencia de `OrigenPatron`. Sincronización de Wave 547 y fila `R-218` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2421 | `AgResultadoPlanCarril.cs`, `AgPlanificadorCarril.cs`, `AgReparadorR6Tests.cs` | Implementado |
| 2422 | `AgPlanificadorPatronSemanal.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 548: Eliminación de setter ambiguo en SemanasRequeridas (Fase 2423)

### Resumen de la Ronda
- **Fase 2423 (T1):** En `AgResultadoPlanCarril.cs:14-23`, supresión definitiva del setter en `SemanasRequeridas`. La propiedad queda exclusivamente como un getter calculado proporcional a `SesionesPorSemana`, eliminando el riesgo de que asignaciones de sesiones brutas generen inconsistencias dimensionales en los consumidores. Sincronización de Wave 548 y fila `R-219` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2423 | `AgResultadoPlanCarril.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 549: Registro diagnóstico de inferencia de origen para días no residuales (Fase 2424)

### Resumen de la Ronda
- **Fase 2424 (T1):** En `AgPlanificadorPatronSemanal.cs:58-75`, estandarización de la clave diagnóstica `ORIGEN_DIA_NO_RESIDUAL` al aplicar la heurística fallback de MarVie para patrones declarados de 1 día no residual (`{2}`, `{3}`, `{4}`) que carecen de `OrigenPatron`. Sincronización de Wave 549 y fila `R-220` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2424 | `AgPlanificadorPatronSemanal.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 550: Integración de catálogo y canónicas sin singleton en AgPlanificadorOleadas (Fases 2425–2427)

### Resumen de la Ronda
- **Fase 2425 (T1):** En `AgPlanificadorOleadas.cs:352-378`, integración efectiva del catálogo determinista y partición sin singleton en la ruta real de `ConstruirReferencia`: detección de origen (`MarVie`/`LunJue`) para inicializar `canonicas` mediante `AgCarriles.ParticionSinSingleton(ranuras, esMarVie)` y resolución de opciones aplicables (`AgParticionCarriles.EsOpcionAplicable`) con respaldo directo (`AgOpcionesParticion.Resolver`) o fallback determinista sin singleton (`ResolverFallbackSinSingleton`), erradicando la degradación a enumeración con bandas singleton en ventanas de 5 días.
- **Fase 2426 (T2):** En `AgPlanificadorOleadasTests.cs:985-1045`, cobertura unitaria de la ruta real para ventanas de 5 días (`LUN-VIE` produciendo 2 bandas sin singleton `{2,3|1,4,5}` o `{1,2,5|3,4}`) y ventanas de 4 días (`4,5,6,7` aplicando catálogo `5,6M|4,6T,7`).
- **Fase 2427 (T3):** Formalización de verificación e integración y sincronización de numeración en `REGISTRY.md` (Wave 550, Fases 2425–2427, fila `R-221`).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2425 | `AgPlanificadorOleadas.cs` | Implementado |
| 2426 | `AgPlanificadorOleadasTests.cs` | Implementado |
| 2427 | `REGISTRY.md` | Implementado |

---

## Ronda 551: Validación de aplicabilidad en opción directa y aserciones de catálogo real (Fases 2428–2429)

### Resumen de la Ronda
- **Fase 2428 (T1):** En `AgPlanificadorOleadas.cs:374-395`, verificación explícita de aplicabilidad (`AgParticionCarriles.EsOpcionAplicable(ranuras, opcionDirecta)`) sobre la opción devuelta por `AgOpcionesParticion.Resolver` antes de asignarla a `opciones`. Si la opción directa no conserva exactamente las ranuras del contexto, no bloquea el fallback e invoca `AgOpcionesParticion.ResolverFallbackSinSingleton` emitiendo registro diagnóstico estructurado (`catalogoFallback=true`).
- **Fase 2429 (T2):** En `AgPlanificadorOleadasTests.cs:985-1080`, aserciones exhaustivas de ruta real sobre la partición: verificación de días exactos en `4,5,6,7` (`{5,6}` y `{4,6,7}` correspondientes a `5,6M|4,6T,7`), pertenencia canónica a particiones válidas sin singleton en 5 días (`2.3|1.4.5` o `1.2.5|3.4`) y activación efectiva del fallback ante opciones provistas inválidas o incompatibles. Sincronización de Wave 551 y fila `R-222` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2428 | `AgPlanificadorOleadas.cs` | Implementado |
| 2429 | `AgPlanificadorOleadasTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 552: Desempate de catálogo admitiendo bandas equivalentes sin empeoramiento (Fase 2430)

### Resumen de la Ronda
- **Fase 2430 (T1):** En `AgParticionCarriles.cs:144-157`, ajuste en la condición de aceptación del catálogo: sustitución de la restricción estricta de menor número de bandas (`x.Ev.Bandas < referencia.Bandas`) por menor o igual (`x.Ev.Bandas <= referencia.Bandas`), permitiendo que opciones de catálogo diseñadas para patrones específicos (como `5,6M|4,6T,7` en ventanas `4,5,6,7` / `JUE-DOM`) sustituyan la canónica cuando igualan en número de bandas (2 vs 2) y no empeoran infactibles ni makespan. Sincronización de Wave 552 y fila `R-223` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2430 | `AgParticionCarriles.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 553: Comparación global catálogo vs enumeración y desempate por catálogo (Fase 2431)

### Resumen de la Ronda
- **Fase 2431 (T1):** En `AgParticionCarriles.cs:120-170, 286-370`, supresión del cortocircuito de retorno inmediato tras evaluar catálogo; ejecución completa de la enumeración contigua y resolución por comparación simétrica (`Comparar(mejorCatalogo, mejorEnumeracion)`). Si la enumeración es estrictamente superior, gana la enumeración bajo la etiqueta `catalogo_fallback`; si el catálogo es estrictamente superior o existe empate exacto (`<= 0`), gana el catálogo bajo la etiqueta `catalogo`. Se añade batería de pruebas en `AgParticionCarrilesOpcionesTests.cs` cubriendo empates y mejoras contiguas. Sincronización de Wave 553 y fila `R-224` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2431 | `AgParticionCarriles.cs`, `AgParticionCarrilesOpcionesTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 554: Cobertura de empate, métrica conjunta de candidatas y catalogo_fallback (Fase 2432)

### Resumen de la Ronda
- **Fase 2432 (T1):** En `AgParticionCarriles.cs:338-356`, unificación de `CandidatasEvaluadas` en `AgSeleccionParticion` para reflejar la suma acumulada de evaluaciones de ambas rutas (`evaluadasOpciones + evaluadas`) en cualquier escenario de selección. En `AgParticionCarrilesOpcionesTests.cs:570-715`, incorporación de pruebas unitarias (`Seleccionar_EmpateConMismoSingleton_MantieneCatalogo` y `Seleccionar_EnumeracionMejorPorSingleton_GanaEnumeracion`) que garantizan la preservación del catálogo en empates exactos, la elección de la enumeración cuando elimina singletons residuales (`catalogo_fallback`), y la fidelidad de la métrica conjunta de candidatas evaluadas. Sincronización de Wave 554 y fila `R-225` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2432 | `AgParticionCarriles.cs`, `AgParticionCarrilesOpcionesTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 555: Política catálogo-only, cobertura sin enumeración, escalado obligatorio y fijación de patrón de 3 días (Fases 2433–2437)

### Resumen de la Ronda
- **Fase 2433 (T1):** En `AgParticionCarriles.cs:97-185`, introducción de `internal const bool PermitirEnumeracion = false` y parámetro `permitirEnumeracion = PermitirEnumeracion` en `Seleccionar`. Cuando es `false`, se omite la llamada a `Particiones()` y la evaluación de contiguas; si no hay opción aceptable de catálogo, retorna `Origen = "sin_catalogo"` y `Elegida = referencia`. En `AgPlanificadorOleadas.cs:450-465`, captura de `sin_catalogo` retornando `Result.Failure("I5_SIN_CATALOGO_AMPLIAR: no hay catálogo para la ventana actual.")`. Erradicación de `origen=enumeracion` y `catalogo_fallback`.
- **Fase 2434 (T2):** En `AgOpcionesParticion.cs:35-155`, cobertura determinista de catálogo para ventanas `4,5,6,7` (`JUE-DOM` retornando `5,6M|4,6T,7`) y resolución de fallbacks mixtos en 5 días (`2,3|1,4,5` y `1,2,5|3,4`) ponderados por conteo de señales de patrones `{1}`, `{2,3}`, `{4,5}` (MarVie) vs `{1,2}`, `{3,4}`, `{5}` (LunJue). `AgCarriles.ParticionSinSingleton` preservada exclusivamente para ventanas de 5 días.
- **Fase 2435 (T3):** En `AgPlanificadorOleadas.cs:91-140`, orquestación de ampliación obligatoria de ranuras ante fallo `I5_SIN_CATALOGO_AMPLIAR` (Opción 2A): bucle de reintento sobre `AgEscaleraVentanas.Siguiente(ventanaActual)` y reconstrucción de ranuras mediante `AgRanuras.PorVentana` hasta `MaxPasos`, evitando el fallback silencioso a enumeración.
- **Fase 2436 (T4):** En `AgPlanificadorPatronSemanal.cs:122-130`, `AgPlanificadorOleadas.cs:645-655` y `AgFabricaIndividuos.cs:424-440`, eliminación de singletons de 1 día en ventanas de 5 días: asignaturas en banda `{1,2,5}` o `{1,4,5}` fijan patrón completo de 3 días en `DiasPatron` y ranuras efectivas (Opción 1A). `AgCoherenciaGrupo.cs` preserva las horas y fechas sin revertir los días del patrón.
- **Fase 2437 (T5):** En `AgParticionCarrilesOpcionesTests.cs` y `AgPlanificadorOleadasTests.cs`, batería de pruebas unitarias validando política catálogo-only (`sin_catalogo` ante ausencia de catálogo, selección limpia sin enumeración), resolución exacta de `5,6M|4,6T,7` en `JUE-DOM`, escalado automático de `VIE-DOM` a `JUE-DOM` y fijación de 3 días en bandas residuales de 5 días. Sincronización de Wave 555 y fila `R-226` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2433 | `AgParticionCarriles.cs`, `AgPlanificadorOleadas.cs` | Implementado |
| 2434 | `AgOpcionesParticion.cs`, `AgCarriles.cs` | Implementado |
| 2435 | `AgPlanificadorOleadas.cs` | Implementado |
| 2436 | `AgPlanificadorPatronSemanal.cs`, `AgPlanificadorOleadas.cs`, `AgFabricaIndividuos.cs` | Implementado |
| 2437 | `AgParticionCarrilesOpcionesTests.cs`, `AgPlanificadorOleadasTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 556: Compatibilidad de enumeración por defecto, visibilidad de detector y sincronización de ranuras expandidas (Fases 2438–2440)

### Resumen de la Ronda
- **Fase 2438 (T1):** En `AgParticionCarriles.cs:103`, restitución del valor por defecto compatible `bool permitirEnumeracion = true` en la signatura de `Seleccionar` para salvaguardar todos los tests históricos de `AgParticionCarrilesTests` y suites de regresión. En `AgPlanificadorOleadas.cs:450-458`, pase explícito de `permitirEnumeracion: false` en el punto de producción para activar de forma quirúrgica la política catálogo-only sin afectar consumidores legados.
- **Fase 2439 (T2):** En `AgPlanificadorPatronSemanal.cs:185`, ajuste de visibilidad de `DetectarEsMarVie` de `private static` a `internal static`, resolviendo el error de compilación CS0122 ocasionado por su invocación en `AgFabricaIndividuos.cs`.
- **Fase 2440 (T3):** En `AgFabricaIndividuos.cs:421-448`, levantamiento del cómputo de `horasEfectivas` previo a la resolución de planes de oleada y derivación de `ranurasPatron` mediante `AgRanuras.PorVentana` filtradas por `bandaCompleta` al expandir singletons de 1 día a 3 días. Guarda condicional para no sobreescribir `ranurasPatron` con las ranuras singleton heredadas de `planOleada.RanurasPatron`. Se añade prueba unitaria en `AgFabricaIndividuosTests.cs`. Sincronización de Wave 556 y fila `R-227` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2438 | `AgParticionCarriles.cs`, `AgPlanificadorOleadas.cs` | Implementado |
| 2439 | `AgPlanificadorPatronSemanal.cs` | Implementado |
| 2440 | `AgFabricaIndividuos.cs`, `AgFabricaIndividuosTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 557: Cobertura de catálogo para MIE-DOM y completitud de escalera (Fase 2441)

### Resumen de la Ronda
- **Fase 2441 (T1):** En `AgOpcionesParticion.cs:40-52, 195-205`, incorporación de resolución determinista de catálogo para ventanas `3,4,5,6,7` (`MIE-DOM` retornando `5,6M|3,4,6T,7` ante franja 6T o `5,6M|3,4,7` / `5,6|3,4,7` según ranuras provistas), garantizando particiones binarias sin singletons que completan el segundo escalón de la escalera `VIE-DOM → JUE-DOM → MIE-DOM`. Pruebas unitarias de resolución y escalado en `AgOpcionesParticionTests.cs` y `AgPlanificadorOleadasTests.cs`. Sincronización de Wave 557 y fila `R-228` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2441 | `AgOpcionesParticion.cs`, `AgOpcionesParticionTests.cs`, `AgPlanificadorOleadasTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 558: Aserción exacta de partición MIE-DOM y cobertura de variante sin 6T (Fase 2442)

### Resumen de la Ronda
- **Fase 2442 (T1):** En `AgPlanificadorOleadasTests.cs:1185-1196`, refuerzo de la aserción de partición resultante al escalar a `MIE-DOM` verificando la clave exacta `"5.6|3.4.6.7"` y la presencia de franja MANANA en carril 0 y TARDE en carril 1 para el sábado (6). En `AgOpcionesParticionTests.cs:335-370`, verificación exacta de carriles y franjas para `5,6M|3,4,6T,7` y adición de caso de prueba unitaria para variante sin sábado tarde `5,6M|3,4,7`. Sincronización de Wave 558 y fila `R-229` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2442 | `AgPlanificadorOleadasTests.cs`, `AgOpcionesParticionTests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 559: Distinción de canónico vs catálogo, log con franjas y ampliación de ventana en R6 (Fases 2443–2445)

### Resumen de la Ronda
- **Fase 2443 (T1):** En `AgParticionCarriles.cs:195-235`, distinción estricta entre partición de catálogo y referencia canónica bajo política catálogo-only (`!permitirEnumeracion`): cuando no existe opción de catálogo aprobada pero la referencia es válida (sin singletons, sin patrones no alojables y sin infactibles), retorna `Origen = "canonico"` en lugar de etiquetarla indebidamente como `catalogo` con cero opciones evaluadas. En `AgPlanificadorOleadas.cs:480`, inclusión de `canonico` en el registro de log diagnóstico.
- **Fase 2444 (T2):** En `AgPlanificadorOleadas.cs:766-773`, generación de la clave de partición preservando tokens de franja para el sábado (`6M` y `6T`), evitando el falso diagnóstico visual de día 6 duplicado (`bandas={5,6M}|{6T,7}`).
- **Fase 2445 (T3):** En `AgPlanificadorCarril.cs:26-55, 145-195` y `AgReparadorR6.cs:73-88`, soporte de reintento con ventana ampliada (`AgEscaleraVentanas.Siguiente`) cuando la planificación de carril falla con `SIN_FECHAS` o `CAPACIDAD`, calculando fechas válidas extendidas dentro de la ventana académica del contexto y emitiendo traza diagnóstica estructurada para eliminar violaciones residuales R6.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2443 | `AgParticionCarriles.cs`, `AgPlanificadorOleadas.cs` | Implementado |
| 2444 | `AgPlanificadorOleadas.cs` | Implementado |
| 2445 | `AgPlanificadorCarril.cs`, `AgReparadorR6.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 560: Corrección de argumento diasVentana en llamada a Seleccionar y cobertura de canonico y ventana ampliada en R6 (Fases 2446–2447)

### Resumen de la Ronda
- **Fase 2446 (T1):** En `AgPlanificadorOleadas.cs:459-467`, supresión del argumento nombrado obsoleto `diasVentana: diasVentana` en la invocación a `AgParticionCarriles.Seleccionar`, alineando la llamada con la firma actual del método y resolviendo el error de compilación CS1739.
- **Fase 2447 (T2):** En `AgParticionCarrilesOpcionesTests.cs:920-960`, adición de prueba unitaria `Seleccionar_VieDomCanonico_ReportaCanonicoNoCatalogo` asegurando que una partición de fin de semana válida sin opciones evaluadas retorne deterministamente `Origen = "canonico"` y no `catalogo`. En `AgReparadorR6Tests.cs:685-705`, incorporación de test `Planificar_ConVentanaAmpliada_CuandoSinFechasEnVentanaOriginal_LograPlanificar` verificando que la provisión de `ventanaAmpliada` desbloquea la planificación ante carencia de fechas en la ventana original. Sincronización de Wave 560 y fila `R-231` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2446 | `AgPlanificadorOleadas.cs` | Implementado |
| 2447 | `AgParticionCarrilesOpcionesTests.cs`, `AgReparadorR6Tests.cs`, `REGISTRY.md` | Implementado |

---

## Ronda 561: Fortalecimiento de pruebas de reparación R6 y ventana ampliada (Fase 2448)

### Resumen de la Ronda
- **Fase 2448 (T1):** En `AgReparadorR6Tests.cs:685-730`, reemplazo de la aserción tautológica previa por pruebas estrictas de no-regresión: `Planificar_ConVentanaAmpliada_VerificaReparacionRealYSinSolape` valida `Factible = true`, `SesionesDisponibles > 0` e invariante de no-solapamiento de fechas efectivas en la cadena de asignaturas (152, 153, 151), y `RepararR6_CadenaSinFechas_AmpliaVentanaYReduceHard` verifica que la ejecución integral de `AgReparadorR6.Ejecutar` reduce efectivamente las infracciones duras (`HardPost < HardPre`). Sincronización de Wave 561 y fila `R-232` en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2448 | `AgReparadorR6Tests.cs`, `REGISTRY.md` | Implementado |














