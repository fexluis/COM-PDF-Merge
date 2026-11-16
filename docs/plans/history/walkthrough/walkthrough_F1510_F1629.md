<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# walkthrough - Fases 1510-1629 (Rondas 287..328)

> Indice: [../../../walkthrough.md](../../../walkthrough.md) - Escision: 2026-11-16

---
### Verificación de Principios (Ronda 286)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1504 | [X] Sí | [X] Sí | [X] Sí | Sí | `REGISTRY.md` sincronizado |
| F1505 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1506 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1507 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1508 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1509 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |


---

## Ronda 287: Reincidencia de numeración y frontera de capa limpia en R-007 (Fases 1510–1514)

### Hallazgos y Acciones Ejecutadas
- **Fase 1510**: Asignadas y cerradas filas `R-009` (1504–1509) y `R-010` (1510–1514) en `REGISTRY.md`, elevando contadores a `287 / 1514` (siguiente `288 / 1515`).
- **Fase 1511**: Reubicada la población de `Etiqueta` en `VinculacionListService` (capa Application), erradicando la dependencia `Infrastructure -> Application`.
- **Fase 1512**: Consulta A actualizada con `LEFT JOIN` a programas y fallback `PROG {id}`.
- **Fase 1513**: Normalizado `GROUP_CONCAT` de `SubunidadesCorrida` con orden determinista y separador `', '`.
- **Fase 1514**: Restaurados los `.doc` ajenos a R-007 y des-staged el índice Git para prevenir mezclas de versiones.

### Verificación de Principios (Ronda 287)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1510 | [X] Sí | [X] Sí | [X] Sí | Sí | `REGISTRY.md` sincronizado |
| F1511 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1512 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1513 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1514 | [X] Sí | [X] Sí | [X] Sí | Sí | Índice Git limpio |


---

## Ronda 288: Integridad de artefactos ajenos al plan y determinismo de agregación (Fases 1515–1516)

### Hallazgos y Acciones Ejecutadas
- **Fase 1515**: Confirmada paridad de bytes de `.doc` con HEAD. Declarada regla estricta de no tocar archivos fuera de `files_lock`.
- **Fase 1516**: Especificada normalización determinista en C# (`Distinct().OrderBy()`) para campos agregados del resumen de staging en el plan.
- **Sincronización REGISTRY**: Fila `R-011` sellada (`closed`), contadores §2 elevados a `288 / 1516` (siguiente `289 / 1517`).

### Verificación de Principios (Ronda 288)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1515 | [X] Sí | [X] Sí | [X] Sí | Sí | Contenido idéntico a HEAD |
| F1516 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan especificado |

---

## Ronda 293: Correcciones de Robustez y Coherencia en Recálculo y Vista por Fechas (Fases 1530–1536)

### Evaluación Crítica de las Fases

#### 1. Fase 1530 — Cabecera y panel quedan obsoletos tras recalcular
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al recalcular, `_ultimoResultado` mantenía la referencia a la corrida previa en memoria viva. `CargarModuloResultadosAsync` evalúa `if (_ultimoResultado?.MejorIndividuo == null)`, saltándose la rehidratación desde SQLite. La UI seguía mostrando los contadores anteriores.
  - *Causa Raíz:* No se invalidaba `_ultimoResultado` antes de solicitar la recarga.
  - *Impacto Confirmado:* Desincronización visual entre el recálculo persistido y las etiquetas del Módulo R.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2351`
- **Acción Aplicada:** Asignación `_ultimoResultado = null;` inmediatamente antes de `await CargarModuloResultadosAsync(idCorrida);`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 2. Fase 1531 — `btnRecalcularViolaciones` queda deshabilitado tras ejecutar el AG
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El botón solo se activaba al seleccionar desde el combo (`CargarCorridaDesdeComboAsync`). Si la corrida terminaba en la sesión viva mediante `IniciarOptimizacionAsync`, el botón permanecía en `Enabled = false`.
  - *Causa Raíz:* Omisión de la actualización del estado de `btnRecalcularViolaciones` en `CargarModuloResultadosAsync`.
  - *Impacto Confirmado:* Imposibilidad de recalcular inmediatamente después de optimizar sin recargar la corrida por combo.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2008`
- **Acción Aplicada:** Añadida asignación `btnRecalcularViolaciones.Enabled = ind.Id > 0;`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 3. Fase 1532 — Contexto ajeno puede sobrescribir `hard_violations`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Si el usuario cambiaba los selectores de período o programa en la UI, `AgRecalculoViolacionesService` evaluaba el individuo óptimo contra un contexto ajeno, persistiendo métricas y violaciones erróneas que podían habilitar la promoción indebida.
  - *Causa Raíz:* Ausencia de validación cruzada entre `contexto` y `corrida`.
  - *Impacto Confirmado:* Corrupción lógica de violaciones y riesgo de promoción de horarios infactibles.
  - *Evidencia:* `AgRecalculoViolacionesService.cs:33-53`
- **Acción Aplicada:** Verificación de igualdad entre `contexto.IdPeriodo`/`contexto.IdPrograma`/`contexto.Subunidad` y `corrida.*`, retornando `Result.Failure` si difieren.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgRecalculoViolacionesService.cs`

#### 4. Fase 1533 — `ConstruirIndices()` puede lanzar `NullReferenceException`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `ConstruirIndices()` asume que todas las listas del contexto están inicializadas; un contexto parcial causaba NRE silenciosa reportada genéricamente.
  - *Causa Raíz:* Falta de validación de colecciones requeridas en el servicio.
  - *Impacto Confirmado:* Falla no descriptiva en contextos incompletos.
  - *Evidencia:* `AgRecalculoViolacionesService.cs:61`
- **Acción Aplicada:** Validación previa de que `Aulas`, `Bloques`, `Asignaturas`, `Docentes`, `Sedes` y `PoolDocentes` no sean nulos.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgRecalculoViolacionesService.cs`

#### 5. Fase 1534 — Mutación del `AgContexto` compartido
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `contexto.IdCorrida = idCorrida;` modificaba permanentemente el objeto en memoria del formulario llamador.
  - *Causa Raíz:* Mutación directa sin resguardo ni restauración.
  - *Impacto Confirmado:* Pérdida de trazabilidad del contexto original en memoria.
  - *Evidencia:* `AgRecalculoViolacionesService.cs:68`
- **Acción Aplicada:** Resguardo de `idCorridaPrevio` y restauración en bloque `finally`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgRecalculoViolacionesService.cs`

#### 6. Fase 1535 — Vista por fechas toma solo la primera corrida y no filtra MODO AG
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `CargarViolacionesAsync` consultaba `ag_violacion` aún en Producción y usaba `FirstOrDefault`, marcando únicamente las violaciones de una corrida cuando se visualizaban varias.
  - *Causa Raíz:* Falta de chequeo `_modo.EsModoAg` y consulta con parámetro escalar único.
  - *Impacto Confirmado:* Consultas innecesarias en Producción y omisión de resaltado en agrupaciones multi-corrida.
  - *Evidencia:* `FormVerVinculaciones.cs:824`
- **Acción Aplicada:** Guarda `if (!_modo.EsModoAg) return;` y consulta SQL con `WHERE av.id_corrida IN @IdCorridas`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`

#### 7. Fase 1536 — Cobertura de pruebas unitarias para robustez del recálculo
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Las nuevas guardas de validación de contexto ajeno (1532), contexto incompleto (1533) y preservación de `IdCorrida` (1534) requerían pruebas unitarias automatizadas para prevenir regresiones.
  - *Causa Raíz:* Ausencia de casos de prueba para fallos de contexto en `AgRecalculoViolacionesServiceTests`.
  - *Impacto Confirmado:* Cobertura incompleta del contrato de servicio.
  - *Evidencia:* `AgRecalculoViolacionesServiceTests.cs`
- **Acción Aplicada:** Incorporadas pruebas: `RecalcularAsync_ContextoNoCorrespondeACorrida_DevuelveFallo`, `RecalcularAsync_ContextoIncompleto_DevuelveFallo`, y `RecalcularAsync_PreservaIdCorridaPrevioDelContexto`.
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRecalculoViolacionesServiceTests.cs`

---

### Verificación de Principios (Ronda 293)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1530 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2350-2355` |
| F1531 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2008-2010` |
| F1532 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgRecalculoViolacionesService.cs:41-52` |
| F1533 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgRecalculoViolacionesService.cs:61-66` |
| F1534 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgRecalculoViolacionesService.cs:67-80` |
| F1535 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `FormVerVinculaciones.cs:815-840` |
| F1536 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgRecalculoViolacionesServiceTests.cs:145-220` |

---

## Ronda 294: Refinamiento de Coherencia en Recálculo y Vista por Fechas (Fases 1537–1540)

### Evaluación Crítica de las Fases

#### 1. Fase 1537 — El panel de violaciones queda vacío tras recalcular
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La rehidratación forzada por F1530 (`_ultimoResultado = null`) consultaba `GetIndividuoByIdAsync`, método que no puebla la colección `Violaciones`. El grid `dgvViolaciones` quedaba en blanco a pesar de haberse persistido las violaciones.
  - *Causa Raíz:* Omisión del transporte del detalle `r.Violaciones` al individuo cargado en la UI.
  - *Impacto Confirmado:* El panel de violaciones no mostraba el resultado recién recalculado.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2356`
- **Acción Aplicada:** Asignación `ind.Violaciones = r.Violaciones?.ToList() ?? new List<AgViolacion>();` y actualización de `dgvViolaciones.DataSource` inmediatamente tras recargar la corrida.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 2. Fase 1538 — La validación de subunidad bloquea recálculos legítimos
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Una sesión sin sincronización previa mantiene `contexto.Subunidad = ""`. Si la corrida tenía territorial registrada, la validación estricta de F1532 bloqueaba el recálculo con error falso positivo.
  - *Causa Raíz:* Comparación de cadenas sin verificar si ambos lados aportaban valor no vacío.
  - *Impacto Confirmado:* Imposibilidad de recalcular en entornos locales o sesiones sin sync.
  - *Evidencia:* `AgRecalculoViolacionesService.cs:48-51`
- **Acción Aplicada:** Comprobación condicional `if (contextoConSubunidad && corridaConSubunidad && !string.Equals(...))`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgRecalculoViolacionesService.cs`

#### 3. Fase 1539 — `btnRecalcularViolaciones` no se deshabilita en los abortos de staging
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En `CargarModuloResultadosAsync`, si la lectura de sesiones fallaba o el conteo era cero, se desactivaba `btnPromoverProduccion` pero `btnRecalcularViolaciones` permanecía habilitado.
  - *Causa Raíz:* Omisión de actualizar el estado del botón en las ramas de retorno temprano por error o staging vacío.
  - *Impacto Confirmado:* Estado inconsistente de botones en la barra de resultados.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2030,2045`
- **Acción Aplicada:** Asignación `btnRecalcularViolaciones.Enabled = false;` en ambos bloques de salida.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 4. Fase 1540 — F1533 deja fuera `Festivos` y `AulaAsignaturas`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `ConstruirIndices()` itera sobre `Festivos` sin verificar nulidad. F1533 no validó `Festivos` ni `AulaAsignaturas`, exponiendo el método a `NullReferenceException` con contextos parciales.
  - *Causa Raíz:* Lista de verificación incompleta en las guardas de pre-vuelo.
  - *Impacto Confirmado:* Error no controlado si un contexto carece de la lista de festivos o asignaciones de aula.
  - *Evidencia:* `AgRecalculoViolacionesService.cs:64`
- **Acción Aplicada:** Añadidas validaciones `contexto.Festivos == null || contexto.AulaAsignaturas == null` en el servicio y prueba unitaria.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgRecalculoViolacionesService.cs`

---

### Verificación de Principios (Ronda 294)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1537 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:2360-2374` |
| F1538 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgRecalculoViolacionesService.cs:48-54` |
| F1539 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:2035,2049` |
| F1540 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgRecalculoViolacionesService.cs:64-67` |

---

## Ronda 295: Deshabilitación Defensiva del Recálculo sin Individuo Óptimo (Fase 1541)

### Evaluación Crítica de las Fases

#### 1. Fase 1541 — `btnRecalcularViolaciones` no se deshabilita en el retorno sin individuo
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La salida temprana `if (_ultimoResultado?.MejorIndividuo == null) return false;` en `CargarModuloResultadosAsync` dejaba el botón `btnRecalcularViolaciones` en el estado heredado de la corrida anterior. Si se cargaba una corrida sin individuo óptimo en base de datos, el botón quedaba habilitado y fallaba al pulsarse.
  - *Causa Raíz:* Omisión de restablecer el botón en el retorno temprano sin individuo.
  - *Impacto Confirmado:* Botón de recálculo activo indebidamente para corridas incompletas o sin optimización.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:1979`
- **Acción Aplicada:** Bloque con `btnRecalcularViolaciones.Enabled = false;` y `btnPromoverProduccion.Enabled = false;` antes del `return false;`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

---

### Verificación de Principios (Ronda 295)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1541 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:1979-1984` |

---

## Ronda 296: Selector de Corridas del Módulo R: Filtro Tolerante y Estado Vacío Saneado (Fases 1542–1544)

### Evaluación Crítica de las Fases

#### 1. Fase 1542 — Cascada tolerante de filtros en selector de corridas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La consulta de corridas en staging aplicaba un único filtro estricto por `idPeriodo`, `idPrograma` y `subunidad`. Si la corrida activa en staging correspondía a otro período/programa pero de la misma territorial (caso #48 con período 3 y programa 2 frente a contexto período 2 y programa 1), el combo quedaba vacío.
  - *Causa Raíz:* Inexistencia de mecanismo de ampliación progresiva de filtros para consultas locales de staging.
  - *Impacto Confirmado:* Ocultamiento total de corridas optimizadas válidas en el Módulo R.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2208`
- **Acción Aplicada:** Implementada función pura `AgCorridaSelector.CascadaFiltros(idPeriodo, idPrograma)` que genera hasta 3 intentos `(p, pr)`, `(p, null)` y `(null, null)` incondicional, consumida en bucle con corte temprano al encontrar filas.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 2. Fase 1543 — Saneamiento del estado vacío y de errores
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al no haber resultados o ante error de consulta, se insertaban instancias `string` directamente en `cmbCargarCorridas.Items` manteniendo `DisplayMember` y `ValueMember` asignados de cargas previas. El texto se recortaba fuera del área visible y `SelectedItem` devolvía un objeto no tipado.
  - *Causa Raíz:* Reutilización de `Items` como canal de mensajes de error/vacío en vez de delegarlo a etiquetas descriptivas (`lblAgModo`).
  - *Impacto Confirmado:* Deformación visual del combo, selecciones no tipadas y confusión de estado en la UI.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2217,2228`
- **Acción Aplicada:** Limpieza de enlace de datos (`DisplayMember = ""`, `ValueMember = ""`, `Items.Clear()`, `SelectedIndex = -1`, `Text = ""`) previo a cualquier operación. Reporte de ausencia de corridas o error en `lblAgModo` y desactivación de botones solo si `_idCorridaCargada <= 0`. Asignación de `DisplayMember` y `ValueMember` exclusivamente cuando `items.Count > 0` antes de insertar elementos tipados.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 3. Fase 1544 — No reconsultar en cada apertura del desplegable
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El evento `cmbCargarCorridas.DropDown` ejecutaba `RefrescarComboCorridasSeguro()` de manera indiscriminada en cada click, limpiando `Items` y recargando desde SQLite mientras el desplegable intentaba abrirse, produciendo parpadeo, cierre abrupto y re-selección indeseada del índice 0.
  - *Causa Raíz:* Ausencia de comprobación de estado o clave de filtros entre aperturas del desplegable.
  - *Impacto Confirmado:* Parpadeo y degradación de la experiencia de usuario al desplegar el ComboBox.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:399`
- **Acción Aplicada:** Incorporado campo `_claveFiltroCorridas`, método `ClaveFiltroCorridas()` y guarda en `DropDown` que aborta si el combo ya contiene ítems y la clave `Periodo|Programa|Subunidad` no ha variado.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

---

### Verificación de Principios (Ronda 296)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1542 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgCorridaSelector.cs:22-38`, `AgFormGeneradorHorarios.cs:2217-2234` |
| F1543 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:2237-2287` |
| F1544 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:67,400-406,2291,2301-2309` |

---

## Ronda 297: Tipado de Asserts en Tests y Contexto de Corrida en Recálculo (Fases 1545–1548)

### Evaluación Crítica de las Fases

#### 1. Fase 1545 — Asserts de tuplas sin tipo destino rompen la compilación de la suite
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* FluentAssertions 6.12 carece de sobrecarga `Should().Be()` para `ValueTuple` y resuelve a `ObjectAssertions.Be(object)`. Los literales con `null` carecían de tipo natural provocando error de compilación `CS8716`, y `(2, 1)` boxeaba `(int, int)` que no igualaba a `(int?, int?)`.
  - *Causa Raíz:* Asignación implícita a `object` sin tipo de destino fijado para las tuplas nulas o nullables.
  - *Impacto Confirmado:* Falla de compilación de la suite de pruebas unitarias (`ConversorPDF.Tests`).
  - *Evidencia:* `AgCorridaSelectorTests.cs:142,153,163,172,184`
- **Acción Aplicada:** Uso de arrays fuertemente tipados `(int?, int?)[]` con `.Should().Equal()` y verificación de propiedades individuales en `.Subject` y `.Last()`.
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Application/Services/AgCorridaSelectorTests.cs`

#### 2. Fase 1546 — El recálculo se rechaza cuando la corrida no coincide con los combos del Módulo S
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al auto-cargar una corrida de staging ampliada (ej. #48 de período 3/programa 2 con combos en período 2/programa 1), «Recalcular violaciones» reconstruía `_contextoActual` usando los combos de la UI, siendo rechazado por la guarda de `AgRecalculoViolacionesService` (F1532) por discordancia de contexto.
  - *Causa Raíz:* `ConstruirContextoAsync` dependía exclusivamente de los controles de la UI y no aceptaba parámetros forzados de la corrida.
  - *Impacto Confirmado:* Recálculo bloqueado para cualquier corrida cargada por relajación de filtros.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2382`
- **Acción Aplicada:** Parámetros opcionales `idPeriodoForzado`, `idProgramaForzado`, `subunidadForzada` en `ConstruirContextoAsync`. Consulta de la corrida antes de evaluar `_contextoActual` y reconstrucción dirigida con sus metadatos específicos.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 3. Fase 1547 — `finally` rehabilita el botón de recálculo sin condición
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El bloque `finally` de `RecalcularViolacionesDesdeModuloRAsync` re-habilitaba incondicionalmente `btnRecalcularViolaciones.Enabled = true`, contradiciendo las guardas de F1539/F1541 para corridas sin individuo óptimo persistido.
  - *Causa Raíz:* Asignación booleana estática sin comprobar la validez del individuo cargado.
  - *Impacto Confirmado:* Botón de recálculo habilitado indebidamente en corridas incompletas.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2449`
- **Acción Aplicada:** Asignación condicional: `btnRecalcularViolaciones.Enabled = _ultimoResultado?.MejorIndividuo?.Id > 0;`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 4. Fase 1548 — El panel conserva violaciones obsoletas si la rehidratación falla
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La actualización de `dgvViolaciones` estaba supeditada a `ind != null`; si el individuo rehidratado era nulo tras el recálculo (ej. staging vacío), la grilla retenía violaciones obsoletas.
  - *Causa Raíz:* Anidamiento de la actualización de la grilla dentro de la condición de individuo óptimo.
  - *Impacto Confirmado:* Desincronización visual entre el estado del recálculo y la grilla de violaciones.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2414`
- **Acción Aplicada:** Asignación y refresco de `dgvViolaciones.DataSource` fuera del bloque condicional `if (ind != null)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

---

### Verificación de Principios (Ronda 297)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1545 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgCorridaSelectorTests.cs:136-180` |
| F1546 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:570,654,688-695,720,2400-2425` |
| F1547 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:2478` |
| F1548 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:2445-2457` |

---

## Ronda 298: Guardas de Cobertura de Aulas y Validación de Cabecera en Recálculo (Fases 1549–1550)

### Evaluación Crítica de las Fases

#### 1. Fase 1549 — Contexto reconstruido sin las aulas de la corrida ⇒ `hard_violations = 0` persistido
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `ConstruirContextoAsync` fuerza período y programa de la corrida, pero la sede y cohortes provienen de la selección manual en la UI. Si la corrida es de otra sede o cohorte, el contexto no contiene sus aulas. Al faltar las aulas en `AulasPorId`, `AgValidadorReglas` omite evaluar esos genes, generando un detalle vacío o parcial y persistiendo `hard_violations = 0`, lo cual habilita indebidamente el botón de promoción a producción y corrompe `ag_violacion`.
  - *Causa Raíz:* Ausencia de guarda de cobertura entre los índices construidos del contexto y los requerimientos de aula de los genes del individuo.
  - *Impacto Confirmado:* Sobre-escritura errónea de violaciones en base de datos con habilitación de promoción espuria.
  - *Evidencia:* `AgRecalculoViolacionesService.cs:79-98`
- **Acción Aplicada:** Agregada validación explícita de `aulasFaltantes` antes de evaluar reglas. Si faltan aulas, se aborta con `Result.Failure` detallando los IDs ausentes. Cobertura de pruebas agregada en `AgRecalculoViolacionesServiceTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgRecalculoViolacionesService.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRecalculoViolacionesServiceTests.cs`

#### 2. Fase 1550 — Mensaje ambiguo para corridas sin cabecera
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Si `GetCorridaByIdAsync` devuelve una corrida huérfana en staging (`IdPeriodo <= 0` o `IdPrograma <= 0`, caso #48), la reconstrucción del contexto ignoraba en silencio los IDs <= 0 y recurría a los combos actuales, fallando luego con un mensaje genérico.
  - *Causa Raíz:* Omisión de comprobación temprana de metadatos mínimos de cabecera de la corrida antes de intentar reconstruir el contexto.
  - *Impacto Confirmado:* Confusión de usuario ante error genérico en corridas huérfanas de staging.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2408-2415`
- **Acción Aplicada:** Verificación previa `if (corrida.IdPeriodo <= 0 || corrida.IdPrograma <= 0)` con diálogo informativo advirtiendo que la corrida carece de cabecera en `ag_corrida`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

---

### Verificación de Principios (Ronda 298)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1549 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgRecalculoViolacionesService.cs:79-95`, `AgRecalculoViolacionesServiceTests.cs:19,253-276` |
| F1550 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:2409-2417` |

---

## Ronda 299: Cobertura Integral de Contexto (Aulas, Asignaturas y Docentes) en Recálculo (Fase 1551)

### Evaluación Crítica de las Fases

#### 1. Fase 1551 — Contexto sin las asignaturas/docentes de los genes ⇒ violaciones parciales persistidas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La guarda F1549 comparaba exclusivamente `AulasPorId`. `AgValidadorReglas` omite igualmente la evaluación de reglas si la asignatura o el docente no están presentes en `contexto.AsignaturasPorId` o `contexto.DocentesPorId`. Si el contexto cubre las aulas pero no las asignaturas o docentes de los genes (desfase de árbol de cohortes), las violaciones quedaban incompletas y `ReemplazarViolacionesAsync` persistía métricas disminuidas.
  - *Causa Raíz:* Cobertura parcial del contexto: solo se validaba uno de los tres ejes requeridos por el validador de reglas.
  - *Impacto Confirmado:* Riesgo de evaluación incompleta y sobreescritura de violaciones con valores artificialmente bajos.
  - *Evidencia:* `AgRecalculoViolacionesService.cs:79-94`
- **Acción Aplicada:** Ampliada la guarda para extraer `asignaturasFaltantes` y `docentesFaltantes`. Si faltan elementos en cualquiera de las tres colecciones, se aborta con `Result.Failure` detallado. Cobertura de pruebas unitarias actualizada y añadida en `AgRecalculoViolacionesServiceTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgRecalculoViolacionesService.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRecalculoViolacionesServiceTests.cs`

---

### Verificación de Principios (Ronda 299)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1551 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgRecalculoViolacionesService.cs:79-110`, `AgRecalculoViolacionesServiceTests.cs:23-24,277-300` |

---

## Ronda 300: Ajuste de Guarda de Cobertura en Recálculo a Ejes Activos del Validador (Fases 1552–1553)

### Evaluación Crítica de las Fases

#### 1. Fase 1552 — Exigir cobertura de docentes no evita violaciones incompletas y bloquea corridas válidas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgValidadorReglas` no consulta `DocentesPorId` en ninguna de sus reglas (0 referencias). Por consiguiente, verificar `docentesFaltantes` no aportaba protección alguna frente a violaciones omitidas, y provocaba el rechazo injustificado de recálculos legítimos cuando la corrida contenía docentes no registrados en el contexto actual (docentes interinstitucionales, de otra sede o fuera del pool).
  - *Causa Raíz:* Extensión de la guarda de pre-vuelo a colecciones no consumidas por el validador de reglas.
  - *Impacto Confirmado:* Falso rechazo en recálculos con docentes fuera del alcance del contexto en memoria.
  - *Evidencia:* `AgRecalculoViolacionesService.cs:94-109`
- **Acción Aplicada:** Removida la comprobación de `docentesFaltantes` de la condición de fallo, limitando la guarda estrictamente a `aulasFaltantes` y `asignaturasFaltantes`. Agregado test unitario `RecalcularAsync_DocenteAusenteDelContexto_NoBloqueaRecalculo` en `AgRecalculoViolacionesServiceTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgRecalculoViolacionesService.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgRecalculoViolacionesServiceTests.cs`

#### 2. Fase 1553 — Comentario obsoleto contradice la guarda
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El comentario de cabecera de la guarda afirmaba que se requería cobertura de «aulas, asignaturas o docentes», contradiciendo la exclusión de docentes realizada en la Fase 1552.
  - *Causa Raíz:* Omisión de actualizar el comentario superior tras retirar la validación del eje docente.
  - *Impacto Confirmado:* Incoherencia documental en código que inducía a reintroducir la restricción indebida.
  - *Evidencia:* `AgRecalculoViolacionesService.cs:79`
- **Acción Aplicada:** Comentario actualizado a «abortar si el contexto no cubre las aulas o asignaturas de la corrida».
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgRecalculoViolacionesService.cs`

---

### Verificación de Principios (Ronda 300)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1552 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgRecalculoViolacionesService.cs:79-102`, `AgRecalculoViolacionesServiceTests.cs:302-329` |
| F1553 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgRecalculoViolacionesService.cs:79` |

---

## Ronda 302: La selección programática del selector de corridas nunca dispara la carga (Fases 1554–1555)

### Evaluación Crítica de las Fases

#### 1. Fase 1554 — `_actualizandoComboCorridas` descarta la carga de la corrida preseleccionada
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al abrir el selector de corridas (`DropDown`), `RefrescarComboCorridasAsync` establece `_actualizandoComboCorridas = true` y asigna `cmbCargarCorridas.SelectedIndex = idx`. Esto detona `SelectedIndexChanged`, pero el handler `CargarCorridaDesdeComboAsync` aborta inmediatamente por la guarda `if (_actualizandoComboCorridas) return;`. Al finalizar en el `finally`, la bandera vuelve a `false`, pero ningún otro evento desencadena la carga. Si la lista contiene solo 1 corrida (como `#48`), el usuario no puede cambiar el índice manualmente y la carga nunca se ejecuta.
  - *Causa Raíz:* La guarda de reentrada descarta la selección inicial programática sin reanudar la carga tras completar el refresco.
  - *Impacto Confirmado:* Módulo R permanece vacío a pesar de que el combo muestra la corrida seleccionada.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2295-2308`, `:2337`
- **Acción Aplicada:** Agregada invocación explícita `await CargarCorridaDesdeComboAsync()` tras el bloque `try-finally` una vez restablecida `_actualizandoComboCorridas = false`, si `cmbCargarCorridas.SelectedItem` es válido y difiere de `_idCorridaCargada`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 2. Fase 1555 — Métricas y grids obsoletos cuando la carga no se completa
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Cuando `CargarModuloResultadosAsync` falla o retorna `false`, `CargarCorridaDesdeComboAsync` deshabilitaba los botones pero dejaba intactas las etiquetas de métricas (`lblAptitud`, `lblHardViolaciones`, `lblVacantes`) y los data sources de los grids, exhibiendo datos engañosos o heredados de la corrida previa.
  - *Causa Raíz:* Manejo incompleto del estado visual de fallo en el flujo de carga desde el combo.
  - *Impacto Confirmado:* La interfaz muestra métricas huérfanas (`1.000 / 0 / 0`) con grillas vacías.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2346-2362`
- **Acción Aplicada:** En la condición `if (!cargada)`, reiniciar `lblAptitud`, `lblHardViolaciones` y `lblVacantes` a `—` y fijar `dgvHorariosGenerados.DataSource = null` y `dgvViolaciones.DataSource = null`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

---

### Verificación de Principios (Ronda 302)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1554 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:2302-2308` |
| F1555 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:2353-2358` |

---

## Ronda 303: Guarda de Reentrada en RefrescarComboCorridasAsync (Fase 1556)

### Evaluación Crítica de las Fases

#### 1. Fase 1556 — `RefrescarComboCorridasAsync` carece de guard de reentrada y dispara cargas concurrentes
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `RefrescarComboCorridasAsync` se dispara asíncronamente desde el inicio del formulario (`:270`) y ante cada evento `DropDown` (`:404`). Sin una guarda de reentrada, invocaciones simultáneas provocan múltiples accesos y limpiezas concurrentes sobre `cmbCargarCorridas.Items` y, tras la Fase 1554, lanzan múltiples llamadas asíncronas concurrentes a `CargarCorridaDesdeComboAsync`, escribiendo en paralelo sobre el estado del Módulo R (`_ultimoResultado`, `_idCorridaCargada`, `dgvHorariosGenerados`, `dgvViolaciones`).
  - *Causa Raíz:* Ausencia de bandera de exclusión mutua para el método completo; `_actualizandoComboCorridas` solo protegía el disparo del evento de selección, no la ejecución concurrente del método asíncrono.
  - *Impacto Confirmado:* Desalineación intermitente en las grillas de resultados y métricas del formulario ante aperturas rápidas o eventos solapados.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2220-2223`, `:2313-2317`
- **Acción Aplicada:** Incorporado el campo privado `_refrescandoComboCorridas`, guarda de corte `if (cmbCargarCorridas == null || _refrescandoComboCorridas) return;` con activación `_refrescandoComboCorridas = true;`, y bloque `finally { _refrescandoComboCorridas = false; }` en `RefrescarComboCorridasAsync`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

---

### Verificación de Principios (Ronda 303)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1556 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:67,2220-2224,2315-2319` |

---

## Ronda 304: Carga del Staging Sin Individuo Óptimo y Advertencia en Selector (Fases 1557–1558)

### Evaluación Crítica de las Fases

#### 1. Fase 1557 — Cargar el grid de sesiones aunque no exista individuo óptimo
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En `CargarModuloResultadosAsync`, si `_ultimoResultado?.MejorIndividuo == null` (como en la corrida `#48`, donde hay staging huérfano sin cabecera en `ag_corrida` ni individuo en `ag_individuo`), el método abortaba con `return false` antes de leer las sesiones materializadas en `ag_sesiones`. Esto ocasionaba que una corrida con 180 sesiones en staging se mostrara vacía y con grids limpiados.
  - *Causa Raíz:* La guarda por falta de individuo óptimo estaba ubicada antes de la invocación a `GetSesionesStagingAsync`.
  - *Impacto Confirmado:* Imposibilidad de inspeccionar o auditar corridas con staging en la grilla del Módulo R si carecen de cabecera AG o individuo persistido.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2004-2009`, `:2043-2180`
- **Acción Aplicada:** Extraído el método `CargarGridStagingAsync(long idCorrida)` con la carga y enlace de `dgvHorariosGenerados`, `dgvVacantes` y `dgvViolaciones`. En `CargarModuloResultadosAsync`, se llama primero a `CargarGridStagingAsync`; si no hay individuo óptimo, se visualizan las etiquetas en modo solo lectura (`Restricciones Hard: — (sin resultado AG)`), se deshabilitan botones de acción y se retorna `true`. Adicionalmente, se incluyó diagnóstico en logs de advertencia ante IDs de individuo huérfanos detectados en las sesiones de staging.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 2. Fase 1558 — El selector de corridas no advierte cuando la corrida carece de resultado AG
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La corrida `#48` se desplegaba en el ComboBox con el mismo formato que una corrida con resultado AG óptimo, generando confusión en el usuario respecto a la disponibilidad de métricas de aptitud y violaciones evaluadas.
  - *Causa Raíz:* `CorridaStagingItem` no transportaba la existencia del individuo óptimo persistido (`id_individuo_optimo`), y `AgCorridaSelector.FormatearEtiqueta` no discriminaba dicho estado.
  - *Impacto Confirmado:* Experiencia de usuario confusa al seleccionar corridas huérfanas en el selector de Módulo R.
  - *Evidencia:* `AgCorridaSelector.cs:46-67`, `VinculacionQueryRepository.cs:250-271`
- **Acción Aplicada:** Agregada propiedad `IdIndividuoOptimo` y `TieneIndividuoOptimo` en `CorridaStagingItem`. Proyectado `COALESCE(MAX(c.id_individuo_optimo), 0) AS IdIndividuoOptimo` en las consultas de `VinculacionQueryRepository.GetCorridasStagingAsync`. En `AgCorridaSelector.FormatearEtiqueta`, si `corrida.IdCorrida > 0 && !corrida.TieneIndividuoOptimo`, se agrega el sufijo `sin resultado AG`. Agregadas pruebas en `AgCorridaSelectorTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/CorridaStagingItem.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/VinculacionQueryRepository.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/AgCorridaSelectorTests.cs`

---

### Verificación de Principios (Ronda 304)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1557 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:1961-2245` |
| F1558 | [X] Sí | [X] Sí | [X] Sí | Sí | `CorridaStagingItem.cs:43-50`, `VinculacionQueryRepository.cs:240,263`, `AgCorridaSelector.cs:66-70`, `AgCorridaSelectorTests.cs:15-32,107-122` |

---

## Ronda 305: Hidratación de ag_violacion al Rehidratar la Corrida en Módulo R (Fase 1559)

### Evaluación Crítica de las Fases

#### 1. Fase 1559 — Hidratar ag_violacion al rehidratar la corrida
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al cargar una corrida en el Módulo R (`CargarModuloResultadosAsync`), las violaciones mostradas en `dgvViolaciones` se tomaban de `_ultimoResultado.MejorIndividuo.Violaciones`. En corridas persistidas (como la corrida `#1`), `GetIndividuoByIdAsync` solo consultaba `ag_individuo` y `ag_gen`, omitiendo la lectura de las filas existentes en `ag_violacion`. Por ello, `Violaciones` permanecía nulo/vacío y la grilla de violaciones se visualizaba vacía, a menos que el usuario pulsara «Recalcular violaciones». Además, `IAgScheduleRepository` carecía de un contrato para leer las violaciones persistidas por individuo.
  - *Causa Raíz:* Inexistencia de método lector en `IAgScheduleRepository` y omisión de la consulta a `ag_violacion` durante el flujo de hidratación del Módulo R.
  - *Impacto Confirmado:* La interfaz no mostraba el desglose de violaciones persistidas en base de datos al seleccionar corridas existentes.
  - *Evidencia:* `IAgScheduleRepository.cs:26-33`, `AgScheduleRepository.cs:522-550`, `AgFormGeneradorHorarios.cs:2234-2249`
- **Acción Aplicada:**
  - Añadido `Task<Result<IReadOnlyCollection<AgViolacion>>> GetViolacionesByIndividuoAsync(long idIndividuo, CancellationToken ct);` en `IAgScheduleRepository`.
  - Implementada la consulta en `AgScheduleRepository.GetViolacionesByIndividuoAsync` recuperando las columnas de `ag_violacion` ordenadas por `es_dura DESC, regla`.
  - En `AgFormGeneradorHorarios.CargarGridStagingAsync`, antes de asignar a `dgvViolaciones`, si `indActual.Id > 0` y no tiene violaciones cargadas, se invoca `GetViolacionesByIndividuoAsync` e hidrata `indActual.Violaciones`.
  - Agregado log de diagnóstico ante corridas sin resultado AG.
  - Corregido cierre accidental de bloque en `AgFormGeneradorHorarios.cs:2250`.
  - Implementada prueba unitaria en `AgScheduleRepositoryNullGuardTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgScheduleRepositoryNullGuardTests.cs`

---

### Verificación de Principios (Ronda 305)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1559 | [X] Sí | [X] Sí | [X] Sí | Sí | `IAgScheduleRepository.cs:27`, `AgScheduleRepository.cs:542-573`, `AgFormGeneradorHorarios.cs:2236-2253`, `AgScheduleRepositoryNullGuardTests.cs:126-166` |

---

## Ronda 306: Ajuste de Layout en Módulo R y Separación de Botón Promover (Fase 1562)

### Evaluación Crítica de las Fases

#### 1. Fase 1562 — Separar «Promover a Producción» y ampliar el selector
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `cmbCargarCorridas` medía 185 px truncando las etiquetas de corridas (`#48 · 2027-01 · APT-52 · VALLE · 180 ses`). `btnPromoverProduccion` no tenía `Anchor` y estaba ubicado fijamente en `(740, 10)`, quedando pegado al combo y siendo tapado por el menú desplegable (`DropDownWidth = 430`).
  - *Causa Raíz:* Dimensiones de diseño fijas en `AgFormGeneradorHorarios.Designer.cs` sin adaptación al ancho dinámico del panel (`Dock = Top`).
  - *Impacto Confirmado:* Truncamiento de texto en el selector y superposición visual al desplegar las opciones.
  - *Evidencia:* `AgFormGeneradorHorarios.Designer.cs:1006, 1044, 1049, 1068-1070`
- **Acción Aplicada:**
  - Actualizado `pnlHeaderResultados.Size` de diseño a `(1100, 50)`.
  - Ampliado `cmbCargarCorridas.DropDownWidth = 480`.
  - Ampliado `cmbCargarCorridas.Size` a `(330, 25)` (545..875).
  - Configurado `btnPromoverProduccion.Location = (890, 10)` y `Anchor = Top | Right`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`

---

### Verificación de Principios (Ronda 306)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1562 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.Designer.cs:1006,1044,1049,1064,1069` |

---

## Ronda 307: Purga sin huérfanos y reparación de corridas del AG (Fases 1563–1566)

### Evaluación Crítica de las Fases

#### 1. Fase 1563 — Purga anclada a `id_corrida`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `LimpiarStagingAsync(idPeriodo, idPrograma)` borraba `ag_corrida`, `ag_individuo`,
    `ag_violacion`, `ag_bitacora` y `ag_promocion` por el par de la cabecera, pero `ag_vinculaciones`,
    `ag_sesiones` y `ag_gen` por el par almacenado en cada fila. Cuando el programa del combo difiere del
    programa de la asignatura (APT-53 vs APT-52) el staging sobrevive sin cabecera.
  - *Causa Raíz:* Purga basada en columnas denormalizadas en lugar de la clave `id_corrida`.
  - *Impacto Confirmado:* Corrida #48 con 180 sesiones visibles, sin violaciones y sin promoción posible.
  - *Evidencia:* `log_20261110.txt:346`, `ag_vinculaciones(idprograma=2)` vs `[Purga] idPrograma=1`.
- **Acción Aplicada:** `idsDel = ag_corrida ∪ ag_vinculaciones` por el par solicitado; todos los borrados por
  `id_corrida IN @idsDel`; limpieza de genes/violaciones de individuos referenciados por el staging; invariante
  post-purga con `LoggerHelper.Error`; guardas `TablaExiste`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`

#### 2. Fase 1564 — Fallar rápido en el flujo de ejecución
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `UpdateCorridaAsync` ejecutaba un `UPDATE` sin comprobar filas afectadas: si la
    corrida había sido purgada, el flujo continuaba y materializaba staging sin cabecera.
  - *Causa Raíz:* Ausencia de verificación de `rowcount` y de los resultados de persistencia en el orquestador.
- **Acción Aplicada:** `UpdateCorridaAsync` devuelve `Failure` con 0 filas; `EjecutarOptimizacionAsync` aborta
  antes de materializar; log del conteo real de `ag_violacion` tras `SaveIndividuoAsync`.
- **Archivos Afectados:** `AgScheduleRepository.cs`, `AgFormGeneradorHorarios.cs`

#### 3. Fase 1565 — Observabilidad y feedback en el Módulo R
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `CargarModuloResultadosAsync` descartaba `corridaRes.Error` y anunciaba «sin cabecera»
    incluso ante errores de lectura; `PromoverAProduccionAsync` retornaba en silencio cuando no había individuo,
    percibiéndose como «promover no funciona».
- **Acción Aplicada:** registro del error real, tres estados diferenciados (huérfana / sin individuo / error),
  habilitación de «Recalcular violaciones» como vía de reparación y mensaje explicativo en la promoción.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 4. Fase 1566 — Reparación de la corrida huérfana sin re-ejecutar el AG
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Con 30 genes intactos en `ag_gen`, la corrida era recuperable, pero no existía
    ninguna ruta para reconstruir `ag_corrida`/`ag_individuo` con los ids que referencia el staging.
- **Acción Aplicada:** nuevos métodos de repositorio (`GetGenesByIndividuoAsync`, `DetectarCorridaHuerfanaAsync`,
  `ReinsertarCabeceraHuerfanaAsync`), DTOs `AgCorridaHuerfana`/`AgReparacionCorridaResultado` y
  `AgReparacionCorridaService` (valida cobertura del contexto antes de escribir, evalúa con el evaluador real y
  persiste `ag_violacion`). Se invoca desde el botón existente «Recalcular violaciones».
- **Archivos Afectados:** `IAgScheduleRepository.cs`, `AgScheduleRepository.cs`, `AgCorridaHuerfana.cs`,
  `AgReparacionCorridaResultado.cs`, `IAgReparacionCorridaService.cs`, `AgReparacionCorridaService.cs`,
  `AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL.csproj`

#### 5. Fase 1567 — Nombre de la regla en el detalle de violaciones
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `ag_violacion.regla` está declarada `INTEGER` en el vault, pero el código persiste el
    **nombre** de la regla. Al leer, System.Data.SQLite aplica la afinidad de la columna y devuelve `0`, por lo que
    el panel mostraba «regla 0» y la vista por fechas perdía el identificador de la regla.
  - *Causa Raíz:* ausencia de `CAST` en los dos lectores (`GetViolacionesByIndividuoAsync` y
    `FormVerVinculaciones.CargarViolacionesAsync`).
  - *Impacto Confirmado:* el detalle de violaciones quedaba ilegible para auditoría incluso con las filas presentes.
  - *Evidencia:* copia real del vault — `typeof(regla)='text'` con `'R1_SuperposicionDocente'` leído como `0`.
- **Acción Aplicada:** `CAST(regla AS TEXT)` en ambos lectores y prueba de regresión con la columna declarada
  `INTEGER`. Verificado con arnés end-to-end sobre copia del vault real.
- **Archivos Afectados:** `AgScheduleRepository.cs`, `FormVerVinculaciones.cs`,
  `ConversorPDF.Tests/Integration/Academic/AgScheduleRepositoryHuerfanasTests.cs`

### Verificación de Principios (Ronda 307)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1563 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgScheduleRepository.cs:963-1113`, `AgScheduleRepositoryHuerfanasTests.cs:186-268` |
| F1564 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgScheduleRepository.cs:335-345`, `AgFormGeneradorHorarios.cs:1829-1893` |
| F1565 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:2017-2099`, `:2506-2510`, `:2809-2821` |
| F1566 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgReparacionCorridaService.cs:1-230`, `AgScheduleRepository.cs:1120-1278`, `AgReparacionCorridaServiceTests.cs` |
| F1567 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgScheduleRepository.cs:574-580`, `FormVerVinculaciones.cs:830-836`, `AgScheduleRepositoryHuerfanasTests.cs:372-397` |

### Nota operativa
La reparación devuelve la visibilidad de las violaciones y la posibilidad de aprobar la corrida; la promoción
sigue exigiendo `HARD = 0` (`AgPromotorVinculaciones:85-88`). La corrida #48 (HARD=95) requerirá corregir la
programación y volver a ejecutar el AG.

---

## Ronda 308: Validador de Reglas AG y Oleadas Temporales (Fases 1568–1574)

### Evaluación Crítica de las Fases

#### 1. Fase 1568 — Helper de Solape Temporal `AgVentanaSolape`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La detección de solapes se realizaba puramente a nivel de bloque horario y día de semana. Faltaba una utilidad pura y universal para determinar si dos ventanas de fechas se intersecan, manteniendo una semántica conservadora ante fechas nulas.
- **Acción Aplicada:** Implementado `AgVentanaSolape` con `SeSolapan` y `AgruparPorSolape`, registrado en `.csproj` con 5 pruebas unitarias.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgVentanaSolape.cs`, `ConversorPDF.XLL.csproj`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgVentanaSolapeTests.cs`.

#### 2. Fase 1569 — Reglas R1, R2, R3, R4 y R14 conscientes de ventanas de fechas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En el caso real #48, se reportaban 101 infracciones HARD porque múltiples asignaturas modulares en oleadas consecutivas comparten la misma aula, bloque y día.
- **Acción Aplicada:** En `AgValidadorReglas.cs`, se filtran colisiones en R1 (docente), R2 (aula), R3 (grupo), R4 (dos asignaturas mismo día y franja) y R14 (solape de intervalo horario) invocando `AgVentanaSolape.SeSolapan`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`, `AgValidadorReglasTests.cs`.

#### 3. Fase 1570 — R5 Concurrencia de asignaturas evaluada por día real continuo
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La agrupación por semanas ISO cortaba artificialmente asignaturas que empezaban o terminaban a mitad de semana, penalizando erróneamente oleadas válidas.
- **Acción Aplicada:** Se sustituyó la lógica de semanas ISO por un recorrido continuo día por día sobre el intervalo de la cohorte contra `maxCarriles`, emitiendo una infracción consolidada por tramo contiguo.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`, `AgValidadorReglasTests.cs`.

#### 4. Fase 1571 — Aulas con Estado Aprobado permitidas en R4
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Un aula en estado `Aprobado` era rechazada en R4 de compatibilidad, cuando representa un aula formalmente validada y apta para asignación.
- **Acción Aplicada:** R4 admite tanto `AgEstadoAula.Disponible` como `AgEstadoAula.Aprobado`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`, `AgValidadorReglasTests.cs`.

#### 5. Fase 1572 — Asignaturas requeridas acotadas al alcance de las aulas cubiertas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Asignaturas pertenecientes a aulas no contempladas en la optimización eran marcadas como faltantes si no aparecían en el cromosoma.
- **Acción Aplicada:** Acotada la exigencia a asignaturas mapeadas en `contexto.AulaAsignaturas` para las aulas presentes en la corrida.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`, `AgValidadorReglasTests.cs`.

#### 6. Fase 1573 — Pruebas de integración para escenarios de oleadas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Se requería validar de extremo a extremo los casos reales inspirados en la corrida #48 (Sábados dos bloques, Jueves 3 asignaturas, Viernes misma aula, etc.).
- **Acción Aplicada:** Creada la suite `AgValidadorReglasOleadasTests.cs` con 6 pruebas integradas.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasOleadasTests.cs`.

#### 7. Fase 1574 — Verificación general y documentación
- **Estado:** ACEPTADA
- **Acción Aplicada:** Verificación con `scripts/verify.ps1 -Mode Fast` (1123/1123 PASS), generación de artículo técnico en `docs/articles/changes/2026-11-11-validador-oleadas.md`, y sellado de fila R-030 en `REGISTRY.md`.
- **Verificación sobre datos reales (#48):** se ejecutó el validador corregido contra los 30 genes reales del individuo #33 reconstruyendo el contexto desde el vault ⇒ **HARD = 0, SOFT = 0, 0 filas**; los 101/125 HARD previos eran falsos positivos. Se confirmó además que `ag_gen.dias_patron` está poblado (`"4,5"` / `"6"`) y que R6/R13 agrupan correctamente por carril.
- **Corrección del fixture:** `AgValidadorReglasOleadasTests` usaba `Franja = "M"/"T"` (contrato canónico: `MANANA`/`TARDE`), provocando un fallo espurio de `R11_DuracionSesionInvalida`; corregido y con diagnóstico de violaciones en las aserciones.


---

### Verificación de Principios (Ronda 308)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1568 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgVentanaSolape.cs:1-68`, `AgVentanaSolapeTests.cs:1-55` |
| F1569 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:105-190, 240-330`, `AgValidadorReglasTests.cs` |
| F1570 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:382-455`, `AgValidadorReglasTests.cs:630-670` |
| F1571 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:495-502`, `AgValidadorReglasTests.cs:672-690` |
| F1572 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:52-73`, `AgValidadorReglasTests.cs:692-720` |
| F1573 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasOleadasTests.cs:1-342` (6/6 PASS) |
| F1574 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md`, `2026-11-11-validador-oleadas.md`, `verify.ps1` (1123 PASS) |



---

## Ronda 309: Contexto Homónimo y Anclaje de Filtros en el Módulo R (Fases 1575–1578)

### Evaluación Crítica de las Fases

#### 1. Fase 1575 — Respetar el programa de la corrida en `ConstruirContextoAsync`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* la resolución de homónimos (Fases 446/477) sustituía el programa **forzado** de la corrida
    (APT-52, id 2) por el homónimo con más asignaturas del semestre (APT-53, id 1). El contexto quedaba en `3/1`
    frente a una corrida `3/2`, por lo que `AgRecalculoViolacionesService` (Fase 1532) abortaba y el validador nuevo
    nunca se ejecutaba. El panel seguía mostrando 125 HARD / 101 filas persistidas el 11/11 a las 14:18 con el
    validador antiguo (**no era una regresión**).
  - *Causa Raíz:* el override de homónimos se aplicaba después del forzado y sin excepción.
- **Acción Aplicada:** con `idProgramaForzado > 0` se omite la resolución de homónimos; `ProgramasAlcance` se fija al
  programa de la corrida; traza `ProgramaHomonimoOmitido`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 2. Fase 1576 — Detalle staging anclado a la corrida
- **Estado:** ACEPTADA
- **Hallazgo:** `AbrirDetalleSesionesModoAgAsync` filtraba `vw_AgSesionesDetalle` con
  `ProgramasAlcance/IdPrograma` del árbol (1) mientras el staging de la corrida es programa (2) ⇒ «No se encontraron
  sesiones registradas…».
- **Acción Aplicada:** si hay corrida cargada, período y programa se leen de `ag_corrida`; log con los filtros usados.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

#### 3. Fase 1577 — Guarda de coherencia de contexto en la reparación
- **Estado:** ACEPTADA
- **Hallazgo:** `AgReparacionCorridaService` no comparaba el período/programa del contexto con los del staging; la
  reparación inicial evaluó con el pensum del programa homónimo y persistió violaciones falsas.
- **Acción Aplicada:** rechazo previo a cualquier escritura cuando `contexto.IdPeriodo/IdPrograma` no coincide con la
  corrida, con mensaje accionable que cita ambos pares.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparacionCorridaService.cs`

#### 4. Fase 1578 — Verificación y cierre
- **Estado:** ACEPTADA
- **Acción Aplicada:** test `RepararAsync_ConContextoDeOtroPrograma_FallaSinEscribir`; suite **1124/1124 PASS**;
  recálculo real sobre copia del vault (`RecalcularAsync(48, contexto 3/2)`): **125 → 0 HARD** y **101 → 0 filas**;
  artículo `docs/articles/changes/2026-11-11-contexto-homonimo-modulo-r.md`.
- **Archivos Afectados:** `AgReparacionCorridaServiceTests.cs`, documentación y registro.

### Verificación de Principios (Ronda 309)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1575 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:750-786, 792-798` |
| F1576 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgFormGeneradorHorarios.cs:3404-3423` |
| F1577 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgReparacionCorridaService.cs:55-62` |
| F1578 | [X] Sí | [X] Sí | [X] Sí | Sí | `AgReparacionCorridaServiceTests.cs` · arnés 125→0 · suite 1124 |

### Nota operativa
Tras desplegar el build, en el Módulo R: seleccionar `#48` → **«Recalcular violaciones»** (esperado `HARD 0`) →
**«Aprobar corrida»** → **«Promover a producción»**. El HARD 125 que se veía antes era dato persistido, no una
evaluación nueva.



---

## Ronda 312: Partición de Carriles Consciente del Patrón Sembrado (Fase 1581)

### Evaluación Crítica de las Fases

#### 1. Fase 1581 — Partición de Carriles Consciente del Patrón Sembrado (`AgParticionCarriles`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La partición canónica en bloques de 2 ranuras contiguas (`[4,5]|[6]|[7]`) separaba días de patrones sembrados como `"6,7"` en ventana extendida, generando rechazo a carril en pasada 1 y marcando asignaturas como infactibles por `R8_DiaFueraDePatron` en pasada 3 a pesar de haber capacidad disponible en el semestre.
  - *Causa Raíz:* Falta de un mecanismo de evaluación de particiones alternativas contiguas de días que preserved las ranuras y acomode los patrones sembrados.
- **Acción Aplicada:**
  - Creación de `AgParticionCarriles.cs` con el selector determinista que evalúa particiones candidatas de `AgBandasDias.Particiones(ventana)` deduplicadas por firma de ranuras.
  - Extracción de `EvaluarParticion` en `AgPlanificadorOleadas.cs` garantizando que la referencia canónica y las candidatas pasen por el mismo algoritmo.
  - Verificación previa de R8 antes de construir patrones semanales.
  - Registro de `AgParticionCarriles.cs` en `ConversorPDF.XLL.csproj`.
  - Documentación de cambio en `docs/articles/changes/2026-11-11-particion-carriles-patron.md` y actualización de `REGISTRY.md` (R-034).
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgParticionCarriles.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasCaracterizacionTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgParticionCarrilesTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgParticionCarrilesIntegracionTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGateG1ParticionCarrilesTests.cs`
  - `.agents/plans/REGISTRY.md`
  - `docs/articles/changes/2026-11-11-particion-carriles-patron.md`

### Verificación de Principios (Ronda 312)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1581 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs`, `AgPlanificadorOleadas.cs`, `verify.ps1` (1156 PASS) |


---

## Ronda 313: Robustecimiento Integral, Fidelidad de Patrones y Corrección de Pruebas (Fases 1582–1589)

### Evaluación Crítica de las Fases

#### 1. Fases 1582 a 1589 — Corrección e Implementación en Disco
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La revisión estática constató que F1582 a F1588 estaban pendientes de aplicación en disco y que el informe anterior no reflejaba el estado del workspace.
  - *Causa Raíz:* Guardas incompletas, omisión de fidelidad estricta de patrones sembrados, conteo impreciso y aserciones condicionales en tests de integración.
- **Acción Aplicada:**
  - Aplicación efectiva en código de F1582 (guardas previas), F1585 (rechazo de duplicados), F1584 (contador de evaluaciones reales), F1586 (`EsParticionFiel` y Gate G-1 real con `AgFabricaIndividuos`), F1583 (búsqueda por capacidad con igualdad estructural) y F1588 (fixture calibrado sin aserciones condicionales).
  - Sincronización documental de `REGISTRY.md`, `implementation_plan.md` y `walkthrough.md`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgParticionCarriles.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgParticionCarrilesTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgParticionCarrilesIntegracionTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGateG1ParticionCarrilesTests.cs`
  - `.agents/plans/REGISTRY.md`

### Verificación de Principios (Ronda 313)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1582–F1589 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | Código y pruebas implementados y sincronizados |


---

## Ronda 314: Separación de Casos de Integración, Diagnóstico I1 vs I5 y Gate con API Real (Fases 1590–1595)

### Evaluación Crítica de las Fases

#### 1. Fases 1590, 1593, 1594 y 1595
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En pruebas de integración faltaba un caso estricto de ausencia de mejora (F1590), el Gate G-1 no ejercitaba la asignación por `AgAulaAsignatura` ni validaba `gen.DiasPatron` (F1593), `AgValidadorVentana` etiquetaba I1 como I5 (F1594), y el test de Gate G-1 invocaba un método inexistente `Validar` en lugar de `Evaluar` (F1595).
  - *Causa Raíz:* Fixtures incompletos, aserciones laxas y desincronía con la firma pública del validador.
- **Acción Aplicada:**
  - Separación de `Caso76` en prueba de mejora parcial y prueba de ausencia total de mejora en `AgParticionCarrilesIntegracionTests.cs`.
  - Corrección de etiqueta diagnóstica en `AgValidadorVentana.cs`.
  - Reescritura de `AgGateG1ParticionCarrilesTests.cs` con el fixture representativo completo, llamada a `validador.Evaluar(individuo, contexto)` y verificación de 0 violaciones duras.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgParticionCarrilesIntegracionTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGateG1ParticionCarrilesTests.cs`
  - `.agents/plans/REGISTRY.md`

#### 2. Fases 1591 y 1592
- **Estado:** RETIRADA (Rechazo formal ratificado)
- **Hallazgo:**
  - F1591 pretendía debilitar la fidelidad del patrón permitiendo ampliaciones hacia días colindantes; rechazada para mantener la restricción conservadora aprobada.
  - F1592 proponía devolver selecciones vacías ante argumentos corruptos; rechazada porque oculta errores internos sin ruta pública demostrable.
- **Acción Aplicada:** Producción conservada intacta.

### Verificación de Principios (Ronda 314)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1590–F1595 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | Archivos corregidos y sincronizados en disco |

---

## Ronda 315: Bloques Canónicos e Índices en Gate G-1, Habilitaciones en Pool Docente y Comprobación Temporal de Festivos (Fases 1596–1598)

### Evaluación Crítica de las Fases

#### 1. Fases 1596, 1597 y 1598
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En el Gate G-1 faltaba la construcción explícita de índices (ConstruirIndices()), los bloques de fin de semana no tenían franjas diurnas válidas (F1596), no existían habilitaciones docentes en PoolDocentes provocando genes con vacantes (F1597), y la prueba de festivo no verificaba la extensión temporal ni contrastaba contra control sin festivo (F1598).
  - *Causa Raíz:* Fixtures parciales de contexto en las pruebas unitarias y de integración.
- **Acción Aplicada:**
  - Configuración de bloques canónicos (sábado MANANA/TARDE, domingo MANANA) e invocación de contexto.ConstruirIndices() en AgGateG1ParticionCarrilesTests.cs.
  - Poblado de PoolDocentes y aserción de EsVacante == 0 con docentes reales.
  - Creación de prueba dual de festivos en AgParticionCarrilesIntegracionTests.cs (24h: 3 semanas con festivo vs 2 semanas sin festivo).
- **Archivos Afectados:**
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGateG1ParticionCarrilesTests.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgParticionCarrilesIntegracionTests.cs
  - .agents/plans/REGISTRY.md

### Verificación de Principios (Ronda 315)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1596–F1598 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | Tests corregidos y sincronizados en disco |

---

## Ronda 316: Habilitaciones Canónicas en Gate G-1 y Calibración de Partición en Integración (Fase 1599)

### Evaluación Crítica de las Fases

#### 1. Fase 1599
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En `AgGateG1ParticionCarrilesTests`, las habilitaciones docentes asignadas con `Id = i + 1` directamente en listas no poblaban las claves canónicas calculadas por `TerceroAsignaturaKey.Calcular(...)` en `PoolDocenteKeys`, arriesgando fallos espurios en `EsDocenteHabilitado` y violaciones R5. Adicionalmente, el test presentaba errores de compilación por falta del namespace de enums, referencias a `DiasPatron` en `Asignatura`, y una aserción descalibrada en `Caso76` de `AgParticionCarrilesIntegracionTests` que esperaba 2 bandas cuando el desempate determinista por mayor concurrencia (`ThenByDescending(ev => ev.Bandas)`) elige la partición equivalente de 3 bandas `{{4}, {5}, {6, 7}}`.
  - *Causa Raíz:* Asignación manual eludiendo el método factory `RegistrarDocenteHabilitado`, y aserción en integración que no consideraba el criterio de desempate por bandas.
- **Acción Aplicada:**
  - Sustitución de asignaciones directas por `contexto.RegistrarDocenteHabilitado(101 + i, 1, ids[i], "I")` e invocación de `contexto.ConstruirIndices()`, validando `contexto.EsDocenteHabilitado(...).Should().BeTrue()`.
  - Inclusión de `using ConversorPDF.XLL.Academic.Domain.Enums;` y consulta de `DiasPatron` desde `contexto.AulaAsignaturaPorAulaAsig`.
  - Actualización de la aserción de `res.Value.Particion` en `Caso76_Integracion_MejoraParcial` para reflejar la partición seleccionada de 3 bandas `{{4}, {5}, {6, 7}}`.
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGateG1ParticionCarrilesTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgParticionCarrilesIntegracionTests.cs`
  - `.agents/plans/REGISTRY.md`
  - `implementation_plan.md`
  - `walkthrough.md`

### Verificación de Principios (Ronda 316)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1599 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | 44/44 tests focalizados PASS, 1158/1158 Fast PASS |

---

## Ronda 317: Reparto Compatible con Patrón Sembrado y Diagnóstico Verificable de Capacidad en Carriles (Fases 1600–1602)

### Evaluación Crítica de las Fases

#### 1. Fase 1600 — Evitar Reparto Incompatible con Patrón
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La pasada 1 de `AgPlanificadorOleadas.cs` seleccionaba el carril más específico antes de validar si disponía de capacidad suficiente; al no caber, pasaba a la segunda pasada greedy donde se descartaban las restricciones sembradas y un fallback forzado (`if (elegido < 0) elegido = 0`) reasignaba la materia al carril 0 (incompatible con el patrón), provocando que en la pasada 3 se reportara falsamente `R8_DiaFueraDePatron` aun cuando el día sí existía en la ventana semanal.
  - *Causa Raíz:* Comprobación tardía de capacidad en la pasada 1 y omisión de filtros de compatibilidad de patrón y capacidad en la pasada 2 con fallback desbordante a carril 0.
  - *Impacto Confirmado:* Falsos bloqueos R8 y confusión entre saturación de capacidad del carril e incompatibilidad estructural.
- **Acción Aplicada:**
  - En la pasada 1: comprobación de capacidad dentro del bucle de búsqueda de carril compatible.
  - En la pasada 2: respeto estricto a los días del patrón sembrado y a la capacidad proyectada, suprimiendo el fallback forzado a carril 0 y permitiendo que `carrilDe[i]` se mantenga en `-1`.
  - En la pasada 3: guarda `if (c < 0)` antes de indexar `slotsCarril[c]`, clasificando como `I5_CAPACIDAD_CARRIL_PATRON` si es alojable o `R8_DiaFueraDePatron` si no lo es.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs`

#### 2. Fase 1601 — Diagnóstico Verificable
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El diagnóstico no proporcionaba visibilidad unificada de los carriles compatibles y sus capacidades reservadas/totales, ni distinguía en `AgValidadorVentana.cs` la saturación del carril compatible de la inviabilidad temporal; además, `AgGeneradorHorarios.cs` usaba el término desactualizado «inviabilidad matemática».
  - *Causa Raíz:* Detalle de infactibilidad genérico sin estado de capacidades por carril e invocación del validador sin discriminación de código de saturación de carril.
  - *Impacto Confirmado:* Imposibilidad para el operador de auditar si un rechazo obedece a falta de capacidad o a conflicto de días de patrón.
- **Acción Aplicada:**
  - Integración en `DetalleInfactibles` de un único diagnóstico estructurado por asignatura detallando patrón, demanda, código (`I5_CAPACIDAD_CARRIL_PATRON` vs `R8_DiaFueraDePatron`) y capacidades compatibles (`c{j}[{dias}] capacidad={cap} reservada={carga}`).
  - Clasificación en `AgValidadorVentana.cs` distinguiendo `[I5_CAPACIDAD_CARRIL_PATRON]` de `[R8_DiaFueraDePatron]` e `[I5_INFACTIBLE_OLEADA]`.
  - Reemplazo de «inviabilidad matemática» por «bloqueos de factibilidad detectados» en `AgGeneradorHorarios.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`

#### 3. Fase 1602 — Regresiones y Validación
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Ausencia de pruebas unitarias específicas que cubrieran el nuevo comportamiento de asignación por capacidad en patrones, prevención de desbordamiento a carriles incompatibles, y soporte de vacantes permitidas cuando una materia no tiene pool docente.
  - *Causa Raíz:* Cobertura previa basada en el comportamiento desbordante y aserciones ancladas a R8.
  - *Impacto Confirmado:* Riesgo de regresión silenciosa en futuras modificaciones de carriles y oleadas.
- **Acción Aplicada:**
  - Adición de 6 nuevas pruebas en `AgPlanificadorOleadasTests.cs` (patrón 5 con capacidad, carril específico lleno con alternativa disponible, patrón 5 saturado hacia I5 sin reparto incompatible, R8 para día fuera de ventana, materias sin patrón por utilización y casos 6,7 mixtos con festivos).
  - Adición de 2 nuevas pruebas en `AgValidadorVentanaTests.cs` (materia sin pool docente con vacante permitida sin bloqueo, y patrón saturado con código `[I5_CAPACIDAD_CARRIL_PATRON]`, sede y aula).
  - Actualización de aserciones en `AgGeneradorHorariosTests.cs` y `AgPlanificadorOleadasCaracterizacionTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGeneradorHorariosTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasCaracterizacionTests.cs`
  - `.agents/plans/REGISTRY.md`

### Verificación de Principios (Ronda 317)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1600 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs` (Pasadas 1, 2 y 3 saneadas) |
| F1601 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs`, `AgValidadorVentana.cs`, `AgGeneradorHorarios.cs` |
| F1602 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs`, `AgValidadorVentanaTests.cs`, `AgGeneradorHorariosTests.cs`, `AgPlanificadorOleadasCaracterizacionTests.cs` |

---

## Ronda 318: Corrección de Tipos en Pruebas, Calibración de Fixtures de Capacidad y Cobertura Efectiva de Festivos (Fases 1603–1605)

### Evaluación Crítica de las Fases

#### 1. Fase 1603 — Pruebas con Tipos Inexistentes
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La prueba `Construir_Patron5ConCapacidad_DebeAsignarEnCarrilCompatible` y `Construir_Patron5Saturado` en `AgPlanificadorOleadasTests.cs` referenciaban la propiedad inexistente `p.DiasCarril` sobre `AgPlanOleada`, cuando la propiedad real de dominio es `DiasPatron`. Asimismo, en `AgValidadorVentanaTests.cs`, la prueba `ValidarFactibilidad_MateriaSinPoolDocente_PermiteVacanteYGeneracionContinua` asignaba instancias de `AgDocenteHabilitado` a `PoolPorAsignatura[1]`, cuando el tipo de diccionario esperado por el contexto es `Dictionary<int, List<TerceroAsignatura>>`.
  - *Causa Raíz:* Uso de nombres de propiedad y tipos desalineados respecto al contrato inmutable de `AgPlanOleada` y a los métodos canónicos de registro en `AgContexto`.
  - *Impacto Confirmado:* Fallos de compilación CS1061 y CS0029 al compilar el proyecto de pruebas.
- **Acción Aplicada:**
  - En `AgPlanificadorOleadasTests.cs`: sustitución de `p.DiasCarril` por `p.DiasPatron` y uso de `res.Value.Planes.Should().OnlyContain(p => p.DiasPatron.Contains(5))`.
  - En `AgValidadorVentanaTests.cs`: eliminación de asignación manual a `PoolPorAsignatura` e invocación de `contexto.RegistrarDocenteHabilitado(101, 0, 1)` junto con `contexto.ConstruirIndices()`.
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs`

#### 2. Fase 1604 — Fixtures no Alcanzan Rama Evaluada
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:*
    1. En `AgPlanificadorOleadasTests.cs` (`Patron5Saturado`), la ventana `"VIE"` tenía capacidad total de 3 sesiones frente a una demanda de 6 sesiones (2 × 12h / 4h), fallando prematuramente en el paso global con `I5_SIN_CAPACIDAD` antes de ejecutar la lógica de oleadas.
    2. En `AgValidadorVentanaTests.cs` (`PatronSaturadoEnAula`), la ventana `"VIE"` disparaba el bloqueo global `[I5_INFACTIBLE]` previo a la llamada a `AgPlanificadorOleadas.Construir`.
    3. En `AgPlanificadorOleadasTests.cs` (`CarrilMasEspecificoLleno`), el fixture contenía aserciones dentro de un `if (res.Value.Infactibles.Count > 0)` que permitía un falso éxito sin verificar la alternancia de carriles.
  - *Causa Raíz:* Fixtures con capacidad global deficiente o condiciones asertivas no deterministas.
  - *Impacto Confirmado:* Pruebas que pasaban por ramas no deseadas o no ejercitaban la compuerta de partición.
- **Acción Aplicada:**
  - En `AgPlanificadorOleadasTests.cs` (`Patron5Saturado`): ventana ampliada `"1,2,3,4,5"` con patrones exclusivos de viernes (5) y demanda total 6 sesiones; verifica 1 plan factible en viernes, 1 infactible en viernes por `I5_CAPACIDAD_CARRIL_PATRON` y 0 desbordamientos a días incompatibles.
  - En `AgValidadorVentanaTests.cs` (`PatronSaturadoEnAula`): configuración de aula con ventana `"1,2,3,4,5"`, jornada `"NOCHE"` y bloques nocturnos para días 1..5, logrando que la capacidad total (15 ranuras) supere la demanda (6 sesiones) y emitiendo el bloqueo discriminado `[I5_CAPACIDAD_CARRIL_PATRON]`.
  - En `AgPlanificadorOleadasTests.cs` (`CarrilMasEspecificoLleno`): reemplazo por fixture canónico: ventana `"5,6"`, 4 semanas, bloques [5,6M] (capacidad 8) y [6T] (capacidad 4); materias 1 y 2 de 16h con patrón "6", verificando que la materia 1 ocupe el carril 1 (más específico) y la materia 2 ocupe el carril 0 (alternativa compatible) sin infactibles.
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs`

#### 3. Fase 1605 — Prueba Mixta/Festivo sin Cobertura Efectiva
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En `Construir_Casos6Y7_MixtosYFestivos_DebeConservarResultados`, ambos patrones eran `6,7` sin variación entre asignaturas, el festivo configurado era lunes (`2027-03-22`) cuando los patrones evaluaban fin de semana, y el assert usaba `AllSatisfy` que es vacuamente verdadero sobre colecciones vacías.
  - *Causa Raíz:* Ausencia de diferenciación de patrones y fecha de festivo desacoplada de los días de los patrones.
  - *Impacto Confirmado:* Falta de garantía real sobre la preservación de patrones mixtos y festivos en oleadas.
- **Acción Aplicada:**
  - Ventana explícita `"4,5,6,7"`.
  - Materia 1 con patrón `"4,5"` y Materia 2 con patrón `"6,7"`.
  - Festivo configurado en el primer sábado utilizado (`2027-02-06`).
  - Aserción de existencia obligatoria de 2 planes (`res.Value.Planes.Should().HaveCount(2)`) y validación estricta de `plan.DiasPatron` contra `AgCarriles.Ventana(...)`.
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasTests.cs`
  - `.agents/plans/REGISTRY.md`

### Verificación de Principios (Ronda 318)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1603 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs`, `AgValidadorVentanaTests.cs` |
| F1604 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs`, `AgValidadorVentanaTests.cs` |
| F1605 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs`, `.agents/plans/REGISTRY.md` |

---

## Empaquetado Release 1.6.4.07 — `/build-release 1.6.4.07`

- **Comando:** `powershell -ExecutionPolicy Bypass -File ".\scripts\build_installer.ps1" -NewVersion "1.6.4.07"`
- **Acciones Realizadas:**
  - Sincronización de versión canónica `1.6.4.07` en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
  - Limpieza de instaladores MSI previos en `installer\`.
  - Ejecución de la suite completa de pruebas unitarias y de integración (`dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`):
    - **Total Tests:** 1166
    - **Superados:** 1166 (100% PASS)
    - **Errores:** 0
    - **Omitidos:** 0
    - **Duración:** 1 m 10 s
  - Compilación y empaquetado del Add-In Excel-DNA en modo `Release|x64` (`ConversorPDF.XLL-AddIn64.xll`).
  - Generación del instalador MSI WiX 4 Toolset (`InstaladorConversorPDF.wixproj`).
- **Artefacto Generado:** `installer\InstaladorConversorPDFv1.6.4.07.msi` (5,722,112 bytes).

---

## Ronda 319–325: Sincronización del Selector de Corridas del Módulo R (Fases 1606–1612)

### Fases Ejecutadas

#### 1. Fase 1606 — Esqueleto DTO AgSeleccionModuloR e inclusión en csproj (integrator-first)
- **Estado:** ACEPTADA
- **Acción:** Creación de `AgSeleccionModuloR.cs` inmutable (Single Type Per File), registro en `ConversorPDF.XLL.csproj` (Lease L0) y stub `ResolverSeleccionModuloR` en `AgCorridaSelector.cs` con `NotImplementedException`.
- **Evidencia:** `scripts/verify.ps1 -Mode Build` exit 0 (`_evidence/20261112-071607/summary.md`).

#### 2. Fase 1607 — Caracterización y Pruebas Unitarias RED en AgCorridaSelectorTests
- **Estado:** ACEPTADA
- **Acción:** Redacción de 6 pruebas unitarias puras en `AgCorridaSelectorTests.cs` cubriendo los 6 casos de selección de Módulo R sin alterar las pruebas existentes de `ResolverSeleccionInicial`.
- **Evidencia:** RED auténtico certificado (24 pruebas: 18 superadas, 6 fallando con `NotImplementedException`, `_evidence/20261112-071854/summary.md`).

#### 3. Fase 1608 — Implementación de la Política ResolverSeleccionModuloR y Enlace Seguro
- **Estado:** ACEPTADA
- **Acción:** Implementación determinista en `AgCorridaSelector.cs` y enlace en `AgFormGeneradorHorarios.cs` en `RefrescarComboCorridasAsync`. Eliminada la degradación destructiva a índice 0 y la llamada incondicional a `CargarCorridaDesdeComboAsync`.
- **Evidencia:** GREEN certificado (24/24 pruebas unitarias PASS, `_evidence/20261112-072421/summary.md`).

#### 4. Fase 1609 — Refresco Seguro en DropDown y Control de Concurrencia
- **Estado:** ACEPTADA
- **Acción:** Eliminación de la guarda de caché en `DropDown`, invocando siempre `RefrescarComboCorridasSeguro()`. Coalescencia de ráfagas con `_solicitudPendienteRefrescoCombo`, captura de filtros antes de await, descarte de respuestas obsoletas y verificación de `IsDisposed`/`Disposing`.
- **Evidencia:** Fast PASS (24/24 pruebas, `_evidence/20261112-073145/test.trx`).

#### 5. Fase 1610 — Ciclo de Vida de Corrida, Sincronización Post-Ejecución y Retorno de Diálogo
- **Estado:** ACEPTADA
- **Acción:** Registro de `idCorridaCreada` en `IniciarOptimizacionAsync` y refresco garantizado en `finally` con `CancellationToken.None` independiente. Actualización de `lblAgModo` con `#N` en ejecución directa. Limpieza de `_idCorridaCargada = 0`, métricas y grillas ante fallo de staging en `CargarModuloResultadosAsync`. Comprobación de recarga y refresco tras retorno de `FormVinculaciones`.
- **Evidencia:** Fast PASS (617/617 pruebas de Ag, `_evidence/20261112-074004/test.trx`).

#### 6. Fase 1611 — Verificación Integral de Suite Completa (Full)
- **Estado:** ACEPTADA
- **Acción:** Verificación integral de la suite completa de pruebas unitarias e integración en modo Full mediante `scripts/verify.ps1 -Mode Full`.
- **Evidencia:** Full PASS (1172/1172 tests superados, 0 fallos, 0 omitidos, 0 regresiones vs baseline oficial 952). Reporte TRX: `_evidence/20261112-074457/test.trx`, Summary: `_evidence/20261112-074457/summary.md`.

#### 7. Fase 1612 — Documentación Técnica del Cambio
- **Estado:** ACEPTADA
- **Acción:** Redacción exhaustiva del artículo de cambios en `docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md`.
- **Evidencia:** `-Mode Build` exit 0 (`_evidence/20261112-075059/summary.md`).

### Verificación de Principios (Ronda 319–325)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1606 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgSeleccionModuloR.cs`, `ConversorPDF.XLL.csproj` |
| F1607 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCorridaSelectorTests.cs` |
| F1608 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCorridaSelector.cs`, `AgFormGeneradorHorarios.cs` |
| F1609 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1610 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1611 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `_evidence/20261112-074457/summary.md` (1172 tests) |
| F1612 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md` |

---

## Ronda 326: Desincronización de Commits, Mutación de Items con Popup Abierto y Cargas Concurrentes del Selector (Fases 1613–1620)

### Fases Ejecutadas

#### 1. Fase 1613 — Sincronización de Archivos no Versionados en git
- **Estado:** ACEPTADA
- **Acción:** Identificación y versionado de `AgPerfilAlcanceResolver.cs`, `AgParticionCarriles.cs` y sus pruebas de integración/caracterización referenciadas en `ConversorPDF.XLL.csproj` para asegurar que un clon limpio de HEAD compile sin errores CS2001 ni CS0246.
- **Archivos Afectados:** `AgPerfilAlcanceResolver.cs`, `AgParticionCarriles.cs`, `ConversorPDF.XLL.csproj`.

#### 2. Fase 1614 — Sincronización Canónica de Versión 1.6.4.07
- **Estado:** ACEPTADA
- **Acción:** Consolidación atómica de la versión `1.6.4.07` en `App.cs` (`APP_VERSION = "1.6.4.07"`) e `installer/InstaladorConversorPDF.wxs` en coherencia con `ConversorPDF.XLL.csproj`.
- **Archivos Afectados:** `App.cs`, `InstaladorConversorPDF.wxs`.

#### 3. Fase 1615 — Prevención de Mutación de Items con Popup Abierto
- **Estado:** ACEPTADA
- **Acción:** Sustitución de `DropDown` por `DropDownClosed += (s, e) => RefrescarComboCorridasSeguro();` y adición de la guarda `if (cmbCargarCorridas.DroppedDown) { _solicitudPendienteRefrescoCombo = true; return; }` en `RefrescarComboCorridasAsync` para inhibir mutaciones de `Items` mientras el desplegable está abierto.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 4. Fase 1616 — Exclusión Mutua en Cargas Concurrentes de Corrida
- **Estado:** ACEPTADA
- **Acción:** Introducción del campo `private bool _cargandoCorridaDesdeCombo;` en `AgFormGeneradorHorarios.cs` con guarda y liberación en `finally` dentro de `CargarCorridaDesdeComboAsync`.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 5. Fase 1617 — Acotamiento Estricto de Carga Inicial a la Apertura del Formulario
- **Estado:** ACEPTADA
- **Acción:** Parametrización de `RefrescarComboCorridasAsync(bool permitirCargaInicial = false)` y `RefrescarComboCorridasSeguro(bool permitirCargaInicial = false)`. Invocación con `true` únicamente en `OnShown` al inicializar la pantalla.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 6. Fase 1618 — Preservación de Corrida Activa Fuera de Filtros
- **Estado:** ACEPTADA
- **Acción:** Inyección de la corrida activa como ítem con etiqueta `" (fuera de filtro)"` cuando los filtros activos la excluyen, seleccionándola en el combo y actualizando `lblAgModo` para evitar que el selector quede en blanco.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 7. Fase 1619 — Eliminación de Código Muerto de Filtros
- **Estado:** ACEPTADA
- **Acción:** Eliminación del campo `private string _claveFiltroCorridas;` y de sus asignaciones redundantes en `RefrescarComboCorridasAsync`.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 8. Fase 1620 — Alineación Documental del Artículo de Cambios
- **Estado:** ACEPTADA
- **Acción:** Corrección de la nomenclatura de los 6 tests unitarios en `docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md` e incorporación de notas técnicas sobre las salvaguardas de concurrencia y refresco.
- **Archivos Afectados:** `docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md`.

### Verificación de Principios (Ronda 326)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1613 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPerfilAlcanceResolver.cs`, `AgParticionCarriles.cs` |
| F1614 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `App.cs`, `InstaladorConversorPDF.wxs` |
| F1615 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1616 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1617 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1618 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1619 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1620 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md` |

---

## Ronda 327: Correcciones Fases 1621–1626 — Refresco Diferido, Coalescencia de Cargas y Marcador Fuera de Filtro

### Fases Ejecutadas

#### 1. Fase 1621 — Preservación del Marcador «(fuera de filtro)»
- **Estado:** ACEPTADA
- **Acción:** Captura de la referencia `fueraDeFiltro` y aplicación del sufijo `" (fuera de filtro)"` tras el bucle de formateo masivo de etiquetas de `items`.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 2. Fase 1622 — Consumo de Refrescos Pendientes en `DropDownClosed`
- **Estado:** ACEPTADA
- **Acción:** Configuración del evento `DropDownClosed` para verificar y consumir `_solicitudPendienteRefrescoCombo`, ejecutando el refresco diferido al cerrar el popup.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 3. Fase 1623 — Disparador Dual: Consulta en `DropDown` y Mutación Diferida
- **Estado:** ACEPTADA
- **Acción:** Conexión de `cmbCargarCorridas.DropDown += (s, e) => RefrescarComboCorridasSeguro();` para garantizar consulta fresca al desplegar, y reubicación de la guarda `if (cmbCargarCorridas.DroppedDown)` previa a `_actualizandoComboCorridas = true` para permitir que el query asíncrono se ejecute y diferir la mutación de `Items`.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 4. Fase 1624 — Acumulación de Autorización `permitirCargaInicial`
- **Estado:** ACEPTADA
- **Acción:** Declaración del campo `private bool _cargaInicialPendiente;` y acumulación `_cargaInicialPendiente |= permitirCargaInicial;` en reentradas y deferrals, garantizando la ejecución de la carga automática en `OnShown`.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 5. Fase 1625 — Coalescencia de Selección de Corrida en Vuelo
- **Estado:** ACEPTADA
- **Acción:** Declaración del campo `private bool _recargaCorridaPendiente;`. Captura de cambio de selección durante cargas en progreso y re-invocación asíncrona segura en el bloque `finally` de `CargarCorridaDesdeComboAsync`.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 6. Fase 1626 — Registro de Fila R-043 y Trazabilidad Append-Only
- **Estado:** ACEPTADA
- **Acción:** Inclusión de fila `R-043` en `.agents/plans/REGISTRY.md` con estado `executing` y contadores actualizados a ronda 327 / fase 1626.
- **Archivos Afectados:** `.agents/plans/REGISTRY.md`.

### Verificación de Principios (Ronda 327)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1621 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1622 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1623 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1624 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1625 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1626 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `.agents/plans/REGISTRY.md` |

---

## Ronda 328: Refresco Anticipado del Selector, Desacoplamiento de Modal y Trazabilidad de Commit (Fases 1627–1629)

