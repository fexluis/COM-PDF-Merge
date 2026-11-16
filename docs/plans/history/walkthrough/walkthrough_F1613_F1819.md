<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# walkthrough - Fases 1613-1819 (Rondas 329..398)

> Indice: [../../../walkthrough.md](../../../walkthrough.md) - Escision: 2026-11-16

---
### Fases Ejecutadas

#### 1. Fase 1627 — Saneamiento de Estado de Fases 1613/1614
- **Estado:** ACEPTADA
- **Acción:** Actualización documental clarificando que los archivos `AgPerfilAlcanceResolver.cs` y `AgParticionCarriles.cs` y sus tests están integrados en el working tree y pendientes de commit en modo Act.
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`.

#### 2. Fase 1628 — Refresco Anticipado en `Enter` y `MouseDown`
- **Estado:** ACEPTADA
- **Acción:** Enlace de eventos `cmbCargarCorridas.Enter` y `cmbCargarCorridas.MouseDown` (botón izquierdo) para iniciar la consulta fresca de forma anticipada y asegurar que la lista esté actualizada en la primera apertura sin duplicar consultas.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 3. Fase 1629 — Desacoplamiento de Modal de Error y Despacho en `BeginInvoke`
- **Estado:** ACEPTADA
- **Acción:** Introducción del campo `private bool _ultimaCargaDesdeComboOk;`. Condicionamiento del despacho de selección pendiente a la carga exitosa de la corrida previa y delegación asíncrona desacoplada mediante `BeginInvoke(new Action(() => { if (!IsDisposed) _ = CargarCorridaDesdeComboAsync(); }));`.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

### Verificación de Principios (Ronda 328)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1627 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `implementation_plan.md`, `walkthrough.md` |
| F1628 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1629 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |

---

## Ronda 329: Reapertura Segura de Popup, Deduplicación de Consultas y Robustez de Handle (Fases 1630–1632)

### Fases Ejecutadas

#### 1. Fase 1630 — Limpieza Incondicional de Flag y Validación de Handle
- **Estado:** ACEPTADA
- **Acción:** En `CargarCorridaDesdeComboAsync`, reseteo incondicional de `_recargaCorridaPendiente = false` en `finally`. Validación de `IsHandleCreated` y captura de `InvalidOperationException` en `BeginInvoke`.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 2. Fase 1631 — Cierre Temporal y Reapertura Asíncrona de Popup
- **Estado:** ACEPTADA
- **Acción:** En `RefrescarComboCorridasAsync`, cierre temporal si el popup está desplegado (`DroppedDown = false`), swap seguro de `Items`, y reapertura asíncrona del popup en `BeginInvoke`, logrando que la lista aparezca fresca en el primer clic.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 3. Fase 1632 — Deduplicación Temporal de Refresco y Retiro de Disparador Redundante
- **Estado:** ACEPTADA
- **Acción:** Deduplicación de llamadas en `RefrescarComboCorridasSeguro` con ventana de 750 ms (`_ultimoRefrescoComboUtc`), retiro del evento redundante `MouseDown`, y marcado de timestamp al completar el swap.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

### Verificación de Principios (Ronda 329)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1630 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1631 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1632 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |

---

## Ronda 330: Desacoplamiento de Reapertura, Flag de Cierre Interno y Precisión de Timestamp (Fases 1633–1635)

### Fases Ejecutadas

#### 1. Fase 1633 — Flag de Inhibición para Reapertura Programática
- **Estado:** ACEPTADA
- **Acción:** Declaración del flag `_reabriendoPopupCombo` y guarda en `cmbCargarCorridas.DropDown` para bloquear refrescos cíclicos originados por la reapertura programática del selector.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 2. Fase 1634 — Flag de Cierre Interno y Reapertura Garantizada en `finally`
- **Estado:** ACEPTADA
- **Acción:** Declaración del flag `_cerrandoPopupInterno` para ignorar `DropDownClosed` durante el swap de colección. Invocación de `ReabrirPopupDiferido()` desde el bloque `finally` para asegurar que el popup se reabra incluso en el camino de lista vacía (`items.Count == 0`).
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 3. Fase 1635 — Precisión de Timestamp al Completar Swap Efectivo
- **Estado:** ACEPTADA
- **Acción:** Reubicación de `_ultimoRefrescoComboUtc = DateTime.UtcNow;` dentro del bloque `try` tras asignar `SelectedIndex`, asegurando que no se bloqueen intentos legítimos de refresco si no hubo swap efectivo.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

### Verificación de Principios (Ronda 330)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1633 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1634 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1635 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |

---

## Ronda 331: Reapertura con Validación de Foco y Guard Contra Cascadas Espurias (Fases 1636–1637)

### Fases Ejecutadas

#### 1. Fase 1636 — Contador de Reapertura y Validación de Foco
- **Estado:** ACEPTADA
- **Acción:** Conversión de `_reabriendoPopupCombo` a contador reentrante `int` y verificación de `cmbCargarCorridas.Focused` antes de restaurar `DroppedDown = true` en `ReabrirPopupDiferido()`.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

#### 2. Fase 1637 — Filtro de Reentrada para Clave en Vuelo
- **Estado:** ACEPTADA
- **Acción:** Declaración de `_claveFiltroEnVuelo` y condicionamiento de `_solicitudPendienteRefrescoCombo = true` exclusivamente cuando `ClaveFiltroCorridas() != _claveFiltroEnVuelo`, previniendo dobles pasadas de `do-while` en interacciones normales de usuario.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`.

### Verificación de Principios (Ronda 331)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1636 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |
| F1637 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` |

---

## Ronda 338: Validación y Corrección de Dependencias, Integridad de Plan Explícito y Estado del Árbol (Fases 1687–1691)

### Fases Ejecutadas

#### 1. Fase 1687 — Olas Compatibles con Dependencias
- **Estado:** ACEPTADA
- **Acción:** Desdoblamiento a 10 olas en serie (`waves: 10`) en el plan de ranuras dinámicas para respetar la causalidad `T1->T2` y `T3->T4`.
- **Archivos Afectados:** `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`.

#### 2. Fase 1688 — Prohibición de Fallback Legado
- **Estado:** ACEPTADA
- **Acción:** Sustitución de `DemoverPlanExplicito()` por replanificación atómica del grupo con rechazo del candidato si falla.
- **Archivos Afectados:** `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`.

#### 3. Fase 1689 — Cobertura Estricta y Descuento de Festivos
- **Estado:** ACEPTADA
- **Acción:** Descuento de festivos en reservas y exigencia de `Result.Failure` estricto en el expansor si la cobertura no coincide exactamente con las horas del gen.
- **Archivos Afectados:** `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`.

#### 4. Fase 1690 — Completitud de Rutas y Leases
- **Estado:** ACEPTADA
- **Acción:** Inclusión de `AgGeneradorHorarios.cs` en T6 y rutas completas con tests para T2–T9.
- **Archivos Afectados:** `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`.

#### 5. Fase 1691 — Precondición y Diagnóstico de Gate
- **Estado:** ACEPTADA
- **Acción:** Actualización de precondición a HEAD `c0d986d` limpio y calibración de diagnóstico de exit code 3 de `verify.ps1`.
- **Archivos Afectados:** `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`.

### Verificación de Principios (Ronda 338)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1687 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md` |
| F1688 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md` |
| F1689 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md` |
| F1690 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md` |
| F1691 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md` |

---

## Ronda 339: Consolidación Documental y Sincronización del Plan de Ranuras Dinámicas (Fase 1692)

### Fases Ejecutadas

#### 1. Fase 1692 — Aplicación Efectiva del Diff en Disco
- **Estado:** ACEPTADA
- **Acción:** Modificación y saneamiento integral de `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`, alineando el texto real con los acuerdos de Ronda 338 (10 olas, rutas completas con tests en `files_lock`, erradicación de `DemoverPlanExplicito` y precondición de árbol limpio).
- **Archivos Afectados:** `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`.

### Verificación de Principios (Ronda 339)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1692 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md` |





## Ronda 340: Rutas del plan y numeración (Fases 1693–1694)

- **Fase 1693 realizada (documental):** corregida ruta de AgGen.cs en T6 a Academic/Domain/Models/Genetica; completada ruta de AgGeneradorHorarios.cs en T10.
- **Fase 1694 realizada (documental):** registradas revisiones 338–339 en R-054/R-055 y esta corrección en R-056. Contadores ronda 340/341, fase 1694/1695. R-047 continúa executing.
- **Verificación:** inspección de archivos; sin compilación ni pruebas. T1–T10 siguen pending y sin reserva. No se implementó código 4→8.

---

# Walkthrough — Empaquetado Release e Instalador 1.6.4.08

## Resumen de Ejecución
- **Comando Ejecutado:** `powershell -ExecutionPolicy Bypass -File ".\scripts\build_installer.ps1" -NewVersion "1.6.4.08"`
- **Acciones Realizadas:**
  1. Limpieza de artefactos e instaladores previos en `installer\`.
  2. Sincronización canónica de versión `1.6.4.08` en `ConversorPDF.XLL/Core/Helpers/App.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj` e `installer/InstaladorConversorPDF.wxs`.
  3. Ejecución integral y aprobación al 100% de la suite de pruebas unitarias (`dotnet test`): 0 fallos, 0 regresiones.
  4. Compilación del Add-In Excel-DNA en modo `Release|x64` y empaquetado del runtime.
  5. Empaquetado del instalador MSI mediante WiX Toolset 4 (`InstaladorConversorPDF.wixproj`).

## Verificación de Artefactos

| Artefacto | Estado | Detalle |
| :--- | :---: | :--- |
| `installer\InstaladorConversorPDFv1.6.4.08.msi` | Generado | Instalador MSI WiX 4 de 5,726,208 bytes (5.46 MB) con auto-activación per-user en Excel. |
| `ConversorPDF.Tests` | 100% Pass | Suite completa ejecutada sin errores ni regresiones. |
| `ConversorPDF.XLL` | Compilado | `Release|x64` empaquetado sin errores CS. |

---

## Ronda 366 — Corrección de Criterios de Mejora, Persistencia de Patrones y Trazabilidad (Fases 1742–1746)

### Resumen de Cambios Aplicados

#### 1. Fase 1742 — Decisión Basada en Mejora Real de Saturación
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs`
- **Cambio:** Incorporación de la propiedad `Mejora` y guarda estricta `bool mejora = !evalDespues.HaySaturacion || !evalDespues.SaturacionCarriles; if (!mejora) continue;`. Impide proponer escalado cuando la saturación de carril permanece inmutable.

#### 2. Fase 1743 — Persistencia y Carga de `origen_patron`
- **Archivos:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgAulaAsignatura.cs`
- **Cambio:** Integración de `origen_patron` en `INSERT` (`'AUTOMATICO'`), `SELECT` hacia el contexto y valor por defecto en la entidad para garantizar round-trip completo.

#### 3. Fase 1744 — Proyección de Makespan, Fecha Fin y Ranuras Ociosas
- **Archivos:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResultadoEvaluacionVentana.cs`
- **Cambio:** Cálculo de `SemanasRequeridas`, `FechaFinMinima`, `RanurasOciosas` y captura de `IdAsignaturasSaturadas` usando la distribución en `AgCarriles`.

#### 4. Fase 1745 — Acotación del Reset de Patrones y Aviso Modal
- **Archivos:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambio:** Reset SQL acotado a `IdAsignaturasBottleneck`, señalización de `MarkDirty()` ante cambios de patrones y advertencia explícita en el diálogo modal pre-vuelo.

#### 5. Fase 1746 — Sincronización Registral
- **Archivo:** `.agents/plans/REGISTRY.md`
- **Cambio:** Corrección de la fila `R-066` apuntando al commit real de cierre `c34bec4`.

### Verificación de Principios (Ronda 366)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1742 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorVentanaAula.cs` |
| F1743 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgScheduleRepository.cs`, `AgFormGeneradorHorarios.cs`, `AgAulaAsignatura.cs` |
| F1744 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs`, `AgResultadoEvaluacionVentana.cs` |
| F1745 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorVentanaAula.cs`, `AgFormGeneradorHorarios.cs` |
| F1746 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md` |


---

## Ronda 367 — Corrección de Evaluación por Carriles Reales, Preservación de Patrones y Bottlenecks (Fases 1747–1752)

### Resumen de Cambios Aplicados

#### 1. Fase 1747 — Evaluación por Carril Real y Mejora Estricta de Semanas Requeridas
- **Archivos:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs`
- **Cambio:** `SaturacionCarriles` calculado sobre la capacidad de ocurrencias reales de cada carril (`AgPlanificadorTemporal.ContarOcurrencias`). `Mejora` exige reducción estricta de `SemanasRequeridas` (`evalDespues.SemanasRequeridas < evalAntes.SemanasRequeridas`) o reducción neta de saturación, impidiendo la propuesta espuria de `JUE-DOM` cuando el domingo permanece saturado (caso 3 materias domingo).

#### 2. Fase 1748 — Asignación Estricta de Carril Conteniendo Todos los Días del Patrón
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** En `ElegirCarril`, cuando el patrón no es `AUTOMATICO`, se valida que todos los días pertenezcan al carril (`d.DiasPatron.All(day => carriles[c].Any(r => r.Dia == day))`). Si ningún carril satisface todos los días, retorna `-1` para registrarlo como fuera de ventana/no acomodable en vez de asignar erróneamente por `Any`.

#### 3. Fase 1749 — `FinDeSemanaValida` con Días de Ventana y Calendario de Festivos
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** Reemplazo de fórmula lineal por cálculo iterativo semanal que verifica días activos de la ventana libres de festivos (`esFestivo`), eliminando el desfase off-by-one.

#### 4. Fase 1750 — Preservación de Patrones `FIJO` y `LEGADO` en Regeneración de Aulas
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Cambio:** Reemplazo de `DELETE FROM ag_aula_asignatura` masivo por borrado selectivo de registros `AUTOMATICO` no seleccionados (`UPPER(COALESCE(origen_patron,''))='AUTOMATICO' AND id_asignatura NOT IN @idAsignaturas`) e inserción con `ON CONFLICT(id_aula, id_asignatura) DO NOTHING`.

#### 5. Fase 1751 — Corrección de Regresión en OrigenPatron por Defecto
- **Archivo:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgAulaAsignatura.cs`
- **Cambio:** Reversión del valor por defecto en la entidad a `"LEGADO"` en constructor y propiedad, delegando la asignación de `"AUTOMATICO"` a las operaciones explícitas de generación.

#### 6. Fase 1752 — Aislamiento Estricto de Bottlenecks para Reset
- **Archivos:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambio:** `res.IdAsignaturasSaturadas` poblado con `noCupieron` en lugar del ordenamiento total. Asignación a `IdAsignaturasBottleneck` y advertencia modal con recuento exacto de materias de cuello de botella a restablecer.

### Verificación de Principios (Ronda 367)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1747 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs`, `AgEscaladorVentanaAula.cs` |
| F1748 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1749 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1750 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgScheduleRepository.cs` |
| F1751 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaAsignatura.cs` |
| F1752 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs`, `AgEscaladorVentanaAula.cs`, `AgFormGeneradorHorarios.cs` |

---

## Ronda 368 — Corrección de Compilación, Bounds de Carril, Capacidad por Ocurrencias, Fecha Fin de Ventana y Preservación de Patrones (Fases 1753–1758)

### Resumen de Cambios Aplicados

#### 1. Fase 1753 — Eliminación de Bloque Huérfano en `AgEvaluadorDemandaVentana`
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** Supresión del bloque huérfano `{ ... }` residual a nivel de clase posterior a `PuedeAcomodar` que provocaba errores de compilación CS1519 y CS0103.

#### 2. Fase 1754 — Guarda ante Retorno `-1` en `ElegirCarril`
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** Verificación defensiva `if (idxCarril < 0)` registrando la materia en `materiasFueraVentana` y omitiendo el acceso indebido a `carriles[idxCarril]`.

#### 3. Fase 1755 — Capacidad Real por Carril y Distribución por Ranuras del Patrón
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** Distribución de carga por carril restringida a ranuras compatibles con el patrón (`slotsPatron`). Cómputo de capacidad real `capCarril[c]` mediante `AgPlanificadorTemporal.ContarOcurrencias` y detección directa de `SaturacionCarriles` si `cargaCarriles[c] > capCarril[c]`.

#### 4. Fase 1756 — Búsqueda Inversa del Último Día Activo de la Ventana en `FinDeSemanaValida`
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** Búsqueda descendente desde el final de la semana para devolver el último día perteneciente a la ventana que no sea festivo en lugar de un sábado genérico.

#### 5. Fase 1757 — Protección de Patrones `FIJO/LEGADO` en Siembra y Aislamiento de Bottlenecks
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs`
- **Cambio:** Cláusula `ON CONFLICT` en `AgFormGeneradorHorarios.cs` condicionada a `UPPER(COALESCE(origen_patron, 'AUTOMATICO')) = 'AUTOMATICO'`. En `AgEscaladorVentanaAula.cs`, eliminación de fallback que asignaba todas las asignaturas del aula ante lista vacía de saturadas.

#### 6. Fase 1758 — Actualización Registral de Rondas y Fases
- **Archivo:** `.agents/plans/REGISTRY.md`
- **Cambio:** Actualización de contadores globales (`ronda`: 368 / 369, `fase`: 1758 / 1759) e incorporación de filas `R-067`, `R-068` y `R-069`.

### Verificación de Principios (Ronda 368)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1753 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1754 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1755 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1756 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1757 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs`, `AgEscaladorVentanaAula.cs` |
| F1758 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md` |

---

## Ronda 369 — Granularidad por Slot/Franja, Poblado de Bottlenecks y Supresión de Reset Masivo (Fases 1759–1762)

### Resumen de Cambios Aplicados

#### 1. Fase 1759 — Control Granular por `(Dia, Franja)`
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** Incorporación de mapa `slotCarga` por `(Dia, Franja)`. Detección precisa de saturación cuando la carga excede las ocurrencias del día específico, impidiendo que carriles multirranura enmascaren cuellos de botella en un solo día.

#### 2. Fase 1760 — Poblado de `IdAsignaturasSaturadas` y Supresión de Reset Incondicional
- **Archivos:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs`
- **Cambio:** Asegurado el llenado de `IdAsignaturasSaturadas` ante saturación por slot o déficit global. En el repositorio `AgEscaladorVentanaAula.cs`, eliminación de la rama `else` que ejecutaba reset global sobre toda el aula/subunidad.

#### 3. Fase 1761 — Semántica de Último Día Hábil en `FinDeSemanaValida`
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** Validación de la lógica de terminación que devuelve el último día hábil de ventana no festivo de la última semana requerida.

#### 4. Fase 1762 — Sincronización Registral
- **Archivo:** `.agents/plans/REGISTRY.md`
- **Cambio:** Incorporación de fila `R-070` y actualización de contadores globales (`ronda`: 369 / 370, `fase`: 1762 / 1763).

### Verificación de Principios (Ronda 369)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1759 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1760 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs`, `AgEscaladorVentanaAula.cs` |
| F1761 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1762 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md` |

---

## Ronda 370 — Exposición de Métricas Temporales, Diagnóstico I5_DIA_SATURADO, Retiro de Bandas y Filtrado de Bottlenecks (Fases 1763–1767)

### Resumen de Cambios Aplicados

#### 1. Fase 1763 — Exposición de Métricas en `Decision`, Modal y Validador Pre-Vuelo
- **Archivos:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Cambio:** Incorporación de `SemanasRequeridas`, `FechaFinMinima` y `RanurasOciosas` en `Decision`. Visualización explícita en el modal pre-vuelo y adición del bloqueo `[I5_DIA_SATURADO]` en `AgValidadorVentana.cs` cuando la demanda excede la capacidad por franja.

#### 2. Fase 1764 — Verificación Conceptual de `FechaFinMinima` con Festivos
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** Semántica de última semana válida consolidada respecto a exclusión de festivos.

#### 3. Fase 1765 — Retiro de `PuedeAcomodar` y Partición por Bandas
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** Eliminación definitiva de la lógica de agrupamiento por bandas y del método de backtracking `PuedeAcomodar`, dejando la evaluación unificada y determinista en `slotCarga`.

#### 4. Fase 1766 — Filtrado Quirúrgico de Asignaturas Saturadas
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** `IdAsignaturasSaturadas` restringido estrictamente a las asignaturas que coinciden con los días de los `slotsSaturados`, evitando reseteos colaterales en asignaturas que no contribuyen al cuello de botella.

#### 5. Fase 1767 — Actualización Registral de Rondas y Fases
- **Archivo:** `.agents/plans/REGISTRY.md`
- **Cambio:** Incorporación de fila `R-071` y avance de contadores globales (`ronda`: 370 / 371, `fase`: 1767 / 1768).

### Verificación de Principios (Ronda 370)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1763 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorVentanaAula.cs`, `AgFormGeneradorHorarios.cs`, `AgValidadorVentana.cs` |
| F1764 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1765 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1766 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1767 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md` |

---

## Ronda 371 — Supresión de Diagnóstico Contradictorio y Priorización en Déficit Global (Fases 1768–1771)

### Resumen de Cambios Aplicados

#### 1. Fase 1768 — Supresión de Oleada Contradictoria tras `I5_DIA_SATURADO`
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Cambio:** Incorporación de `continue;` tras registrar `[I5_DIA_SATURADO]`, eliminando el mensaje contradictorio que sugería escalar la ventana a `JUE-DOM` cuando la saturación es inviable de resolver por escalado.

#### 2. Fase 1769 — Validación de Desplazamiento por Festivos
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** Verificación de la consistencia temporal ante domingos festivos intermedios.

#### 3. Fase 1770 — Priorización por Carga en Déficit Global
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs`
- **Cambio:** `IdAsignaturasSaturadas` selecciona las asignaturas ordenadas por demanda decreciente (`OrderByDescending(d => d.Sesiones)`).

#### 4. Fase 1771 — Sincronización Registral
- **Archivo:** `.agents/plans/REGISTRY.md`
- **Cambio:** Incorporación de fila `R-072` y avance de contadores globales (`ronda`: 371 / 372, `fase`: 1771 / 1772).

### Verificación de Principios (Ronda 371)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1768 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs` |
| F1769 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1770 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs` |
| F1771 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md` |

---

## Ronda 372 — Reducción de Materias Fuera de Ventana en Mejora (Fases 1772–1774)

### Resumen de Cambios Aplicados

#### 1. Fase 1772 — Inclusión de MateriasFueraVentana en Criterio de Mejora
- **Archivos:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResultadoEvaluacionVentana.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs, ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs
- **Cambio:** Incorporación de la propiedad MateriasFueraVentana y agregada la condición evalDespues.MateriasFueraVentana.Count < evalAntes.MateriasFueraVentana.Count a Mejora. Permite proponer escalados legítimos (como LUN-JUE → LUN-VIE para materias de viernes) aun cuando persista saturación horaria en el nuevo escalón.

#### 2. Fase 1773 — Diagnóstico de Cobertura y Suite de Pruebas
- **Archivo:** Árbol de trabajo y artefactos documentales.
- **Cambio:** Identificación de suites de prueba pendientes para ejecución una vez aprobada la fase de revisión.

#### 3. Fase 1774 — Actualización Registral de Rondas y Fases
- **Archivo:** .agents/plans/REGISTRY.md
- **Cambio:** Incorporación de fila R-073 y avance de contadores globales (
onda: 372 / 373, ase: 1774 / 1775).

### Verificación de Principios (Ronda 372)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1772 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgResultadoEvaluacionVentana.cs, AgEvaluadorDemandaVentana.cs, AgEscaladorVentanaAula.cs |
| F1773 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | implementation_plan.md |
| F1774 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | REGISTRY.md |

---

## Ronda 373 — Horizonte Acotado de Fecha Fin, Autoridad del Planificador e Idempotencia (Fases 1775–1779)

### Resumen de Cambios Aplicados

#### 1. Fase 1775 — Límite de Horizonte Temporal en FinDeSemanaValida
- **Archivos:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs, ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs
- **Cambio:** Integración de cota superior finita limite en FinDeSemanaValida, firma nullable DateTime? y manejo explícito de  no cabe dentro del plazo ante sobrepaso de horizonte.

#### 2. Fase 1776 — Supresión de Restricción Rígida en Patrones AUTOMATICO
- **Archivos:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs
- **Cambio:** DiasPatron ignorado y sustituido por vacío/nulo cuando OrigenPatron == AUTOMATICO en demandas, fábrica de individuos y oleadas, permitiendo que el motor redistribuya libremente.

#### 3. Fase 1777 — Primacía del Planificador de Oleadas
- **Archivo:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs
- **Cambio:** Ejecución previa de AgPlanificadorOleadas.Construir. El diagnóstico de [I5_DIA_SATURADO] solo se emite si el planificador formal declara infactibilidad.

#### 4. Fase 1778 — Idempotencia y Protección ante Decisiones Obsoletas
- **Archivo:** ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs
- **Cambio:** Verificación previa en transacción: si el aula ya está en Hacia no reescribe; si difiere de Desde aborta con excepción descriptiva; actualización acotada con AND COALESCE(dias_ventana, '') = @Desde.

#### 5. Fase 1779 — Registro y Trazabilidad
- **Archivo:** .agents/plans/REGISTRY.md
- **Cambio:** Incorporación de fila R-074 y actualización de contadores globales (
onda: 373 / 374, ase: 1779 / 1780).

### Verificación de Principios (Ronda 373)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1775 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEvaluadorDemandaVentana.cs, AgValidadorVentana.cs, AgFormGeneradorHorarios.cs |
| F1776 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEvaluadorDemandaVentana.cs, AgFabricaIndividuos.cs, AgValidadorVentana.cs |
| F1777 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgValidadorVentana.cs |
| F1778 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEscaladorVentanaAula.cs |
| F1779 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | REGISTRY.md |

---

## Ronda 374 — Métricas del Plan Efectivo, Confirmadas en Idempotencia y Huella de Entrada (Fases 1780–1783)

### Resumen de Cambios Aplicados

#### 1. Fase 1780/1782 — Proyección Posterior en Decision y Desacople de Pendiente
- **Archivos:** ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs, ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs
- **Cambio:** Decision almacena las métricas de evalDespues (semanas, fecha fin mínima, ranuras ociosas y demanda pendiente calculada). El modal pre-vuelo utiliza d.DemandaPendiente en vez de fórmulas abstractas.

#### 2. Fase 1781 — Separación de Aulas Confirmadas y Modificadas en Repositorio
- **Archivo:** ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs
- **Cambio:** Clase ResultadoAplicacion que distingue Confirmadas de Modificadas. La verificación en AplicarAsync devuelve las confirmadas (incluyendo aulas que ya estaban en Hacia), garantizando compatibilidad con la guarda plicadas == esperadas de la UI.

#### 3. Fase 1783 — Huella de Entrada en Decisión
- **Archivo:** ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs
- **Cambio:** Propiedad HuellaEntrada para verificación de inmutabilidad entre evaluación y persistencia.

#### 4. Numeración y Registro
- **Archivo:** .agents/plans/REGISTRY.md
- **Cambio:** Incorporación de fila R-075 y avance de contadores globales (
onda: 374 / 375, ase: 1783 / 1784).

### Verificación de Principios (Ronda 374)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1780 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEscaladorVentanaAula.cs |
| F1781 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEscaladorVentanaAula.cs |
| F1782 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEscaladorVentanaAula.cs, AgFormGeneradorHorarios.cs |
| F1783 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEscaladorVentanaAula.cs, REGISTRY.md |

---

## Ronda 375 — Rollback con Contadores Vacíos, Demanda Pendiente por Ranuras y Validación de Huella/Idempotencia (Fases 1784–1787)

### Resumen de Cambios Aplicados

#### 1. Fase 1784 — Cero Confirmadas en Excepción
- **Archivo:** ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs
- **Cambio:** En catch, retorno de
ew ResultadoAplicacion() en lugar de devolver contadores incrementados antes del fallo transaccional.

#### 2. Fase 1785 — Demanda Pendiente Real por Capacidad de Slot
- **Archivo:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorDemandaVentana.cs
- **Cambio:** DemandaPendiente computa la sumatoria de sesiones que no pudieron acomodarse en sus respectivos slots frente a la capacidad del calendario.

#### 3. Fase 1786 — Verificación de Huella Pre-Mutación
- **Archivo:** ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs
- **Cambio:** Comprobación de d.HuellaEntrada dentro de la transacción contra la configuración persistida antes de aplicar cambios.

#### 4. Fase 1787 — Idempotencia Completa Incluyendo Nodo
- **Archivo:** ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs
- **Cambio:** Si el aula ya posee dias_ventana == Hacia, se verifica que el nodo en g_arbol_nodo también coincida antes de darla por confirmada.

#### 5. Numeración y Registro
- **Archivo:** .agents/plans/REGISTRY.md
- **Cambio:** Incorporación de fila R-076 y avance de contadores globales (
onda: 375 / 376, ase: 1787 / 1788).

### Verificación de Principios (Ronda 375)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1784 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEscaladorVentanaAula.cs |
| F1785 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEvaluadorDemandaVentana.cs |
| F1786 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEscaladorVentanaAula.cs |
| F1787 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgEscaladorVentanaAula.cs, REGISTRY.md |


## Ronda 376 propuesta — Fases 1788–1793 — 2026-11-13

Implementación editada: plan compartido con sesiones fechadas; aprovechamiento de ranuras sobrantes; fábrica consume las mismas fechas; fecha proyectada sin promesa de mínimo; huella completa de contexto/BD; idempotencia y contadores posteriores al commit; recarga de UI y caché por contenido. Pruebas de aceptación y consistencia añadidas.

Estado: sin compilación ni ejecución de pruebas por restricción del usuario. Fases 1792–1793 sin evidencia de aprobación. No se declara release listo ni smoke superado. Detalle vigente: docs/plans/2026-11-12-escalado-tres-casos-revision-ronda-376.md. Sin nueva reserva en REGISTRY.

---

## Ronda 398 — Sincronización Documental, Trazabilidad e Invariantes del Planificador con Límite de Ventana (Fases 1815–1819)

### Resumen de Cambios Aplicados

#### 1. Fase 1815 — Citas por Símbolo y Estado Implementado
- **Archivo:** docs/plans/2026-11-15-planificador-limite-ventana-design.md
- **Cambio:** Marcado como IMPLEMENTADO, citas de líneas sustituidas por referencias canónicas por símbolo (`EmpaquetarPendientes`, `RedFlujo`, `carrilPorId`, `ValidarFechado`, `AgCarriles.cs:13`) y recuento a 64 casos.

#### 2. Fase 1816 — Reversión Quirúrgica sin Destruir Fix
- **Archivo:** docs/plans/2026-11-15-planificador-limite-ventana-design.md
- **Cambio:** Paso 4 de orden de trabajo modificado para restringir reversión a símbolos nuevos (`EmpaquetarPendientes`, `RedFlujo`, `asignado`) evitando `git checkout --`.

#### 3. Fase 1817 — Criterio de Cierre con Re-baseline Autorizado
- **Archivo:** docs/plans/2026-11-15-planificador-limite-ventana-design.md
- **Cambio:** Aserciones marcadas con excepción explícita aprobada en `AgPrioridadHorasTests:157-165` para `CompletarReferencia_WhenMayorCargaPendiente`; familia de 64 casos pendiente de corrida (sin evidencia todavía).

#### 4. Fase 1818 — Envelope y Registro Formal en Libro
- **Archivo:** docs/plans/2026-11-15-planificador-limite-ventana-design.md, .agents/plans/REGISTRY.md
- **Cambio:** Inclusión de `AgPlanificadorRanurasFallbackTests.cs` en superficies. Creación de fila `R-080` en `REGISTRY.md` y avance de contadores (`ronda: 398 / 399`, `fase: 1819 / 1820`).

#### 5. Fase 1819 — Invariantes de Respaldo en Tabla
- **Archivo:** docs/plans/2026-11-15-planificador-limite-ventana-design.md
- **Cambio:** Adición formal en la tabla de invariantes de las aserciones de respaldo sin mutación cubiertas por `AgPlanificadorRanurasFallbackTests:20-56` y `:59-83`.

### Verificación de Principios (Ronda 398)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1815 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/plans/2026-11-15-planificador-limite-ventana-design.md` |
| F1816 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/plans/2026-11-15-planificador-limite-ventana-design.md` |
| F1817 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/plans/2026-11-15-planificador-limite-ventana-design.md` |
| F1818 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/plans/2026-11-15-planificador-limite-ventana-design.md`, `.agents/plans/REGISTRY.md` |
| F1819 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/plans/2026-11-15-planificador-limite-ventana-design.md` |

---

