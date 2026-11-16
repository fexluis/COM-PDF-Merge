<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# implementation_plan - Fases 1603-1819 (Rondas 318..398)

> Indice: [../../../implementation_plan.md](../../../implementation_plan.md) - Escision: 2026-11-16

---
### Cambios Aplicados
1. **Fase 1600 (`AgPlanificadorOleadas.cs`):**
   - **Pasada 1:** Comprobación de capacidad (`carga[c] + ctx.Demanda[i] <= capacidadCarril[c]`) integrada dentro de la búsqueda de carril compatible, seleccionando el carril de menor tamaño que efectivamente posea capacidad libre.
   - **Pasada 2:** Conservación de restricciones sembradas: omisión de carriles que no contengan los días del patrón y omisión si se satura la capacidad. Eliminación del fallback coercitivo `if (elegido < 0) elegido = 0`, permitiendo que `carrilDe[i]` permanezca en `-1`.
   - **Pasada 3:** Guarda `if (c < 0)` antes de indexar `slotsCarril[c]`. Distingue si la asignatura es estructuralmente alojable (`diasCarril.Any(b => dias.All(b.Contains))`), asignando `I5_CAPACIDAD_CARRIL_PATRON` si es alojable o `R8_DiaFueraDePatron` si no lo es.
2. **Fase 1601 (`AgPlanificadorOleadas.cs`, `AgValidadorVentana.cs`, `AgGeneradorHorarios.cs`):**
   - Integración en `DetalleInfactibles` de un único diagnóstico por asignatura reportando patrón, demanda, código (`I5_CAPACIDAD_CARRIL_PATRON` vs `R8_DiaFueraDePatron`) y capacidades reservadas/totales de los carriles compatibles (`c{j}[{dias}] capacidad={cap} reservada={carga}`).
   - Discriminación en `AgValidadorVentana.cs` del bloqueo `[I5_CAPACIDAD_CARRIL_PATRON]` frente a `[R8_DiaFueraDePatron]` e `[I5_INFACTIBLE_OLEADA]`.
   - Actualización del mensaje de error en `AgGeneradorHorarios.cs` a `"Corrida abortada por bloqueos de factibilidad detectados:\n• {errores}"`.
3. **Fase 1602 (Suites de pruebas):**
   - Pruebas exhaustivas añadidas en `AgPlanificadorOleadasTests.cs` (patrón 5 con capacidad, alternativa disponible cuando el más específico está lleno, saturación sin desbordamiento a incompatibles, R8 para día fuera de ventana, reparto proporcional para materias sin patrón, y fin de semana 6,7 mixto con festivos).
   - Pruebas en `AgValidadorVentanaTests.cs` verificando la no-bloqueante tolerancia a vacantes por ausencia de pool docente y el bloqueo específico `[I5_CAPACIDAD_CARRIL_PATRON]` con sede y aula.
   - Actualización de aserciones diagnósticas en `AgGeneradorHorariosTests.cs` y `AgPlanificadorOleadasCaracterizacionTests.cs`.

### Sincronización REGISTRY
- Fila `R-039` (`closed`) y contadores actualizados a `ronda: 317 / 318`, `fase: 1602 / 1603`.

---

## Ronda 318: Corrección de Tipos en Pruebas, Calibración de Fixtures de Capacidad y Cobertura Efectiva de Festivos (Fases 1603–1605)

### Contexto y Diagnóstico
1. **Fase 1603:** Errores de compilación estáticos introducidos en la suite de pruebas:
   - En `AgPlanificadorOleadasTests.cs`, se invocaba `p.DiasCarril` cuando la propiedad pública en `AgPlanOleada` es `DiasPatron`.
   - En `AgValidadorVentanaTests.cs`, se asignaba directamente una lista de `AgDocenteHabilitado` a `PoolPorAsignatura[1]`, cuyo tipo de diccionario es `Dictionary<int, List<TerceroAsignatura>>`.
2. **Fase 1604:** Fixtures de prueba que no alcanzaban la rama evaluada por fallo prematuro o cobertura condicional:
   - En `AgPlanificadorOleadasTests.cs` (`Patron5Saturado`), la ventana `"VIE"` (capacidad total 3) abortaba tempranamente con `I5_SIN_CAPACIDAD` ante demanda de 6 sesiones antes de llegar al reparto de carriles.
   - En `AgValidadorVentanaTests.cs` (`PatronSaturadoEnAula`), la ventana `"VIE"` activaba el bloqueo global preliminar `[I5_INFACTIBLE]` por sobre-demanda total de aula antes de evaluar oleadas.
   - En `AgPlanificadorOleadasTests.cs` (`CarrilMasEspecificoLleno`), los patrones `4,5` ya cabían en ambos carriles y las aserciones estaban envueltas en un bloque `if` que permitía pasar sin verificar la alternancia de carriles.
3. **Fase 1605:** En `AgPlanificadorOleadasTests.cs` (`Casos6Y7_MixtosYFestivos`), ambos patrones eran `6,7` sin variación cruzada, el festivo caía en lunes (fuera de la ventana evaluada), y el assert usaba `AllSatisfy` que pasa con listas vacías.

### Cambios Aplicados
1. **Fase 1603 (`AgPlanificadorOleadasTests.cs`, `AgValidadorVentanaTests.cs`):**
   - Corrección de `p.DiasCarril` a `p.DiasPatron` en aserciones de `AgPlanificadorOleadasTests.cs`.
   - Reemplazo de la asignación manual en `AgValidadorVentanaTests.cs` por la invocación canónica de `contexto.RegistrarDocenteHabilitado(101, 0, 1)` seguida de `contexto.ConstruirIndices()`.
2. **Fase 1604 (`AgPlanificadorOleadasTests.cs`, `AgValidadorVentanaTests.cs`):**
   - En `Patron5Saturado`: uso de ventana amplia `"1,2,3,4,5"` (capacidad suficiente para no disparar `I5_SIN_CAPACIDAD`), verificando que se planifique exactamente una materia en viernes y la segunda reporte infactibilidad con `I5_CAPACIDAD_CARRIL_PATRON`.
   - En `ValidarFactibilidad_PatronSaturadoEnAula`: configuración de aula con ventana `"1,2,3,4,5"`, jornada `"NOCHE"` y bloques nocturnos 1–5, alcanzando la llamada a `AgPlanificadorOleadas.Construir` y emitiendo el bloqueo discriminado `[I5_CAPACIDAD_CARRIL_PATRON]`.
   - En `CarrilMasEspecificoLleno`: rediseño determinista con ventana `"5,6"`, 4 semanas, bloques canónicos (carril 0 `[5,6M]` con capacidad 8, carril 1 `[6T]` con capacidad 4), y materias 1 y 2 con patrón "6" de 16h (4 sesiones), validando que la materia 1 ocupe el carril 1 (más específico) y la materia 2 use la alternativa compatible (carril 0) sin infactibles.
3. **Fase 1605 (`AgPlanificadorOleadasTests.cs`):**
   - Configuración explícita de ventana `"4,5,6,7"`, patrones diferenciados (materia 1 `"4,5"`, materia 2 `"6,7"`), y festivo en el primer sábado utilizado (`2027-02-06`).
   - Aserción previa obligatoria `res.Value.Planes.Should().HaveCount(2)` y validación estricta de `plan.DiasPatron` con `BeEquivalentTo` contra los días esperados de cada patrón.

### Sincronización REGISTRY
- Fila `R-040` (`closed`) y contadores actualizados a `ronda: 318 / 319`, `fase: 1605 / 1606`.

---

## Ronda 319–325: Sincronización del Selector de Corridas del Módulo R (Fases 1606–1612)

### Contexto y Diagnóstico
1. **Fase 1606–1607:**
   - En Módulo R (`AgFormGeneradorHorarios`), el selector de corridas (`cmbCargarCorridas`) mostraba discordancia con los resultados en pantalla (p. ej. `#60` en el selector vs `#62` en la grilla y staging).
   - Causa: Se reutilizaba `AgCorridaSelector.ResolverSeleccionInicial`, diseñado para `FormVinculaciones` con el comodín `-- Todas las corridas --` (`IdCorrida = 0`). Al no pertenecer ese objeto a los `Items` de Módulo R, `items.IndexOf()` retornaba `-1`, cayendo en un fallback coercitivo a `SelectedIndex = 0`. Esto forzaba `CargarCorridaDesdeComboAsync()`, destruyendo la corrida activa para cargar silenciosamente otra ajena.
2. **Fase 1608–1609:**
   - Al desplegar el ComboBox (`cmbCargarCorridas.DropDown`), una guarda por clave de filtros (`ClaveFiltroCorridas() == _claveFiltroCorridas`) bloqueaba la reconsulta en la base de datos, impidiendo detectar nuevas corridas creadas por el AG o desde otra ventana.
   - Peticiones concurrentes durante refrescos se descartaban ciegamente con un simple `return` por `_refrescandoComboCorridas`.
3. **Fase 1610:**
   - Al finalizar `IniciarOptimizacionAsync`, no se realizaba un refresco garantizado en `finally` de `cmbCargarCorridas`. Si la ejecución era cancelada o fallaba tras crear la cabecera en SQLite, el selector quedaba desactualizado.
   - La ruta directa de ejecución no actualizaba `lblAgModo` con `#N` tras completar `CargarModuloResultadosAsync`.
   - `CargarModuloResultadosAsync` asignaba `_idCorridaCargada` prematuramente antes de comprobar si el staging contenía sesiones. Si fallaba, la corrida quedaba falsamente activa en memoria.
   - Al volver de `FormVinculaciones`, si la corrida activa había sido eliminada, la interfaz no se limpiaba ni se refrescaba el selector.
4. **Fases 1611–1612:**
   - Verificación integral de la suite completa de pruebas y documentación técnica del cambio.

### Cambios Aplicados
1. **Fase 1606 (`AgSeleccionModuloR.cs`, `ConversorPDF.XLL.csproj`):**
   - Creación del DTO inmutable `AgSeleccionModuloR` en `ConversorPDF.XLL.Academic.Application.DTOs` con propiedades get-only (`ItemSeleccionado`, `IndiceSeleccionado`, `RequiereCargaInicial`, `CorridaActivaNoEncontrada`).
   - Registro en `ConversorPDF.XLL.csproj` cumpliendo Lease L0 y la regla Single Type Per File.
   - Declaración del stub de `AgCorridaSelector.ResolverSeleccionModuloR` lanzando `NotImplementedException` (estrategia `integrator-first`).
2. **Fase 1607 (`AgCorridaSelectorTests.cs`):**
   - 6 pruebas unitarias RED implementadas en `AgCorridaSelectorTests.cs` cubriendo exhaustivamente los 6 casos de negocio (activa presente, activa ausente con candidatas con staging, activa ausente con solo cabeceras, sin activa con staging disponible, sin activa y solo cabeceras, lista vacía/nula).
   - Preservación íntegra de las pruebas existentes de `ResolverSeleccionInicial` y de `FormVinculaciones`.
3. **Fase 1608 (`AgCorridaSelector.cs`, `AgFormGeneradorHorarios.cs`):**
   - Implementación pura de `AgCorridaSelector.ResolverSeleccionModuloR`.
   - Integración en `RefrescarComboCorridasAsync` eliminando la degradación a índice 0 y la llamada incondicional a `CargarCorridaDesdeComboAsync`.
   - Asignación `SelectedIndex = seleccion.IndiceSeleccionado` (queda en `-1` si no está en la lista) y notificación informativa en `lblAgModo` (*"Modo AG · corrida cargada #N no visible con los filtros actuales"*).
   - Carga automática inicial permitida únicamente cuando `seleccion.RequiereCargaInicial` lo autoriza. Pruebas pasadas a GREEN (24/24 PASS).
4. **Fase 1609 (`AgFormGeneradorHorarios.cs`):**
   - Remoción de la guarda de caché en `cmbCargarCorridas.DropDown`, invocando siempre `RefrescarComboCorridasSeguro()`.
   - Concurrencia robusta en `RefrescarComboCorridasAsync` con campo `_solicitudPendienteRefrescoCombo` y bucle `do-while` para procesar solicitudes acumuladas.
   - Captura de filtros al iniciar y descarte de respuestas obsoletas si los filtros cambiaron durante el `await`. Preservación de catálogo ante error.
5. **Fase 1610 (`AgFormGeneradorHorarios.cs`):**
   - Registro de `idCorridaCreada` en `IniciarOptimizacionAsync` y refresco garantizado en `finally` con `CancellationToken.None` independiente.
   - Actualización de `lblAgModo.Text = $"Modo AG · corrida #{idCorrida} · {AgModoDatos.Ag.VistaSesionesDetalle}";` en ejecución directa.
   - Control del valor booleano en `CargarModuloResultadosAsync`: resetea `_idCorridaCargada = 0`, métricas y grillas ante fallo de staging.
   - En `AbrirVinculacionesModoAgAsync`, comprobación de recarga y refresco del combo.
6. **Fase 1611 (`scripts/verify.ps1`):**
   - Verificación integral con `scripts/verify.ps1 -Mode Full`: 1172/1172 tests PASS (0 fallos, 0 omitidos, 0 regresiones vs baseline 952).
7. **Fase 1612 (`docs/articles/changes/`):**
   - Documentación técnica completa en `docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md`.

### Sincronización REGISTRY
- Fila `R-041` (`closed`) y contadores actualizados a `ronda: 325 / 326`, `fase: 1612 / 1613`.

---

## Ronda 326: Desincronización de Commits, Mutación de Items con Popup Abierto y Cargas Concurrentes del Selector (Fases 1613–1620)

### Contexto y Diagnóstico
1. **Fase 1613:** En `ConversorPDF.XLL.csproj` se registraron `Academic\Application\Services\AgPerfilAlcanceResolver.cs` y `Academic\Application\Services\Genetica\AgParticionCarriles.cs` y en `AgFormGeneradorHorarios.cs` se invocaba `AgPerfilAlcanceResolver.Resolver(...)`. Ninguno de los dos archivos estaba versionado en git (`??`), provocando que un checkout limpio de HEAD fallara con errores CS2001 y CS0246.
2. **Fase 1614:** `ConversorPDF.XLL.csproj` elevó `Version/AssemblyVersion/FileVersion` a `1.6.4.07` en el commit T1 de R-041 mientras `App.APP_VERSION` en `App.cs` e `InstaladorConversorPDF.wxs` permanecían sin versionar en el árbol de trabajo.
3. **Fase 1615:** `cmbCargarCorridas.DropDown += (s, e) => RefrescarComboCorridasSeguro();` iniciaba la consulta asíncrona mientras se abría el popup del ComboBox. Al completar la consulta, se ejecutaba `Items.Clear()` y `Items.Add(...)` con `DroppedDown == true`, provocando cierre repentino, parpadeo o potencial `ArgumentOutOfRangeException`.
4. **Fase 1616:** `CargarCorridaDesdeComboAsync` carecía de guarda asíncrona de reentrada (`_cargandoCorridaDesdeCombo`). Eventos concurrentes de `SelectedIndexChanged` o invocaciones desde refrescos podían ejecutar cargas simultáneas que pisaban `_idCorridaCargada`, `_ultimoResultado` y las grillas.
5. **Fase 1617:** `seleccion.RequiereCargaInicial` se evaluaba en cada refresco. Si una corrida fallaba al cargar staging o en optimización (`_idCorridaCargada = 0`), el refresco post-ejecución en `finally` resolvía `RequiereCargaInicial = true` y cargaba automáticamente otra corrida ajena, ocultando el error recién emitido.
6. **Fase 1618:** Cuando la corrida activa quedaba excluida por los filtros de período/programa, `seleccion.CorridaActivaNoEncontrada` dejaba `SelectedIndex = -1`, mostrando el ComboBox en blanco y violando el criterio 1 (representar la misma corrida en selector, etiqueta y grilla).
7. **Fase 1619:** El campo privado `_claveFiltroCorridas` era código muerto tras retirar la guarda del `DropDown`, pues solo se escribía y nunca se leía.
8. **Fase 1620:** En `docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md:111-115`, los nombres de las pruebas unitarias tenían discrepancias de nomenclatura contra `AgCorridaSelectorTests.cs`.

### Cambios Aplicados
1. **Fase 1613 (`AgPerfilAlcanceResolver.cs`, `AgParticionCarriles.cs`, `ConversorPDF.XLL.csproj`):**
   - Versionado de los archivos de soporte de rondas previas (`AgPerfilAlcanceResolver.cs`, `AgParticionCarriles.cs` y sus tests de integración y caracterización) para garantizar que un clon limpio de HEAD compile sin errores CS2001 ni CS0246.
2. **Fase 1614 (`App.cs`, `InstaladorConversorPDF.wxs`):**
   - Versionado coherente del bump `1.6.4.07` en `App.cs` e `InstaladorConversorPDF.wxs` alineado con `ConversorPDF.XLL.csproj`.
3. **Fase 1615 (`AgFormGeneradorHorarios.cs`):**
   - Sustitución de `DropDown` por `DropDownClosed += (s, e) => RefrescarComboCorridasSeguro();`.
   - Guarda defensiva en `RefrescarComboCorridasAsync`: `if (cmbCargarCorridas.DroppedDown) { _solicitudPendienteRefrescoCombo = true; return; }` para asegurar que nunca se manipulen `Items` mientras el popup esté desplegado.
4. **Fase 1616 (`AgFormGeneradorHorarios.cs`):**
   - Introducción de la guarda `private bool _cargandoCorridaDesdeCombo;` en `CargarCorridaDesdeComboAsync` con `try/finally` asegurando exclusión mutua de cargas concurrentes.
5. **Fase 1617 (`AgFormGeneradorHorarios.cs`):**
   - Parametrización `RefrescarComboCorridasAsync(bool permitirCargaInicial = false)` y `RefrescarComboCorridasSeguro(bool permitirCargaInicial = false)`.
   - Invocación con `permitirCargaInicial: true` exclusivamente en `OnShown` al inicializar la pantalla. Refrescos posteriores, post-ejecución y retorno de diálogos usan `false` por defecto.
6. **Fase 1618 (`AgFormGeneradorHorarios.cs`):**
   - Inyección de la corrida activa como ítem con etiqueta `" (fuera de filtro)"` cuando los filtros activos la excluyen, seleccionándola en el combo y actualizando `lblAgModo` informativamente sin dejar el ComboBox en blanco.
7. **Fase 1619 (`AgFormGeneradorHorarios.cs`):**
   - Eliminación del campo muerto `_claveFiltroCorridas` y de sus asignaciones obsoletas en `RefrescarComboCorridasAsync`.
8. **Fase 1620 (`docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md`):**
   - Corrección de la nomenclatura de los 6 tests unitarios para alinearlos exactamente con `AgCorridaSelectorTests.cs` y documentación de las salvaguardas de `DropDownClosed`, `_cargandoCorridaDesdeCombo` y `permitirCargaInicial`.

### Sincronización REGISTRY
- Fila `R-042` (`closed`) y contadores actualizados a `ronda: 326 / 327`, `fase: 1620 / 1621`.

---

## Ronda 327: Correcciones Fases 1621–1626 — Refresco Diferido, Coalescencia de Cargas y Marcador Fuera de Filtro

### Contexto y Diagnóstico
1. **Fase 1621:** En `AgFormGeneradorHorarios.cs`, la reinyección de la corrida activa asignaba `activa.Etiqueta = ... + " (fuera de filtro)"`, pero la iteración subsiguiente `foreach (var it in items) it.Etiqueta = AgCorridaSelector.FormatearEtiqueta(it);` reformateaba todos los ítems borrando el sufijo antes de llegar a la UI.
2. **Fase 1622:** La guarda `if (cmbCargarCorridas.DroppedDown)` marcaba `_solicitudPendienteRefrescoCombo = true` y retornaba antes de `_refrescandoComboCorridas = true`. Al no haber reintento al cerrar el desplegable, los refrescos post-ejecución ocurridos con el popup abierto se descartaban y el selector no mostraba la corrida `#N`.
3. **Fase 1623:** Sustituir `DropDown` por `DropDownClosed` impedía cumplir el Criterio 2 (*"al desplegar el ComboBox se realiza una consulta fresca"*). Se requiere que `DropDown` inicie la consulta asíncrona y que la mutación de `Items` espere a que el popup esté cerrado, consumiéndose en `DropDownClosed`.
4. **Fase 1624:** Si `OnShown` invocaba `RefrescarComboCorridasSeguro(permitirCargaInicial: true)` mientras `_refrescandoComboCorridas == true`, la invocación en curso completaba su ciclo con `permitirCargaInicial == false`, perdiéndose la autorización de carga automática al abrir el formulario.
5. **Fase 1625:** `_cargandoCorridaDesdeCombo` descartaba `SelectedIndexChanged` si el usuario elegía otra corrida mientras una carga estaba en curso, desalineando el combo (`#B`) y la grilla/métricas (`#A`).
6. **Fase 1626:** R-042 requería adición de fila `R-043` en `REGISTRY.md` para rastrear las correcciones aplicadas y mantener append-only el registro.

### Cambios Aplicados
1. **Fase 1621 (`AgFormGeneradorHorarios.cs`):**
   - Captura de `fueraDeFiltro` y aplicación del sufijo `" (fuera de filtro)"` estrictamente después del bucle `foreach (var it in items) it.Etiqueta = AgCorridaSelector.FormatearEtiqueta(it);`.
2. **Fase 1622 & 1623 (`AgFormGeneradorHorarios.cs`):**
   - Registro dual en `cmbCargarCorridas`: `DropDown` invoca la consulta fresca y `DropDownClosed` consume solicitudes pendientes (`if (_solicitudPendienteRefrescoCombo) { _solicitudPendienteRefrescoCombo = false; RefrescarComboCorridasSeguro(); }`).
   - Traslado de la guarda de inhibición de `DroppedDown` inmediatamente antes de la manipulación de `Items` (`_actualizandoComboCorridas = true`), permitiendo que el query asíncrono corra y difiriendo el swap de la colección al cierre del popup.
3. **Fase 1624 (`AgFormGeneradorHorarios.cs`):**
   - Introducción del campo privado `_cargaInicialPendiente` y acumulación mediante `_cargaInicialPendiente |= permitirCargaInicial;` tanto en la guarda de reentrada como en el deferral por `DroppedDown`.
4. **Fase 1625 (`AgFormGeneradorHorarios.cs`):**
   - Introducción del campo privado `_recargaCorridaPendiente`. En `CargarCorridaDesdeComboAsync`, si hay carga en curso y la selección cambió, se marca `_recargaCorridaPendiente = true`, despachando la recarga en el bloque `finally`.
5. **Fase 1626 (`.agents/plans/REGISTRY.md`):**
   - Inclusión de fila `R-043` (`closed`, gate `Full`) y actualización de contadores a `ronda: 327 / 328`, `fase: 1626 / 1627`.

---

## Ronda 328: Refresco Anticipado del Selector, Desacoplamiento de Modal y Trazabilidad de Commit (Fases 1627–1629)

### Contexto y Diagnóstico
1. **Fase 1627:** En `implementation_plan.md` se declararon resueltas las Fases 1613/1614, pero al estar en modo revisión sin compilar ni ejecutar herramientas git de mutación de historial, los archivos `AgPerfilAlcanceResolver.cs` y `AgParticionCarriles.cs` permanecen untracked (`??`) y `App.cs`/`*.wxs` modificados en el working tree sin commit, requiriendo su inclusión formal en el commit atómico de cierre en modo Act.
2. **Fase 1628:** Disparar la consulta en `DropDown` dejaba la lista visible sin actualizar en la primera apertura porque la guarda `DroppedDown` difería el swap hasta `DropDownClosed`. Se requiere anticipar la consulta a eventos previos (`Enter` y `MouseDown` con botón izquierdo) para que la consulta concluya antes o durante el despliegue.
3. **Fase 1629:** Si la carga de una corrida fallaba mostrando un modal de error (`MessageBox`), el bloque `finally` despachaba incondicionalmente la recarga pendiente encolada en vuelo. Si la segunda corrida también fallaba, se disparaba un segundo modal de error consecutivo y el encadenamiento síncrono/recursivo anidaba llamadas en la pila.

### Cambios Aplicados
1. **Fase 1627 (`implementation_plan.md`, `walkthrough.md`):**
   - Corrección de la declaración de estado de 1613/1614: documentadas como preparadas en árbol de trabajo y pendientes de commit atómico en modo Act.
2. **Fase 1628 (`AgFormGeneradorHorarios.cs`):**
   - Enlace anticipado de eventos en `cmbCargarCorridas`: `Enter += (s, e) => RefrescarComboCorridasSeguro();` y `MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) RefrescarComboCorridasSeguro(); };`, manteniendo `DropDown` y `DropDownClosed` como respaldos defensivos.
3. **Fase 1629 (`AgFormGeneradorHorarios.cs`):**
   - Introducción del campo `private bool _ultimaCargaDesdeComboOk;`. En `CargarCorridaDesdeComboAsync`, registro de `_ultimaCargaDesdeComboOk = cargada;` en el flujo exitoso y reseteo a `false` en los bloques `catch`.
   - En el bloque `finally`, condicionamiento del despacho pendiente a `_ultimaCargaDesdeComboOk == true` y ejecución desacoplada fuera de la pila mediante `BeginInvoke(new Action(() => { if (!IsDisposed) _ = CargarCorridaDesdeComboAsync(); }));`.
4. **Fase 1626/1629 Sincronización REGISTRY:**
   - Inclusión de fila `R-044` (`closed`, gate `Full`) en `REGISTRY.md` con contadores actualizados a `ronda: 328 / 329`, `fase: 1629 / 1630`.

---

## Ronda 329: Reapertura Segura de Popup, Deduplicación de Consultas y Robustez de Handle (Fases 1630–1632)

### Contexto y Diagnóstico
1. **Fase 1630:** `_recargaCorridaPendiente` solo se limpiaba si `_ultimaCargaDesdeComboOk == true`. Ante un fallo de carga, el flag permanecía en `true`, causando que la siguiente carga exitosa disparara una recarga espuria. Asimismo, `BeginInvoke` sin validar `IsHandleCreated` provocaba `InvalidOperationException` al cerrar el formulario.
2. **Fase 1631:** La consulta asíncrona iniciada en `Enter`/`DropDown` finalizaba cuando `DroppedDown == true`. La guarda previa difería el swap a `DropDownClosed`, por lo que en clic único el usuario seguía viendo la lista obsoleta en la primera apertura.
3. **Fase 1632:** La multiplicidad de disparadores (`Enter`, `MouseDown`, `DropDown`) ejecutaba hasta 3 consultas y llamadas a `Items.Clear()` en rápida sucesión, generando parpadeo visual y saturación de consultas a SQLite.

### Cambios Aplicados
1. **Fase 1630 (`AgFormGeneradorHorarios.cs`):**
   - En `CargarCorridaDesdeComboAsync`, el flag `_recargaCorridaPendiente` se limpia siempre incondicionalmente en `finally`.
   - Se valida `IsHandleCreated` y se envuelve `BeginInvoke` en `try/catch (InvalidOperationException)` para prevenir excepciones durante el desensamble del handle.
2. **Fase 1631 (`AgFormGeneradorHorarios.cs`):**
   - Sustitución de la guarda abortiva por cierre temporal (`bool reabrirPopup = cmbCargarCorridas.DroppedDown; if (reabrirPopup) cmbCargarCorridas.DroppedDown = false;`), mutación segura de `Items` mientras `DroppedDown == false`, y reapertura asíncrona del popup en `BeginInvoke` tras completar el swap.
3. **Fase 1632 (`AgFormGeneradorHorarios.cs`):**
   - Introducción del campo `private DateTime _ultimoRefrescoComboUtc = DateTime.MinValue;` y deduplicación temporal (ventana de 750 ms) en `RefrescarComboCorridasSeguro(..., bool omitirSiReciente = false)`.
   - Retiro del evento redundante `MouseDown` en `cmbCargarCorridas` y activación de `omitirSiReciente: true` en `Enter` y `DropDown`.
4. **Sincronización REGISTRY:**
   - Inclusión de fila `R-045` (`closed`, gate `Full`) y actualización de contadores a `ronda: 329 / 330`, `fase: 1632 / 1633`.

---

## Ronda 330: Desacoplamiento de Reapertura, Flag de Cierre Interno y Precisión de Timestamp (Fases 1633–1635)

### Contexto y Diagnóstico
1. **Fase 1633:** Reabrir el popup mediante `cmbCargarCorridas.DroppedDown = true` disparaba el evento `DropDown`, cuyo ciclo de refresco solo se contenía si el hilo STA procesaba el evento en menos de 750 ms. Una pausa mayor generaba un bucle indefinido de apertura/cierre de popup.
2. **Fase 1634:** Cerrar programáticamente el popup elevaba `DropDownClosed`, consumiendo `_solicitudPendienteRefrescoCombo` e iniciando una segunda cascada completa e innecesaria de refresco. Además, cuando `items.Count == 0`, la sentencia `continue` saltaba la reapertura del popup, dejándolo cerrado ante el usuario.
3. **Fase 1635:** `_ultimoRefrescoComboUtc` se asignaba en el bloque `finally` de `RefrescarComboCorridasAsync`, suprimiendo refrescos legítimos posteriores durante 750 ms incluso tras errores de consulta o listas vacías donde no hubo swap efectivo.

### Cambios Aplicados
1. **Fase 1633 (`AgFormGeneradorHorarios.cs`):**
   - Introducción del flag `private bool _reabriendoPopupCombo;`.
   - Guarda `if (_reabriendoPopupCombo) return;` en `cmbCargarCorridas.DropDown` para inhibir refrescos originados por reaperturas programáticas.
2. **Fase 1634 (`AgFormGeneradorHorarios.cs`):**
   - Introducción del flag `private bool _cerrandoPopupInterno;`.
   - Envoltorio de `DroppedDown = false` con `_cerrandoPopupInterno = true / finally false` y guarda `if (_cerrandoPopupInterno) return;` en `cmbCargarCorridas.DropDownClosed`.
   - Reubicación de la reapertura diferida mediante helper `ReabrirPopupDiferido()` en el bloque `finally` del swap para asegurar que se ejecute siempre, inclusive si `items.Count == 0`.
3. **Fase 1635 (`AgFormGeneradorHorarios.cs`):**
   - Reubicación de `_ultimoRefrescoComboUtc = DateTime.UtcNow;` estrictamente tras la asignación de `SelectedIndex` dentro del `try`, registrando el timestamp exclusivamente tras swaps efectivos.
4. **Sincronización REGISTRY:**
   - Inclusión de fila `R-046` (`closed`, gate `Full`) y actualización de contadores a `ronda: 330 / 331`, `fase: 1635 / 1636`.

---

## Ronda 331: Reapertura con Validación de Foco y Guard Contra Cascadas Espurias (Fases 1636–1637)

### Contexto y Diagnóstico
1. **Fase 1636:** La reapertura asíncrona mediante `BeginInvoke` en `ReabrirPopupDiferido()` se ejecutaba incluso si el usuario había hecho clic en otro control o ventana, robando el foco. Adicionalmente, el uso de un `bool` simple para `_reabriendoPopupCombo` dejaba desprotegidas reaperturas concurrentes o solapadas.
2. **Fase 1637:** En `RefrescarComboCorridasAsync`, la guarda de reentrada marcaba `_solicitudPendienteRefrescoCombo = true` incondicionalmente cada vez que colisionaban eventos (p. ej. `Enter` y `DropDown` por el mismo clic), forzando una segunda iteración completa de cascada de filtros y repoblado del combo sin cambio alguno en los filtros.

### Cambios Aplicados
1. **Fase 1636 (`AgFormGeneradorHorarios.cs`):**
   - Transformación de `_reabriendoPopupCombo` a contador reentrante (`private int _reabriendoPopupCombo;`) con incrementos y decrementos balanceados en `try/finally`.
   - Comprobación estricta de `cmbCargarCorridas.Focused` antes de asignar `cmbCargarCorridas.DroppedDown = true` dentro del delegado `BeginInvoke`.
2. **Fase 1637 (`AgFormGeneradorHorarios.cs`):**
   - Introducción del campo `private string _claveFiltroEnVuelo;`. En la guarda de reentrada, solo se marca `_solicitudPendienteRefrescoCombo = true` si `ClaveFiltroCorridas() != _claveFiltroEnVuelo`.
   - Limpieza garantizada de `_claveFiltroEnVuelo = null;` en el bloque `finally` de `RefrescarComboCorridasAsync`.
3. **Sincronización REGISTRY:**
   - Inclusión de fila `R-047` (`executing`, gate `Full`) y actualización de contadores a `ronda: 331 / 332`, `fase: 1637 / 1638`.

---

## Ronda 338: Validación y Corrección de Dependencias, Integridad de Plan Explícito y Estado del Árbol (Fases 1687–1691)

### Contexto y Diagnóstico
1. **Fase 1687:** Olas incompatibles con dependencias causales en el plan: Wave 1 agrupaba T1 y T2 simultáneamente a pesar de que T2 declara `depends_on: [T1]`; idéntica anomalía en Wave 2 con T3 y T4.
2. **Fase 1688:** La propuesta de degradar con `DemoverPlanExplicito()` (`PlanExplicito = false`) ante la invalidación de fechas activaba el expansor semanal tradicional con `AddDays(7)`, destruyendo la reserva explícita.
3. **Fase 1689:** Tolerar déficit por festivos quebrantaba la cobertura estricta (`HorasGen == FechasGen.Count * unidad`).
4. **Fase 1690:** Omisión de `AgGeneradorHorarios.cs` en `files_lock` de T6 y ausencia de suites de test en rutas de locks.
5. **Fase 1691:** Precondición de git desactualizada y error conceptual al asociar exit code 3 únicamente a `SUITE_SHRUNK`.

### Cambios Aplicados
1. **Olas y Dependencias (Fase 1687):** Desdoble a 10 olas en serie (`waves: 10`) con inicio secuencial `Wave n: Tn; iniciar después de T(n-1)`.
2. **Prohibición de Fallback Legado (Fase 1688):** Ante fechas invalidadas, replanificación atómica del grupo (`ReplanificarGrupo(grupo, contexto)`); si no es factible, `Result.Failure` sin degradación.
3. **Cobertura Estricta (Fase 1689):** Descuento de festivos en reservas del planificador; `Result.Failure` ante déficit y suite `Expansor_DeficitPorFestivo_Falla`.
4. **Completitud de Leases (Fase 1690):** Inclusión de `AgGeneradorHorarios.cs` en T6 y rutas canónicas completas con tests en T2–T9.
5. **Precondición y Diagnóstico (Fase 1691):** HEAD `c0d986d` verificado limpio; exit code 3 diagnosticado consultando `summary.md`.

---

## Ronda 339: Consolidación Documental y Sincronización del Plan de Ranuras Dinámicas (Fase 1692)

### Contexto y Diagnóstico
- **Fase 1692:** Desincronización entre las correcciones aprobadas de la Ronda 338 y el contenido físico en disco del plan `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md` (el archivo en disco aún conservaba `waves: 8`, omisiones en `files_lock` y referencias a `DemoverPlanExplicito`).

### Cambios Aplicados
- **Sincronización en Disco:** Aplicación efectiva de las 5 correcciones de la Ronda 338 directamente sobre `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`.
- **Verificación de Diff:** 10 olas consolidadas, rutas completas con tests para T2–T9, prohibición formal de `DemoverPlanExplicito` y precondición de árbol limpio certificada.


## Ronda 340: Rutas del plan y numeración (Fases 1693–1694)

- **Fase 1693 realizada (documental):** corregida ruta de AgGen.cs en T6 a Academic/Domain/Models/Genetica; completada ruta de AgGeneradorHorarios.cs en T10.
- **Fase 1694 realizada (documental):** registradas revisiones 338–339 en R-054/R-055 y esta corrección en R-056. Contadores ronda 340/341, fase 1694/1695. R-047 continúa executing.
- **Verificación:** inspección de archivos; sin compilación ni pruebas. T1–T10 siguen pending y sin reserva. No se implementó código 4→8.

---

## Ronda 366: Corrección de Criterios de Mejora, Persistencia de Origen de Patrón, Makespan y Alcance de Reset (Fases 1742–1746)

### Contexto y Diagnóstico
1. **Fase 1742:** La decisión en `AgEscaladorVentanaAula.Evaluar` evaluaba únicamente crecimiento de capacidad total (`CapacidadDespues > CapacidadAntes`), ignorando si la saturación de carril persistía inalterada (p. ej. 3 materias de 48h en domingo donde JUE-DOM no reduce la sobrecarga dominical).
2. **Fase 1743:** El `INSERT` y `SELECT` de `ag_aula_asignatura` omitían la columna `origen_patron`, forzando a que todas las filas leídas o sembradas quedaran con valor por defecto `LEGADO`, impidiendo que el reseteo transaccional reconociera los patrones automáticos.
3. **Fase 1744:** `AgEvaluadorDemandaVentana` no proyectaba `SemanasRequeridas`, `FechaFinMinima` ni `RanurasOciosas` a partir de la distribución en carriles (`AgCarriles.Cantidad`), limitándose a bin-packing sin trazabilidad temporal.
4. **Fase 1745:** `AgEscaladorVentanaAula.AplicarAsync` ejecutaba un reseteo de patrones sobre el aula completa sin acotación a las asignaturas causantes del cuello de botella y sin reflejar aviso en el diálogo modal interactivo.
5. **Fase 1746:** Inconsistencia de puntero en fila `R-066` de `REGISTRY.md` que referenciaba el commit de T7 en lugar del commit de cierre T8 (`c34bec4`).

### Cambios Aplicados
1. **Fase 1742 (`AgEscaladorVentanaAula.cs`):** Introducción de `Mejora` en `Decision` y verificación estricta `bool mejora = !evalDespues.HaySaturacion || !evalDespues.SaturacionCarriles;`. Si la saturación de carril no disminuye, la propuesta se descarta.
2. **Fase 1743 (`AgScheduleRepository.cs`, `AgFormGeneradorHorarios.cs`, `AgAulaAsignatura.cs`):** Inclusión de `origen_patron` con valor `'AUTOMATICO'` en las consultas `INSERT` de siembra y vinculación, mapeo en el `SELECT` del contexto académico y valor por defecto inicializado a `AUTOMATICO` en la entidad.
3. **Fase 1744 (`AgEvaluadorDemandaVentana.cs`, `AgResultadoEvaluacionVentana.cs`):** Adición de cálculo de `RanurasOciosas`, `SemanasRequeridas` y `FechaFinMinima` proyectada sobre `carriles`, registrando `IdAsignaturasSaturadas`.
4. **Fase 1745 (`AgEscaladorVentanaAula.cs`, `AgFormGeneradorHorarios.cs`):** Acotación del reseteo SQL a `IdAsignaturasBottleneck` en `ag_aula_asignatura`, activación de `MarkDirty()` ante cambios de patrones aunque el aula conserve ID, y adición de aviso explícito de recálculo de patrones en el diálogo modal interactivo.
5. **Fase 1746 (`REGISTRY.md`):** Sincronización de la fila `R-066` apuntando al commit de cierre `c34bec4`.

---

## Ronda 367: Saturación Estricta de Carril, Compatibilidad Total de Patrón, Makespan con Festivos y Preservación en Regeneración (Fases 1747–1752)

### Contexto y Diagnóstico
1. **Fase 1747:** La saturación de carril no consideraba `SemanasRequeridas` en `Mejora`, por lo que 3 materias fijas en Domingo proponían `JUE-DOM` erróneamente sin reducir la duración requerida.
2. **Fase 1748:** `ElegirCarril` usaba `Any` aceptando carriles con coincidencia parcial de días, enmascarando saturación y asignando sesiones a días fuera del patrón.
3. **Fase 1749:** `FinDeSemanaValida` calculaba una fecha lineal fija sin descontar festivos ni recorrer semanas válidas.
4. **Fase 1750:** `InsertAulaAsignaturasAsync` ejecutaba `DELETE` indiscriminado sobre el aula perdiendo patrones `FIJO/LEGADO`.
5. **Fase 1751:** El valor por defecto de `OrigenPatron` en la entidad debía permanecer en `"LEGADO"` para no clasificar datos antiguos como automáticos ni quebrar pruebas existentes.
6. **Fase 1752:** `IdAsignaturasBottleneck` no identificaba quirúrgicamente las asignaturas sin cabida y el conteo del modal difería de las asignaturas reseteadas.

### Cambios Aplicados
1. **Fase 1747 (`AgEvaluadorDemandaVentana.cs`, `AgEscaladorVentanaAula.cs`):** `Mejora` condicionada a `evalDespues.SemanasRequeridas < evalAntes.SemanasRequeridas || (!evalDespues.HaySaturacion && evalAntes.HaySaturacion) || (!evalDespues.SaturacionCarriles && evalAntes.SaturacionCarriles)`.
2. **Fase 1748 (`AgEvaluadorDemandaVentana.cs`):** `ElegirCarril` exige `d.DiasPatron.All(day => carriles[c].Any(r => r.Dia == day))` para patrones no automáticos; si no hay compatibilidad total devuelve `-1`.
3. **Fase 1749 (`AgEvaluadorDemandaVentana.cs`):** `FinDeSemanaValida` itera semanas calendario verificando días hábiles no festivos en la ventana activa.
4. **Fase 1750 (`AgScheduleRepository.cs`):** `InsertAulaAsignaturasAsync` reemplaza el `DELETE` masivo por borrado selectivo exclusivo de `AUTOMATICO` no seleccionadas e `INSERT ... ON CONFLICT DO NOTHING`.
5. **Fase 1751 (`AgAulaAsignatura.cs`):** `OrigenPatron` restaurado a `"LEGADO"` en constructor y propiedad.
6. **Fase 1752 (`AgEvaluadorDemandaVentana.cs`, `AgEscaladorVentanaAula.cs`, `AgFormGeneradorHorarios.cs`):** `IdAsignaturasSaturadas` captura asignaturas que no cupieron (`noCupieron`); conteo modal y reseteo en memoria alineados exactamente a `IdAsignaturasBottleneck`.

---

## Ronda 368: Corrección de Compilación, Indexación Segura de Carriles, Saturación por Ocurrencias, Fecha Fin de Ventana y Preservación en Siembra (Fases 1753–1758)

### Contexto y Diagnóstico
1. **Fase 1753:** Bloque huérfano `{ ... }` a nivel de clase en `AgEvaluadorDemandaVentana.cs` tras el método `PuedeAcomodar` causaba errores de compilación CS1519 y CS0103.
2. **Fase 1754:** Cuando `ElegirCarril` retornaba `-1` para materias con patrones no compatibles con ningún carril, el llamador intentaba acceder a `carriles[-1]`, generando `ArgumentOutOfRangeException`.
3. **Fase 1755:** La saturación de carriles en `AgEvaluadorDemandaVentana.cs` dependía de bandas agrupadas en vez de contrastar la carga contra la capacidad real de ocurrencias del calendario (`capCarril[c]`), permitiendo que 3 materias fijas en domingo sobre `JUE-DOM` pasaran como no saturadas al juntar sábado y domingo en una sola banda. Además, las sesiones se repartían en todas las ranuras del carril en vez de limitarse a las ranuras compatibles con el patrón de la asignatura.
4. **Fase 1756:** `FinDeSemanaValida` calculaba `cursor.AddDays(-1)` devolviendo siempre el sábado de la última semana, en lugar del último día activo de la ventana académica libre de festivos (domingo para `VIE-DOM`).
5. **Fase 1757:** En la siembra de `AgFormGeneradorHorarios.cs`, la cláusula `ON CONFLICT(id_aula, id_asignatura) DO UPDATE` actualizaba `dias_patron` indiscriminadamente sin validar `origen_patron`, sobrescribiendo patrones de asignaturas `FIJO` o `LEGADO`. Asimismo, en `AgEscaladorVentanaAula.cs`, ante cuellos de botella vacíos, `IdAsignaturasBottleneck` caía incorrectamente a todas las asignaturas del aula.
6. **Fase 1758:** Desactualización del registro de rondas y fases en `REGISTRY.md` con respecto a las revisiones 366, 367 y 368.

### Cambios Aplicados
1. **Fase 1753 (`AgEvaluadorDemandaVentana.cs`):** Eliminación total del bloque huérfano de código residual posterior a `PuedeAcomodar`.
2. **Fase 1754 (`AgEvaluadorDemandaVentana.cs`):** Verificación de `idxCarril < 0`. Si es menor que cero, se añade la asignatura a `materiasFueraVentana` y se omite la indexación sobre `carriles`. Si existen materias fuera de ventana, se marca `HaySaturacion = true`, `SaturacionCarriles = true` y se retorna el resultado de evaluación inmediatamente.
3. **Fase 1755 (`AgEvaluadorDemandaVentana.cs`):** Distribución de sesiones acotada a `slotsPatron` (`Where(r => d.DiasPatron.Contains(r.Dia))`). Evaluación de saturación directa calculando `capCarril[c]` con `AgPlanificadorTemporal.ContarOcurrencias` y marcando saturación si `cargaCarriles[c] > capCarril[c]`.
4. **Fase 1756 (`AgEvaluadorDemandaVentana.cs`):** En `FinDeSemanaValida`, búsqueda descendente desde el final de la semana (`cursor.AddDays(-7 + d)`) para retornar exactamente el último día de la ventana académica no festivo.
5. **Fase 1757 (`AgFormGeneradorHorarios.cs`, `AgEscaladorVentanaAula.cs`):** En `sqlInsertAsig`, actualización condicional de `dias_patron` restringida a registros con `origen_patron = 'AUTOMATICO'`. En `AgEscaladorVentanaAula.cs`, asignación directa de `evalAntes.IdAsignaturasSaturadas ?? new List<int>()` sin fallback expansivo al aula completa.
6. **Fase 1758 (`REGISTRY.md`):** Registro de filas `R-067`, `R-068` y `R-069` en estado `closed` y actualización de contadores a `ronda: 368 / 369`, `fase: 1758 / 1759`.

---

## Ronda 369: Granularidad por Slot/Franja, Poblado de Bottlenecks, Retención de Festivo y Registro (Fases 1759–1762)

### Contexto y Diagnóstico
1. **Fase 1759:** La saturación calculada a nivel de carril agrupaba las ranuras de todo el carril, ocultando la sobrecarga en una franja específica de un solo día (por ejemplo, 2 materias en viernes `{5}` en carril `[5,6M]` sumaban 24 frente a 34 nominal, enmascarando que el viernes con 17 ranuras estaba saturado).
2. **Fase 1760:** En casos de déficit global (`demandaTotal > capacidadTotal`) o saturación por slot/carril, `res.IdAsignaturasSaturadas` quedaba vacío, lo que provocaba que `AplicarAsync` cayera en una rama `else` que reseteaba masivamente todas las materias automáticas del aula/subunidad.
3. **Fase 1761:** La documentación y fixtures de prueba requerían explicitar que `FechaFinMinima` corresponde al último día no festivo de la última semana válida de la ventana (y no a un fin de semana genérico), y que un festivo en el día activo del patrón (ej. domingo 15 de agosto de 2027) desplaza la fecha de terminación al 17 de octubre de 2027.
4. **Fase 1762:** Estado del registro de revisiones sin reflejar la ronda 369 ni el consumo de las fases 1759 a 1762.

### Cambios Aplicados
1. **Fase 1759 (`AgEvaluadorDemandaVentana.cs`):** Implementación de diccionario `slotCarga` indexado por `(r.Dia, r.Franja)` para ranuras reales del calendario. Reparto fraccional exacto entre los slots compatibles con el patrón del carril y detección directa de saturación si `slotCarga[(Dia, Franja)] > AgPlanificadorTemporal.ContarOcurrencias(Dia, ...)`.
2. **Fase 1760 (`AgEvaluadorDemandaVentana.cs`, `AgEscaladorVentanaAula.cs`):** Poblado garantizado de `res.IdAsignaturasSaturadas` con todas las asignaturas participantes en escenarios de sobrecarga o déficit. En `AgEscaladorVentanaAula.cs`, eliminación de la rama `else` de reseteo masivo incondicional, limitando el reseteo exclusivamente a listas no vacías de cuellos de botella (`IdAsignaturasBottleneck`).
3. **Fase 1761 (`AgEvaluadorDemandaVentana.cs`):** Formalización de la semántica de `FinDeSemanaValida` que devuelve el último día hábil de la ventana académica libre de festivos para la cantidad de semanas requeridas calculada.
4. **Fase 1762 (`REGISTRY.md`):** Adición de la fila `R-070` (`closed`, sin build/tests) y avance de contadores a `ronda: 369 / 370`, `fase: 1762 / 1763`.

---

## Ronda 370: Exposición de Semanas/Fecha Fin, Diagnóstico I5_DIA_SATURADO, Retiro de Bandas y Aislamiento de Bottlenecks (Fases 1763–1767)

### Contexto y Diagnóstico
1. **Fase 1763:** `SemanasRequeridas`, `FechaFinMinima` y `RanurasOciosas` no se propagaban a `AgEscaladorVentanaAula.Decision`, impidiendo que el modal informativo del generador y el validador pre-vuelo expusieran el diagnóstico y la proyección temporal. En particular, ante saturación no escalable (caso 3 domingos), no se emitía `I5_DIA_SATURADO`.
2. **Fase 1764:** Formalización de la expectativa de 36 semanas con festivos del calendario donde un domingo festivo desplaza la terminación al 17 de octubre de 2027.
3. **Fase 1765:** El método `PuedeAcomodar` basado en partición por bandas continuaba ejecutándose innecesariamente tras el cálculo por slot, introduciendo complejidad combinatoria y riesgo de bloqueo en pre-vuelo.
4. **Fase 1766:** Ante saturación por slot, `res.IdAsignaturasSaturadas` capturaba todas las asignaturas de la demanda sin filtrar cuáles recaían efectivamente sobre los slots saturados, ampliando innecesariamente el reset de patrones.
5. **Fase 1767:** Desactualización del registro de numeración para la ronda 370 y fases 1763–1767.

### Cambios Aplicados
1. **Fase 1763 (`AgEscaladorVentanaAula.cs`, `AgFormGeneradorHorarios.cs`, `AgValidadorVentana.cs`):** Adición de `SemanasRequeridas`, `FechaFinMinima` y `RanurasOciosas` a `Decision`, inclusión en el texto de las líneas de detalle del modal de confirmación, y emisión del bloqueo `[I5_DIA_SATURADO]` en `AgValidadorVentana.cs` ante saturación de ventana o carril.
2. **Fase 1764 (`AgEvaluadorDemandaVentana.cs`):** Mantenimiento de la semántica de semanas y fecha fin sobre el calendario real considerando festivos.
3. **Fase 1765 (`AgEvaluadorDemandaVentana.cs`):** Eliminación completa del algoritmo de bandas y del método recursivo `PuedeAcomodar`, consolidando la evaluación sobre el mapa `slotCarga`.
4. **Fase 1766 (`AgEvaluadorDemandaVentana.cs`):** En saturación por slot, `IdAsignaturasSaturadas` filtra quirúrgicamente solo aquellas asignaturas cuyos `DiasPatron` intersectan con los días de los `slotsSaturados`.
5. **Fase 1767 (`REGISTRY.md`):** Adición de la fila `R-071` (`closed`, sin build/tests) y avance de contadores a `ronda: 370 / 371`, `fase: 1767 / 1768`.

---

## Ronda 371: Supresión de Diagnóstico Contradictorio, Acotación en Déficit Global y Registro (Fases 1768–1771)

### Contexto y Diagnóstico
1. **Fase 1768:** Al detectar saturación en franja e informar `[I5_DIA_SATURADO]`, la ejecución en `AgValidadorVentana.cs` continuaba evaluando oleadas, emitiendo un bloqueo redundante y contradictorio `[I5_INFACTIBLE_OLEADA]` que sugería erróneamente escalar la ventana a 'JUE-DOM'.
2. **Fase 1769:** Verificación y formalización de que ante un domingo festivo del calendario académico (ej. 15 de agosto de 2027), las 36 semanas requeridas extienden la fecha final exactamente al 17 de octubre de 2027.
3. **Fase 1770:** En casos de déficit global (`demandaTotal > capacidadTotal`), `IdAsignaturasSaturadas` capturaba todas las materias de la demanda sin criterio de prioridad; se requería priorizar las asignaturas por mayor volumen horario para un reseteo ordenado.
4. **Fase 1771:** Desactualización del registro de numeración para la ronda 371 y fases 1768–1771.

### Cambios Aplicados
1. **Fase 1768 (`AgValidadorVentana.cs`):** Adición de instrucción `continue;` inmediatamente posterior a la emisión del bloqueo `[I5_DIA_SATURADO]`, impidiendo la ejecución del bloque subsiguiente de oleadas genéricas que propondría un escalado espurio.
2. **Fase 1769 (`AgEvaluadorDemandaVentana.cs`):** Consolidación de la regla de desplazamiento por festivos en `FinDeSemanaValida`.
3. **Fase 1770 (`AgEvaluadorDemandaVentana.cs`):** En déficit global, `IdAsignaturasSaturadas` selecciona las asignaturas ordenadas descendentemente por número de sesiones requeridas.
4. **Fase 1771 (`REGISTRY.md`):** Adición de la fila `R-072` (`closed`, sin build/tests) y avance de contadores a `ronda: 371 / 372`, `fase: 1771 / 1772`.

---

## Ronda 372: Reducción de Materias Fuera de Ventana como Criterio de Mejora y Registro (Fases 1772–1774)

### Contexto y Diagnóstico
1. **Fase 1772:** Si una materia con patrón fijo quedaba fuera de la ventana inicial (por ejemplo, 2 materias en viernes `{5}` sobre `LUN-JUE`), `AgEvaluadorDemandaVentana` reportaba `materiasFueraVentana = 2` y `SemanasRequeridas = 0`. Al evaluar el escalón ampliado (ej. `LUN-VIE`), el patrón encajaba pero requería 24 ranuras (frente a 17 disponibles), quedando saturado (`SemanasRequeridas = 24`). Dado que la condición `Mejora` solo comparaba `evalDespues.SemanasRequeridas < evalAntes.SemanasRequeridas` (`24 < 0` falso) o resolución total de saturación, el escalador descartaba la propuesta a pesar de resolver la exclusión de ventana.
2. **Fase 1773:** Estado de los tests de matriz y verificación estricta de las correcciones previas.
3. **Fase 1774:** Constatación del alcance de reseteo ante déficit global, requiriendo listar y controlar las asignaturas participantes.

### Cambios Aplicados
1. **Fase 1772 (`AgResultadoEvaluacionVentana.cs`, `AgEvaluadorDemandaVentana.cs`, `AgEscaladorVentanaAula.cs`):** Adición de la propiedad `MateriasFueraVentana` a `AgResultadoEvaluacionVentana` y a `AgEscaladorVentanaAula.Decision`. Integración de la cláusula de mejora `evalDespues.MateriasFueraVentana.Count < evalAntes.MateriasFueraVentana.Count` en `Mejora`, habilitando la propuesta de escalado para ventanas con patrones fijos inicialmente fuera de rango.
2. **Fase 1773/1774 (`REGISTRY.md`):** Adición de la fila `R-073` (`closed`, sin build/tests) y avance de contadores a `ronda: 372 / 373`, `fase: 1774 / 1775`.

---

## Ronda 373: Horizonte Acotado de Fecha Fin, Autoridad del Planificador, Idempotencia de Escalado y Registro (Fases 1775–1779)

### Contexto y Diagnóstico
1. **Fase 1775:** `FinDeSemanaValida` iteraba día a día sin límite superior. Ante un calendario enteramente festivo o con festivos infinitos, entraba en ciclo infinito. La propuesta de restringir la búsqueda a `fin` truncaba el cálculo de extensión mínima (makespan) más allá del semestre. Se requería una cota superior finita razonable (`limite = fin.AddDays(...)`) y firma `DateTime?` con fallback en UI y validador a "no cabe dentro del plazo".
2. **Fase 1776:** A pesar de tener `origen_patron`, los patrones marcados como `AUTOMATICO` seguían extrayéndose en `ObtenerDemandasAsignaturas`, en `AgFabricaIndividuos.ObtenerDiasPatron` y en `AgValidadorVentana.patronesAula`, actuando como restricciones rígidas de día en el motor genético y el validador en lugar de redistribuirse libremente.
3. **Fase 1777:** `AgValidadorVentana` anteponía la heurística de `AgEvaluadorDemandaVentana` emitiendo `[I5_DIA_SATURADO]` e interrumpiendo el flujo con `continue` antes de invocar `AgPlanificadorOleadas.Construir`. El planificador es la autoridad algorítmica: solo debe diagnosticarse saturación cuando el plan resulte infactible.
4. **Fase 1778:** `AgEscaladorVentanaAula.AplicarAsync` actualizaba `ag_aula` sin verificar si `dias_ventana` seguía coincidiendo con `Desde`, sobrescribiendo cambios concurrentes. Además, una reaplicación idempotente fallaba si no se detectaba que el aula ya se encontraba en el estado objetivo `Hacia`.
5. **Fase 1779:** Desactualización del registro de numeración para la ronda 373 y fases 1775–1779.

### Cambios Aplicados
1. **Fase 1775 (`AgEvaluadorDemandaVentana.cs`, `AgValidadorVentana.cs`, `AgFormGeneradorHorarios.cs`):** Parámetro `DateTime limite` añadido a `FinDeSemanaValida`, retorno `DateTime?` ante agotamiento del límite, y despliegue del texto "no cabe dentro del plazo" en modal y validaciones cuando el resultado es nulo.
2. **Fase 1776 (`AgEvaluadorDemandaVentana.cs`, `AgFabricaIndividuos.cs`, `AgValidadorVentana.cs`):** Ignorado sistemático de `DiasPatron` cuando `OrigenPatron == "AUTOMATICO"` tanto en la extracción de demandas, en la fábrica de individuos genéticos como en la construcción de patrones para oleadas.
3. **Fase 1777 (`AgValidadorVentana.cs`):** `AgPlanificadorOleadas.Construir` se ejecuta primero; el bloqueo `[I5_DIA_SATURADO]` solo se emite cuando `!planFactible` (`planRes.IsFailure || Infactibles.Count > 0`).
4. **Fase 1778 (`AgEscaladorVentanaAula.cs`):** Guarda SQL `AND COALESCE(dias_ventana, '') = @Desde`, consulta previa en transacción para omitir idempotentemente si ya es `Hacia`, y lanzamiento de `InvalidOperationException` si `dias_ventana != Desde`.
5. **Fase 1779 (`REGISTRY.md`):** Adición de la fila `R-074` (`closed`, sin build/tests) y avance de contadores a `ronda: 373 / 374`, `fase: 1779 / 1780`.

---

## Ronda 374: Métricas del Plan Efectivo en Modal, Confirmadas en Idempotencia y Huella de Entrada (Fases 1780–1783)

### Contexto y Diagnóstico
1. **Fase 1780:** `AgEscaladorVentanaAula.Evaluar` requería unificación metodológica con el planificador de oleadas para garantizar que las propuestas de escalado correspondan a mejoras que el planificador materializa efectivamente.
2. **Fase 1781:** La comprobación de idempotencia en `AplicarAsync` saltaba con `continue` cuando la ventana ya estaba en `Hacia`, devolviendo 0 filas actualizadas. En la UI, la verificación `aplicadas != esperadas` interpretaba erróneamente la reaplicación como fallo.
3. **Fase 1782:** El modal pre-vuelo calculaba la demanda pendiente de forma abstracta (`demanda - capacidad total`) y presentaba las métricas previas en lugar de las proyectadas para la ventana destino.
4. **Fase 1783:** Falta de captura de huella de configuración (`HuellaEntrada`) para garantizar que la demanda, jornada y capacidad no hayan cambiado entre la evaluación y la aplicación.

### Cambios Aplicados
1. **Fase 1780/1782 (`AgEscaladorVentanaAula.cs`, `AgFormGeneradorHorarios.cs`):** Las propiedades `SemanasRequeridas`, `FechaFinMinima`, `RanurasOciosas`, `DemandaPendiente` y `MateriasFueraVentana` de `Decision` ahora almacenan los resultados de `evalDespues`, reflejando exactamente el estado posterior al escalado. En la UI, el texto de demanda pendiente utiliza directamente `d.DemandaPendiente`.
2. **Fase 1781 (`AgEscaladorVentanaAula.cs`):** Introducción de `ResultadoAplicacion` con `Confirmadas` y `Modificadas`. `AplicarDetalladoAsync` incrementa `resultado.Confirmadas++` tanto si se actualiza la base de datos como si ya se encontraba idempotentemente en `Hacia`. `AplicarAsync` retorna `resultado.Confirmadas`, satisfaciendo la condición `aplicadas == esperadas` de la UI.
3. **Fase 1783 (`AgEscaladorVentanaAula.cs`):** Adición de `HuellaEntrada` en `Decision` (`diasVentana|jornada|demanda|capacidad`) para rastrear cambios en la configuración entre la evaluación y la persistencia transaccional.
4. **Numeración y Registro (`REGISTRY.md`):** Adición de la fila `R-075` (`closed`, sin build/tests) y avance de contadores a `ronda: 374 / 375`, `fase: 1783 / 1784`.

---

## Ronda 375: Rollback con Contadores Vacíos, Demanda Pendiente por Ranuras No Acomodadas y Validación de Huella/Idempotencia (Fases 1784–1787)

### Contexto y Diagnóstico
1. **Fase 1784:** En caso de excepción tras incrementar `resultado.Confirmadas` (por ejemplo, en el reseteo transaccional de patrones), la transacción se revertía pero el bloque `catch` devolvía el objeto con contadores positivos, permitiendo que la UI asumiera éxito.
2. **Fase 1785:** `DemandaPendiente` calculada como `Math.Max(0, demandaTotal - capacidadTotal)` ocultaba demandas insatisfechas cuando había capacidad global pero los slots/franjas específicos estaban saturados (ej. materias dominicales).
3. **Fase 1786:** `HuellaEntrada` se asignaba pero no se verificaba contra la base de datos dentro de la transacción previa a la mutación.
4. **Fase 1787:** La comprobación de idempotencia confirmaba el aula si `dias_ventana == Hacia` sin validar que el nodo del árbol (`ag_arbol_nodo`) también estuviera actualizado en la ventana destino.

### Cambios Aplicados
1. **Fase 1784 (`AgEscaladorVentanaAula.cs`):** En el bloque `catch`, se retorna `new ResultadoAplicacion()` (con 0 confirmadas y 0 modificadas) garantizando que un fallo o aborto transaccional no reporte confirmaciones espurias.
2. **Fase 1785 (`AgEvaluadorDemandaVentana.cs`):** Cálculo de `DemandaPendiente` como sumatoria de sesiones no acomodadas en cada slot respecto a su capacidad real (`demandas.Sum(d => Math.Max(0, d.Sesiones - sesionesUbicadasPorAsignatura[d.IdAsignatura]))`).
3. **Fase 1786 (`AgEscaladorVentanaAula.cs`):** Validación transaccional de `HuellaEntrada` comprobando consistencia de ventana y jornada antes de ejecutar la actualización.
4. **Fase 1787 (`AgEscaladorVentanaAula.cs`):** Verificación de idempotencia completa: si `actual == d.Hacia`, se verifica adicionalmente que `ag_arbol_nodo.dias_ventana == d.Hacia` cuando se especifica `idArbol > 0`.
5. **Numeración y Registro (`REGISTRY.md`):** Adición de la fila `R-076` (`closed`, sin build/tests) y avance de contadores a `ronda: 375 / 376`, `fase: 1787 / 1788`.


## Ronda 376 propuesta — Fases 1788–1793 — 2026-11-13

Implementación editada: plan compartido con sesiones fechadas; aprovechamiento de ranuras sobrantes; fábrica consume las mismas fechas; fecha proyectada sin promesa de mínimo; huella completa de contexto/BD; idempotencia y contadores posteriores al commit; recarga de UI y caché por contenido. Pruebas de aceptación y consistencia añadidas.

Estado: sin compilación ni ejecución de pruebas por restricción del usuario. Fases 1792–1793 sin evidencia de aprobación. No se declara release listo ni smoke superado. Detalle vigente: docs/plans/2026-11-12-escalado-tres-casos-revision-ronda-376.md. Sin nueva reserva en REGISTRY.

---

## Ronda 398: Sincronización Documental, Trazabilidad e Invariantes del Planificador con Límite de Ventana (Fases 1815–1819)

### Contexto y Diagnóstico
1. **Fase 1815:** El plan v2 (`2026-11-15-planificador-limite-ventana-design.md`) se redactó contra HEAD previo antes de que se implementara la solución con `EmpaquetarPendientes` y `RedFlujo`. Mantenía citas de líneas obsoletas/muertas (`carrilDe` en `:425`, `:241-258`, etc.) y figuraba como «Sin implementar».
2. **Fase 1816:** El paso 4 del orden de trabajo prescribía revertir con `git checkout --` los archivos del planificador ante fallo, lo cual destruiría la implementación ya presente y funcional de `EmpaquetarPendientes`, `RedFlujo` y la re-baseline autorizada.
3. **Fase 1817:** El criterio de cierre exigía «Ninguna aserción debilitada ni re-baselineada», lo que colisionaba con el re-baseline autorizado en `AgPrioridadHorasTests:157-165` (fechas de asignaturas 2 y 3 por decisión humana para respetar cuotas compactas).
4. **Fase 1818:** Las superficies omitían `AgPlanificadorRanurasFallbackTests.cs` y `.agents/plans/REGISTRY.md` carecía de fila para registrar esta unidad (última asignación histórica R-079).
5. **Fase 1819:** Los invariantes de respaldo sin mutación estaban clasificados como sostenidos «por inspección», cuando en realidad cuentan con pruebas directas en `AgPlanificadorRanurasFallbackTests:20-56` y `:59-83`.

### Cambios Aplicados
1. **Fase 1815 (`docs/plans/2026-11-15-planificador-limite-ventana-design.md`):** Estado actualizado a `IMPLEMENTADO (empaquetado con flujo máximo de horizonte mínimo + respaldo al relleno)`. Reemplazo de citas por líneas obsoletas por citas por símbolo (`reparto por carril en Construir`, `carrilPorId` en `:472-479`, relleno de respaldo en `AgPlanificadorRanuras.cs:433-439`, sustitución por `EmpaquetarPendientes` `:160-245`, `ValidarFechado` `:531-605`, huella 2 ranuras en `AgCarriles.cs:13`, y objetivo de 64 casos).
2. **Fase 1816 (`docs/plans/2026-11-15-planificador-limite-ventana-design.md`):** Paso 4 de orden de trabajo modificado para prohibir `git checkout --` y ordenar la reversión quirúrgica exclusivamente de los símbolos nuevos (`EmpaquetarPendientes`, `RedFlujo`, `asignado`), borrado de `AgPlanificadorRanurasFallbackTests.cs` y restauración de la expectativa en `AgPrioridadHorasTests`.
3. **Fase 1817 (`docs/plans/2026-11-15-planificador-limite-ventana-design.md`):** Criterio de cierre actualizado explicitando la excepción autorizada para `CompletarReferencia_WhenMayorCargaPendiente`; familia de 64 casos pendiente de corrida (sin evidencia todavía).
4. **Fase 1818 (`docs/plans/2026-11-15-planificador-limite-ventana-design.md`, `.agents/plans/REGISTRY.md`):** Adición de `AgPlanificadorRanurasFallbackTests.cs` en las superficies previstas. Incorporación de la fila `R-080` en `REGISTRY.md` y avance de contadores a `ronda: 398 / 399`, `fase: 1819 / 1820`.
5. **Fase 1819 (`docs/plans/2026-11-15-planificador-limite-ventana-design.md`):** Inclusión en la tabla de invariantes de las filas correspondientes a `AgPlanificadorRanurasFallbackTests:20-56` (respaldo sin mutación cuando día de patrón no tiene hueco) y `:59-83` (día de patrón ausente en ventana pero presente en ranuras).

---

