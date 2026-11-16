<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# walkthrough - Fases 1342-1509 (Rondas 264..286)

> Indice: [../../../walkthrough.md](../../../walkthrough.md) - Escision: 2026-11-16

---
### Evaluación Crítica de las Fases

#### 1. Fase 1342 — Propiedad `SelectedIndex` y Carga Tipada
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `SelectedIndexSet` no existe en WinForms ni en los call sites; los sitios reales usan asignación de propiedad (`cmbProgramas.SelectedIndex = 0`).
- **Acción Aplicada:** Implementar `public int SelectedIndex { get; set; }` y carga tipada `Cargar<T>` sin binding `DataSource`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Controls/CheckedComboBox.cs`, `AgFormGeneradorHorarios.cs`.

#### 2. Fase 1343 — Modo de Selección Único y Disparo de Evento
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `SelectedIndexChanged` sin invocación congelaba el recálculo de ventana semestral en `OnVentanaChanged`.
- **Acción Aplicada:** Introducir enum `ModoSeleccion { Unico, Multiple }`, configurar `chkPeriodos.Modo = ModoSeleccion.Unico` y emitir `SelectedIndexChanged` al alternar.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Controls/CheckedComboBox.cs`, `AgFormGeneradorHorarios.cs:233, 491-497`.

#### 3. Fase 1344 — Compensación de Altura en `SplitterDistance`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `grpSeleccion.Size` es recalculado por `Dock = Fill`; además omitió `pnlArbolToolbar` (34 px), dejando el árbol en 151 px.
- **Acción Aplicada:** `pnlFiltros` a 66 px (2 filas en y=4 y y=36) y `splitContenedorSim.SplitterDistance = 390` (+20 px), preservando 189 px útiles de árbol.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs:177, 441`.

#### 4. Fase 1345 — Uso Canónico de `ctx.Open()` en Repositorio
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `EncryptedSQLiteContext` no expone propiedad pública `Connection`.
- **Acción Aplicada:** Obtener la conexión mediante `using (var conn = ctx.Open())` para ejecutar la transacción Dapper.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgFiltroSeleccionRepository.cs`.

#### 5. Fase 1346 — Registro de Migración V78 en `SqliteHelper`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `MigrationV78_AddAgFiltroSeleccion` no estaba registrada en el array de migraciones activas de `SqliteHelper.cs`.
- **Acción Aplicada:** Agregar `new MigrationV78_AddAgFiltroSeleccion()` a continuación de `MigrationV77`.
- **Archivos Afectados:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs:378`.

#### 6. Fase 1347 — Saneamiento de Régimen de Estados
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Reaparición de marcas positivas prematuras en rondas documentales sin código físico.
- **Acción Aplicada:** Estado estricto `ESPECIFICADA` para todo el backlog no implementado y guiones `—` en solución/evidencia de fases pendientes.
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`.

---

### Verificación de Principios (Ronda 240)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1342 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1343 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1344 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1345 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1346 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1347 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `implementation_plan.md`, `walkthrough.md` |

---

## 🔧 Ronda 241: Parches 1342–1347 Registrados Sin Código y 6 Defectos en las Correcciones "Aceptadas" (Fases 1348 a 1353)

### Diagnóstico y Contexto
- **Régimen de Estados:** Todo el backlog (Fases 1311 a 1353) se mantiene en `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Evitar disponer conexión compartida (`AcademicRepository.cs:25`): usar `Func<IDbConnection>` sin `using` local sobre `Open()`.
  - Definición unificada, coherente y compilable de `CheckedComboBox` (evita duplicidad y símbolos huérfanos).
  - Posicionamiento exacto de etiquetas y controles en `pnlFiltros` de 66 px eliminando colisiones.
  - Declaración e inicialización explícita de controles en `AgFormGeneradorHorarios.Designer.cs` y desvinculación de combos antiguos.
  - Sincronización del modelo de persistencia `ag_filtro_seleccion` con la restricción de período único (`Modo = ModoSeleccion.Unico`).
  - Accesores canónicos `ProgramaPrimario()` y `ProgramasSeleccionados()` para desacoplar los 6 call sites sensibles de `chkProgramas`.

### Evaluación Crítica de las Fases

#### 1. Fase 1348 — Conexión Compartida sin `using` en Repositorio
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `using (var conn = ctx.Open())` dispone la conexión compartida del contexto; `SyncScheduler.ForcePersistVaultAsync` falla al cerrar con `ObjectDisposedException`.
- **Acción Aplicada:** Inyectar `Func<IDbConnection>` en `AgFiltroSeleccionRepository` y operar sin `using` sobre la conexión compartida del contexto.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgFiltroSeleccionRepository.cs`.

#### 2. Fase 1349 — Definición Unificada de `CheckedComboBox`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Snippets previos divergentes generaban conflicto de tipos en `_items` y métodos auxiliares no declarados.
- **Acción Aplicada:** Consolidar clase interna `Item` (`Obj`, `Text`, `Value`), `_checks` (`HashSet<int>`), `_raiseEvents`, `IndiceUnico()`, `RefrescarTexto()` y propiedad `SelectedItem`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Controls/CheckedComboBox.cs`.

#### 3. Fase 1350 — Ubicación de Etiquetas sin Solapamiento
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Controles y etiquetas preexistentes (`lblPeriodo`, `lblPrograma`, `lblSede`) colisionaban en `x=15`, `x=225`, `x=575` con la nueva distribución.
- **Acción Aplicada:** Reubicar explícitamente etiquetas (`lblTipoPrograma` 15,7; `lblPeriodo` 190,7; `lblPrograma` 455,7; `lblSede` 15,39) y controles en 2 filas limpias.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`.

#### 4. Fase 1351 — Declaración de Controles en Designer
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Controles nuevos referenciados en layout carecían de declaración de campo e instanciación en `InitializeComponent`.
- **Acción Aplicada:** Declarar campos privados, instanciar con sus modos (`Unico`/`Multiple`), registrarlos en `pnlFiltros.Controls.Add` y remover combos antiguos.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs:120-130, 431-437`.

#### 5. Fase 1352 — Reconciliación de Persistencia de Período Único
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `chkPeriodos.Modo = Unico` provocaba discrepancia con persistencia multi-fila `PERIODO`, perdiendo selecciones en cada guardado.
- **Acción Aplicada:** Restringir guardado a 1 registro `PERIODO` vía `chkPeriodos.Primario<Periodo>()` y restaurar con `FirstOrDefault()`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 6. Fase 1353 — Accesores de Programa para Call Sites Sensibles
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** 6 call sites leían `SelectedItem` arbitrario al ser `chkProgramas` multiselección, comprometiendo contexto y corrida AG.
- **Acción Aplicada:** Introducir helpers `ProgramaPrimario()` y `ProgramasSeleccionados()`, direccionando cada call site según su semántica real.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:568-570, 1648-1649, 2087, 2908-2909, 3908`.

---

### Verificación de Principios (Ronda 241)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1348 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1349 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1350 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1351 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1352 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1353 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 🔧 Ronda 242: Parches 1348–1353 Registrados Sin Código y 6 Defectos Residuales (Fases 1354 a 1359)

### Diagnóstico y Contexto
- **Régimen de Estados:** Todo el backlog (Fases 1311 a 1359) permanece en `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Consumo efectivo de selecciones en `PoblarArbolContextoAcademico` mediante `FiltrarPeriodos`, `FiltrarSedes` y `FiltrarProgramas`.
  - Unificación de autoridad en la selección de nodos nivel 6 del árbol (combos actúan estrictamente como pre-filtros).
  - Definición explícita de `private string SubunidadActual() => SyncContext.ActiveSubunidad ?? string.Empty;`.
  - Superficie interactiva de `CheckedComboBox` con botón desplegable, `CheckedListBox` alojado en `ToolStripDropDown`.
  - Incorporación de pruebas unitarias SDK-style sin dependencias de compilación manual.
  - Dimensionamiento fijo (280×18) y `AutoEllipsis = true` con `ToolTip` en etiquetas dinámicas del panel de filtros.

### Evaluación Crítica de las Fases

#### 1. Fase 1354 — Consumo de Selecciones en Población de Árbol
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `PoblarArbolContextoAcademico` (L2822-2845) cargaba directamente todos los catálogos activos, dejando inoperantes los combos y el botón `btnFiltrarArbol`.
- **Acción Aplicada:** Incorporar helpers `FiltrarPeriodos`, `FiltrarSedes` y `FiltrarProgramas` que consultan `CheckedComboBox` para construir el árbol.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2822-2845`.

#### 2. Fase 1355 — Autoridad Única en Selección de Árbol
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Generar doble fuente de verdad entre `ProgramasSeleccionados()` (combo) y `progsSel` (nodos nivel 6 del árbol) producía divergencia en la construcción de contexto.
- **Acción Aplicada:** Mantener autoridad absoluta en el árbol para `ConstruirContextoAsync` y limitar combos a rol de pre-filtro de visualización.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:679-682, 822`.

#### 3. Fase 1356 — Declaración de `SubunidadActual()`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** El método `SubunidadActual()` referenciado en snippets de persistencia no existía en el proyecto, produciendo CS0103.
- **Acción Aplicada:** Declarar `private string SubunidadActual() => SyncContext.ActiveSubunidad ?? string.Empty;`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 4. Fase 1357 — Superficie Interactiva en `CheckedComboBox`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** La especificación de F1349 carecía de controles visuales interactivos para desplegar y alternar checks.
- **Acción Aplicada:** Incorporar `_boton` (`Button`), `_lista` (`CheckedListBox`), `_host` (`ToolStripControlHost`) y `_drop` (`ToolStripDropDown`).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Controls/CheckedComboBox.cs`.

#### 5. Fase 1358 — Pruebas Unitarias para Lógica Pura
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Ausencia de arnés de verificación automatizada para componentes no-UI introducidos.
- **Acción Aplicada:** Declarar suites unitarias en proyecto SDK-style (`ConversorPDF.Tests`): `AgTerceroFiltroTests`, `AgFiltroSeleccionRepositoryTests` y `CheckedComboBoxTests`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/*.cs`.

#### 6. Fase 1359 — Dimensionamiento Fijo y AutoEllipsis en Etiquetas
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Etiquetas dinámicas con `AutoSize = true` en `x=500` y `x=800` desbordaban el ancho del panel (1086 px).
- **Acción Aplicada:** Fijar `AutoSize = false`, `AutoEllipsis = true`, `Size = (280, 18)` y tooltip explicativo.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs:486-488`.

---

### Verificación de Principios (Ronda 242)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1354 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1355 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1356 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1357 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1358 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1359 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 🔧 Ronda 243: Fases 1354–1359 Solo Documentadas — 5 Defectos Residuales (Fases 1360 a 1364)

### Diagnóstico y Contexto
- **Régimen de Estados:** Todo el backlog (Fases 1311 a 1364) permanece en `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Preservación de combinaciones marcadas (`Checked = true` en nivel 6) mediante `RecolectarClavesMarcadas` y `RestaurarMarcas` tras repoblar el árbol, con `treeContextoAcademico.EndUpdate()` en bloque `finally`.
  - Prevención de recursión infinita en `CheckedComboBox` mediante bandera `_sincronizandoLista` en el evento `ItemCheck`.
  - Corrección de `AutoClose = true` y cierre accesible mediante `Keys.Enter` y `Keys.Escape` en el dropdown.
  - Consumo efectivo de `cmbTipoPrograma` en `FiltrarProgramas` delegando en `AgNivelAcademicoCatalogo.Desde(p.Modalidad)`.
  - Sustitución de `_toolTip` por el componente real del Designer `toolTipConteo`.

### Evaluación Crítica de las Fases

#### 1. Fase 1360 — Preservación de Combinaciones en Repoblado de Árbol
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `PoblarArbolContextoAcademico` ejecuta `Nodes.Clear()` y no restaura marcas; al filtrar, `ObtenerCombinacionesSeleccionadas` devuelve 0 y el contexto degrada en silencio a 1 programa arbitrario.
- **Acción Aplicada:** Recolectar claves de nivel 6 antes del repoblado con `RecolectarClavesMarcadas`, envolver en `try/finally` para `EndUpdate()`, restaurar marcas con `RestaurarMarcas` y advertir si el filtro descarta todo.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2822-2938, 678-682, 750-751`.

#### 2. Fase 1361 — Guarda Anti-Recursión en `ItemCheck`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `SincronizarLista` dentro de `ItemCheck` llamaba `Items.Add(..., bool)`, lo cual re-disparaba `ItemCheck`, provocando recursión infinita o excepción por colección modificada.
- **Acción Aplicada:** Introducir bandera `_sincronizandoLista` en `SincronizarLista` e ignorar eventos en `OnItemCheck` mientras la bandera esté activa.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Controls/CheckedComboBox.cs`.

#### 3. Fase 1362 — Autocierre Funcional de Desplegable
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `AutoClose = false` dejaba el desplegable flotante e inamovible tras seleccionar, impidiendo interactuar con el resto de la interfaz.
- **Acción Aplicada:** Configurar `_drop.AutoClose = true`, refrescar texto al evento `Closed` y soportar cierre mediante `Keys.Enter`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Controls/CheckedComboBox.cs`.

#### 4. Fase 1363 — Consumo de `cmbTipoPrograma` en Filtrado
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `cmbTipoPrograma` no era consumido en ningún flujo de filtrado, dejando inoperante la distinción entre Pregrado y Posgrado.
- **Acción Aplicada:** En `FiltrarProgramas`, discriminar por `AgNivelAcademicoCatalogo.Desde(p.Modalidad)` según el índice seleccionado (1: Pregrado, 2: Posgrado).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 5. Fase 1364 — Identificador Canónico `toolTipConteo`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Referencia errónea a `_toolTip` producía error de compilación CS0103.
- **Acción Aplicada:** Utilizar el componente existente en el Designer `toolTipConteo`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs:80, 1268`, `AgFormGeneradorHorarios.cs`.

---

### Verificación de Principios (Ronda 243)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1360 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1361 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1362 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1363 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1364 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 🔧 Ronda 244: Fases 1360–1364 No Aplicadas — Restauración de Selección Inválida (Fases 1365 a 1368)

### Diagnóstico y Contexto
- **Régimen de Estados:** Todo el backlog (Fases 1311 a 1368) permanece en `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Sustitución de `RecolectarClavesMarcadas`/`RestaurarMarcas` por `RecolectarRutasMarcadas` y `RestaurarRutas`, materializando recursivamente nodos lazy vía `CargarHijosLazy(actual)`.
  - Clave compuesta jerárquica por ruta (`Nivel:Id:Codigo/...`) para evitar colisiones en niveles 4 a 6 donde `Id = 0`.
  - Supresión de cascada `AfterCheck` durante la restauración mediante `_isUpdatingTreeChecks = true` en bloque `try/finally`.
  - Preservación de ramas expandidas mediante `RecolectarExpandidos` y `RestaurarExpandidos` para evitar el colapso del árbol al filtrar.

### Evaluación Crítica de las Fases

#### 1. Fase 1365 — Restauración Jerárquica con Materialización Lazy
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Los niveles 4 a 6 del árbol son cargados bajo demanda (`CargarHijosLazy`); un barrido eager sobre un árbol recién poblado no encuentra nodos de nivel 6 (`restauradas = 0`).
- **Acción Aplicada:** Registrar snapshot de rutas (`Nivel:Id:Codigo/...`) con `RecolectarRutasMarcadas` y materializar ramas intermedias mediante `CargarHijosLazy(actual)` en `RestaurarRutas`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2822-2938, 258, 2951-2969`.

#### 2. Fase 1366 — Clave Compuesta Jerárquica para Desambiguar Ramas
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** En niveles 4 a 6, `Id = 0`; comparar solo `(Id, Codigo)` causa colisiones masivas entre grupos y modalidades homónimas de diferentes programas.
- **Acción Aplicada:** Utilizar la ruta jerárquica completa y resolver nodos de manera contextual dentro de `actual.Nodes` vía `BuscarNodo`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2965`, `AgNodoArbolKey.cs:12-24`.

#### 3. Fase 1367 — Supresión de Eventos `AfterCheck` en Restauración
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Asignar `node.Checked = true` dispara `TreeContextoAcademico_AfterCheck`, ejecutando `AplicarExclusividadModalidad` y `PropagarCheckHaciaAbajo` que alteran la modalidad seleccionada.
- **Acción Aplicada:** Envolver `RestaurarRutas` bajo la guarda `_isUpdatingTreeChecks = true` con restauración en bloque `finally`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:3354-3385, 3387-3400`.

#### 4. Fase 1368 — Preservación de Expansión y Vista del Árbol
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `Nodes.Clear()` colapsa todas las ramas abiertas obligando al usuario a re-expandir manualmente el árbol tras cada filtrado.
- **Acción Aplicada:** Implementar `RecolectarExpandidos` y `RestaurarExpandidos` preservando el estado visual y la navegación del usuario.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2828-2938`.

---

### Verificación de Principios (Ronda 244)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1365 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1366 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1367 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1368 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.203 Ronda 245 — Fases 1365–1368 No Aplicadas — Restauración Incompleta y Expansión Masiva (Fases 1369 a 1372)

### Evaluación Crítica de las Fases

#### 1. Fase 1369 — Restauración Parcial Silenciosa y Reporte de Fallidas
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `CargarHijosLazy` genera semestres fijos `I..X` y grupos `A..D`. Snapshots con grupos `"E"`, `"ÚNICO"` o semestres superiores fallan silenciosamente (`BuscarNodo` devuelve `null`), reduciendo el conteo de combinaciones sin aviso al usuario.
- **Acción Aplicada:** Agregar parámetro de salida `out int fallidas` en `RestaurarRutas`, contabilizar rutas inalcanzables y emitir advertencia con `LoggerHelper.WarningLimitado`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2973-2990`.

#### 2. Fase 1370 — Recálculo Ascendente de Ancestros y Glifos Post-Restauración
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Con `_isUpdatingTreeChecks = true`, `ActualizarEstadoPadre` y `RefrescarEstadoNodo` no se disparan para ancestros. El nivel 6 queda `Checked`, pero padres quedan desmarcados con glifos vacíos, provocando que cualquier interacción posterior borre las modalidades restauradas.
- **Acción Aplicada:** Incorporar `RefrescarCadenaRestaurada` ejecutando `RefrescarEstadoNodo` y `ActualizarEstadoPadre` de abajo hacia arriba antes de `EndUpdate()`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:3354-3385, 3435-3450, 3079-3083`.

#### 3. Fase 1371 — Predicado de Expansión Real `!n.IsExpanded`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Usar `Nodes.Count > 0` como predicado clasifica como expandidos a todos los nodos que contienen el hijo dummy `"Cargando..."`, expandiendo masivamente el árbol completo (~10.000 nodos) y congelando la UI.
- **Acción Aplicada:** Filtrar en `RecolectarExpandidos` evaluando estrictamente `if (!n.IsExpanded) continue;` y recorrer la ruta con carga lazy controlada en `RestaurarExpandidos`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:258, 2951-2969`.

#### 4. Fase 1372 — Límite de Split, `TryParse` y Saneamiento de Separador en Rutas
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Nodos cuyo `Codigo` contiene `:` (ej. ciudades o descripciones de programas) fragmentan el arreglo en más de 3 partes, truncando la búsqueda; `int.Parse` sin captura puede arrojar `FormatException` con el árbol bloqueado en `BeginUpdate`.
- **Acción Aplicada:** Utilizar `Split(new[] { ':' }, 3)`, validar niveles e IDs con `int.TryParse`, y sanear bidireccionalmente el caracter `:` en serialización y búsqueda.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2866, 2884`.

---

### Verificación de Principios (Ronda 245)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1369 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1370 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1371 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1372 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.204 Ronda 246 — Fases 1369–1372 No Aplicadas — Dimensiones de Decisión Fabricadas en el Árbol (Fases 1373 a 1375)

### Evaluación Crítica de las Fases

#### 1. Fase 1373 — Supresión de Dimensiones Fabricadas en `CargarHijosLazy`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `CargarHijosLazy` inicializa ciegamente todo nodo de nivel 6 con `MANANA/4/MAR-VIE`, anulando las validaciones I1 e I6 en `AgSeleccionMapper` y produciendo genes físicos para modalidades incompatibles (ej. virtuales).
- **Acción Aplicada:** Resolver dimensiones exclusivamente desde selección persistida (`AgNodoArbolKeyMapper.DesdeDto`) o contexto resuelto; si no hay fuente, asignar vacíos/0 para gatillar `I1_JORNADA_VACIA` / `I6_HORAS_DIA_INVALIDO` de forma accionable.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2996-3011`, `AgSeleccionMapper.cs:30-54`.

#### 2. Fase 1374 — Snapshot Completo con DTO de Contexto y Restauración de Dimensiones
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Guardar únicamente rutas jerárquicas descarta `Jornada`, `HorasDia`, `DiasVentana` y `AulaRef`. La restauración lazy reinyecta valores dummy y corrompe las selecciones del usuario tras un filtrado.
- **Acción Aplicada:** Snapshot basado en `List<AgContextoSeleccionDto>` obtenido de `ObtenerCombinacionesSeleccionadas()`; al recorrer la ruta, invocar `AgNodoArbolKeyMapper.Aplicar(dto, k6)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:3455-3490`, `AgNodoArbolKeyMapper.cs:39-61`.

#### 3. Fase 1375 — Soporte de Dimensiones en Migración V78 (`ag_filtro_seleccion`)
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** La tabla `ag_filtro_seleccion` de la spec F1333 solo almacena `(subunidad, tipo, id_ref)`, impidiendo guardar y restaurar `Jornada`, `HorasDia`, `DiasVentana` y `AulaRef` en frío.
- **Acción Aplicada:** Extender el DDL de `MigrationV78_AddAgFiltroSeleccion` incorporando `codigo`, `aula_ref`, `jornada`, `horas_dia`, `dias_ventana`, `orden` e índice único con clave compuesta completa.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV78_AddAgFiltroSeleccion.cs`.

---

### Verificación de Principios (Ronda 246)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1373 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1374 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1375 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.205 Ronda 247 — Fases 1373–1375 No Aplicadas — API Inexistente y Regresión que Bloquea la Generación (Fases 1376 a 1378)

### Evaluación Crítica de las Fases

#### 1. Fase 1376 — Corrección de Firma y Nombre de Método en `AgNodoArbolKeyMapper`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** La spec F1374 invocaba `AgNodoArbolKeyMapper.Aplicar(dto, k6)`. El método en disco es `AplicarDimensiones(key, dto)` con retorno `Result` y orden de parámetros invertido, causando `CS0117` y `CS1503`.
- **Acción Aplicada:** Corregir llamada a `AgNodoArbolKeyMapper.AplicarDimensiones(k6, dto)` agregando la directiva de namespace `ConversorPDF.XLL.Academic.Application.Mappers`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 2. Fase 1377 — Exclusión de Nodos ante Fallo de Dimensiones
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `AplicarDimensiones` aplica fallback `MANANA/4/MAR-VIE` y retorna `Result.Failure`. Ignorar el resultado marcaría el nodo con constantes fabricadas, reintroduciendo el vicio de F1373.
- **Acción Aplicada:** Verificar `aplicado.IsFailure`; en caso de fallo, registrar advertencia con `LoggerHelper.WarningLimitado`, incrementar `fallidas++` y omitir el check del nodo.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 3. Fase 1378 — Resolución de Dimensiones desde Catálogo de Aulas (`ag_aula`)
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Al blanquear las dimensiones en `CargarHijosLazy` sin controles UI, todas las combinaciones fallaban en `AgSeleccionMapper.DesdeNodo` con `I1_JORNADA_VACIA` / `I6_HORAS_DIA_INVALIDO`, bloqueando la generación pese a que `ag_aula` almacena las dimensiones requeridas.
- **Acción Aplicada:** Indexar aulas en `_aulasPorRef` durante la carga y en `CargarHijosLazy` resolver las dimensiones desde `aula.Jornada`, `aula.HorasDia` y `aula.DiasVentana` del aula correspondiente a `AgAulaFactory.ConstruirAulaRef`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2996-3011, 686-696`.

---

### Verificación de Principios (Ronda 247)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1376 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1377 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1378 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.206 Ronda 248 — Fases 1376–1378 No Aplicadas — Caché Vacía y Variables Inexistentes (Fases 1379 a 1381)

### Evaluación Crítica de las Fases

#### 1. Fase 1379 — Precarga de Aulas para Resolución de Dimensiones en Árbol
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `_aulasPorRef` solo se poblaba al generar horarios (`ConstruirContextoAsync`), mientras que el árbol se construye al abrir el formulario (`CargarCatalogosAsync`). Al expandir niveles, el diccionario estaba vacío y las combinaciones quedaban con dimensiones vacías.
- **Acción Aplicada:** Incorporar `CargarAulasParaArbolAsync()` ejecutándose antes de `PoblarArbolContextoAcademico()` en `CargarCatalogosAsync` y tras cada ejecución exitosa de "Generar Aulas".
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:393-476, 686-696, 2951-3011`.

#### 2. Fase 1380 — Reconstrucción de Contexto Ancestral `RutaHasta` en Carga Lazy
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** F1378 invocaba `AgAulaFactory.ConstruirAulaRef` utilizando `idPeriodo`, `idSede`, `idPrograma` y `semestre`, variables no declaradas en el ámbito local de `CargarHijosLazy` donde solo se dispone de `node` y `key` (nivel 5), generando 4 errores `CS0103`.
- **Acción Aplicada:** Crear método auxiliar `RutaHasta(node)` que remonta los nodos padre recopilando `(IdPeriodo, IdSede, IdPrograma, Semestre, Grupo)` antes de construir `AulaRef`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2951-3011`.

#### 3. Fase 1381 — Índice Complementario Normalizado por Modalidad (`_aulasPorClaveNorm`)
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Búsqueda textual por `AulaRef` exacto falla ante registros históricos donde la modalidad aún figura como `"VIRTUAL"` en vez de `"PLATAFORMA"` (previo a migración V66), dejando combinaciones válidas sin dimensiones.
- **Acción Aplicada:** Mantener índice `_aulasPorClaveNorm` con modalidad normalizada vía `ModalidadTutoriaCatalog.NormalizarTutoria` como fallback si el `AulaRef` exacto no coincide.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

---

### Verificación de Principios (Ronda 248)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1379 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1380 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1381 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.207 Ronda 249 — Fases 1379–1381 No Aplicadas — `DataSource` Incompatible y Estado Inicial de Filtros (Fases 1382 a 1385)

### Evaluación Crítica de las Fases

#### 1. Fase 1382 — Reemplazo de Asignaciones a `DataSource` en `CheckedComboBox`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `CheckedComboBox` no implementa `DataSource`, `DisplayMember` ni `ValueMember`. El código actual en `CargarCatalogosAsync` asigna estas tres propiedades a `cmbPeriodos`, `cmbProgramas` y `cmbSedes`, causando 9 errores de compilación `CS1061`.
- **Acción Aplicada:** Reemplazar las asignaciones por `cmbPeriodos.Cargar(_periodos, p => p.Nombre, p => p.Id)`, `cmbProgramas.Cargar(_programas, p => p.Nombre, p => p.Id)` y `cmbSedes.Cargar(_sedes, s => s.Nombre, s => s.Id)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:459-469`.

#### 2. Fase 1383 — Corrección de Selección Inicial en Filtros Múltiples ("Vacío = Todos")
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `cmbProgramas.SelectedIndex = 0` y `cmbSedes.SelectedIndex = 0` seleccionan forzosamente el primer elemento de cada catálogo, restringiendo el árbol a un único programa y sede por defecto y violando la regla de "vacío = todos".
- **Acción Aplicada:** Mantener `SelectedIndex = 0` exclusivamente en `cmbPeriodos` (modo único) y retirar la selección forzada en `cmbProgramas` y `cmbSedes`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:471-473`.

#### 3. Fase 1384 — Invalidación de Caché de Aulas ante Cambio de Contexto
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Al cambiar de período o subunidad, `_aulasPorRef` y `_aulasPorClaveNorm` mantenían las aulas del período previo, asociando dimensiones obsoletas o erróneas a las combinaciones del nuevo período.
- **Acción Aplicada:** Crear `InvalidarAulasCache()` e invocarlo en `OnVentanaChanged`, al cambiar de perfil de árbol y tras la generación de aulas.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:380-391, 477`.

#### 4. Fase 1385 — Granularidad y Precarga Acotada de `ag_aula`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Consultar todas las aulas del período sin filtro de sede genera latencia en períodos extensos.
- **Acción Aplicada:** Acotar la precarga inicial a la sede activa del período seleccionado y delegar la carga integral a `ConstruirContextoAsync`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

---

### Verificación de Principios (Ronda 249)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1382 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1383 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1384 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1385 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.208 Ronda 250 — Fases 1382–1385 No Aplicadas — Designer Sin Migrar y Selección Múltiple Colapsada (Fases 1386 a 1389)

### Evaluación Crítica de las Fases

#### 1. Fase 1386 — Migración de Tipos y Propiedades en `AgFormGeneradorHorarios.Designer.cs`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** El archivo `Designer.cs` continúa instanciando y declarando `cmbPeriodos`, `cmbProgramas` y `cmbSedes` como `System.Windows.Forms.ComboBox` y asignando `DropDownStyle` y `FormattingEnabled`, propiedades inexistentes en `CheckedComboBox`, impidiendo el uso de su API y generando 6 errores `CS1061`.
- **Acción Aplicada:** Declarar e instanciar los controles como `CheckedComboBox`, remover `DropDownStyle` y `FormattingEnabled`, y configurar sus modos de selección (`ModoSeleccion.Unico` / `Multiple`).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs:59-63, 456-458, 476-478, 496-502, 1232-1236`.

#### 2. Fase 1387 — Soporte de Marcado Programático en Restauración de Perfiles (`Marcar`/`SeleccionarTodos`)
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** La restauración de perfiles realizaba asignaciones a `SelectedItem` (`cmbPeriodos.SelectedItem = per`), propiedad sin setter en `CheckedComboBox`, causando 3 errores de compilación `CS0200/CS1061` y colapsando selecciones múltiples a un único ítem.
- **Acción Aplicada:** Agregar métodos `Marcar(object)` y `SeleccionarTodos(IEnumerable<int>)` en `CheckedComboBox` y utilizarlos en la restauración de perfil.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:3904, 3917, 3925`, `CheckedComboBox.cs`.

#### 3. Fase 1388 — Adaptación de 16 Call Sites de `SelectedItem` a `Seleccionados<T>`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** 16 puntos de consumo en el formulario leían `SelectedItem as Programa` o `SelectedItem as Sede`, ignorando la multiselección activa y forzando una selección singleton.
- **Acción Aplicada:** Migrar los puntos de lectura al patrón `cmbProgramas.Seleccionados<Programa>()` y `cmbSedes.Seleccionados<Sede>()`, aplicando la convención "vacío = todos" si el conteo es cero.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:568-570, 1647-1649, 2086-2087, 2908-2909, 3521, 3694, 3790, 4271`.

#### 4. Fase 1389 — Evento de Compatibilidad `SelectedIndexChanged` en `CheckedComboBox`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** El formulario suscribe `cmbPeriodos.SelectedIndexChanged` para actualizar la ventana semestral, pero `CheckedComboBox` no exponía dicho evento, causando `CS1061`.
- **Acción Aplicada:** Exponer el evento `SelectedIndexChanged` mapeado internamente a `SeleccionChanged`, disparándolo tras cualquier alteración de marcas.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:233`, `CheckedComboBox.cs`.

---

### Verificación de Principios (Ronda 250)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1386 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1387 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1388 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1389 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.209 Ronda 251 — Fases 1386–1389 No Aplicadas — Registro de Migración, Idempotencia V78 y Duplicación de `FormTerceros` (Fases 1390 a 1392)

### Evaluación Crítica de las Fases

#### 1. Fase 1390 — Registro de Migración V78 en Ruta Real y Sintaxis Válida
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** La entrada V77 en `SqliteHelper.cs` (L378) no tiene coma final. Insertar V78 sin añadir dicha coma produce el error `CS1002`. Además, el backlog citaba una ruta inexistente (`Academic\Infrastructure\Persistence\SqliteHelper.cs`) en lugar de `Core\Helpers\SqliteHelper.cs`.
- **Acción Aplicada:** Registrar `MigrationV78_AddAgFiltroSeleccion` en `ConversorPDF.XLL\Core\Helpers\SqliteHelper.cs:378-380` con coma de separación tras V77.
- **Archivos Afectados:** `ConversorPDF.XLL\Core\Helpers\SqliteHelper.cs`.

#### 2. Fase 1391 — Idempotencia en V78 con Guard `PRAGMA table_info` (`IMigrationTolerante`)
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `ALTER TABLE ... ADD COLUMN` sin verificación previa lanza excepción por columna duplicada si la base de datos ya fue migrada parcialmente, bloqueando el inicio completo del Add-In en `EnsureSchemaUpToDate`.
- **Acción Aplicada:** Implementar `IMigration, IMigrationTolerante` en `MigrationV78_AddAgFiltroSeleccion` y emplear `PRAGMA table_info` antes de agregar columnas adicionales.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\Infrastructure\Persistence\Migrations\MigrationV78_AddAgFiltroSeleccion.cs`.

#### 3. Fase 1392 — Reutilización de `FormTerceros` Existente para Evitar Duplicidad
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** En `Academic\UI\Forms\FormTerceros.cs` ya existe un formulario funcional de 850 líneas con repositorios y sincronización de terceros enlazado a dos formularios activos. Crear `AgFormTerceros` introduce dos fuentes de verdad para `PoolDocenteKeys` y divergencia de validaciones.
- **Acción Aplicada:** Reutilizar `FormTerceros.cs` integrando `AgFiltroSeleccionRepository` para canalizar el filtro de selección y descartar la creación de `AgFormTerceros`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.cs`, `FormVerVinculaciones.cs`, `FormVinculaciones.cs`.

---

### Verificación de Principios (Ronda 251)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1390 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1391 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1392 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.210 Ronda 252 — Fases 1390–1392 No Aplicadas — Guard de Tabla Ausente y Clave de Filtro No Disponible (Fases 1393 a 1394)

### Evaluación Crítica de las Fases

#### 1. Fase 1393 — Guard Previo de Tabla en `AgregarColumnaSiNoExiste` (Paridad V77)
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `PRAGMA table_info` sobre una tabla no creada retorna 0 filas, interpretándose erróneamente como columna ausente e intentando ejecutar `ALTER TABLE` sobre una tabla inexistente, arrojando `SQLiteException: no such table` y bloqueando el arranque del Add-In.
- **Acción Aplicada:** Incorporar comprobación previa en `sqlite_master` antes de inspeccionar columnas y renombrar a `AgregarColumnaSiNoExiste` conforme a `MigrationV77`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\Infrastructure\Persistence\Migrations\MigrationV78_AddAgFiltroSeleccion.cs`.

#### 2. Fase 1394 — Contexto de Filtro Desacoplado y Exposición de `IdSede` en `SesionListItem`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `SesionListItem` no exponía `IdSede` y `FormTerceros` carecía de constructor para recibir dimensiones de filtrado. El llamado propuesto en F1392 introducía 6 errores `CS0103` y dejaba el filtro sin parámetros válidos.
- **Acción Aplicada:** Agregar `IdSede` en `SesionListItem.cs` y crear `AgContextoFiltro` en `FormTerceros.cs` como parámetro opcional del constructor.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\Domain\Models\SesionListItem.cs`, `Academic\UI\Forms\FormTerceros.cs`, `FormVerVinculaciones.cs:415-417`.

---

### Verificación de Principios (Ronda 252)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1393 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1394 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.211 Ronda 253 — Cierre de Brechas A1–A11 y Escenarios Operativos (Fases 1395 a 1401)

### Evaluación Crítica de las Fases

#### 1. Fase 1395 — Declaración e Inicialización de `cmbTipoPrograma` y Reflow de `pnlFiltros` a 66px
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `cmbTipoPrograma` no existía en `AgFormGeneradorHorarios.Designer.cs` y `pnlFiltros` continuaba en 46px (1 sola fila), impidiendo alojar el selector de nivel académico y causando colisión con las demás etiquetas.
- **Acción Aplicada:** Declarar e inicializar `cmbTipoPrograma` en `Designer.cs`, configurar layout de dos filas en `pnlFiltros` a 66px de altura y posicionar las etiquetas de acuerdo a la grilla definida.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs`.

#### 2. Fase 1396 — Integración de Filtrado de Programas por Tipo Académico
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** La selección en `cmbTipoPrograma` no filtraba la lista visible en `cmbProgramas` ni las ramas del árbol de contexto académico.
- **Acción Aplicada:** Conectar evento `SelectedIndexChanged` de `cmbTipoPrograma` para filtrar `_programas` mediante `AgNivelAcademicoCatalogo.Desde(p.Modalidad)` y repoblar el árbol preservando las combinaciones activas.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\AgFormGeneradorHorarios.cs:2822-2845, 459-473`.

#### 3. Fase 1397 — Deducción de Horas en Nivel 6 desde Modalidad del Programa
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `CargarHijosLazy` asignaba de forma estática `HorasDia = 4` a todas las modalidades de posgrado y pregrado indiscriminadamente.
- **Acción Aplicada:** Resolver las horas por día consultando la modalidad del programa ancestro (Pregrado = 4h, Posgrado = 8h) antes de registrar la clave de nivel 6.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\AgFormGeneradorHorarios.cs:2998-3003`.

#### 4. Fase 1398 — Advertencia No Modal ante Selección de 8h en Pregrado
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** No existía aviso interactivo si un usuario seleccionaba una modalidad de 8 horas dentro de una cohorte de pregrado, arriesgando bloqueos tardíos en la fase de pre-vuelo 1114.
- **Acción Aplicada:** Incorporar advertencia informativa no modal en `TreeContextoAcademico_AfterCheck` preservando inalterados los algoritmos de `AgUnidadesSemanales` y pre-vuelo.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\AgFormGeneradorHorarios.cs:3354`.

#### 5. Fase 1399 — Acciones Masivas y Búsqueda sobre Vista Filtrada en `FormTerceros`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Ausencia de botones masivos (`btnSeleccionarTodos`, `btnBorrarSeleccion`, `btnBuscar`, `chkSoloHabilitadas`) operando sobre la vista filtrada en `FormTerceros`.
- **Acción Aplicada:** Incorporar barra de herramientas y controles en `FormTerceros.Designer.cs` implementando los manejadores sobre la lista visible `_tercerosVista`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.cs`, `FormTerceros.Designer.cs`.

#### 6. Fase 1400 — Unicidad por Subunidad y Sentencia Upsert en V78
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** La falta de índice único compuesto con subunidad provocaba registros duplicados en `ag_filtro_seleccion` ante guardados sucesivos.
- **Acción Aplicada:** Crear índice único sobre `(idperiodo, idsede, idprograma, semestre, grupo, modalidad, IFNULL(subunidad,''))` e implementar persistencia con `INSERT ... ON CONFLICT DO UPDATE`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\Infrastructure\Persistence\Migrations\MigrationV78_AddAgFiltroSeleccion.cs`.

#### 7. Fase 1401 — Pre-marcado Automático en `FormTerceros` y Flujo Previo a `ConstruirIndices`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Al abrir `FormTerceros` los docentes previamente guardados no aparecían marcados, y en el generador de horarios `AplicarAlcanceTerceros` no se ejecutaba antes de `ConstruirIndices` (L1005).
- **Acción Aplicada:** Pre-marcar los terceros al cargar el formulario según la selección persistida y enlazar la invocación de `AplicarAlcanceTerceros` en `AgFormGeneradorHorarios.cs:1005`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.cs`, `AgFormGeneradorHorarios.cs:1005`.

---

### Verificación de Principios (Ronda 253)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1395 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1396 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1397 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1398 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1399 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1400 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1401 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.212 Ronda 254 — Corrección de Dominio y Retractación (Fases 1402 a 1407)

### Evaluación Crítica de las Fases

#### 1. Fase 1402 — Filtrado de Pool de Docentes Previo a Construcción de Índices
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `FiltrarPoolDocentesPorTerceros` debe ejecutarse antes de `ConstruirIndices` en el pipeline genético para asegurar que las restricciones de terceros activos se apliquen al grafo de asignación.
- **Acción Aplicada:** Integrar llamada a `FiltrarPoolDocentesPorTerceros` en `AgContexto.cs` y `AgFormGeneradorHorarios.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgContexto.cs`, `AgFormGeneradorHorarios.cs`.

#### 2. Fase 1403 — Nivel Académico desde Modalidad
- **Estado:** RETIRADA (Falso positivo)
- **Diagnóstico Crítico:** `programas.modalidad` ya almacena `"PREGRADO"` y `"POSGRADO"` en el esquema productivo (`AcademicSchema.sql:23`, vistas canónicas). No se requiere alterar la tabla `programas` ni agregar la columna `nivel_academico`.
- **Acción Aplicada:** Retirada la fase 1403 del backlog de cambios físicos de BD.
- **Archivos Afectados:** Ninguno.

#### 3. Fase 1404 — Reutilización de `txtBuscar` y `cmbTerritorial` en `FormTerceros`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Evitar duplicar controles de filtrado ya existentes en `FormTerceros.Designer.cs`, incorporando únicamente los dos botones masivos faltantes (`btnSeleccionarTodos` y `btnBorrarSeleccion`).
- **Acción Aplicada:** Ajustar `FormTerceros.Designer.cs` reutilizando los componentes de búsqueda existentes.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.Designer.cs`.

#### 4. Fase 1405 — Índice Único Plano y Deduplicación en V78
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** El índice compuesto de `ag_filtro_seleccion` debe ser plano y tolerar subunidades nulas mediante deduplicación previa en memoria antes de indexar.
- **Acción Aplicada:** Configurar índice único plano y sentencias upsert en `MigrationV78_AddAgFiltroSeleccion.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\Infrastructure\Persistence\Migrations\MigrationV78_AddAgFiltroSeleccion.cs`.

#### 5. Fase 1406 — Catálogo `AgTipoProgramaCatalog` para Nivel Académico
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Distinguir conceptualmente el nivel académico (`PREGRADO`/`POSGRADO` en `programas.modalidad`) de la modalidad de tutoría (`Presencial`/`Plataforma`/`Mixta` en `ModalidadTutoriaCatalog`), evitando cruces de validación erróneos en `FiltrarProgramas`.
- **Acción Aplicada:** Crear `AgTipoProgramaCatalog.cs` encapsulando constantes `TODOS`, `PREGRADO`, `POSGRADO` y helper `Nivel(modalidadPrograma)`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\Domain\Models\AgTipoProgramaCatalog.cs`, `AgFormGeneradorHorarios.cs:444-452`.

#### 6. Fase 1407 — Asignación de Horas en Nivel 6 Desacoplada de Tutoría
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `HorasDia` en nivel 6 debe derivarse exclusivamente del nivel académico del programa (Pregrado = 4h, Posgrado = 8h), mientras que la modalidad de tutoría define únicamente el aula y las características de vinculación.
- **Acción Aplicada:** En `CargarHijosLazy`, resolver `HorasPorNivelAcademico(prog?.Modalidad)` y aplicar prioridad: `aula.HorasDia` (si existe y > 0) > nivel académico > 4h por defecto.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\AgFormGeneradorHorarios.cs:2951-3011`.

---

### Verificación de Principios (Ronda 254)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1402 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1403 | [X] N/A (Retirada) | [X] Sí | [X] Sí | — | — |
| F1404 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1405 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1406 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1407 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.213 Ronda 255 — `dgvTerceros` Sin Columna de Selección, DDL de V78 Indefinido y Selección Perdida al Refiltrar (Fases 1408 a 1410)

### Evaluación Crítica de las Fases

#### 1. Fase 1408 — Inserción de `colSeleccion` en `dgvTerceros` y Tracking con `HashSet<int>`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `dgvTerceros` sólo contenía columnas de texto. Intentar leer `row.Cells["colSeleccion"]` arrojaba `ArgumentException`. Además, `Tercero` no poseía campo de selección.
- **Acción Aplicada:** Insertar `DataGridViewCheckBoxColumn` (`colSeleccion`) en `dgvTerceros`, rastrear IDs en `HashSet<int> _seleccionados` mediante `CellValueChanged` e implementar `MarcarVisibles(bool)` para `btnSeleccionarTodos` y `btnBorrarSeleccion`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.cs:51-62`, `FormTerceros.Designer.cs`.

#### 2. Fase 1409 — Definición Completa de DDL para V78 con Soporte de `actualizado`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** La tabla `ag_filtro_seleccion` carecía de sentencia DDL `CREATE TABLE` formal en el plan, provocando fallo de columna inexistente en el upsert `actualizado = datetime('now')`.
- **Acción Aplicada:** Establecer DDL formal con columnas completas (`idperiodo`, `idsede`, `idprograma`, `semestre`, `grupo`, `modalidad`, `subunidad`, `creado`, `actualizado`) e índice único coincidente en `MigrationV78_AddAgFiltroSeleccion.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\Infrastructure\Persistence\Migrations\MigrationV78_AddAgFiltroSeleccion.cs`.

#### 3. Fase 1410 — Preservación de Selección en `AplicarFiltroTerceros` tras Rebind de `_bs`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Al teclear en `txtBuscar` o cambiar territorial, `_bs.DataSource = filtered.ToList()` regenera las filas del grid perdiendo las marcas en pantalla.
- **Acción Aplicada:** Iterar `dgvTerceros.Rows` tras asignar `DataSource` en `AplicarFiltroTerceros` reponiendo `row.Cells["colSeleccion"].Value = _seleccionados.Contains(t.Id)`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.cs:183-206`.

---

### Verificación de Principios (Ronda 255)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1408 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1409 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1410 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.214 Ronda 256 — Auditoría de Cobertura v4, Layout de Filtros y Sincronía de Grilla (Fases 1411 a 1414)

### Evaluación Crítica de las Fases

#### 1. Fase 1411 — Commit Inmediato con `CurrentCellDirtyStateChanged` en `dgvTerceros`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `DataGridViewCheckBoxColumn` retiene la confirmación del valor hasta que el foco abandona la celda o fila. `CellValueChanged` no se disparaba inmediatamente en el primer clic, desfasando el estado de `_seleccionados`.
- **Acción Aplicada:** Manejar `CurrentCellDirtyStateChanged` en `dgvTerceros` ejecutando `CommitEdit(DataGridViewDataErrorContexts.Commit)` para reflejar el cambio al instante.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.cs:51-62`.

#### 2. Fase 1412 — Precarga de `_seleccionInicial` en `_seleccionados`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Al instanciar `FormTerceros`, la primera ejecución de `AplicarFiltroTerceros` desmarcaba todas las casillas porque `_seleccionados` se inicializaba vacío, violando el pre-marcado requerido (A10).
- **Acción Aplicada:** Cargar los IDs persistidos en `_seleccionados` antes del primer llamado a `AplicarFiltroTerceros` bajo el flag de guarda `_cargandoGrid`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.cs:135, 183-206`.

#### 3. Fase 1413 — Refactor de Layout en `pnlFiltros` a 66px Sin Colisiones
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** En `AgFormGeneradorHorarios.Designer.cs`, los controles de filtro ocupaban un único renglón hasta x=870. Las coordenadas propuestas en F1313 pisaban `lblPrograma` (x=208 vs 225) y `cmbSedes` (x=751 vs 620-870) dentro de un panel de 46px.
- **Acción Aplicada:** Establecer `pnlFiltros.Size = (1086, 66)` y organizar los controles en dos filas limpias (y=6 y y=34) respetando márgenes y jerarquía visual.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs:439-501`.

#### 4. Fase 1414 — Advertencia No Modal Mediante `lblResumenFiltros` y `toolTipConteo`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Ubicar un control estático `lblAvisoHorasNivel` en `pnlCabecerasNivel6` (20px) generaba solapamientos con las cabeceras calculadas en resize (`AgFormGeneradorHorarios.cs:3294`).
- **Acción Aplicada:** Eliminar `lblAvisoHorasNivel` del Diseñador y emitir el aviso de pregrado mediante `lblResumenFiltros.Text` y `toolTipConteo`, limpiándolo al seleccionar 4 horas.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs`, `AgFormGeneradorHorarios.cs:3294`.

---

### Verificación de Principios (Ronda 256)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1411 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1412 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1413 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1414 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.215 Ronda 257 — Cascada Geométrica en `grpSeleccion` y Habilitación de Edición en `dgvTerceros` (Fases 1415 a 1416)

### Evaluación Crítica de las Fases

#### 1. Fase 1415 — Cascada de Desplazamiento Vertical en Hijos de `grpSeleccion`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** F1413 omitía que `pnlArbolToolbar` se localiza en y=65 (34px de alto). Al expandir `pnlFiltros` a 66px (bottom en y=85), se ocluían los primeros 20px de la barra de herramientas del árbol.
- **Acción Aplicada:** Desplazar en cascada: `pnlArbolToolbar` a y=85, `pnlCabecerasNivel6` a y=119 y `treeContextoAcademico` a y=139 (alto 146px, bottom 285px dentro de `grpSeleccion.Height = 288`).
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs:287-289, 439-441`.

#### 2. Fase 1416 — Habilitación de Edición en `dgvTerceros` (`ReadOnly = false`)
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `dgvTerceros.ReadOnly = true` en `FormTerceros.Designer.cs:113` anulaba completamente la interacción con `colSeleccion`, impidiendo que el evento `CurrentCellDirtyStateChanged` o el método `CommitEdit` surtieran efecto.
- **Acción Aplicada:** Configurar `dgvTerceros.ReadOnly = false`, marcar cada columna de texto con `col.ReadOnly = true`, ajustar `colSeleccion.ReadOnly = false` y activar `EditMode = EditOnEnter`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.Designer.cs:113`, `FormTerceros.cs:51-62`.

---

### Verificación de Principios (Ronda 257)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1415 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1416 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.216 Ronda 258 — Constructor Extendido, Panel de Selección Inferior y DisplayName (Fases 1417 a 1419)

### Evaluación Crítica de las Fases

#### 1. Fase 1417 — Sobrecarga de Constructor en `FormTerceros` para Inyección de Selección
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** F1412 invocaba `_seleccionInicial` en el cuerpo del formulario, pero dicho campo no estaba declarado y el único constructor disponible era `FormTerceros(Func<IDbConnection>)`, arrojando `CS0103`.
- **Acción Aplicada:** Declarar `_modoSeleccion`, `_seleccionInicial`, `_cargarSeleccion` y sobrecargar el constructor delegando al base para conservar sin cambios las invocaciones desde `FormVerVinculaciones` y `FormVinculaciones`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.cs:25-37`.

#### 2. Fase 1418 — Incorporación de `pnlSeleccion` (Dock Bottom) en `FormTerceros`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Las coordenadas de botones masivos provenientes de F1318 asumían el contenedor de Opción A y no encajaban en `FormTerceros` (donde `pnlTop` ya cuenta con 7 controles y `pnlDetails` ocupa `Dock.Fill`).
- **Acción Aplicada:** Agregar `pnlSeleccion` con `Dock = DockStyle.Bottom` (40px) visible solo cuando `_modoSeleccion = true`, albergando `btnSeleccionarTodos`, `btnBorrarSeleccion`, `btnGuardarSeleccion` y `lblContadorSeleccion`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\UI\Forms\FormTerceros.Designer.cs:152-216, 711`, `FormTerceros.cs`.

#### 3. Fase 1419 — Propiedad Calculada `Tercero.DisplayName` y Enlace a `dgvTerceros`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `dgvTerceros` enlazaba a `Nombre`, omitiendo la residencia requerida por el contrato A9 (`NOMBRE (RESIDENCIA)`). `Tercero` no disponía de propiedad de presentación combinada.
- **Acción Aplicada:** Añadir la propiedad de sólo lectura `DisplayName => string.IsNullOrWhiteSpace(Residencia) ? (Nombre ?? "") : $"{Nombre} ({Residencia})"` en `Tercero.cs` y enlazarla a la primera columna de texto en `SetupGrid`.
- **Archivos Afectados:** `ConversorPDF.XLL\Academic\Domain\Models\Tercero.cs`, `FormTerceros.cs:51-58`.

---

### Verificación de Principios (Ronda 258)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1417 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1418 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1419 | [X] Sí | [X] Sí | [X] Sí | — | — |

---

## 157.217 Plan Consolidado v5 (Handoff) — Filtro de Árbol AG + CheckedComboBox + Selección de Terceros

- **Decisión Estructural Final:** Opción A ratificada (creación de `AgFormTerceros.cs` y `.Designer.cs`; `FormTerceros.cs` permanece inalterado).
- **Régimen de Trazabilidad:** 16 rondas de auditoría (Fases 1311 a 1394 y parches 1402 a 1421) consolidadas en 12 pasos ejecutables (P0 a P11).
- **Gates P0 en Disco Verificados:**
  - `Academic\UI\Controls\`: No existe (`Test-Path = False`, P0 creará el directorio).
  - `SqliteHelper.cs`: Línea 378 es `new MigrationV77_AddPernoctaAgVinculaciones()` sin coma (P4 registrará V78).
  - `AgFormGeneradorHorarios.Designer.cs`: Anclas de `ComboBox` y `Location` confirmadas (P5 aplicará 66px y cascada).
  - `FormTerceros.cs`: Intacto y protegido contra mutaciones.
- **Estado Global:** Consolidado y aceptado para transición a modo de ejecución física.






































---

## Ronda 264: Corregir RenombrarSolicitudes y rango #YY (Fases 1426–1431)

### Hallazgos y Acciones Ejecutadas
- **Fase 1426 (T6)**: Agregado `using System.Linq;` en `ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs`.
- **Fase 1427 (T1)**: Corregido `TryParseLimite` en `ViaticoNumeroComision.cs` para admitir números de años anteriores con longitud >= 5 o sufijo de vigencia sin invertir años pasados.
- **Fase 1428 (T1)**: Sustituido retorno contradictorio en `EnRango` para procesar sin filtro ante límites no interpretables (`return true`).
- **Fase 1429 (T1)**: Desacoplado año de vigencia con `AnioVigenciaPorDefecto` y pruebas dinamizadas contra `App.APP_VIGENCIA`.
- **Fase 1430 (T2/T5)**: Delegado feedback UI al formulario `FormGestionDocumental.BtnUnificar_Click` y retirados `MessageBox` bloqueantes del procesador en capa Core.
- **Fase 1431 (T7/T8)**: Actualizado `docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosGestionProcessor.yml` y sincronizados contadores de `.agents/plans/REGISTRY.md` (Ronda 264 / Fase 1431).

### Verificación de Principios (Ronda 264)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1426 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1427 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1428 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1429 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1430 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1431 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |


---

## Ronda 265: Correcciones Ronda 264 y desacople total de WinForms (Fases 1432–1436)

### Hallazgos y Acciones Ejecutadas
- **Fase 1432 (T2/T5)**: `RenombrarSolicitudes` modificado para retornar `int` (conteo o `-1` en error); `ViaticosGestionProcessor` despojado de `System.Windows.Forms` y `MessageBox`; `FormGestionDocumental` asume feedback visual de éxito, lista vacía y error.
- **Fase 1433 (T5)**: `UnificarSolicitudes` retorna `-1` en excepciones/directorio inexistente; `FormGestionDocumental` y `RibbonController` distinguen error (`-1`) de ausencia de expedientes (`0`).
- **Fase 1434 (T7)**: Dinamizados los valores en asserts de `EnRango` en `ViaticoNumeroComisionTests.cs` mediante `AnioVigenciaActual % 100` y `(AnioVigenciaActual - 1) % 100`.
- **Fase 1435 (T8)**: Actualizado `docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosGestionProcessor.yml` reflejando retorno `System.Int32` en `RenombrarSolicitudes`.
- **Fase 1436 (T8)**: Creado plan `.agents/plans/2026-10-30-gestion-documental-rango-yy-plan.md` y actualizada fila R-002 en `REGISTRY.md` con contadores actualizados (Ronda 265 / Fase 1436).

### Verificación de Principios (Ronda 265)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1432 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1433 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1434 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1435 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1436 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |

---

## Ronda 266: Procesador base, interfaces, servicio y formulario UI (Fases 1437–1438)

### Hallazgos y Acciones Ejecutadas
- **Fase 1437 (T1)**: Creados `ResultadoRenombrado.cs` y `MotivoRenombrado.cs` bajo `ConversorPDF.XLL/Core/Models/` y añadidos a `ConversorPDF.XLL.csproj`. En `ViaticosGestionProcessor.cs`, se implementó `CarpetasDeSolicitudes` para escanear raíz y subcarpeta `Solicitudes`, se restringió `Estandarizar` a `*.pdf`, se añadieron contadores estructurados y se preservó `RenombrarSolicitudes` como fachada sobre `RenombrarSolicitudesConDetalle`.
- **Fase 1438 (T2)**: Declarado `RenombrarSolicitudesConDetalle` en `IViaticosGestionProcessor` e `IViaticosService`, implementado en `ViaticosService` y conectado a `FormGestionDocumental.BtnRenombrar_Click`, sustituyendo el mensaje falso de "sin comisiones" por el reporte fiel con desglose de examinados, renombrados, colisiones y motivos.

### Verificación de Principios (Ronda 266)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1437 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1438 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |

---

## Ronda 267: Deduplicación N0–N3 y FindCC endurecido (Fases 1439–1440)

### Hallazgos y Acciones Ejecutadas
- **Fase 1439 (T3)**: `GetSolicitudes` sanitizado con `Trim()`. `RenombrarFila` migrado a política N0 (agrupación y deduplicación por número y cédula normalizada), N1 (metadatos de PDF), N2 (OCR validado) y N3 (ambigua sin modificar). Creados helpers `IdentificarPorMetadatos`, `IdentificarPorOcr` y `NormalizarDocumento`.
- **Fase 1440 (T4)**: `FindCC` actualizado a firma con `ICollection<string> cedulasCandidatas`, restringiendo las coincidencias al conjunto de cédulas válidas para la comisión y eliminando falsos positivos provocados por rubros presupuestales (ej. `0505023`).

### Verificación de Principios (Ronda 267)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1439 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1440 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |

---

## Ronda 268: Pruebas de regresión, documentación y cierre R-003 (Fases 1441–1442)

### Hallazgos y Acciones Ejecutadas
- **Fase 1441 (T5)**: 7 pruebas unitarias nuevas agregadas a `ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs` validando subcarpeta `Solicitudes`, dedupe idéntico sin OCR, OCR con comisiones compartidas, rechazo de texto ajeno, detección de colisión sin sobrescritura, prefijos no numéricos y control de errores I/O.
- **Fase 1442 (T6)**: Actualizado `docs/articles/gestion-documental.md` con política N0–N3 y tabla de motivos; creada bitácora `docs/articles/changes/2026-10-31-gestion-documental-carpeta-y-desambiguacion.md`; sincronizado `docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosGestionProcessor.yml`; cerrada fila `R-003` en `REGISTRY.md` con contadores actualizados (Ronda 268 / Fase 1442).

### Verificación de Principios (Ronda 268)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1441 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1442 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |

---

## Ronda 269: Revisión de la implementación Ronda 266–268 (Fases 1443–1450)

### Hallazgos y Acciones Ejecutadas
- **Fase 1443**: Columna `verify` en `REGISTRY.md` fila R-003 ajustada a `pendiente (gate Full)` para reflejar que la suite física no fue ejecutada bajo el paso restrictivo de `/ronda-review`.
- **Fase 1444**: Manejador `BtnRenombrar_Click` en `FormGestionDocumental.cs` actualizado con rama `else if (r.YaCorrectos > 0)` para feedback fiel en re-ejecuciones sin cambios.
- **Fase 1445**: En `ViaticosGestionProcessor.cs`, `GetSolicitudes` descarta registros con cédula vacía; en `RenombrarFila` se valida no-nulidad de `CC` y `Nom` antes de renombrar devolviendo `Ambigua` si faltan.
- **Fase 1446**: En `ViaticosGestionProcessor.Estandarizar` se agregaron trazas `LoggerHelper.Debug` ante renombrados y `LoggerHelper.Warning` ante excepciones.
- **Fase 1447**: `r.Detalle` almacena claves compuestas `Path.GetFileName(carpeta) + "/" + Path.GetFileName(file)` para evitar sobreescritura entre raíz y `Solicitudes`.
- **Fase 1448**: Regex de `FindCC` ampliado con `(?<!\d)\d[\d\.,\s]{4,18}\d(?!\d)` filtrando longitudes entre 6 y 11 dígitos, tolerando comas de miles y evitando falsos ambiguos.
- **Fase 1449**: Unificados encabezados `## Ronda` en `implementation_plan.md` y agregado `RenombrarSolicitudesConDetalle` en `docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosService.yml`.
- **Fase 1450**: §10 de `.agents/plans/2026-10-31-gestion-documental-carpeta-solicitudes-fix-plan.md` actualizado consagrando la invariante de balance como criterio rector de aceptación.

### Verificación de Principios (Ronda 269)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1443 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1444 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1445 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1446 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1447 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1448 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1449 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1450 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |

---

## Ronda 270: Verificación de los parches 1443–1450 (regresiones y cierres) (Fases 1451–1456)

### Hallazgos y Acciones Ejecutadas
- **Fase 1451**: `REGISTRY.md` fila `R-003` actualizado de `closed` a `executing` con `verify = pendiente (gate Full)` para preservar la integridad del contrato de certificación.
- **Fase 1452**: En `ViaticosGestionProcessor.GetSolicitudes` se implementó contador `sinCedula` y emisión de trazas `LoggerHelper.Warning` por cada fila omitida, reportándose en el resumen de `CONSULTAS`.
- **Fase 1453**: En `ViaticosGestionProcessor.RenombrarFila`, la validación de no-vacíos se movió a N0 con `Where(!string.IsNullOrWhiteSpace(c.CC) && !string.IsNullOrWhiteSpace(c.Nom))` retornando `SinComision`, removiendo la guarda inadecuada que clasificaba como `Ambigua`.
- **Fase 1454**: Normalizada la clave de `r.Detalle` mediante `Path.GetFileName(carpeta.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))` para evitar cadenas vacías ante separadores finales.
- **Fase 1455**: `FormGestionDocumental.cs` actualizado para incluir el desglose de archivos pendientes (sin comisión, ambiguos, colisiones, prefijo inválido) en la rama `YaCorrectos > 0`.
- **Fase 1456**: Normalizados encabezados a `## Ronda` en `implementation_plan.md` y sincronizada referencia de tipo `ResultadoRenombrado` en `IViaticosService.yml`.

### Verificación de Principios (Ronda 270)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1451 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1452 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1453 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1454 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1455 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1456 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |

---

## Ronda 271: Verificación de parches 1451–1456 y desfases de trazabilidad (Fases 1457–1460)

### Hallazgos y Acciones Ejecutadas
- **Fase 1457**: Creado `MotivoRenombrado.SinDatos` ("SIN_DATOS_EN_CONSULTAS") y propiedad `SinDatos` en `ResultadoRenombrado`. En `ViaticosGestionProcessor.RenombrarFila`, N0 separa comisiones inexistentes (`SinComision`) de comisiones existentes incompletas (`SinDatos`). En `GetSolicitudes`, se agrupó la advertencia por filas sin cédula en un único aviso para prevenir saturación de logs.
- **Fase 1458**: Actualizados contadores en §2 de `.agents/plans/REGISTRY.md` a Ronda 271 y Fase 1460 (siguiente 272 / 1461), restableciendo la sincronía append-only.
- **Fase 1459**: Actualizado `FormGestionDocumental.BtnRenombrar_Click` en la rama `else` para diagnosticar dinámicamente si el motivo principal es la falta de comisiones en `CONSULTAS` o errores de nomenclatura en archivos.
- **Fase 1460**: Creados los esquemas de API DocFX `docs/api/ConversorPDF.XLL.Core.Models.ResultadoRenombrado.yml` y `ConversorPDF.XLL.Core.Models.MotivoRenombrado.yml`, y registrados en `ConversorPDF.XLL.Core.Models.yml`.

### Verificación de Principios (Ronda 271)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1457 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1458 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1459 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1460 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |

---

## Ronda 272: Incremento de versión a 1.6.3.83 e instalador Wix (Fase 1461)

### Hallazgos y Acciones Ejecutadas
- **Fase 1461**: Actualizada la versión global del producto a `1.6.3.83` en `ConversorPDF.XLL/Core/Helpers/App.cs` (`APP_VERSION`), `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (`Version`, `AssemblyVersion`, `FileVersion`) e `installer/InstaladorConversorPDF.wxs` (`Version` de Package y `DefaultVersion` de plantillas `filTemplate1` y `filTemplate2`). Sincronizados contadores en §2 de `.agents/plans/REGISTRY.md` a Ronda 272 y Fase 1461.

### Verificación de Principios (Ronda 272)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1461 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |


---

## Ronda 276: Revisión del bump a 1.6.3.83 (Ronda 272 / Fase 1461) y reserva de R-005 (Fases 1468–1472)

### Hallazgos y Acciones Ejecutadas
- **Fase 1468**: Se fija 1.6.3.83 como versión objetivo compartida de R-003 y R-005, prohibiendo la generación y publicación del instalador MSI antes de certificar gate verde para ambas capacidades.
- **Fase 1469**: En `installer/InstaladorConversorPDF.wxs`, se preserva `DefaultVersion="1.6.3.82"` en componentes `filTemplate1` (`GF-FO-023.xlsx`) y `filTemplate2` (`FV-FO-033.xlsx`), evitando purgas y sobrescrituras innecesarias de archivos locales del usuario en releases sin alteraciones de plantillas.
- **Fase 1470**: Añadida la fila `R-004` en `REGISTRY.md` §3 formalizando la tarea de mantenimiento del bump a 1.6.3.83 en estado `closed` con `verify = pendiente`.
- **Fase 1471**: Registrada la fila `R-005` en `REGISTRY.md` §3 para el plan `2026-11-01-verificacion-solicitudes-plan.md` en estado `reserved` (`Wave 1..3` → `Ronda 273–275`, `T1..T6` → `Fase 1462–1467`), y contadores de §2 actualizados a Ronda 276 y Fase 1472.
- **Fase 1472**: Actualizada la versión de partida y notas de empaquetado en `.agents/plans/2026-10-31-gestion-documental-carpeta-solicitudes-fix-plan.md` (líneas 50 y 903), y generada la bitácora `docs/articles/changes/2026-10-31-version-1.6.3.83.md`.

### Verificación de Principios (Ronda 276)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1468 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1469 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1470 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1471 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1472 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |


---

## Ronda 277: Verificación de parches 1468–1472 (bump 1.6.3.83 / WiX / registro) (Fases 1473–1476)

### Hallazgos y Acciones Ejecutadas
- **Fase 1473**: En `installer/InstaladorConversorPDF.wxs`, se cambió `On="install"` a `On="uninstall"` en `PurgeOldTemplate1` (`GF-FO-023.xlsx`) y `PurgeOldTemplate2` (`FV-FO-033.xlsx`), asegurando que las plantillas locales del usuario no se purguen durante actualizaciones de versión.
- **Fase 1474**: Creado físicamente en disco el plan `.agents/plans/2026-11-01-verificacion-solicitudes-plan.md` con bloque `numbering:` (R-005, ronda_base 273, fase_base 1462, 3 waves, 6 tasks), `files_lock` y criterios v2, dejando la reserva en estado plenamente ejecutable.
- **Fase 1475**: En `.agents/plans/REGISTRY.md`, fila `R-004` corregida de `closed` a `executing` con `verify = pendiente (gate conjunto)` y `HEAD = pendiente`, condicionando su cierre definitivo al pase exitoso del gate de build y pruebas.
- **Fase 1476**: En `docs/articles/changes/2026-10-31-version-1.6.3.83.md`, el alcance de 1.6.3.83 fue redactado en términos condicionales a la certificación de gate de R-003+R-005. Sincronizados títulos en `implementation_plan.md` y `walkthrough.md`.

### Verificación de Principios (Ronda 277)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1473 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1474 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1475 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1476 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |


---

## Ronda 278: Revisión del plan R-005 y verificación de 1473–1476 (Fases 1477–1482)

### Hallazgos y Acciones Ejecutadas
- **Fase 1477**: Corregidos en T3 del plan los nombres de controles a `txtViaticoInicial` y `txtViaticoFinal`, especificando la suscripción a `this.Load += FormGestionDocumental_Load` en el constructor de `FormGestionDocumental`.
- **Fase 1478**: Acotado el alcance de búsqueda en D-2 y T2 a `folder` (carpeta del libro de Excel) y subcarpeta `Solicitudes` mediante el helper existente `CarpetasDeSolicitudes(folder)`, prohibiendo explícitamente el recorrido del árbol de OneDrive.
- **Fase 1479**: Reescrito `.agents/plans/2026-11-01-verificacion-solicitudes-plan.md` completando el envelope de ejecución (`baseline: 952`, `verify`, `hot_paths_L0`, `stop_conditions`), tabla de olas y los 11 campos obligatorios por tarea (T1–T6).
- **Fase 1480**: En T6 del plan se estipula auditar y registrar el estado de las fases 1457–1460 en R-003 antes del sellado, y se define explícitamente que ante fallo del gate las filas permanecen en `executing` y abren ronda de fix.
- **Fase 1481**: Corregidas columnas de lectura en T2 a `CONSULTAS_NUMERO_COMISION` (B=2), `CONSULTAS_ESTADO` (C=3), `CONSULTAS_INFORME` (H=8) y `CONSULTAS_CEDULA` (J=10), retirando la referencia a columna F.
- **Fase 1482**: En `.agents/plans/REGISTRY.md`, normalizada fecha de `R-005` a `2026-11-01` y descripción de `R-004` como `(mantenimiento sin plan: bump 1.6.3.83)`. Contadores de §2 actualizados a Ronda 278 y Fase 1482.

### Verificación de Principios (Ronda 278)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1477 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1478 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1479 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1480 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1481 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1482 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |


---

## Ronda 279: Revisión de Ronda 278 (Fases 1477–1482) y cobertura funcional de R-005 (Fases 1483–1487)

### Hallazgos y Acciones Ejecutadas
- **Fase 1483**: Aceptada corrección de marcas erróneas en H (`SIN SOLICITUD`/`FALTA SOLICITUD` cuando existe archivo físico) con limpieza de celda, contador `CeldasLimpiadas` y guardado condicional.
- **Fase 1484**: Aceptada política de `_Duplicados` mediante `MismoContenido` para desahogar los 14 archivos con colisión idéntica a `Solicitudes\_Duplicados\` con desambiguación guid hex.
- **Fase 1485**: Aceptada normalización de prefijos de 6 dígitos que inician en "20" con validación de coincidencia por token de nombre en `CONSULTAS`.
- **Fase 1486**: Aceptada normalización obligatoria con `Norm()` (`Trim().ToUpperInvariant()`) en lectura y evaluación de columnas C y H.
- **Fase 1487**: Aceptada preservación de marcas operativas (`ERROR`, `SIN INFORME`, `DUPLICADA`) vía regla D-5b para evitar su sobrescritura.

### Verificación de Principios (Ronda 279)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1483 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1484 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1485 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1486 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |
| F1487 | [X] Sí | [X] Sí | [X] Sí | Sí | Sí |

## Ronda 279: Duplicados en `_Duplicados` y auditoría de INFORME (Fases 1483–1487)

### Contexto y Diagnóstico
- La corrida del 31/10 14:03 devolvió `renombrados=0 · yaCorrectos=381 · colisiones=14 · sinComision=28 · sinDatos=0 · ambiguos=0 · prefijoInvalido=11`.
- Las 14 “colisiones” eran duplicados exactos; 2 prefijos de 6 dígitos correspondían a comisiones existentes; 2 filas tenían `SIN SOLICITUD` con archivo presente.

### Cambios Aplicados
1. `Fase 1483`: `App.FolderNames.DUPLICADOS`, `MotivoRenombrado.Duplicado`, `ResultadoRenombrado.Duplicados`, `MismoContenido` y movimiento a `Solicitudes\_Duplicados`.
2. `Fase 1484`: `EstadoComision`, `ResultadoVerificacionSolicitudes`, `RangoSolicitudes`, `SaveActiveWorkbook` y `VerificarSolicitudesPresentes` (protección VERIFICADO/REINTEGRO, estados terminales, limpieza de marcas, `SIN SOLICITUD`, guardado condicional).
3. `Fase 1485`: `Norm`, `EsMarcaFaltante`, `SinAcentos`/`CoincideNombre` y remapeo de prefijos de 6 dígitos.
4. `Fase 1486`: 12 pruebas nuevas (21 en la clase).
5. `Fase 1487`: documentación y cierre `R-006`; `R-005` → `abandoned`.

### Verificación
- `verify.ps1 -Mode Fast -Filter ViaticosGestionProcessorTests` → 21/21 PASS.
- `verify.ps1 -Mode Full` → exit 0, 999 pruebas, 0 fallos (baseline 952).


---

## Ronda 282: Contratos de Arquitectura y Repositorio para Corridas en Staging AG (Fases 1488–1490)

### Hallazgos y Acciones Especificadas
- **Fase 1488 (T1)**: ESPECIFICADA (no implementada) — Creación de tipos limpios `CorridaStagingItem`, `AgStagingResumen`, `AgCorridaSelector`, `AgFiltrosPolicy` y su registro en `ConversorPDF.XLL.csproj`.
- **Fase 1489 (T2)**: ESPECIFICADA (no implementada) — Consulta SQL `UNION ALL` en `VinculacionQueryRepository` para reflejar fielmente corridas normales y huérfanas con staging, además de `GetResumenStagingAsync`.
- **Fase 1490 (T3)**: ESPECIFICADA (no implementada) — Extracción de `BuildWhere` estático para garantizar paridad estricta entre filas listadas y contadores de status.

### Verificación de Principios (Ronda 282)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1488 | [ ] Pendiente | [ ] Pendiente | [ ] Pendiente | Pendiente | — |
| F1489 | [ ] Pendiente | [ ] Pendiente | [ ] Pendiente | Pendiente | — |
| F1490 | [ ] Pendiente | [ ] Pendiente | [ ] Pendiente | Pendiente | — |

---

## Ronda 283: UI FormVinculaciones, Filtros Desacoplados y Layout Robusto (Fases 1491–1494)

### Hallazgos y Acciones Especificadas
- **Fase 1491 (T4)**: ESPECIFICADA (no implementada) — Desactivación de preselección forzada de periodo y subunidad de `config` cuando `EsModoAg == true`.
- **Fase 1492 (T5)**: ESPECIFICADA (no implementada) — Eliminación de SQL crudo en `FormVinculaciones` y sustitución por `AgCorridaSelector` con preservación de selección.
- **Fase 1493 (T6)**: ESPECIFICADA (no implementada) — Eliminación de `ExecuteScalarAsync` y adición de `btnAjustarFiltrosCorrida` para sincronización instantánea de filtros con la corrida.
- **Fase 1494 (T7)**: ESPECIFICADA (no implementada) — Ajuste de altura de `pnlFilters` a 105 px y reubicación en fila 3 para prevenir solapes en DPI escalado, junto al enlace en `AgFormGeneradorHorarios.cs`.

### Verificación de Principios (Ronda 283)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1491 | [ ] Pendiente | [ ] Pendiente | [ ] Pendiente | Pendiente | — |
| F1492 | [ ] Pendiente | [ ] Pendiente | [ ] Pendiente | Pendiente | — |
| F1493 | [ ] Pendiente | [ ] Pendiente | [ ] Pendiente | Pendiente | — |
| F1494 | [ ] Pendiente | [ ] Pendiente | [ ] Pendiente | Pendiente | — |

---

## Ronda 284: Suite de Pruebas RED→GREEN y Sello de Calidad (Fase 1495)

### Hallazgos y Acciones Especificadas
- **Fase 1495 (T8)**: ESPECIFICADA (no implementada) — Suite de pruebas exhaustivas (unitarias e integración en memoria) cubriendo desambiguación, selección inicial, formateo y consistencia de filtros.

### Verificación de Principios (Ronda 284)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1495 | [ ] Pendiente | [ ] Pendiente | [ ] Pendiente | Pendiente | — |

---

## Ronda 285: Falso positivo de certificación y defectos de contrato en R-007 (Fases 1496–1503)

### Hallazgos y Acciones Ejecutadas
- **Fase 1496**: Revertido falso positivo de certificación de Fases 1488–1495 en `walkthrough.md`, `implementation_plan.md` y contadores de `REGISTRY.md` §2 (`281 / 1487`, siguiente `282 / 1488`, `R-007` en `reserved`). Tareas T2–T5 del plan reestablecidas a `pending`.
- **Fase 1497**: Corregido contrato de `CorridaStagingItem` en el plan: `TieneStaging => TotalSesiones > 0;` (propiedad de expresión) y `Etiqueta { get; set; }` para permitir asignación por el selector.
- **Fase 1498**: Resuelta dependencia invertida y colisión de lease T2/T3: extracción de `BuildWhere` estático trasladada a T2 (primero); T3 acotada a delegación en `IVinculacionListService` / `VinculacionListService` con `depends_on: [T2]`.
- **Fase 1499**: Desacopladas consultas de `GetResumenStagingAsync`: consulta A para descriptores reales de corrida (`PeriodosCorrida`, `SubunidadesCorrida`, `SesionesCorrida`, `VincCorrida` sin filtros UI) y consulta B para `SesionesConFiltros` con `BuildWhere`.
- **Fase 1500**: Añadido manejo de centinela `idCorridaVal <= 0` en T6 para `-- Todas las corridas --`, evitando reporte erróneo de corrida #0.
- **Fase 1501**: Erradicado riesgo de `async void` en `DropDown` de `cmbCorrida` (T5): delegado `(s, e) => _ = CargarCorridasSafeAsync()` con captura de excepciones para proteger el hilo STA de Excel.
- **Fase 1502**: Asegurado layout en T7 con `pnlFilters.MinimumSize = new Size(0, 105);` y anclajes explícitos para escalas DPI 100/125/150 %.
- **Fase 1503**: Incorporada tabla formal de olas (`Wave 1..3` → `Ronda 282..284`) y campo obligatorio `- **fase**:` (1488..1495) en cada tarea T1..T8 del plan.

### Verificación de Principios (Ronda 285)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1496 | [X] Sí | [X] Sí | [X] Sí | Sí | `REGISTRY.md`, docs sincronizados |
| F1497 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1498 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1499 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1500 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1501 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1502 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |
| F1503 | [X] Sí | [X] Sí | [X] Sí | Sí | Plan corregido |


---

## Ronda 286: Sincronización de numeración y frontera de capa en R-007 (Fases 1504–1509)

### Hallazgos y Acciones Ejecutadas
- **Fase 1504**: Actualizada numeración en `REGISTRY.md` (§2 `ronda 285/286`, `fase 1503/1504`) y formalizada fila `R-008` (`closed`) protegiendo el rango 1496–1503 contra reutilización.
- **Fase 1505**: Reubicados `CorridaStagingItem` y `AgStagingResumen` en `Academic.Domain.Models.Genetica`, respetando la frontera de capa limpia.
- **Fase 1506**: Incorporada subconsulta de `ProgramasCorrida` en la Consulta A de `GetResumenStagingAsync`.
- **Fase 1507**: Ajustado orden determinista y separador `', '` en `PeriodosCorrida`.
- **Fase 1508**: Establecida población obligatoria de `Etiqueta` en `GetCorridasStagingAsync` en origen de datos.
- **Fase 1509**: Explicitados anclajes `Top | Left` y ancho fijo 320 px en controles de corrida para DPI 100/125/150 %.

