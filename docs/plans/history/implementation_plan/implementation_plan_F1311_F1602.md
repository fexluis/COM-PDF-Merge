<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# implementation_plan - Fases 1311-1602 (Rondas 264..317)

> Indice: [../../../implementation_plan.md](../../../implementation_plan.md) - Escision: 2026-11-16

---
### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (1311 a 1364) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Preservación de combinaciones marcadas (`Checked = true` en nivel 6) mediante `RecolectarClavesMarcadas` y `RestaurarMarcas` tras repoblar el árbol, con `treeContextoAcademico.EndUpdate()` en bloque `finally`.
  - Prevención de recursión infinita en `CheckedComboBox` mediante bandera `_sincronizandoLista` en el evento `ItemCheck`.
  - Corrección de `AutoClose = true` y cierre accesible mediante `Keys.Enter` y `Keys.Escape` en el dropdown.
  - Consumo efectivo de `cmbTipoPrograma` en `FiltrarProgramas` delegando en `AgNivelAcademicoCatalogo.Desde(p.Modalidad)`.
  - Sustitución de `_toolTip` por el componente real del Designer `toolTipConteo`.

---

## 🔧 Ronda 244: Fases 1360–1364 No Aplicadas — Restauración de Selección Inválida (Fases 1365 a 1368)

### Contexto y Diagnóstico
- **Régimen de Estados:** Todo el backlog (Fases 1311 a 1368) permanece en `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Sustitución de `RecolectarClavesMarcadas`/`RestaurarMarcas` por `RecolectarRutasMarcadas` y `RestaurarRutas`, materializando recursivamente nodos lazy vía `CargarHijosLazy(actual)`.
  - Clave compuesta jerárquica por ruta (`Nivel:Id:Codigo/...`) para evitar colisiones en niveles 4 a 6 donde `Id = 0`.
  - Supresión de cascada `AfterCheck` durante la restauración mediante `_isUpdatingTreeChecks = true` en bloque `try/finally`.
  - Preservación de ramas expandidas mediante `RecolectarExpandidos` y `RestaurarExpandidos` para evitar el colapso del árbol al filtrar.

---

## 🔧 Ronda 245: Fases 1365–1368 No Aplicadas — Restauración Incompleta y Expansión Masiva (Fases 1369 a 1372)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1372) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Detección y contabilización de combinaciones no recuperables (`fallidas` out parameter) cuando grupos/semestres exceden el esqueleto estático (`A..D`, `I..X`).
  - Recálculo ascendente de ancestros y glifos (`RefrescarCadenaRestaurada` ejecutando `RefrescarEstadoNodo` y `ActualizarEstadoPadre`) con `_isUpdatingTreeChecks = true` para evitar inconsistencias visuales y borrado accidental por clics posteriores.
  - Corrección de predicado en `RecolectarExpandidos` evaluando `!n.IsExpanded` en lugar de `Nodes.Count > 0` (evitando expansión inadvertida de ~10.000 nodos dummy y congelamiento de UI).
  - Robustecimiento del parseo de rutas jerárquicas usando `Split(new[] { ':' }, 3)` con `int.TryParse` y saneamiento bidireccional de `:` en códigos de ciudad o programa.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1369 | Restauración Parcial Silenciosa — Grupos Fuera de `{A,B,C,D}` | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1370 | Padres y Glifos Desmarcados Tras Restaurar — Árbol Incoherente | Runtime | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1371 | Expansión Masiva por Predicado `Nodes.Count > 0` — Congelamiento UI | Runtime | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1372 | `Split(':')` Sin Límite y Parseo Frágil — `FormatException` | Compilation/Runtime | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |

---

## 🔧 Ronda 246: Fases 1369–1372 No Aplicadas — Dimensiones de Decisión Fabricadas en el Árbol (Fases 1373 a 1375)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1375) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Erradicación de dimensiones inventadas en `CargarHijosLazy` (`MANANA/4/MAR-VIE`); si no provienen de selección previa ni de contexto/modalidad resuelta, fallar explícitamente (`I1_JORNADA_VACIA`, `I6_HORAS_DIA_INVALIDO`).
  - Sustitución de snapshot plano por `List<AgContextoSeleccionDto>` de `ObtenerCombinacionesSeleccionadas()` preservando `AulaRef`, `Jornada`, `HorasDia` y `DiasVentana` al restaurar mediante `AgNodoArbolKeyMapper.Aplicar`.
  - Ampliación de la definición de `ag_filtro_seleccion` en la migración V78 agregando columnas `codigo`, `aula_ref`, `jornada`, `horas_dia`, `dias_ventana`, `orden` e índice único para permitir persistencia real y completa de combinaciones de nivel 6.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1373 | `CargarHijosLazy` Fabrica Jornada/Horas/Días (`MANANA/4/MAR-VIE`) | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs`, `AgSeleccionMapper.cs` |
| F1374 | Snapshot de Restauración Pierde Dimensiones Reales | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs`, `AgNodoArbolKeyMapper.cs` |
| F1375 | Tabla V78 `ag_filtro_seleccion` Requiere Columnas de Dimensión | Logic (persistencia) | ESPECIFICADA | `MigrationV78_AddAgFiltroSeleccion.cs` |

---

## 🔧 Ronda 247: Fases 1373–1375 No Aplicadas — API Inexistente y Regresión que Bloquea la Generación (Fases 1376 a 1378)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1378) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Corrección de firma y orden de argumentos en llamada a `AgNodoArbolKeyMapper.AplicarDimensiones(key, dto)` con namespace explícito `ConversorPDF.XLL.Academic.Application.Mappers`.
  - Tratamiento explícito de `Result.Failure` retornado por `AplicarDimensiones`: ante dimensiones incompletas, descartar el nodo (`fallidas++`, no marcar `Checked`) impidiendo que el fallback defensivo reintroduzca constantes.
  - Resolución de dimensiones en `CargarHijosLazy` a partir del catálogo indexado de aulas (`_aulasPorRef[aulaRef]`) derivado de `ag_aula`, eliminando el bloqueo de generación sin inventar datos.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1376 | F1374 Invoca `Aplicar` Inexistente — Firma Invertida y CS0117 | Compilation | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1377 | `AplicarDimensiones` Devuelve `Failure` con Fallback — F1374 Aceptaba Fallback | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1378 | F1373 Bloquea Generación Sin Control UI — Fuente Real `ag_aula` Ignorada | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |

---

## 🔧 Ronda 248: Fases 1376–1378 No Aplicadas — Caché Vacía y Variables Inexistentes (Fases 1379 a 1381)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1381) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Precarga de aulas (`CargarAulasParaArbolAsync`) antes de invocar `PoblarArbolContextoAcademico` en `CargarCatalogosAsync` y refresco tras "Generar Aulas", garantizando que `_aulasPorRef` contenga datos al expandir nivel 5.
  - Implementación de `RutaHasta(node)` para reconstruir la tupla jerárquica `(IdPeriodo, IdSede, IdPrograma, Semestre, Grupo)` desde los ancestros, resolviendo 4 errores `CS0103` en `CargarHijosLazy`.
  - Incorporación de diccionario complementario `_aulasPorClaveNorm` con modalidad canónica (`NormalizarTutoria`) como fallback tolerante a registros históricos con modalidades renombradas (ej. V66 `"VIRTUAL"` vs `"PLATAFORMA"`).

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1379 | `_aulasPorRef` Vacío al Construir el Árbol — F1378 No Resuelve Dimensiones | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1380 | F1378 Usa Variables Inexistentes en `CargarHijosLazy` — CS0103 | Compilation | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1381 | Índice Solo por `AulaRef` Exacto — Filas con Modalidad Legacy No Coinciden | Warning | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |

---

## 🔧 Ronda 249: Fases 1379–1381 No Aplicadas — `DataSource` Incompatible y Estado Inicial de Filtros (Fases 1382 a 1385)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1385) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Sustitución de asignaciones a propiedades inexistentes (`DataSource`, `DisplayMember`, `ValueMember`) en `CheckedComboBox` por llamadas canónicas `cmb.Cargar(items, textSelector, valueSelector)`, evitando 9 errores de compilación `CS1061`.
  - Eliminación de `SelectedIndex = 0` para `cmbProgramas` y `cmbSedes`, asegurando que la convención "vacío = todos" no restrinja indebidamente la visualización inicial del árbol a 1 programa y 1 sede.
  - Definición del ciclo de vida e invalidación de caché de aulas (`InvalidarAulasCache()`) ante cambios de período, subunidad y tras regenerar aulas.
  - Optimización de carga granular de aulas para evitar latencia excesiva y consumos de memoria innecesarios al abrir el formulario.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1382 | `DataSource`/`DisplayMember`/`ValueMember` No Existen en `CheckedComboBox` | Compilation | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1383 | `SelectedIndex = 0` Restringe el Árbol a 1 Programa/1 Sede por Defecto | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1384 | `_aulasPorRef` No Se Invalida al Cambiar Período/Subunidad | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1385 | Carga Completa de `ag_aula` en Cada Apertura — Latencia y Memoria | Warning | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |

---

## 🔧 Ronda 250: Fases 1382–1385 No Aplicadas — Designer Sin Migrar y Selección Múltiple Colapsada (Fases 1386 a 1389)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1389) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Migración del archivo `AgFormGeneradorHorarios.Designer.cs` instanciando `CheckedComboBox` y eliminando propiedades exclusivas de `ComboBox` (`DropDownStyle`, `FormattingEnabled`), configurando los modos de selección (`ModoSeleccion.Unico` / `Multiple`).
  - Incorporación de métodos `Marcar(object)` y `SeleccionarTodos(IEnumerable<int>)` en `CheckedComboBox` para soportar la restauración fiel de perfiles sin asumir setter de `SelectedItem`.
  - Corrección de las 16 lecturas que asumían un único `SelectedItem as Programa/Sede`, migrándolas al patrón `Seleccionados<T>` con fallback "vacío = todos".
  - Exposición del evento de compatibilidad `SelectedIndexChanged` en `CheckedComboBox` mapeado a `SeleccionChanged`, garantizando que `OnVentanaChanged` se ejecute al marcar o desmarcar ítems.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1386 | Designer Aún Declara `System.Windows.Forms.ComboBox` — 6 × CS1061 | Compilation | ESPECIFICADA | `AgFormGeneradorHorarios.Designer.cs` |
| F1387 | `SelectedItem = …` en Restauración de Perfil No Existe en `CheckedComboBox` | Compilation | ESPECIFICADA | `AgFormGeneradorHorarios.cs`, `CheckedComboBox.cs` |
| F1388 | 16 Lecturas `SelectedItem as Programa/Sede` Colapsan Selección Múltiple | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1389 | `SelectedIndexChanged` No Está en la API de `CheckedComboBox` | Compilation | ESPECIFICADA | `AgFormGeneradorHorarios.cs`, `CheckedComboBox.cs` |

---

## 🔧 Ronda 251: Fases 1386–1389 No Aplicadas — Registro de Migración, Idempotencia V78 y Duplicación de `FormTerceros` (Fases 1390 a 1392)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1392) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Registro seguro de `MigrationV78_AddAgFiltroSeleccion` en `ConversorPDF.XLL\Core\Helpers\SqliteHelper.cs:378-380` agregando la coma faltante en la entrada V77 previa para evitar error de sintaxis `CS1002`.
  - Implementación tolerante e idempotente de `MigrationV78_AddAgFiltroSeleccion` implementando `IMigration, IMigrationTolerante` y utilizando `PRAGMA table_info` antes de emitir `ALTER TABLE ... ADD COLUMN` para evitar excepciones `duplicate column name` que bloqueen el arranque de la aplicación.
  - Reutilización de `FormTerceros.cs` (850 líneas) existente en `Academic\UI\Forms\`, descartando la duplicación innecesaria de un segundo formulario `AgFormTerceros` y canalizando el filtro de terceros a través de `FormTerceros` con `AgFiltroSeleccionRepository`.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1390 | Registro de V78 Rompe la Lista y Ruta del Backlog Es Incorrecta | Compilation | ESPECIFICADA | `ConversorPDF.XLL\Core\Helpers\SqliteHelper.cs` |
| F1391 | V78 Sin Guard de Columna Duplicada — Excepción al Arrancar | Runtime | ESPECIFICADA | `MigrationV78_AddAgFiltroSeleccion.cs` |
| F1392 | `AgFormTerceros` Duplica `FormTerceros` Existente | Logic | ESPECIFICADA | `FormTerceros.cs`, `FormVerVinculaciones.cs`, `FormVinculaciones.cs` |

---

## 🔧 Ronda 252: Fases 1390–1392 No Aplicadas — Guard de Tabla Ausente y Clave de Filtro No Disponible (Fases 1393 a 1394)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1394) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Alineación de `AgregarColumnaSiNoExiste` en `MigrationV78_AddAgFiltroSeleccion` con el patrón productivo de `MigrationV77`, verificando previamente la existencia de la tabla en `sqlite_master` antes de ejecutar `PRAGMA table_info` y `ALTER TABLE`.
  - Exposición de la propiedad `IdSede` en `SesionListItem.cs` y creación del DTO desacoplado `AgContextoFiltro` en `FormTerceros.cs` con constructor opcional, permitiendo la parametrización exacta de filtros sin romper invocaciones previas.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1393 | Helper de F1391 Sin Guard de Tabla y Nombre No Canónico | Runtime | ESPECIFICADA | `MigrationV78_AddAgFiltroSeleccion.cs` |
| F1394 | F1392 Usa Clave Inexistente — `IdSede` Ausente en `SesionListItem` y Form | Logic | ESPECIFICADA | `SesionListItem.cs`, `FormTerceros.cs`, `FormVerVinculaciones.cs` |

---

## 🔧 Ronda 253: Cierre de Brechas A1–A11 y Escenarios Operativos (Fases 1395 a 1401)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1401) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Creación explícita de `cmbTipoPrograma` (`ComboBox` simple, `TODOS/PREGRADO (n)/POSGRADO (n)`) en `AgFormGeneradorHorarios.Designer.cs` y rediseño de `pnlFiltros` a 2 filas (66px).
  - Filtrado y repoblación del árbol por `cmbTipoPrograma` delegando en `AgNivelAcademicoCatalogo.Desde(p.Modalidad)` preservando selecciones activas.
  - Asignación de horas en nivel 6 según modalidad del programa (Pregrado = 4h, Posgrado = 8h), erradicando el valor constante `HorasDia = 4`.
  - Notificación de advertencia no modal en `TreeContextoAcademico_AfterCheck` al marcar 8h en Pregrado, manteniendo pre-vuelo 1114 y `AgUnidadesSemanales` inalterados.
  - Implementación de controles masivos (`btnSeleccionarTodos`, `btnBorrarSeleccion`, `btnBuscar`, `chkSoloHabilitadas`) operando sobre `_tercerosVista` en `FormTerceros.cs`.
  - Índice de unicidad por subunidad y clave compuesta en `MigrationV78_AddAgFiltroSeleccion.cs` con sentencia `INSERT ... ON CONFLICT DO UPDATE`.
  - Pre-marcado automático de terceros en `FormTerceros` desde registros persistidos y ejecución de `AplicarAlcanceTerceros` previo a `ConstruirIndices` (L1005).

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1395 | Designer: `cmbTipoPrograma` y Layout a 2 Filas (66px) en `pnlFiltros` | Compilation/UI | ESPECIFICADA | `AgFormGeneradorHorarios.Designer.cs` |
| F1396 | `FiltrarProgramas` por `cmbTipoPrograma` con Preservación de Selección | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1397 | Horas Nivel 6 Derivadas de `Programa.Modalidad` (Preg 4h / Posg 8h) | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1398 | Advertencia No Modal al Marcar 8h en Pregrado (Invariantes 1114/4h Intactos) | Logic/UX | ESPECIFICADA | `AgFormGeneradorHorarios.cs` |
| F1399 | Fila de Acciones Masivas y Filtro en `FormTerceros` sobre Vista Filtrada | Logic/UI | ESPECIFICADA | `FormTerceros.cs`, `FormTerceros.Designer.cs` |
| F1400 | V78: Unicidad por Subunidad y Upsert Idempotente en Persistencia | Logic | ESPECIFICADA | `MigrationV78_AddAgFiltroSeleccion.cs` |
| F1401 | Pre-marcado en `FormTerceros` y Ejecución Previa a `ConstruirIndices` | Logic | ESPECIFICADA | `FormTerceros.cs`, `AgFormGeneradorHorarios.cs` |

---

## 🔧 Ronda 254: Corrección de Dominio y Retractación (Fases 1402 a 1407)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1407) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Evidencia de Dominio:**
  - `programas.modalidad` = `"PREGRADO"` / `"POSGRADO"` (`AcademicSchema.sql:23`, vistas `p.modalidad AS ProgramaTipo`, filtro `VinculacionQueryRepository.cs:111-115`).
  - `ModalidadTutoriaCatalog` = `Presencial` / `Plataforma` / `Mixta` aplica exclusivamente a `ag_aula.modalidadVinculacion` y `AgAulaRef`.
  - La migración V78 **no** altera la tabla `programas`.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1402 | `FiltrarPoolDocentesPorTerceros` antes de `ConstruirIndices` | Compilation | ESPECIFICADA | `AgContexto.cs`, `AgFormGeneradorHorarios.cs` |
| F1403 | ~~Nivel académico desde modalidad~~ | Logic | **RETIRADA (falso positivo)** | — |
| F1404 | Reutilizar `txtBuscar`/`cmbTerritorial`; solo 2 botones masivos | Logic/UI | ESPECIFICADA | `FormTerceros.Designer.cs` |
| F1405 | Índice único plano + dedupe antes de indexar (V78) | Runtime | ESPECIFICADA | `MigrationV78_AddAgFiltroSeleccion.cs` |
| F1406 | `AgTipoProgramaCatalog` sobre `programas.modalidad` (PREGRADO/POSGRADO) | Logic | ESPECIFICADA | `AgTipoProgramaCatalog.cs`, `AgFormGeneradorHorarios.cs:444-452` |
| F1407 | Horas nivel 6 = 4h Pregrado / 8h Posgrado, separadas de tutoría | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.cs:2951-3011` |

---

## 🔧 Ronda 255: `dgvTerceros` Sin Columna de Selección, DDL de V78 Indefinido y Selección Perdida al Refiltrar (Fases 1408 a 1410)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1410) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Inserción de `DataGridViewCheckBoxColumn` (`colSeleccion`) en posición 0 de `dgvTerceros` y sincronización con conjunto persistente `HashSet<int> _seleccionados` ante eventos `CellValueChanged`.
  - Definición explícita del DDL de `ag_filtro_seleccion` en `MigrationV78_AddAgFiltroSeleccion.cs` conteniendo todas las columnas de la clave compuesta más `creado` y `actualizado` para soportar `DO UPDATE SET actualizado = datetime('now')`.
  - Restauración automática del estado de selección en `AplicarFiltroTerceros` tras cada asignación a `_bs.DataSource`, garantizando que la búsqueda textual o cambio de territorial no desmarque docentes previamente seleccionados.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1408 | Inserción de `colSeleccion` en `dgvTerceros` y Tracking con `HashSet<int>` | Runtime | ESPECIFICADA | `FormTerceros.cs`, `FormTerceros.Designer.cs` |
| F1409 | Definición Completa de DDL para V78 con Soporte de `actualizado` | Runtime | ESPECIFICADA | `MigrationV78_AddAgFiltroSeleccion.cs` |
| F1410 | Preservación de Selección en `AplicarFiltroTerceros` tras Rebind de `_bs` | Logic | ESPECIFICADA | `FormTerceros.cs` |

---

## 🔧 Ronda 256: Auditoría de Cobertura v4, Layout de Filtros y Sincronía de Grilla (Fases 1411 a 1414)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1414) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Decisión de Contenedor de Terceros:** Opción B confirmada (reutilizar `FormTerceros.cs` con sobrecarga de `modoSeleccion` y `AgContextoFiltro`, preservando compatibilidad con sus 2 llamadores activos).
- **Defectos Críticos Subsanados en Especificación:**
  - Incorporación de `CurrentCellDirtyStateChanged` con `CommitEdit(Commit)` en `dgvTerceros` para forzar confirmación inmediata del check sin esperar cambio de fila.
  - Inicialización de `_seleccionados` desde `_seleccionInicial` antes de la primera llamada a `AplicarFiltroTerceros`, evitando borrado de marcas al abrir el formulario (A10).
  - Corrección del layout de `pnlFiltros` en `AgFormGeneradorHorarios.Designer.cs` ampliando a 66px y distribuyendo los controles en dos filas no colisionantes.
  - Supresión del control redundante `lblAvisoHorasNivel` en `pnlCabecerasNivel6` (evitando solapamiento con cabeceras dinámicas calculadas) y canalización de la advertencia no modal a través de `lblResumenFiltros` y `toolTipConteo`.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1411 | `CurrentCellDirtyStateChanged` con `CommitEdit` en `dgvTerceros` | Logic | ESPECIFICADA | `FormTerceros.cs` |
| F1412 | Pre-carga de `_seleccionInicial` en `_seleccionados` al Abrir `FormTerceros` | Logic | ESPECIFICADA | `FormTerceros.cs` |
| F1413 | Layout de 2 Filas en `pnlFiltros` (66px) Sin Colisiones Geométricas | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.Designer.cs` |
| F1414 | Supresión de `lblAvisoHorasNivel` y Canalización en `lblResumenFiltros` | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.Designer.cs`, `AgFormGeneradorHorarios.cs` |

---

## 🔧 Ronda 257: Cascada Geométrica en `grpSeleccion` y Habilitación de Edición en `dgvTerceros` (Fases 1415 a 1416)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1416) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Cascada de desplazamiento vertical en `grpSeleccion`: al crecer `pnlFiltros` a 66px (bottom 85), desplazar `pnlArbolToolbar` a y=85, `pnlCabecerasNivel6` a y=119, y `treeContextoAcademico` a y=139 con altura 146px, manteniendo el fondo total en 285px sin recortar ni ocluir la barra de herramientas del árbol.
  - Habilitación de edición en la grilla de terceros estableciendo `dgvTerceros.ReadOnly = false`, aplicando `col.ReadOnly = true` sobre las 4 columnas de texto y configurando `EditMode = EditOnEnter` para permitir el marcado inmediato de `colSeleccion`.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1415 | Cascada de Desplazamiento Vertical en Hijos de `grpSeleccion` | Logic | ESPECIFICADA | `AgFormGeneradorHorarios.Designer.cs` |
| F1416 | `dgvTerceros.ReadOnly = false` con Columnas de Texto Protegidas | Runtime | ESPECIFICADA | `FormTerceros.Designer.cs`, `FormTerceros.cs` |

---

## 🔧 Ronda 258: Constructor Extendido, Panel de Selección Inferior y DisplayName (Fases 1417 a 1419)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (Fases 1311 a 1419) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Sobrecarga de constructor en `FormTerceros.cs` con parámetros `(bool modoSeleccion, IReadOnlyCollection<int> seleccionInicial, Func<string, IReadOnlyCollection<int>> cargarSeleccion)`, delegando al constructor primario para preservar intactos a los llamadores `FormVerVinculaciones` y `FormVinculaciones`.
  - Creación de `pnlSeleccion` (`Dock = DockStyle.Bottom`, 40px) en `FormTerceros.Designer.cs`, visible exclusivamente en `modoSeleccion`, alojando los botones masivos (`btnSeleccionarTodos`, `btnBorrarSeleccion`, `btnGuardarSeleccion`) y el contador `lblContadorSeleccion` sin distorsionar `pnlTop` ni `pnlDetails`.
  - Exposición de la propiedad calculada `DisplayName` en `Tercero.cs` para cumplir la convención A9 (`NOMBRE (RESIDENCIA)`) y enlace como `DataPropertyName` en la columna de nombre de `dgvTerceros`.

| Fase | Título | Severidad | Estado | Archivos Afectados |
|---|---|---|---|---|
| F1417 | Sobrecarga de Constructor en `FormTerceros` para Inyección de Selección | Compilation | ESPECIFICADA | `FormTerceros.cs` |
| F1418 | Incorporación de `pnlSeleccion` (Dock Bottom) en `FormTerceros` | Logic | ESPECIFICADA | `FormTerceros.Designer.cs`, `FormTerceros.cs` |
| F1419 | Propiedad Calculada `Tercero.DisplayName` y Enlace a `dgvTerceros` | Logic | ESPECIFICADA | `Tercero.cs`, `FormTerceros.cs` |

---

## 🎯 Plan Consolidado v5 (Handoff) — Filtro de Árbol AG + CheckedComboBox + Selección de Terceros (Fases 1311 a 1421)

### 1. Decisiones Vinculantes de Arquitectura (D1 a D10)
- **D1 (Opción A):** Creación de `AgFormTerceros` nuevo (`.cs` y `.Designer.cs`). `FormTerceros.cs` queda 100% intacto y no se modifica.
- **D2:** Retirada de modificación a `SesionListItem`. `AgFormTerceros` opera directamente con `Tercero.Departamento` y `SyncContext.ActiveSubunidad`.
- **D3:** Modo de selección múltiple (`CheckedComboModo.Multiple`) en `cmbPeriodos`, `cmbSedes` y `cmbProgramas`. `cmbTipoPrograma` es un `ComboBox` nativo (`DropDownList`).
- **D4 & D5:** Migración V78 con índice plano `ux_ag_filtro_seleccion(subunidad, tipo, refid)` con columnas informativas `semestre, grupo, modalidad`. Normalización de nulos y deduplicación antes de indexar.
- **D6:** Regla de consumo de selección: contexto = `Primario<T>()`; enumeraciones = `Seleccionados<T>()`.
- **D7:** Precedencia de horas en nivel 6: `aula.HorasDia > 0` > `HorasPorNivelAcademico(prog?.Modalidad)` > 4h. No se blanquea `Jornada` ni `DiasVentana`.
- **D8:** Liberación de genes fijados cuyos docentes queden fuera del alcance antes de `ConstruirIndices()`, registrando traza preventiva.
- **D9:** Designer editado estáticamente sin llamadas en tiempo de ejecución a `Controls.Clear()`.
- **D10:** Repositorio transaccional inyectando `_connectionFactory`, abriendo y cerrando conexiones locales sin capturar la conexión global de `context.Open()`.

### 2. Inventario de Archivos a Crear / Modificar
1. `ConversorPDF.XLL\Academic\Domain\Models\AgTipoProgramaCatalog.cs` (Nuevo)
2. `ConversorPDF.XLL\Academic\Domain\Models\AgFiltrosSeleccion.cs` (Nuevo)
3. `ConversorPDF.XLL\Academic\Domain\Interfaces\IAgFiltroSeleccionRepository.cs` (Nuevo)
4. `ConversorPDF.XLL\Academic\Application\Services\AgFiltroArbolResolver.cs` (Nuevo)
5. `ConversorPDF.XLL\Academic\Application\Services\AgTerceroFiltro.cs` (Nuevo)
6. `ConversorPDF.XLL\Academic\Application\Services\AgTerceroSeleccion.cs` (Nuevo)
7. `ConversorPDF.XLL\Academic\Application\DTOs\AgTerceroSeleccionDto.cs` (Nuevo)
8. `ConversorPDF.XLL\Academic\Application\DTOs\AgDimensionesSeleccion.cs` (Nuevo)
9. `ConversorPDF.XLL\Academic\Infrastructure\Persistence\Repositories\AgFiltroSeleccionRepository.cs` (Nuevo)
10. `ConversorPDF.XLL\Academic\Infrastructure\Persistence\Migrations\MigrationV78_AddAgFiltroSeleccion.cs` (Nuevo)
11. `ConversorPDF.XLL\Academic\UI\Controls\CheckedComboBox.cs` (Nuevo)
12. `ConversorPDF.XLL\Academic\UI\Forms\AgFormTerceros.cs` + `.Designer.cs` (Nuevos)
13. `ConversorPDF.XLL\Core\Helpers\SqliteHelper.cs` (Modificar L378: registro V78)
14. `ConversorPDF.XLL\Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` (Modificar: 2 filas 66px, controles, cascada)
15. `ConversorPDF.XLL\Academic\UI\Forms\AgFormGeneradorHorarios.cs` (Modificar: integración de filtros, árbol y alcance)
16. `ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgContexto.cs` (Modificar: `FiltrarPoolDocentesPorTerceros`)
17. `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (Modificar: agregar `<Compile Include>` para los nuevos archivos)





















































---

## Ronda 264: Corregir RenombrarSolicitudes y rango #YY (Fases 1426–1431)

## Contexto y Diagnóstico
- **Fase 1426 (T6)**: `ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs` no importaba `System.Linq`, impidiendo la compilación de `Select`.
- **Fase 1427 (T1)**: `TryParseLimite` en `ViaticoNumeroComision.cs` trataba cualquier límite que no terminara en el sufijo vigente como consecutivo del año actual, desvirtuando límites de años anteriores (ej. `12625`).
- **Fase 1428 (T1)**: Rama con `return false` en `EnRango` contradecía la regla "límite no interpretable ⇒ sin filtro".
- **Fase 1429 (T1)**: Año de vigencia literal `2026` duplicado sin desacoplar de `App.APP_VIGENCIA`.
- **Fase 1430 (T2/T5)**: Falta de feedback al usuario en `BtnUnificar_Click` y `MessageBox` bloqueantes en capa Core (`ViaticosGestionProcessor`).
- **Fase 1431 (T7/T8)**: Numeración global desfasada, artefactos sin registrar y YAML docfx desactualizado.

## Cambios Aplicados
1. Importación de `System.Linq` en `ViaticosGestionProcessorTests.cs`.
2. Detección precisa de formatos `#YY` en `TryParseLimite` (`sLimite.Length >= 3 && (sLimite.EndsWith(vigenciaSuffix) || sLimite.Length >= 5)`).
3. Fallback `return true` (sin filtro) en `EnRango` ante límites no interpretables.
4. Definición de `AnioVigenciaPorDefecto = 2026` y pruebas con año dinámico derivado de `App.APP_VIGENCIA`.
5. Feedback en `FormGestionDocumental.BtnUnificar_Click` y remoción de `MessageBox` en capa `Core`.
6. Regeneración de `docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosGestionProcessor.yml` y actualización de `.agents/plans/REGISTRY.md` a Ronda 264 / Fase 1431.


---

## Ronda 265: Correcciones Ronda 264 y desacople total de WinForms (Fases 1432–1436)

## Contexto y Diagnóstico
- **Fase 1432 (T2/T5)**: Presencia residual de `MessageBox` y `using System.Windows.Forms;` en `ViaticosGestionProcessor.cs`. Ausencia de retorno numérico y falta de aviso ante excepciones.
- **Fase 1433 (T5)**: Retorno `0` al fallar `UnificarSolicitudes` enmascaraba errores reales como "sin resultados".
- **Fase 1434 (T7)**: Tests de `EnRango` mantenían valores literales anclados a `26`.
- **Fase 1435 (T8)**: DocFX `IViaticosGestionProcessor.yml` requería actualización al nuevo retorno `int` de `RenombrarSolicitudes`.
- **Fase 1436 (T8)**: Inexistencia de archivo de plan para R-002 y estado sin sellar en `REGISTRY.md`.

## Cambios Aplicados
1. Firma de `RenombrarSolicitudes` actualizada a `int` (`-1` error, `0` vacía, `>0` éxito) en `IViaticosGestionProcessor`, `IViaticosService`, `ViaticosService` y `ViaticosGestionProcessor`.
2. Remoción total de `using System.Windows.Forms;` y llamadas `MessageBox` en `ViaticosGestionProcessor.cs`.
3. Retorno `-1` en caso de error en `UnificarSolicitudes` y manejo en `FormGestionDocumental.cs` y `RibbonController.cs`.
4. Dinamización de valores de prueba en `ViaticoNumeroComisionTests.cs` mediante `AnioVigenciaActual % 100`.
5. Esquema docfx actualizado en `docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosGestionProcessor.yml`.
6. Creación formal del plan `.agents/plans/2026-10-30-gestion-documental-rango-yy-plan.md` y sincronización en `.agents/plans/REGISTRY.md` a Ronda 265 / Fase 1436 con HEAD `3bb7b18`.

---

## Ronda 266: Procesador base, interfaces, servicio y formulario UI (Fases 1437–1438)

## Contexto y Diagnóstico
- **Fase 1437 (T1)**: `RenombrarSolicitudes` solo procesaba la carpeta raíz ignorando la subcarpeta `Solicitudes` donde residen los PDFs de comisiones. `Estandarizar` renombraba cualquier archivo del directorio (no solo `*.pdf`). Ausencia de contadores detallados y catálogo de motivos.
- **Fase 1438 (T2)**: El formulario `FormGestionDocumental` emitía falsamente el mensaje "No se encontraron comisiones para el rango indicado" cuando el conteo de renombrados era 0 pese a existir registros coincidentes en `CONSULTAS`. Falta de exposición de desglose detallado en interfaces y servicio.

## Cambios Aplicados
1. Creación de `ResultadoRenombrado.cs` y `MotivoRenombrado.cs` bajo `ConversorPDF.XLL.Core.Models` e inclusión en `ConversorPDF.XLL.csproj`.
2. Escaneo dual de carpeta raíz y subcarpeta `Solicitudes` mediante helper `CarpetasDeSolicitudes`, restricción de `Estandarizar` a `*.pdf`, trazabilidad `Debug` por archivo y resumen `[RENOMBRAR]`.
3. Adición de `RenombrarSolicitudesConDetalle` en `IViaticosGestionProcessor`, `IViaticosService` y `ViaticosService`.
4. Reemplazo del manejador `BtnRenombrar_Click` en `FormGestionDocumental` con desglose explícito de comisiones en rango, PDFs examinados, renombrados, colisiones, ambiguos, sin comisión y errores.

---

## Ronda 267: Deduplicación N0–N3 y FindCC endurecido (Fases 1439–1440)

## Contexto y Diagnóstico
- **Fase 1439 (T3)**: Filas duplicadas con cédula idéntica en `CONSULTAS` generaban ambigüedad artificial y forzaban lectura OCR innecesaria. Carencia de normalización de espacios (`Trim`) en lectura de comisiones.
- **Fase 1440 (T4)**: `FindCC` capturaba números de rubro presupuestal (`0505023`) u otros códigos institucionales mediante expresiones regulares genéricas, causando falsos positivos de asignación.

## Cambios Aplicados
1. Pipeline de 4 niveles en `ViaticosGestionProcessor`: N0 (deduplicación por `Num, CC`), N1 (metadatos internos), N2 (OCR validado), N3 (ambigua sin modificar).
2. Sanitización con `Trim()` en lectura de columnas de comisión y cédula en `GetSolicitudes`, y helper `NormalizarDocumento`.
3. Sobrecarga de `FindCC(string f, ICollection<string> cedulasCandidatas)` restringiendo el match exclusivamente al conjunto de cédulas válidas para la comisión.

---

## Ronda 268: Pruebas de regresión, documentación y cierre R-003 (Fases 1441–1442)

## Contexto y Diagnóstico
- **Fase 1441 (T5)**: Necesidad de validar la no-regresión de la carpeta `Solicitudes`, dedupe sin OCR, OCR con comisiones compartidas, rechazo de rubros ajenos, detección de colisiones y manejo de prefijos inválidos.
- **Fase 1442 (T6)**: Sincronización de documentación técnica, bitácora de cambios, esquema DocFX y cierre formal del registro `R-003`.

## Cambios Aplicados
1. 7 pruebas unitarias nuevas en `ViaticosGestionProcessorTests.cs` cubriendo todos los escenarios del flujo de renombrado y su invariante de balance.
2. Actualización de `docs/articles/gestion-documental.md` con política N0–N3, catálogo de motivos y requisitos de `CONSULTAS`.
3. Creación de bitácora `docs/articles/changes/2026-10-31-gestion-documental-carpeta-y-desambiguacion.md`.
4. Actualización de `docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosGestionProcessor.yml`.
5. Cierre de fila `R-003` en `.agents/plans/REGISTRY.md` con contadores actualizados a Ronda 268 y Fase 1442.

---

## Ronda 269: Revisión de la implementación Ronda 266–268 (Fases 1443–1450)

## Contexto y Diagnóstico
- **Fase 1443**: Fila `R-003` cerrada con `verify = aprobado` sin ejecución previa de `scripts/verify.ps1 -Mode Full` debido a la restricción del Paso 6 (no compilar/ejecutar en `/ronda-review`).
- **Fase 1444**: Re-ejecución con `Renombrados == 0` y `YaCorrectos > 0` arrojaba advertencia de "no se renombró ningún archivo... verifique carpeta Solicitudes" (falso negativo de idempotencia).
- **Fase 1445**: Filas en `CONSULTAS` con cédula vacía producían candidatos incompletos y riesgo de generar nombres de destino malformados.
- **Fase 1446**: `Estandarizar` mutaba nombres sin dejar traza en logs `Debug`/`Warning`.
- **Fase 1447**: Clave de `r.Detalle` usaba únicamente el nombre de archivo, causando colisiones entre archivos de raíz y subcarpeta `Solicitudes`.
- **Fase 1448**: Patrón 2 de `FindCC` no toleraba comas de miles en cédulas ni acotaba longitud (6–11 dígitos).
- **Fase 1449**: Encabezados inconsistentes en bitácora y ausencia de `RenombrarSolicitudesConDetalle` en `IViaticosService.yml`.
- **Fase 1450**: Criterio fijo 231 obsoleto frente al universo ampliado de 419 PDFs en `Solicitudes`.

## Cambios Aplicados
1. `REGISTRY.md`: Columna `verify` en `R-003` ajustada a `pendiente (gate Full)` para reflejar fielmente el estado pre-compilación.
2. `FormGestionDocumental.cs`: Añadida rama `else if (r.YaCorrectos > 0)` informando estado sin cambios cuando los archivos ya son conformes.
3. `ViaticosGestionProcessor.cs`: Filtro `!string.IsNullOrEmpty(cc)` en `GetSolicitudes`; guarda en `RenombrarFila` retornando `Ambigua` si faltan CC o Nom; trazas en `Estandarizar`; clave discriminada `carpeta/archivo` en `r.Detalle`; regex y acotamiento de longitud en `FindCC`.
4. `implementation_plan.md`: Normalizados encabezados a `## Ronda`.
5. `IViaticosService.yml`: Declarado miembro `RenombrarSolicitudesConDetalle`.
6. Plan `2026-10-31-…-fix-plan.md` §10: Invariante obligatoria formalizada como criterio rector de aceptación.

---

## Ronda 270: Verificación de los parches 1443–1450 (regresiones y cierres) (Fases 1451–1456)

## Contexto y Diagnóstico
- **Fase 1451**: `REGISTRY.md` mantenía estado `closed` en `R-003` pese a que `verify` estaba en `pendiente (gate Full)`.
- **Fase 1452**: En `GetSolicitudes`, el descarte de filas sin cédula no emitía traza de advertencia ni contabilizaba `sinCedula` en el resumen de `CONSULTAS`.
- **Fase 1453**: En `RenombrarFila`, la guarda posterior devolvía `Ambigua` cuando faltaba nombre, distorsionando la métrica de ambigüedad.
- **Fase 1454**: `Path.GetFileName(carpeta)` devolvía cadena vacía si la carpeta terminaba con separador de directorios, degradando la clave de `r.Detalle`.
- **Fase 1455**: Rama `YaCorrectos > 0` en `FormGestionDocumental.cs` omitía el desglose de pendientes (sin comisión, ambiguos, colisiones, prefijo inválido).
- **Fase 1456**: Inconsistencia de encabezados `# Ronda 264` y `# Ronda 265` en `implementation_plan.md` y referencia pendiente en `IViaticosService.yml`.

## Cambios Aplicados
1. `REGISTRY.md`: Estado de `R-003` ajustado a `executing` con `verify = pendiente (gate Full)`.
2. `ViaticosGestionProcessor.cs`: Contador `sinCedula` y traza `LoggerHelper.Warning` por fila en `GetSolicitudes`; filtrado en N0 descartando candidatos sin `CC` o `Nom` con retorno de `SinComision`; normalización de carpeta con `TrimEnd` para `r.Detalle`.
3. `FormGestionDocumental.cs`: Desglose completo de archivos pendientes en la rama `YaCorrectos > 0`.
4. `implementation_plan.md`: Normalizados encabezados `# Ronda 264:` y `# Ronda 265:` a nivel `## Ronda`.
5. `IViaticosService.yml`: Añadida referencia completa a `ResultadoRenombrado`.

---

## Ronda 271: Verificación de parches 1451–1456 y desfases de trazabilidad (Fases 1457–1460)

## Contexto y Diagnóstico
- **Fase 1457**: Descarte de filas por `Nom` vacío en N0 enmascaraba registros incompletos como `SIN_COMISION_EN_CONSULTAS`; avisos warning por fila en `GetSolicitudes` inundaban el log.
- **Fase 1458**: Contadores de §2 en `REGISTRY.md` permanecían desfasados en Ronda 268 / Fase 1442 con riesgo de reutilización indebida de identificadores globales.
- **Fase 1459**: La rama final de `FormGestionDocumental.cs` diagnosticaba de forma estática "Verifique carpeta Solicitudes" ignorando la causa real por ausencia de filas en `CONSULTAS`.
- **Fase 1460**: Ausencia física de los archivos DocFX `.yml` para `ResultadoRenombrado` y `MotivoRenombrado` bajo `docs/api/`.

## Cambios Aplicados
1. `ViaticosGestionProcessor.cs`: Agregado motivo `MotivoRenombrado.SinDatos`, contador `r.SinDatos` en resumen y bucle; agrupación de warning en `GetSolicitudes` a un único resumen de advertencia; separación en N0 entre comisiones inexistentes (`SinComision`) e incompletas (`SinDatos`).
2. `ResultadoRenombrado.cs` y `MotivoRenombrado.cs`: Incorporada propiedad y constante `SinDatos`.
3. `REGISTRY.md`: Contadores globales actualizados a Ronda 271 y Fase 1460.
4. `FormGestionDocumental.cs`: Desglose dinámico de causales en la rama `else` distinguiendo comisiones faltantes en `CONSULTAS` de problemas de nomenclatura de archivos.
5. `docs/api/`: Creados `ConversorPDF.XLL.Core.Models.ResultadoRenombrado.yml` y `ConversorPDF.XLL.Core.Models.MotivoRenombrado.yml`, y registrados en `ConversorPDF.XLL.Core.Models.yml`.

---

## Ronda 272: Incremento de versión a 1.6.3.83 e instalador Wix (Fase 1461)

## Contexto y Diagnóstico
- **Fase 1461**: Incremento oficial de versión de producto a `1.6.3.83` tras la consolidación de la política de desambiguación N0–N3, soporte dual para la subcarpeta `Solicitudes`, modelos de auditoría y correcciones defensivas de gestión documental.

## Cambios Aplicados
1. `App.cs`: `APP_VERSION` actualizado a `"1.6.3.83"`.
2. `ConversorPDF.XLL.csproj`: Propiedades `<Version>`, `<AssemblyVersion>` y `<FileVersion>` actualizadas a `1.6.3.83`.
3. `InstaladorConversorPDF.wxs`: Atributos `Version` del paquete y `DefaultVersion` de componentes de plantillas (`GF-FO-023.xlsx` y `FV-FO-033.xlsx`) actualizados a `1.6.3.83`.
4. `REGISTRY.md`: Contadores globales actualizados a Ronda 272 y Fase 1461.


---

## Ronda 276: Revisión del bump a 1.6.3.83 (Ronda 272 / Fase 1461) y reserva de R-005 (Fases 1468–1472)

## Contexto y Diagnóstico
- **Fase 1468**: La versión 1.6.3.83 se elevó antes de cerrar formalmente `R-003` con gate y antes de implementar `R-005`. Se adopta la decisión de fijar 1.6.3.83 como versión objetivo sin empaquetar MSI hasta certificar gate verde.
- **Fase 1469**: `DefaultVersion="1.6.3.83"` en `filTemplate1` y `filTemplate2` forzaba purga incondicional de plantillas no modificadas. Se preserva `DefaultVersion="1.6.3.82"` en `InstaladorConversorPDF.wxs`.
- **Fase 1470**: Ronda 272 / Fase 1461 carecía de fila en `REGISTRY.md` §3. Se añade formalmente la fila `R-004` (mantenimiento bump).
- **Fase 1471**: La reserva 273–275 / 1462–1467 existía solo en memoria/chat. Se formaliza en `REGISTRY.md` como fila `R-005` (`reserved`).
- **Fase 1472**: Deriva documental en el plan de fix de gestión documental (`2026-10-31-...-fix-plan.md`) y ausencia de bitácora para 1.6.3.83.

## Cambios Aplicados
1. `InstaladorConversorPDF.wxs`: `DefaultVersion` de `filTemplate1` y `filTemplate2` fijado en `"1.6.3.82"` con comentarios explicativos para evitar purgas innecesarias.
2. `REGISTRY.md`: Añadida fila `R-004` (`closed`) para el bump de versión y fila `R-005` (`reserved`) para `2026-11-01-verificacion-solicitudes-plan.md`; contadores actualizados a Ronda 276 y Fase 1472.
3. `2026-10-31-gestion-documental-carpeta-solicitudes-fix-plan.md`: Actualizada tabla de partida (§0) y tarea de cierre (§8) registrando el bump a 1.6.3.83 y la prohibición de generar MSI antes del gate.
4. `docs/articles/changes/2026-10-31-version-1.6.3.83.md`: Creada bitácora oficial de cambios para la versión 1.6.3.83.


---

## Ronda 277: Verificación de parches 1468–1472 (bump 1.6.3.83 / WiX / registro) (Fases 1473–1476)

## Contexto y Diagnóstico
- **Fase 1473**: En `installer/InstaladorConversorPDF.wxs`, `<RemoveFile On="install">` anulaba el beneficio de restaurar `DefaultVersion="1.6.3.82"`, eliminando el archivo local en cada actualización.
- **Fase 1474**: La fila `R-005` en `REGISTRY.md` carecía de plan materializado en disco (`2026-11-01-verificacion-solicitudes-plan.md`).
- **Fase 1475**: La fila `R-004` en `REGISTRY.md` figuraba erróneamente como `closed` con `verify = pendiente` y `HEAD` desfasado.
- **Fase 1476**: Redacción incondicional en bitácora de versión 1.6.3.83 y referencia cruzada a R-004 en títulos de Ronda 276.

## Cambios Aplicados
1. `installer/InstaladorConversorPDF.wxs`: Cambiado `On="install"` a `On="uninstall"` en `PurgeOldTemplate1` y `PurgeOldTemplate2`, preservando las plantillas de usuario durante actualizaciones.
2. `.agents/plans/2026-11-01-verificacion-solicitudes-plan.md`: Creado el plan formal en disco con bloque `numbering:` (R-005, ronda_base 273, fase_base 1462, 3 waves, 6 tasks), `files_lock` y criterios v2.
3. `.agents/plans/REGISTRY.md`: Fila `R-004` ajustada a estado `executing` con `verify = pendiente (gate conjunto)` y `HEAD = pendiente`. Fila `R-005` con fecha sincronizada. Contadores actualizados a Ronda 277 y Fase 1476.
4. `docs/articles/changes/2026-10-31-version-1.6.3.83.md`: Redactado alcance en calidad de condicionado a gate verde de R-003+R-005 e incorporada documentación del cambio en `RemoveFile`.
5. `implementation_plan.md` y `walkthrough.md`: Normalizados títulos de Ronda 276 referenciando la reserva de R-005.


---

## Ronda 278: Revisión del plan R-005 y verificación de 1473–1476 (Fases 1477–1482)

## Contexto y Diagnóstico
- **Fase 1477**: En T3, el plan referenciaba controles inexistentes (`txtComisionInicial`/`txtComisionFinal`). Los controles reales en `FormGestionDocumental.Designer.cs` son `txtViaticoInicial` y `txtViaticoFinal`, requiriendo suscribir `Load += FormGestionDocumental_Load`.
- **Fase 1478**: D-2/T2 redactaban "raíz de OneDrive" con riesgo de barrido masivo en vigencias ajenas. Se acota a la carpeta del libro (`folder`) y subcarpeta `Solicitudes` vía `CarpetasDeSolicitudes(folder)`.
- **Fase 1479**: El plan `2026-11-01-verificacion-solicitudes-plan.md` omitía campos contractuales obligatorios (envelope con baseline 952, verify, hot paths L0, parallel execution groups, y 11 campos por tarea).
- **Fase 1480**: T6 declaraba sellar R-003 sin auditar ni registrar formalmente la deuda de Fases 1457–1460, ni declarar la condición de aborto si el gate falla.
- **Fase 1481**: T2 listaba columna `F` (inexistente en reglas) y omitía columna `C` (`CONSULTAS_ESTADO = 3`).
- **Fase 1482**: Fechas y descripciones de `R-004` y `R-005` desincronizadas en `REGISTRY.md`.

## Cambios Aplicados
1. `.agents/plans/2026-11-01-verificacion-solicitudes-plan.md`: Reescrito bajo el contrato formal completo de swarm (envelope, tabla de olas, 11 campos obligatorios por tarea T1–T6, controles reales `txtViaticoInicial`/`txtViaticoFinal`, suscripción `Load`, búsqueda acotada por `CarpetasDeSolicitudes`, columnas B, C, H, J de `ConstantsHelper` y auditoría de deuda de 1457–1460 previo al cierre).
2. `.agents/plans/REGISTRY.md`: Fechas y notas normalizadas en §3 (`R-004` con nota de mantenimiento sin plan; `R-005` con fecha `2026-11-01`). Contadores globales de §2 actualizados a Ronda 278 y Fase 1482.


---

## Ronda 279: Revisión de Ronda 278 (Fases 1477–1482) y cobertura funcional de R-005 (Fases 1483–1487)

## Contexto y Diagnóstico
- **Fase 1483**: `VerificarSolicitudesPresentes` no limpia marcas erróneas en H si `(B, J)` existe y `H` tiene `SIN SOLICITUD` o `FALTA SOLICITUD`, perpetuando las 2 filas reportadas.
- **Fase 1484**: Los 14 archivos con contenido idéntico al destino en `Solicitudes` retornan `Colision` en vez de moverse a `Solicitudes\_Duplicados\`.
- **Fase 1485**: Los 2 archivos con prefijo de 6 dígitos (`202426`, `202626`) quedan como `PREFIJO_INVALIDO` a pesar de tener candidato único y coincidencia de nombre en `CONSULTAS` (`20426`, `20626`).
- **Fase 1486**: Comparaciones literales de estado en columnas C y H sin normalización (`Trim().ToUpperInvariant()`) ni contemplación de variantes (`FALTA SOLICITUD`).
- **Fase 1487**: Falta de política explícita para 13 filas con marcas operativas (`ERROR`, `SIN INFORME`, `DUPLICADA`), con riesgo de sobrescritura indebida por `SIN SOLICITUD`.

## Cambios Aplicados / Especificados
1. `ViaticosGestionProcessor.cs`: Limpieza de marca errónea con `SetCellValue(ws, "H" + i, string.Empty)`, incremento de `r.CeldasLimpiadas++` y guardado condicional si `CeldasEscritas + CeldasLimpiadas > 0`.
2. `ViaticosGestionProcessor.cs`: Función `MismoContenido(file, destino)` (longitud + lectura secuencial byte a byte), traslado a subcarpeta `_Duplicados` con sufijo hex determinista si colisiona, y exclusión de escaneo recursivo.
3. `ViaticosGestionProcessor.cs`: Normalización determinista de prefijos de 6 dígitos que inician en "20" a comisiones de 5 dígitos con validación por token de nombre (`CoincideNombre`).
4. `ViaticosGestionProcessor.cs`: Helper estático `Norm(s)` para comparaciones invariantes a cultura y espacios en celdas de estado/informe.
5. `ViaticosGestionProcessor.cs`: Regla D-5b para preservar marcas operativas distintas de vacías/faltantes y sumar a `OmitidasOtrasMarcas`.
6. `.agents/plans/REGISTRY.md`: Asignación formal de rango 1483–1487 para materializar R-006 (`_Duplicados` + auditoría de H).

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

### Contexto y Diagnóstico
- En Modo AG, `FormVinculaciones` filtraba por `config.id_periodo` y `config.subunidad` (ocultos en UI), bloqueando 244 sesiones de la corrida #44 (periodo 2027-01).
- Además, el combo tomaba `MAX(id)` (#46, vacía/error), ignoraba corridas huérfanas con staging (#45) y calculaba contadores con SQL crudo desincronizado de la grilla.

### Cambios Especificados
1. **Fase 1488 (T1)**: DTOs `CorridaStagingItem`, `AgStagingResumen` y servicios puros `AgCorridaSelector`, `AgFiltrosPolicy`. Registro de `<Compile Include>` en `ConversorPDF.XLL.csproj`.
2. **Fase 1489 (T2)**: Métodos `GetCorridasStagingAsync` y `GetResumenStagingAsync` en `IVinculacionQueryRepository` y `VinculacionQueryRepository` con consulta compuesta `UNION ALL` que captura corridas huérfanas y previene exclusión de staging.
3. **Fase 1490 (T3)**: Extracción de helper estático `BuildWhere` en `VinculacionQueryRepository` para unificar el predicado del conteo y listado, y delegación en `IVinculacionListService` / `VinculacionListService`.

---

## Ronda 283: UI FormVinculaciones, Filtros Desacoplados y Layout Robusto (Fases 1491–1494)

### Contexto y Diagnóstico
- `FormVinculaciones` ejecutaba consultas prematuras con los filtros de producción de `config`, tenía SQL inline en `CargarCorridas` y presentaba solapamiento visual en `pnlFilters` sobre los botones de acción.

### Cambios Especificados
1. **Fase 1491 (T4)**: En `CargarCatalogosAsync`, condicionar la aplicación de periodo y subunidad de `config` a `AgFiltrosPolicy.AplicarPeriodoConfig(_modo)` y `AgFiltrosPolicy.AplicarSubunidadConfig(_modo)`.
2. **Fase 1492 (T5)**: Selector dinámico de corridas desacoplado en `FormVinculaciones` con soporte para corrida inicial (`idCorridaInicial`), carga ordenada y refresco en `DropDown`.
3. **Fase 1493 (T6)**: Reemplazo de SQL crudo por `GetResumenStagingAsync` y adición del botón `btnAjustarFiltrosCorrida` para sincronizar los filtros visibles con la corrida seleccionada en un clic.
4. **Fase 1494 (T7)**: Rediseño de layout en `FormVinculaciones.Designer.cs` (tercera fila para controles de corrida, `pnlFilters.Height 80 -> 105`) y paso de `_idCorridaCargada` desde `AgFormGeneradorHorarios.cs`.

---

## Ronda 284: Suite de Pruebas RED→GREEN y Sello de Calidad (Fase 1495)

### Contexto y Diagnóstico
- Garantizar que ninguna regresión afecte Modo Producción (952 pruebas baseline) y certificar el comportamiento de la selección de corridas y filtros en Modo AG.

### Cambios Especificados
1. **Fase 1495 (T8)**: Pruebas unitarias para `AgCorridaSelector` y `AgFiltrosPolicy`, integración SQLite en memoria para `VinculacionQueryRepositoryTests`, mocks en `VinculacionListServiceTests`, bitácora documental `2026-11-09-corridas-ag-formvinculaciones.md` y cierre en `REGISTRY.md`.


---

## Ronda 285: Falso positivo de certificación y defectos de contrato en R-007 (Fases 1496–1503)

### Contexto y Diagnóstico
- La Ronda 284 cerró con un falso positivo de certificación al declarar ACEPTADAS e implementadas las fases 1488–1495 sin código materializado en disco.
- Adicionalmente, el plan presentaba defectos críticos de contrato en C# (`CorridaStagingItem`), dependencia invertida T2/T3, semántica cruzada en `GetResumenStagingAsync`, manejo de centinela 0, riesgo de `async void` en UI y omisión de trazabilidad `- **fase**:`.

### Cambios Aplicados
1. **Fase 1496**: Reversión de consumo en `REGISTRY.md` (§2 vuelto a `281 / 1487`, siguiente `282 / 1488`; fila `R-007` en `reserved`), `walkthrough.md` y tareas del plan restablecidas a `pending`.
2. **Fase 1497**: Corrección de firma en `CorridaStagingItem`: `public bool TieneStaging => TotalSesiones > 0;` y `public string Etiqueta { get; set; }`.
3. **Fase 1498**: Resolución de dependencia invertida: `BuildWhere` estático se extrae en T2; T3 se delimita a delegación de servicios y declara `depends_on: [T2]`.
4. **Fase 1499**: Separación semántica de `GetResumenStagingAsync`: consulta A sin filtros UI para descriptores de la corrida y consulta B con `BuildWhere` para el conteo de la grilla.
5. **Fase 1500**: Manejo de centinela `idCorridaVal <= 0` en T6 para prevenir mensaje erróneo de corrida #0 al seleccionar "-- Todas las corridas --".
6. **Fase 1501**: Eliminación de `async void` en evento `DropDown` de `cmbCorrida` mediante `CargarCorridasSafeAsync()` con manejo de excepciones para proteger el STA thread.
7. **Fase 1502**: Declaración de `pnlFilters.MinimumSize = new Size(0, 105)` y anclajes en T7 para estabilidad en resoluciones y escalas DPI 100/125/150%.
8. **Fase 1503**: Incorporación de tabla de mapeo de olas y adición de `- **fase**: 1488` a `1495` en cada tarea T1..T8 del plan.


---

## Ronda 286: Sincronización de numeración y frontera de capa en R-007 (Fases 1504–1509)

### Contexto y Diagnóstico
- La Ronda 285 corrigió los defectos de contrato iniciales pero dejó desincronizada la numeración en `REGISTRY.md` (§2 no registraba 1496–1503, arriesgando reciclaje indebido de IDs).
- El contrato del plan ubicaba DTOs en `Application` siendo devueltos por `Domain.Interfaces` (violación de capas limpias).
- Consultas incompletas en Consulta A (`ProgramasCorrida` omitida, orden/separador de `GROUP_CONCAT`), población tardía de `Etiqueta` y anclaje ambiguo de controles de corrida.

### Cambios Aplicados
1. **Fase 1504**: Sincronización en `REGISTRY.md` §2 (`ronda: 285 / 286`, `fase: 1503 / 1504`) e inserción de la fila `R-008` (`closed`) que sella la revisión de corrección Ronda 285 (Fases 1496–1503).
2. **Fase 1505**: Reubicación de `CorridaStagingItem` y `AgStagingResumen` en `ConversorPDF.XLL.Academic.Domain.Models.Genetica` para preservar la regla arquitectónica de independencia de Domain.
3. **Fase 1506**: Integración de subconsulta SQL para `ProgramasCorrida` con orden determinista en la Consulta A de `GetResumenStagingAsync`.
4. **Fase 1507**: Normalización de `GROUP_CONCAT` en `PeriodosCorrida` (`ORDER BY IdPeriodo`, separador `', '`).
5. **Fase 1508**: Población garantizada de `Etiqueta` en origen (`GetCorridasStagingAsync`) antes de retornar a capas superiores.
6. **Fase 1509**: Definición de anclajes explícitos (`Top | Left`, ancho fijo 320 px) para controles de corrida en T7.


---

## Ronda 287: Reincidencia de numeración y frontera de capa limpia en R-007 (Fases 1510–1514)

### Contexto y Diagnóstico
- La numeración en `REGISTRY.md` no había avanzado para cubrir el rango 1504–1509 de la revisión Ronda 286.
- Población de `Etiqueta` dentro de `VinculacionQueryRepository` introducía una dependencia inversa indebida `Infrastructure -> Application` hacia `AgCorridaSelector`.
- La consulta SQL de `ProgramasCorrida` usaba `INNER JOIN` (riesgo de NULL por programas faltantes) y `SubunidadesCorrida` carecía de separador uniforme y orden determinista.
- El índice de Git presentaba estados mixtos staged/unstaged (`AM`/`MM`).

### Cambios Aplicados
1. **Fase 1510**: Normalización formal en `REGISTRY.md`: creación y cierre de `R-009` (Fases 1504–1509) y `R-010` (Fases 1510–1514), actualizando §2 a `ronda: 287 / 288`, `fase: 1514 / 1515`.
2. **Fase 1511**: Traslado de la población de `Etiqueta` a `VinculacionListService.GetCorridasStagingAsync` (`Application -> Domain`), manteniendo `VinculacionQueryRepository` como mero mapeador sin dependencias a `Application`.
3. **Fase 1512**: Consulta A ajustada con `LEFT JOIN programas pr` y fallback `'PROG ' || d.IdPrograma` para evitar omisiones de programas.
4. **Fase 1513**: Normalización de `SubunidadesCorrida` en Consulta A con filtro de no nulos, orden determinista y separador `', '`.
5. **Fase 1514**: Saneamiento del árbol Git: des-staged de archivos de plantilla `.doc` y unificación del estado de trabajo sin derivas de índice.


---

## Ronda 288: Integridad de artefactos ajenos al plan y determinismo de agregación (Fases 1515–1516)

### Contexto y Diagnóstico
- Los archivos `.doc` en `ExcelTemplates` no pertenecían a R-007; se constató paridad total con HEAD y se formaliza la regla de protección para archivos fuera de `files_lock`.
- La agregación `GROUP_CONCAT` en SQLite no garantiza formalmente el orden en consultas con múltiples periodos; se establece normalización determinista en C# o acotación a periodo único para pruebas.

### Cambios Aplicados
1. **Fase 1515**: Verificada integridad de plantillas `.doc` en `ExcelTemplates` contra HEAD; establecida regla de abstención de checkout/restore en rutas fuera de `files_lock`.
2. **Fase 1516**: Documentada pauta de determinismo en memoria (C#) para `PeriodosCorrida`, `SubunidadesCorrida` y `ProgramasCorrida`, garantizando estabilidad en suites de prueba con múltiples registros.
3. **Sincronización REGISTRY**: Incorporada fila `R-011` (`closed`) y contadores actualizados a `ronda: 288 / 289`, `fase: 1516 / 1517`.


---

## Ronda 293: Correcciones de Robustez y Coherencia en Recálculo y Vista por Fechas (Fases 1530–1536)

### Contexto y Diagnóstico
Revisión manual exhaustiva sobre la entrega R-015 (HEAD `e53d2f5`) detectó 6 hallazgos críticos de coherencia en UI, lógica de validador y visualización de violaciones:
1. **Fase 1530**: Al recalcular en el Módulo R, `_ultimoResultado` retenía en memoria la instancia previa de `MejorIndividuo`, provocando que `CargarModuloResultadosAsync` omitiera la rehidratación desde SQLite y los labels mantuvieran violaciones obsoletas.
2. **Fase 1531**: El botón `btnRecalcularViolaciones` solo se habilitaba en la carga por combo (`CargarCorridaDesdeComboAsync`); tras ejecutar el AG directamente en la sesión viva, permanecía en `Enabled = false`.
3. **Fase 1532**: `AgRecalculoViolacionesService` no validaba que el contexto en memoria (`IdPeriodo`, `IdPrograma`, `Subunidad`) coincidiera con la corrida que se pretendía recalcular, arriesgando sobreescritura de violaciones con datos ajenos.
4. **Fase 1533**: `ConstruirIndices()` en `AgContexto` carecía de validación de colecciones requeridas no nulas (`Aulas`, `Bloques`, `Asignaturas`, `Docentes`, `Sedes`, `PoolDocentes`), arriesgando `NullReferenceException`.
5. **Fase 1534**: Mutación de `contexto.IdCorrida = idCorrida` sobre la instancia compartida del llamador sin bloque `try/finally` para restaurar el valor previo.
6. **Fase 1535**: En `FormVerVinculaciones.cs`, `CargarViolacionesAsync` ejecutaba queries a `ag_violacion` aún en modo Producción, y seleccionaba únicamente `FirstOrDefault(id => id > 0)`, omitiendo el resaltado cuando la vista agrupaba múltiples corridas.
7. **Fase 1536**: Suite de pruebas unitarias (`AgRecalculoViolacionesServiceTests`) actualizada con cobertura explícita para rechazo por contexto ajeno (1532), contexto incompleto (1533) y preservación de `IdCorrida` previo (1534).

### Cambios Aplicados
1. **Fase 1530 (`AgFormGeneradorHorarios.cs`)**: Asignación `_ultimoResultado = null;` inmediatamente antes de `await CargarModuloResultadosAsync(idCorrida);` forzando la rehidratación fidedigna de las violaciones recalculadas.
2. **Fase 1531 (`AgFormGeneradorHorarios.cs`)**: En `CargarModuloResultadosAsync`, asignación `btnRecalcularViolaciones.Enabled = ind.Id > 0;` garantizando disponibilidad del botón tanto desde la ejecución directa del AG como tras rehidratar desde persistencia.
3. **Fase 1532 (`AgRecalculoViolacionesService.cs`)**: Guardas de coherencia entre `contexto` y `corrida` (`IdPeriodo`, `IdPrograma`, `Subunidad`), retornando `Result.Failure` explicativo si difieren.
4. **Fase 1533 (`AgRecalculoViolacionesService.cs`)**: Validación previa de colecciones no nulas en `contexto` antes de disparar `ConstruirIndices()`, retornando fallo claro de negocio.
5. **Fase 1534 (`AgRecalculoViolacionesService.cs`)**: Manejo con variable `idCorridaPrevio` y bloque `try/finally` para salvaguardar y restaurar el `IdCorrida` del contexto compartido.
6. **Fase 1535 (`FormVerVinculaciones.cs`)**: Guarda inmediata `if (!_modo.EsModoAg) return;` y consulta SQL adaptada con `IN @IdCorridas` soportando `Distinct().ToList()` de todas las corridas visibles.
7. **Fase 1536 (`AgRecalculoViolacionesServiceTests.cs`)**: Tres pruebas unitarias añadidas (`RecalcularAsync_ContextoNoCorrespondeACorrida_DevuelveFallo`, `RecalcularAsync_ContextoIncompleto_DevuelveFallo`, `RecalcularAsync_PreservaIdCorridaPrevioDelContexto`).
8. **Sincronización REGISTRY**: Incorporada fila `R-016` (`closed`) y actualizados contadores a `ronda: 293 / 294`, `fase: 1536 / 1537`.


---

## Ronda 294: Refinamiento de Coherencia en Recálculo y Vista por Fechas (Fases 1537–1540)

### Contexto y Diagnóstico
Revisión manual sobre la ronda 293 identificó 4 defectos adicionales:
1. **Fase 1537**: Tras forzar rehidratación (`_ultimoResultado = null`), `GetIndividuoByIdAsync` no carga la colección de violaciones, dejando `dgvViolaciones` vacío en la UI aunque las métricas escalares estuvieran actualizadas.
2. **Fase 1538**: La comprobación estricta de `Subunidad` rechazaba recálculos legítimos cuando la sesión arrancaba sin sync (`contexto.Subunidad = ""` o nulo) y la corrida tenía una territorial fijada.
3. **Fase 1539**: `btnRecalcularViolaciones` se habilitaba antes de verificar staging; ante fallo de lectura o conteo 0 de sesiones, `btnPromoverProduccion` se deshabilitaba pero `btnRecalcularViolaciones` permanecía activo.
4. **Fase 1540**: `AgContexto.ConstruirIndices()` itera sobre `Festivos` sin null-check y requiere `AulaAsignaturas`; F1533 no validaba estas dos colecciones, permitiendo potenciales `NullReferenceException`.

### Cambios Aplicados
1. **Fase 1537 (`AgFormGeneradorHorarios.cs`)**: Asignación de `ind.Violaciones` y enlace a `dgvViolaciones.DataSource` inmediatamente después de `CargarModuloResultadosAsync` para reflejar el detalle recalculado devuelto por el servicio.
2. **Fase 1538 (`AgRecalculoViolacionesService.cs`)**: Comparación condicional de territorial: solo evalúa coincidencia si ambos lados (`contexto` y `corrida`) aportan un valor no vacío.
3. **Fase 1539 (`AgFormGeneradorHorarios.cs`)**: En `CargarModuloResultadosAsync`, deshabilitación explícita `btnRecalcularViolaciones.Enabled = false;` en los retornos tempranos de `sesionesResult.IsFailure` y `sesionesResult.Value.Count == 0`.
4. **Fase 1540 (`AgRecalculoViolacionesService.cs` y tests)**: Validación añadida para `contexto.Festivos == null || contexto.AulaAsignaturas == null` en guardas previas del servicio y prueba unitaria en `AgRecalculoViolacionesServiceTests.cs`.
5. **Sincronización REGISTRY**: Incorporada fila `R-017` (`closed`) y actualizados contadores a `ronda: 294 / 295`, `fase: 1540 / 1541`.


---

## Ronda 295: Deshabilitación Defensiva del Recálculo sin Individuo Óptimo (Fase 1541)

### Contexto y Diagnóstico
En `AgFormGeneradorHorarios.cs`, la salida temprana `if (_ultimoResultado?.MejorIndividuo == null) return false;` (línea 1979) no desactivaba `btnRecalcularViolaciones`. Si una corrida previa había habilitado el botón y luego se cargaba una corrida sin individuo óptimo en base de datos, el botón permanecía en `Enabled = true`, provocando al presionarlo una excepción de negocio ("no tiene individuo óptimo persistido").

### Cambios Aplicados
1. **Fase 1541 (`AgFormGeneradorHorarios.cs:1979-1984`)**: Bloque explícito `{ btnRecalcularViolaciones.Enabled = false; btnPromoverProduccion.Enabled = false; return false; }` ante ausencia de individuo óptimo.
2. **Sincronización REGISTRY**: Incorporada fila `R-018` (`closed`) y actualizados contadores a `ronda: 295 / 296`, `fase: 1541 / 1542`.


---

## Ronda 296: Selector de Corridas del Módulo R: Filtro Tolerante y Estado Vacío Saneado (Fases 1542–1544)

### Contexto y Diagnóstico
En el vault real:
- `ag_corrida`: corrida #1 (período 2 · programa 1 · VALLE · 18 vinc).
- `ag_vinculaciones`: corrida #48 (período 3 · programa 2 · VALLE · 15 vinc, sin cabecera).
- `config/ActiveSubunidad`: período 2 · VALLE.

`VinculacionQueryRepository.GetCorridasStagingAsync` aplica igualdad estricta en `idperiodo`, `idprograma` y `UPPER(TRIM(subunidad))`. El Módulo R le pasaba los tres; con `período 2 · programa 1` no encontraba la corrida #48, resultando en 0 filas y combo con un ítem de tipo string (`Items.Add("— sin corridas … —")`), provocando recorte visual y selección no tipada.

1. **Fase 1542**: Un solo intento estricto de filtro período+programa dejaba el selector vacío aunque existieran corridas con staging en la territorial.
2. **Fase 1543**: Estado vacío o de error insertaba cadenas `string` en `cmbCargarCorridas.Items` conservando `DisplayMember`/`ValueMember` tipados.
3. **Fase 1544**: El evento `cmbCargarCorridas.DropDown` reconsultaba SQLite y limpiaba `Items` en cada apertura del desplegable aunque los filtros no hubieran cambiado.

### Cambios Aplicados
1. **Fase 1542 (`AgCorridaSelector.cs` y `AgFormGeneradorHorarios.cs`)**:
   - `AgCorridaSelector.CascadaFiltros(idPeriodo, idPrograma)`: función pura que retorna cascada tolerante `(p, pr)`, `(p, null)` y `(null, null)` incondicional (deduplicada). La territorial se mantiene fija.
   - En `AgFormGeneradorHorarios.RefrescarComboCorridasAsync`: iteración sobre la cascada con salida anticipada al primer intento no vacío, marcando `ampliado = true` si se relajaron filtros.
2. **Fase 1543 (`AgFormGeneradorHorarios.cs`)**:
   - Limpieza previa de enlace de datos (`DisplayMember = string.Empty`, `ValueMember = string.Empty`, `Items.Clear()`, `SelectedIndex = -1`, `Text = string.Empty`).
   - Si `error != null` o `items.Count == 0`, no se agregan strings a `Items`; se actualiza `lblAgModo` y se desactivan botones defensivamente solo si `_idCorridaCargada <= 0`. Si `ampliado` y `_idCorridaCargada <= 0`, se notifica en `lblAgModo`.
   - `DisplayMember` y `ValueMember` se asignan solo cuando `items.Count > 0` antes de poblar `Items`.
3. **Fase 1544 (`AgFormGeneradorHorarios.cs`)**:
   - Campo `_claveFiltroCorridas` y método `ClaveFiltroCorridas()` para cachear la combinación `Periodo|Programa|Subunidad`.
   - En `cmbCargarCorridas.DropDown`: omisión de reconsulta si `Items.Count > 0` y la clave de filtros coincide con `_claveFiltroCorridas`. Asignación de `_claveFiltroCorridas` al término del refresco exitoso.
4. **Pruebas (`AgCorridaSelectorTests.cs`)**:
   - 4 tests unitarios de cascada (`ConPeriodoYPrograma_DevuelveTresIntentos`, `SinPrograma_DevuelveDosIntentos`, `SinPeriodo_DevuelveDosIntentos`, `TodoNulo_DevuelveUnIntento`) y 1 theory `SiempreTerminaAmplio`.
5. **Sincronización REGISTRY**:
   - Fila `R-019` (`closed`) y contadores actualizados a `ronda: 296 / 297`, `fase: 1544 / 1545`.


---

## Ronda 297: Tipado de Asserts en Tests y Contexto de Corrida en Recálculo (Fases 1545–1548)

### Contexto y Diagnóstico
Revisión manual sobre la entrega R-019 (Ronda 296) identificó 4 defectos:
1. **Fase 1545 (Compilation)**: FluentAssertions 6.12 carece de sobrecarga `Should().Be()` para `ValueTuple` y resuelve a `ObjectAssertions.Be(object)`. Los literales con `null` carecían de tipo natural provocando error de compilación `CS8716`, y `(2, 1)` boxeaba `(int, int)` que no igualaba a `(int?, int?)`.
2. **Fase 1546 (Logic)**: Al auto-cargar una corrida ampliada (ej. #48 de período 3/programa 2 con combos en período 2/programa 1), «Recalcular violaciones» reconstruía `_contextoActual` usando los combos de la UI, siendo rechazado por la guarda de `AgRecalculoViolacionesService` (F1532) por discordancia de contexto.
3. **Fase 1547 (Warning)**: El bloque `finally` de `RecalcularViolacionesDesdeModuloRAsync` re-habilitaba incondicionalmente `btnRecalcularViolaciones.Enabled = true`, contradiciendo las guardas de F1539/F1541 para corridas sin individuo óptimo persistido.
4. **Fase 1548 (Warning)**: La actualización de `dgvViolaciones` estaba supeditada a `ind != null`; si el individuo rehidratado era nulo tras el recálculo (ej. staging vacío), la grilla retenía violaciones obsoletas.
5. **Fase 1545-bis (Warning)**: En `RefrescarComboCorridasAsync`, la salida temprana por `items.Count == 0` no fijaba `_claveFiltroCorridas`, repitiendo las 3 consultas en cada desplegado subsiguiente.

### Cambios Aplicados
1. **Fase 1545 (`AgCorridaSelectorTests.cs`)**:
   - Uso de arrays tipados explícitamente `(int?, int?)[]` (`EsperadoPeriodoYPrograma`, etc.) con `.Should().Equal()`, evitando inferencia a `object`.
   - Extracción de `.Subject` y `.Last()` con aserciones escalares sobre `IdPeriodo` e `IdPrograma`.
2. **Fase 1546 (`AgFormGeneradorHorarios.cs`)**:
   - Parámetros opcionales en `ConstruirContextoAsync(ct, idPeriodoForzado, idProgramaForzado, subunidadForzada)`.
   - Consulta previa de la corrida mediante `_scheduleRepository.GetCorridaByIdAsync(idCorrida)` antes de evaluar `_contextoActual`. Si el contexto difiere en período o programa, se reconstruye forzando los metadatos de la corrida.
   - Sincronización de `ActualizarVentanaSemestralAsync(periodoSel)` con el período forzado.
3. **Fase 1547 (`AgFormGeneradorHorarios.cs`)**:
   - En el `finally` de `RecalcularViolacionesDesdeModuloRAsync`, asignación condicional: `btnRecalcularViolaciones.Enabled = _ultimoResultado?.MejorIndividuo?.Id > 0;`.
4. **Fase 1548 (`AgFormGeneradorHorarios.cs`)**:
   - Asignación y refresco de `dgvViolaciones.DataSource` fuera del bloque condicional `if (ind != null)`, garantizando limpieza y coherencia del panel.
5. **Fase 1545-bis (`AgFormGeneradorHorarios.cs`)**:
   - Asignación `_claveFiltroCorridas = ClaveFiltroCorridas();` en la salida temprana de `items.Count == 0`.
6. **Sincronización REGISTRY**:
   - Fila `R-020` (`closed`) y contadores actualizados a `ronda: 297 / 298`, `fase: 1548 / 1549`.


---

## Ronda 298: Guardas de Cobertura de Aulas y Validación de Cabecera en Recálculo (Fases 1549–1550)

### Contexto y Diagnóstico
Revisión manual sobre la entrega R-020 (Ronda 297) detectó 2 riesgos críticos:
1. **Fase 1549 (Logic)**: Al reconstruir el contexto forzando período y programa de la corrida, la selección de sede y cohortes sigue dependiendo de `cmbSedes`/árbol. Si la corrida pertenece a otra sede/cohorte, el contexto no contiene sus aulas. Como `AgValidadorReglas` omite evaluar genes cuyas aulas no están en `AulasPorId`, `individuo.Violaciones` resultaba vacío o parcial, persistiendo `hard_violations = 0` y habilitando indebidamente la promoción a producción.
2. **Fase 1550 (Warning)**: Para corridas huérfanas en staging sin cabecera en `ag_corrida` (`IdPeriodo <= 0` o `IdPrograma <= 0`), la reconstrucción ignoraba los valores forzados en silencio y fallaba con un mensaje genérico de error de contexto.

### Cambios Aplicados
1. **Fase 1549 (`AgRecalculoViolacionesService.cs` y tests)**:
   - Guarda explícita que extrae `aulasFaltantes` comparando las aulas con `IdAula > 0` de los genes del individuo contra `contexto.AulasPorId`.
   - Si existen aulas faltantes, retorna `Result.Failure` explicativo indicando los IDs ausentes y requiriendo seleccionar la sede/cohortes correspondientes para evitar corrupción de `ag_violacion`.
   - Suite de pruebas actualizada: `AgRecalculoViolacionesServiceTests.cs` incorpora `RecalcularAsync_ContextoNoCubreAulasDeLaCorrida_DevuelveFallo` y fixtures con cobertura de aulas base.
2. **Fase 1550 (`AgFormGeneradorHorarios.cs`)**:
   - Validación temprana previa a la reconstrucción del contexto: si `corrida.IdPeriodo <= 0 || corrida.IdPrograma <= 0`, muestra advertencia clara informando que la corrida carece de cabecera en `ag_corrida`.
3. **Sincronización REGISTRY**:
   - Fila `R-021` (`closed`) y contadores actualizados a `ronda: 298 / 299`, `fase: 1550 / 1551`.


---

## Ronda 299: Cobertura Integral de Contexto (Aulas, Asignaturas y Docentes) en Recálculo (Fase 1551)

### Contexto y Diagnóstico
Revisión manual sobre la entrega R-021 (Ronda 298) detectó:
1. **Fase 1551 (Logic)**: La guarda F1549 comparaba exclusivamente `AulasPorId`. `AgValidadorReglas` omite igualmente la evaluación de reglas si la asignatura o el docente no están presentes en `contexto.AsignaturasPorId` o `contexto.DocentesPorId`. Si el contexto reconstruido cubre las aulas pero no las asignaturas o docentes de los genes (desfase de árbol de cohortes), las violaciones quedaban parciales y `ReemplazarViolacionesAsync` persistía métricas disminuidas.

### Cambios Aplicados
1. **Fase 1551 (`AgRecalculoViolacionesService.cs` y tests)**:
   - Ampliada la guarda de cobertura para calcular `asignaturasFaltantes` (`IdAsignatura > 0` vs `AsignaturasPorId`) y `docentesFaltantes` (`IdTercero > 0` vs `DocentesPorId`).
   - Si cualquiera de los tres conjuntos tiene elementos, aborta con `Result.Failure` descriptivo indicando las cantidades e IDs faltantes.
   - Fixture de tests actualizado con asignaturas y docentes base en `ContextoVacio()`, y agregado test `RecalcularAsync_ContextoSinAsignaturasDeLaCorrida_DevuelveFallo`.
2. **Sincronización REGISTRY**:
   - Fila `R-022` (`closed`) y contadores actualizados a `ronda: 299 / 300`, `fase: 1551 / 1552`.


---

## Ronda 300: Ajuste de Guarda de Cobertura en Recálculo a Ejes Activos del Validador (Fases 1552–1553)

### Contexto y Diagnóstico
Revisión manual sobre la entrega R-022 (Ronda 299) identificó:
1. **Fase 1552 (Logic)**: `AgValidadorReglas` no consulta `DocentesPorId` en ninguna de sus reglas (0 referencias). Por consiguiente, verificar `docentesFaltantes` no aportaba protección alguna frente a violaciones omitidas, y provocaba el rechazo injustificado de recálculos legítimos cuando la corrida contenía docentes no registrados en el contexto actual (docentes interinstitucionales, de otra sede o fuera del pool).
2. **Fase 1553 (Warning)**: Comentario obsoleto en cabecera de la guarda que mencionaba docentes pese a haber sido excluidos en F1552.

### Cambios Aplicados
1. **Fase 1552 (`AgRecalculoViolacionesService.cs` y tests)**:
   - Removida la comprobación de `docentesFaltantes` de la guarda bloqueante en `AgRecalculoViolacionesService.cs`, restringiendo la validación estricta a los dos ejes efectivamente consultados por las reglas (`AulasPorId` y `AsignaturasPorId`).
   - Agregada prueba unitaria `RecalcularAsync_DocenteAusenteDelContexto_NoBloqueaRecalculo` en `AgRecalculoViolacionesServiceTests.cs`.
2. **Fase 1553 (`AgRecalculoViolacionesService.cs`)**:
   - Comentario alineado con la regla activa: «abortar si el contexto no cubre las aulas o asignaturas de la corrida».
3. **Sincronización REGISTRY**:
   - Fila `R-023` (`closed`) y contadores actualizados a `ronda: 300 / 301`, `fase: 1553 / 1554`.


---

## Ronda 302: La selección programática del selector de corridas nunca dispara la carga (Fases 1554–1555)

### Contexto y Diagnóstico
Revisión manual de la integración 1542–1552 contra las capturas de ejecución (Wave 301):
1. **Fase 1554 (Logic)**: `_actualizandoComboCorridas` descarta la carga de la corrida preseleccionada. El refresco del combo fija `SelectedIndex` dentro del bloque con `_actualizandoComboCorridas = true`; el handler `SelectedIndexChanged` retorna inmediatamente por guard (`:2329`) y no existía carga posterior tras el `finally`. Al haber un único ítem en el combo (`#48`), el usuario no puede forzar un cambio de selección manual y la corrida nunca se carga en el Módulo R.
2. **Fase 1555 (Warning)**: Métricas y grids obsoletos en Módulo R cuando la carga no se completa. Si la carga falla o es rechazada, `lblAptitud`, `lblHardViolaciones`, `lblVacantes`, `dgvHorariosGenerados` y `dgvViolaciones` mantenían los valores de una corrida previa o estado no representativo (p. ej. `1.000 / 0 / 0`).

### Cambios Aplicados
1. **Fase 1554 (`AgFormGeneradorHorarios.cs`)**:
   - Forzada la invocación explícita `await CargarCorridaDesdeComboAsync();` inmediatamente después del bloque `try-finally` que libera `_actualizandoComboCorridas = false`, siempre que `cmbCargarCorridas.SelectedItem` sea un `CorridaStagingItem` válido con `IdCorrida != _idCorridaCargada`.
2. **Fase 1555 (`AgFormGeneradorHorarios.cs`)**:
   - En la guarda `if (!cargada)` de `CargarCorridaDesdeComboAsync`, limpiar explícitamente etiquetas de métricas a sus valores neutrales (`—`) y desvincular los data sources de `dgvHorariosGenerados` y `dgvViolaciones`.
3. **Sincronización REGISTRY**:
   - Fila `R-024` (`closed`) y contadores actualizados a `ronda: 302 / 303`, `fase: 1555 / 1556`.


---

## Ronda 303: Guarda de Reentrada en RefrescarComboCorridasAsync (Fase 1556)

### Contexto y Diagnóstico
Revisión manual sobre la entrega Wave 302 / Ronda 302 (Fases 1554–1555):
1. **Fase 1556 (Logic)**: `RefrescarComboCorridasAsync` carecía de una guarda de ejecución concurrente. El refresco del selector se dispara desde dos vías asíncronas sin serialización: en la carga inicial (`AgFormGeneradorHorarios.cs:270`) y en cada evento `DropDown` (`:404`). Si el usuario interactúa con el combo mientras aún corre la carga inicial, ambas ejecuciones manipulaban concurrentemente los ítems de `cmbCargarCorridas` y, tras la incorporación de la Fase 1554, ambas detonaban `CargarCorridaDesdeComboAsync` en paralelo, generando una condición de carrera que podía dejar las colecciones de `dgvHorariosGenerados`, `dgvViolaciones` y `_idCorridaCargada` desalineadas.

### Cambios Aplicados
1. **Fase 1556 (`AgFormGeneradorHorarios.cs`)**:
   - Declarado el campo `private bool _refrescandoComboCorridas;`.
   - Agregada guarda al inicio de `RefrescarComboCorridasAsync`: `if (cmbCargarCorridas == null || _refrescandoComboCorridas) return;` con asignación `_refrescandoComboCorridas = true;`.
   - Reestructurado el manejo de excepciones para incorporar un bloque `finally { _refrescandoComboCorridas = false; }` que garantiza la liberación de la guarda en cualquier camino de salida o error.
2. **Sincronización REGISTRY**:
   - Fila `R-025` (`closed`) y contadores actualizados a `ronda: 303 / 304`, `fase: 1556 / 1557`.


---

## Ronda 304: Carga del Staging Sin Individuo Óptimo y Advertencia en Selector (Fases 1557–1558)

### Contexto y Diagnóstico
Revisión manual y diagnóstico con datos reales de vault (#48 con 180 sesiones en staging pero sin cabecera en `ag_corrida` ni fila en `ag_individuo`):
1. **Fase 1557 (Logic)**: En `CargarModuloResultadosAsync`, el early-return por falta de `MejorIndividuo` ocurría antes de consultar `GetSesionesStagingAsync`. Corridas con staging materializado pero sin resultado AG (o staging huérfano) retornaban `false`, limpiando grids y abortando la visualización.
2. **Fase 1558 (Warning)**: El selector de corridas (`AgCorridaSelector.FormatearEtiqueta`) no indicaba si una corrida disponía o no de resultado genético persistido, induciendo al usuario a esperar violaciones y métricas inexistentes.

### Cambios Aplicados
1. **Fase 1557 (`AgFormGeneradorHorarios.cs`)**:
   - Extraído el método `private async Task<bool> CargarGridStagingAsync(long idCorrida)` para desacoplar la carga de grillas de la rehidratación del individuo óptimo.
   - En `CargarModuloResultadosAsync`, se invoca `CargarGridStagingAsync` en primer orden.
   - Si `_ultimoResultado?.MejorIndividuo == null`, se fija el estado visual de "staging sin resultado AG (solo lectura)", se deshabilitan botones de acción/promoción/recálculo, y se retorna `true` permitiendo inspeccionar las sesiones materializadas en `dgvHorariosGenerados`.
   - Agregada traza de diagnóstico en `CargarGridStagingAsync` que advierte si las sesiones referencian `IdIndividuo` huérfanos sin correspondencia en `ag_individuo`.
2. **Fase 1558 (`CorridaStagingItem.cs`, `VinculacionQueryRepository.cs`, `AgCorridaSelector.cs`, `AgCorridaSelectorTests.cs`)**:
   - Agregada propiedad `IdIndividuoOptimo` y getter `TieneIndividuoOptimo` a `CorridaStagingItem`.
   - En `VinculacionQueryRepository.GetCorridasStagingAsync`, proyectado `COALESCE(MAX(c.id_individuo_optimo), 0) AS IdIndividuoOptimo` (y en la rama UNION sin staging).
   - En `AgCorridaSelector.FormatearEtiqueta`, si `corrida.IdCorrida > 0 && !corrida.TieneIndividuoOptimo`, se añade la indicación `"sin resultado AG"`.
   - Pruebas unitarias actualizadas en `AgCorridaSelectorTests.cs` validando el formateo de etiquetas con y sin resultado AG.
3. **Sincronización REGISTRY**:
   - Fila `R-026` (`closed`) y contadores actualizados a `ronda: 304 / 305`, `fase: 1558 / 1559`.


---

## Ronda 305: Hidratación de ag_violacion al Rehidratar la Corrida en Módulo R (Fase 1559)

### Contexto y Diagnóstico
Revisión manual y diagnóstico del panel «Violaciones de Reglas» al seleccionar corridas en Módulo R:
1. **Fase 1559 (Logic)**: `dgvViolaciones` se alimenta únicamente desde `_ultimoResultado.MejorIndividuo.Violaciones`. Al rehidratar una corrida con resultado existente (como `#1` con 46 violaciones en SQLite), `GetIndividuoByIdAsync` solo recuperaba columnas de individuo y genes en `ag_gen`, omitiendo la lectura de `ag_violacion`. Como consecuencia, el panel de violaciones permanecía vacío salvo que el usuario ejecutara manualmente «Recalcular violaciones». Además, `IAgScheduleRepository` no exponía método lector para `ag_violacion`.

### Cambios Aplicados
1. **Fase 1559 (`IAgScheduleRepository.cs`, `AgScheduleRepository.cs`, `AgFormGeneradorHorarios.cs`, `AgScheduleRepositoryNullGuardTests.cs`)**:
   - Agregado a `IAgScheduleRepository` el método `Task<Result<IReadOnlyCollection<AgViolacion>>> GetViolacionesByIndividuoAsync(long idIndividuo, CancellationToken ct);`.
   - Implementado `GetViolacionesByIndividuoAsync` en `AgScheduleRepository` consultando `ag_violacion` filtrado por `id_individuo` y ordenado por `es_dura DESC, regla`.
   - En `AgFormGeneradorHorarios.CargarGridStagingAsync`, si `indActual.Id > 0` y su colección `Violaciones` está vacía o nula, invocar `GetViolacionesByIndividuoAsync` y asignarla a `indActual.Violaciones` antes de mapear hacia `dgvViolaciones`.
   - Añadido log explicativo ante corridas sin resultado AG: `[AG] Corrida sin resultado AG: no hay violaciones persistidas que mostrar.`.
   - Corregido cierre accidental de llave detectado en `AgFormGeneradorHorarios.cs:2250`.
   - Agregada prueba unitaria `GetViolacionesByIndividuoAsync_WhenViolationsExist_ReturnsOrderedByEsDuraDescAndRegla` en `AgScheduleRepositoryNullGuardTests.cs`.
2. **Sincronización REGISTRY**:
   - Fila `R-027` (`closed`) y contadores actualizados a `ronda: 305 / 306`, `fase: 1559 / 1560`.


---

## Ronda 306: Ajuste de Layout en Módulo R y Separación de Botón Promover (Fase 1562)

### Contexto y Diagnóstico
Revisión manual del layout del panel superior `pnlHeaderResultados` en Módulo R:
1. **Fase 1562 (GUI / Layout)**: `cmbCargarCorridas` poseía un ancho fijo de 185 px que truncaba visualmente las etiquetas descriptivas de las corridas (`#48 · 2027-01 · APT-52 · VALLE · 180 ses`). Además, `btnPromoverProduccion` estaba ubicado en posición fija `Location = (740, 10)` sin propiedad `Anchor`, provocando que en un contenedor de 1100 px quedara pegado al combo y fuera solapado por el desplegable del selector (`DropDownWidth = 430`).

### Cambios Aplicados
1. **Fase 1562 (`AgFormGeneradorHorarios.Designer.cs`)**:
   - Dimensionado de diseño de `pnlHeaderResultados.Size` actualizado a `(1100, 50)`.
   - Ampliado el ancho desplegable `cmbCargarCorridas.DropDownWidth = 480`.
   - Ampliado el control `cmbCargarCorridas.Size` de 185 a 330 px (`Size = (330, 25)`), cubriendo de X=545 a X=875.
   - Reubicado `btnPromoverProduccion` a `Location = (890, 10)` y configurado `Anchor = Top | Right`, garantizando 15 px de separación mínima en resolución base y anclaje al borde derecho al maximizar.
2. **Sincronización REGISTRY**:
   - Fila `R-028` (`closed`) y contadores actualizados a `ronda: 306 / 307`, `fase: 1562 / 1563`.

---

## Ronda 307: Purga sin huérfanos y reparación de corridas del AG (Fases 1563–1566)

### Contexto y Diagnóstico
La corrida **#48** (2027-01 · APT-52 · VALLE · 180 sesiones) quedó visible en el Módulo R pero sin
violaciones y sin posibilidad de promoción. Evidencia reconstructiva:

1. `log_20261109.txt:1522`: `FinCorridaAG IdCorrida=48 | Estado=Completada | Hard=95 | Vinculaciones=15 | Sesiones=180`
   y `19:51:24` «Corrida #48 marcada como Aprobada».
2. `log_20261110.txt:346`: `[Purga] vinc 33->33 | ses 325->325 | gen 70->70 | corridas 2->1 | idPeriodo=3 idPrograma=1 corridasPeriodo=48`.
3. Vault en vivo: `ag_vinculaciones` de la corrida 48 con `idprograma=2`; `ag_gen` con 30 genes del individuo 33
   (`idprograma=2`, aulas 12/30/31); **sin** fila 48 en `ag_corrida` y **sin** fila 33 en `ag_individuo`.
4. Catálogo `programas`: `id=1 → APT-53` y `id=2 → APT-52` con el **mismo nombre**.

**Causa raíz:** `AgFabricaIndividuos` estampa el gen con `asig.IdPrograma` (programa real de la asignatura),
mientras `AgCorrida.IdPrograma` proviene del combo del formulario. `LimpiarStagingAsync(idPeriodo, idPrograma)`
borraba la cabecera/individuo/violaciones por el par de `ag_corrida` `(3,1)` y el staging por el par de cada
fila `(3,2)`, dejando la corrida huérfana. Consecuencias: el panel de violaciones queda vacío (no hay
individuo que consultar), «Aprobar corrida»/«Promover» se deshabilitan y `PromoverAProduccionAsync`
retornaba en silencio.

### Cambios Aplicados
1. **Fase 1563 (`AgScheduleRepository.cs`)** — `LimpiarStagingAsync(int, int)`:
   - `idsDel` = cabeceras del par ∪ `DISTINCT id_corrida` del staging del par.
   - Todos los borrados se ejecutan por `id_corrida IN @idsDel` (staging y metadatos genéticos).
   - Limpieza de `ag_gen`/`ag_violacion` de individuos referenciados por el staging purgado aunque
     `ag_individuo` ya no exista.
   - Invariante post-purga con `LoggerHelper.Error` ante hijos residuales; guardas `TablaExiste`.
2. **Fase 1564 (`AgScheduleRepository.cs`, `AgFormGeneradorHorarios.cs`)**:
   - `UpdateCorridaAsync` falla si el `UPDATE` afecta 0 filas (antes era un no-op silencioso).
   - `EjecutarOptimizacionAsync` aborta antes de materializar staging cuando la cabecera no se consolidó.
   - Verificación del conteo real de `ag_violacion` tras `SaveIndividuoAsync`.
3. **Fase 1565 (`AgFormGeneradorHorarios.cs`)**:
   - `CargarModuloResultadosAsync` registra `corridaRes.Error`/`indRes.Error` y distingue cabecera ausente
     (huérfana), cabecera sin individuo y error de lectura; habilita «Recalcular violaciones» como vía de
     reparación y explica el estado en `lblAgModo`.
   - `PromoverAProduccionAsync` informa la causa en lugar de retornar en silencio.
4. **Fase 1566 (nuevos contratos + servicio)**:
   - `IAgScheduleRepository.GetGenesByIndividuoAsync`, `DetectarCorridaHuerfanaAsync`,
     `ReinsertarCabeceraHuerfanaAsync` (inserción con id explícito e idempotente).
   - `AgCorridaHuerfana`, `AgReparacionCorridaResultado`, `IAgReparacionCorridaService`,
     `AgReparacionCorridaService` (valida cobertura del contexto antes de escribir, evalúa con
     `AgEvaluadorAptitud` y persiste violaciones con `ReemplazarViolacionesAsync`).
   - `RecalcularViolacionesDesdeModuloRAsync` invoca la reparación cuando detecta staging sin cabecera.
   - Registro de los 4 archivos nuevos en `ConversorPDF.XLL.csproj`.

### Pruebas
- `AgScheduleRepositoryHuerfanasTests.cs` (nuevo): 6 casos de purga/reparación sobre SQLite temporal.
- `AgReparacionCorridaServiceTests.cs` (nuevo): 4 casos con Moq (cobertura, éxito, cabecera existente, sin genes).
- `scripts/verify.ps1 -Mode Fast -Filter "…AgScheduleRepositoryPurga|…AgScheduleRepositoryReparacion|…AgReparacionCorrida"` → 10/10 PASS.

### Sincronización REGISTRY
- Fila `R-029` (`executing`) y contadores actualizados a `ronda: 307 / 308`, `fase: 1566 / 1567`.

#### Fase 1567 — Nombre de regla en `ag_violacion` (detectada en la verificación end-to-end)
- **Diagnóstico:** `ag_violacion.regla` está declarada `INTEGER` pero almacena el nombre de la regla
  (`'R1_SuperposicionDocente'`, …). El proveedor SQLite aplica la afinidad de la columna al leer y devolvía `0`,
  de modo que el panel mostraba «regla 0» y la vista por fechas perdía el identificador.
- **Acción:** `CAST(regla AS TEXT)` en `AgScheduleRepository.GetViolacionesByIndividuoAsync` y en
  `FormVerVinculaciones.CargarViolacionesAsync`; prueba de regresión con la columna declarada `INTEGER`.
- **Verificación:** arnés end-to-end sobre copia del vault real: `[HARD] R1_SuperposicionDocente`,
  `[SOFT] S1_HuecosHorarios`.

---

## Ronda 308: Validador de Reglas AG y Oleadas Temporales (Fases 1568–1574)

### Resumen Ejecutivo
El validador de restricciones duras del algoritmo genético (`AgValidadorReglas`) marcaba colisiones espurias (p. ej. 101 HARD en la corrida #48) porque contrastaba únicamente la plantilla semanal (`DiaSemana`, `IdBloqueHorario`) ignorando las ventanas de fechas estimadas (`FechaInicioEstimada`, `FechaFinEstimada`). En cohortes avanzadas, múltiples asignaturas rotan en el mismo día y bloque en la misma aula física de manera modular (oleadas consecutivas sin solape real).

### Detalle de Fases

1. **Fase 1568 (`AgVentanaSolape.cs`, `AgVentanaSolapeTests.cs`)**:
   - Creación del helper estático puro `AgVentanaSolape` con método `SeSolapan(DateTime? iniA, DateTime? finA, DateTime? iniB, DateTime? finB)` y `AgruparPorSolape(IEnumerable<T>)`.
   - Semántica conservadora: cualquier extremo nulo asume solape (`true`).
   - Registro en `ConversorPDF.XLL.csproj`.
2. **Fase 1569 (`AgValidadorReglas.cs`)**:
   - Ajuste de colisiones R1 (docente), R2 (aula), R3 (grupo), R4 (dos asignaturas mismo día y franja) y R14 (bloques solapados en horas) para filtrar únicamente cuando `AgVentanaSolape.SeSolapan` sea verdadero.
3. **Fase 1570 (`AgValidadorReglas.cs`)**:
   - Refactorización de `R5_MaxAsignaturasSemana` de semanas ISO a conteo continuo día a día sobre el rango `[min(FechaInicio), max(FechaFin)]` de la cohorte contra `maxCarriles`.
4. **Fase 1571 (`AgValidadorReglas.cs`)**:
   - Admisión de `AgEstadoAula.Disponible` y `AgEstadoAula.Aprobado` en `R4_CompatibilidadAula`.
5. **Fase 1572 (`AgValidadorReglas.cs`)**:
   - Acotamiento de asignaturas requeridas al alcance de las aulas que participan en la corrida (`contexto.AulaAsignaturas`).
6. **Fase 1573 (`AgValidadorReglasOleadasTests.cs`)**:
   - Suite de pruebas de regresión con 6 escenarios de oleadas reales (caso #48), ventanas disyuntas, y controles cruzados.
7. **Fase 1574 (Verificación, Cierre y Documentación)**:
   - Verificación global `scripts/verify.ps1 -Mode Fast` (1123/1123 tests PASS).
   - Documentación técnica en `docs/articles/changes/2026-11-11-validador-oleadas.md`.
   - Cierre de asignación R-030 en `REGISTRY.md`.
   - **Verificación sobre datos reales (#48):** con el validador corregido y los 30 genes reales del individuo #33
     (contexto reconstruido desde `ag_aula`, `ag_bloque_horario`, `asignaturas` y pool docente) el resultado es
     **HARD = 0, SOFT = 0, 0 filas**; se confirman 0 solapes reales por aula, cohorte y docente, y que
     `ag_gen.dias_patron` está poblado (`"4,5"` / `"6"`), por lo que R6/R13 agrupan bien por carril.
   - **Corrección del fixture (F1573):** `AgValidadorReglasOleadasTests` usaba `Franja = "M"/"T"` en lugar del
     contrato canónico `MANANA`/`TARDE` (`AgFranjaJornada`), lo que producía un fallo espurio de
     `R11_DuracionSesionInvalida`; corregido, con diagnóstico de violaciones incluido en las aserciones.



---

## Ronda 309: Contexto Homónimo y Anclaje de Filtros en el Módulo R (Fases 1575–1578)

### Contexto y Diagnóstico
Con el validador de oleadas desplegado (Ronda 308) la UI seguía mostrando `125 HARD` y fallaban dos acciones:

1. **Fase 1575:** `ConstruirContextoAsync` forzaba el programa de la corrida (id 2) y después la resolución de
   homónimos (Fases 446/477) lo sustituía por el programa homónimo con más asignaturas del semestre (id 1)
   ⇒ `contexto=3/1` frente a `corrida=3/2`. `AgRecalculoViolacionesService` (Fase 1532) aborta y el validador nuevo
   nunca se ejecuta; el panel releyó las 101 filas persistidas el 11/11 a las 14:18 con el validador antiguo.
2. **Fase 1576:** `AbrirDetalleSesionesModoAgAsync` filtraba `vw_AgSesionesDetalle` por el programa del árbol (1),
   mientras el staging es programa (2) ⇒ «No se encontraron sesiones registradas…».
3. **Fase 1577:** `AgReparacionCorridaService` no validaba el período/programa del contexto, por lo que la
   reparación inicial evaluó con el pensum equivocado.
4. **Fase 1578:** faltaba prueba de regresión del contexto y verificación end-to-end del recálculo.

### Cambios Aplicados
1. **Fase 1575 (`AgFormGeneradorHorarios.cs`)**: con `idProgramaForzado > 0` se omite la resolución de homónimos;
   `ctx.ProgramasAlcance` se fija al programa de la corrida; traza `ProgramaHomonimoOmitido`. La lista
   `homonimosProg` se conserva a nivel de método porque también filtra aulas.
2. **Fase 1576 (`AgFormGeneradorHorarios.cs`)**: con corrida cargada, `idPeriodo`/`idPrograma` del detalle se leen de
   `ag_corrida` (log `[AG] Detalle staging: filtros anclados a la corrida #N`).
3. **Fase 1577 (`AgReparacionCorridaService.cs`)**: guarda que rechaza contextos con período/programa distinto al de
   la corrida **antes de escribir**.
4. **Fase 1578**: test `RepararAsync_ConContextoDeOtroPrograma_FallaSinEscribir`; suite **1124/1124 PASS**;
   recálculo real sobre copia del vault: **125 → 0** en `ag_corrida`/`ag_individuo` y **101 → 0** filas en
   `ag_violacion`; artículo `docs/articles/changes/2026-11-11-contexto-homonimo-modulo-r.md`.

### Sincronización REGISTRY
- Fila `R-031` (`closed`) y contadores actualizados a `ronda: 309 / 310`, `fase: 1578 / 1579`.



---

## Ronda 312: Partición de Carriles Consciente del Patrón Sembrado (Fase 1581)

### Contexto y Diagnóstico
La partición canónica en bloques fijos de 2 ranuras contiguas fragmentaba patrones sembrados (`dias_patron`, ej. `"6,7"` en `"4,5,6,7"`), impidiendo que las materias fueran asignadas al carril correspondiente y disparando fallos espurios de `R8_DiaFueraDePatron` e infactibles innecesarios.

### Cambios Aplicados (Fase 1581)
1. **Extracción de evaluador único:** En `AgPlanificadorOleadas.cs`, se extrajo la lógica de asignación y oleadas al método privado `EvaluarParticion(ContextoEvaluacion ctx, string diasVentana, IReadOnlyList<IReadOnlyList<AgRanura>> bandas)`.
2. **Reordenamiento defensivo:** Se verificó la restricción R8 antes de invocar la construcción de `AgPatronSemanal`, evitando cómputos descartados y preservando la salida textual exacta de `DetalleInfactibles`.
3. **Componente selector:** Se creó `AgParticionCarriles.cs` con el método puro `Seleccionar(...)`, aplicando filtros estructurales y las 4 condiciones de aceptación conservadora estricta (reducción de no alojables, preservación de alojables previos, no pérdida de asignaturas planificadas y reducción estricta de infactibles).
4. **Registro en proyecto:** Inclusión de `AgParticionCarriles.cs` en `ConversorPDF.XLL.csproj`.
5. **Gate G-1 y Cobertura:** Suites de caracterización, unitarios del selector, integración y Gate G-1 (`AgPlanificadorOleadasCaracterizacionTests`, `AgParticionCarrilesTests`, `AgParticionCarrilesIntegracionTests`, `AgGateG1ParticionCarrilesTests`). Suite general: **1156/1156 PASS**.

### Sincronización REGISTRY
- Fila `R-034` (`closed`) y contadores actualizados a `ronda: 312 / 313`, `fase: 1581 / 1582`.


---

## Ronda 313: Robustecimiento Integral, Fidelidad de Patrones y Corrección de Pruebas (Fases 1582–1589)

### Contexto y Diagnóstico
La revisión en disco constató discrepancias entre el informe previo y el estado del código:
1. **F1582:** El selector llamaba al evaluador ante listas vacías pudiendo acceder fuera de rango.
2. **F1583:** El selector retornaba temprano si `PatronesNoAlojables == 0` y el filtro `>=` descartaba alternativas con igual adecuación estructural que aportaran capacidad para materias de alta demanda.
3. **F1584:** `CandidatasEvaluadas` devolvía `candidatas.Count` en lugar del número real de llamadas al evaluador.
4. **F1585:** Ausencia de validación de `IdAsignatura` repetidos, provocando crash posterior en `AgFabricaIndividuos`.
5. **F1586:** Ausencia de validación de fidelidad exacta entre las ranuras asignadas y el patrón sembrado, ampliando los días permitidos en la fábrica.
6. **F1588:** Pruebas con `if (res.IsSuccess)` que pasaban sin validar el selector y festivo fuera de rango.

### Cambios Aplicados
1. **Fase 1582:** Guardas estrictas previas al delegado en `AgParticionCarriles.cs` y en `EvaluarParticion`.
2. **Fase 1585:** Guarda `I1_ASIGNATURA_DUPLICADA` en `AgPlanificadorOleadas.Construir`.
3. **Fase 1584:** Contador exacto `evaluadas` tras cada llamada a `evaluar(bandas)`.
4. **Fase 1586:** Delegado de admisibilidad `EsParticionFiel` exigiendo coincidencia estricta de días entre `plan.RanurasPatron` y `dias_patron`.
5. **Fase 1583:** Condición `noAloj > referencia.PatronesNoAlojables` y continuación de búsqueda cuando `PatronesNoAlojables == 0` si existen infactibles.
6. **Fase 1588:** Eliminación de `if (res.IsSuccess)` en `Caso76` y reubicación del festivo en `Caso78`.
7. **Fase 1589:** Sincronización íntegra de artefactos e implementación en disco.

### Sincronización REGISTRY
- Fila `R-035` (`closed`) y contadores actualizados a `ronda: 313 / 314`, `fase: 1589 / 1590`.


---

## Ronda 314: Separación de Casos de Integración, Diagnóstico I1 vs I5 y Gate con API Real (Fases 1590–1595)

### Contexto y Diagnóstico
1. **Fase 1590:** En `AgParticionCarrilesIntegracionTests`, el fixture de dos materias de 8h en W=1 reducía infactibles de 2 a 1 (mejora parcial con banda `{6,7}`), no probaba la ausencia total de mejora.
2. **Fases 1591 y 1592:** Se ratificó el rechazo a relajar la fidelidad del patrón (conservando `SetEquals`) y el rechazo a ocultar errores en guardas internas con retornos vacíos, manteniendo la estricta robustez del diseño.
3. **Fase 1593:** En `AgGateG1ParticionCarrilesTests`, no se comprobaba la concordancia de `gen.DiasPatron` con el patrón sembrado, y el fixture no asociaba `AgAulaAsignatura` con sus respectivos `DiasPatron`.
4. **Fase 1594:** En `AgValidadorVentana.cs`, el fallo `I1_ASIGNATURA_DUPLICADA` se reportaba bajo la etiqueta errónea `[I5_INFACTIBLE_OLEADA]`.
5. **Fase 1595:** En `AgGateG1ParticionCarrilesTests`, se llamaba a un método inexistente `validador.Validar` y se evaluaba `v.Regla` como string en vez de su API real `validador.Evaluar(individuo, contexto)` sobre el enum `AgReglaViolacion`.

### Cambios Aplicados
1. **Fase 1590:** Separación en dos pruebas explícitas: `Caso76_Integracion_MejoraParcial_DebeReducirInfactiblesYConservarFactible` (2 asignaturas de 8h, reduce de 2 a 1 infactible) y `Caso76_Integracion_AusenciaDeMejora_DebeConservarReferencia` (1 asignatura de 16h, 0 planes, 1 infactible, conserva las 3 bandas canónicas).
2. **Fases 1591 y 1592:** Código de producción intacto; preservada la fidelidad estricta y las excepciones tempranas ante entradas corruptas.
3. **Fase 1593 y 1595:** Refactorización integral de `AgGateG1ParticionCarrilesTests` configurando `AgAulaAsignatura` con sus patrones y bloques por franja/horario, verificando que `gen.DiasPatron` no se amplíe (`SetEquals`), que `gen.DiaSemana` pertenezca al patrón sembrado, cobertura completa de materias (38, 40, 42), y validación de reglas con `validador.Evaluar(individuo, contexto)` confirmando 0 violaciones duras R8, R5 y R13.
4. **Fase 1594:** En `AgValidadorVentana.cs`, clasificación diferenciada `I1_ASIGNATURA_DUPLICADA` vs `I5_INFACTIBLE_OLEADA`.

### Sincronización REGISTRY
- Fila `R-036` (`closed`) y contadores actualizados a `ronda: 314 / 315`, `fase: 1595 / 1596`.


---

## Ronda 315: Bloques Canónicos e Índices en Gate G-1, Habilitaciones en Pool Docente y Comprobación Temporal de Festivos (Fases 1596–1598)

### Contexto y Diagnóstico
1. **Fase 1596:** En `AgGateG1ParticionCarrilesTests`, no se invocaba `contexto.ConstruirIndices()` tras agregar datos a las listas, y los bloques de fin de semana estaban configurados erróneamente en franja `NOCHE` en lugar del contrato canónico (sábado `MANANA/TARDE`, domingo `MANANA`).
2. **Fase 1597:** No se asociaban registros a `contexto.PoolDocentes`, por lo que el algoritmo genético generaba genes con `EsVacante == 1` y la comprobación de la regla R5 resultaba inoperante.
3. **Fase 1598:** En `AgParticionCarrilesIntegracionTests`, la prueba de festivo con 16h no verificaba la extensión temporal de oleadas ni comparaba contra un control sin festivo.

### Cambios Aplicados
1. **Fase 1596:** Configuración en `AgGateG1ParticionCarrilesTests` de bloques 61 (sábado MANANA), 62 (sábado TARDE) y 71 (domingo MANANA) con horas de inicio y fin canónicas, y llamada obligatoria a `contexto.ConstruirIndices()` con asserts sobre `AulaAsignaturaPorAulaAsig` y `BloquesPorId`.
2. **Fase 1597:** Inclusión de registros en `contexto.PoolDocentes` para las materias 38, 40 y 42 asociadas a los docentes 101, 102 y 103, con aserción estricta de que todos los genes generados tengan `EsVacante == 0`, `IdTercero > 0` e `IdBloqueHorario > 0`.
3. **Fase 1598:** Escenario de 24h (6 sesiones) en `AgParticionCarrilesIntegracionTests`:
   - Con festivo (`2027-02-06`): requiere 3 semanas y `FechaFin == 2027-02-25`.
   - Control sin festivo: requiere 2 semanas y `FechaFin == 2027-02-18`.

### Sincronización REGISTRY
- Fila `R-037` (`closed`) y contadores actualizados a `ronda: 315 / 316`, `fase: 1598 / 1599`.

---

## Ronda 316: Habilitaciones Canónicas en Gate G-1 y Calibración de Partición en Integración (Fase 1599)

### Contexto y Diagnóstico
1. **Fase 1599:** En `AgGateG1ParticionCarrilesTests`, las habilitaciones docentes se asignaban manualmente con `Id = i + 1`, sin calcular las claves de habilitación esperadas en `PoolDocenteKeys`. `EsDocenteHabilitado` requiere el cálculo de hash mediante `TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, semestre)` realizado por `contexto.RegistrarDocenteHabilitado`.
2. **Correcciones de Compilación y Tipos en Gate G-1:**
   - Incorporación del namespace `ConversorPDF.XLL.Academic.Domain.Enums` para resolver `AgEstadoAula` y `AgOrigenVentana`.
   - Remoción de la propiedad inexistente `DiasPatron` sobre el modelo de dominio `Asignatura`.
   - Consulta de fidelidad de patrón contra `contexto.AulaAsignaturaPorAulaAsig[(aula.Id, gen.IdAsignatura)].DiasPatron`.
3. **Calibración del Desempate en Caso 7.6 (`AgParticionCarrilesIntegracionTests`):**
   - El ordenamiento determinista de particiones candidatas incluye `ThenByDescending(ev => ev.Bandas)`.
   - Entre la partición de 2 bandas `{{4,5}, {6,7}}` y la de 3 bandas `{{4}, {5}, {6,7}}`, ambas con 1 infactible, 0 patrones no alojables y makespan 0, el selector escoge deterministamente la de mayor paralelismo (3 bandas). La aserción fue ajustada para validar `{{4}, {5}, {6,7}}`.

### Cambios Aplicados
1. **`AgGateG1ParticionCarrilesTests.cs`:**
   - Invocación de `contexto.RegistrarDocenteHabilitado(101 + i, 1, ids[i], "I")` para docentes 101, 102, 103 sobre asignaturas 38, 40, 42.
   - Verificación con `contexto.EsDocenteHabilitado(101 + i, 1, ids[i], "I").Should().BeTrue()`.
   - Verificación de ausencia de violaciones duras R8, R5 y R13 tras validación completa con `validador.Evaluar(individuo, contexto)`.
2. **`AgParticionCarrilesIntegracionTests.cs`:**
   - Aserción de `res.Value.Particion` en `Caso76_Integracion_MejoraParcial` adaptada a 3 bandas (`{{4}, {5}, {6, 7}}`).

### Evidencia y Verificación
- **Pruebas Focalizadas:** 44/44 PASS (`AgGateG1ParticionCarrilesTests`, `AgParticionCarrilesIntegracionTests`, `AgParticionCarrilesTests`, `AgPlanificadorOleadasTests`).
- **Suite Completa Fast:** 1158/1158 tests PASS (0 fallos, 0 omitidos).
- **Sincronización REGISTRY:** Fila `R-038` (`closed`) y contadores actualizados a `ronda: 316 / 317`, `fase: 1599 / 1600`.

---

## Ronda 317: Reparto Compatible con Patrón Sembrado y Diagnóstico Verificable de Capacidad en Carriles (Fases 1600–1602)

### Contexto y Diagnóstico
1. **Fase 1600:** En `AgPlanificadorOleadas.cs`, la primera pasada de asignación de patrón elegía el carril más específico antes de verificar si tenía capacidad suficiente; al fallar la capacidad, el carril caía a renormalizados y en la segunda pasada (greedy por utilización relativa) se ignoraban las restricciones de patrón sembrado y el fallback forzaba `if (elegido < 0) elegido = 0`. Esto enviaba asignaturas con patrón (ej. viernes "5") a carriles incompatibles (ej. [1,2] o [3,4]), disparando falsos positivos de `R8_DiaFueraDePatron` aunque el día existiera en la ventana.
2. **Fase 1601:** Ausencia de diagnóstico unificado y discriminado entre incompatibilidad de patrón (`R8_DiaFueraDePatron`) y saturación de capacidad del carril compatible (`I5_CAPACIDAD_CARRIL_PATRON`). El texto en `AgGeneradorHorarios.cs` referenciaba «inviabilidad matemática» en lugar del término preciso «bloqueos de factibilidad detectados».
3. **Fase 1602:** Requerimiento de batería de pruebas de regresión y validación cubriendo asignación compatible con capacidad, desvío a carriles compatibles alternativos, aislamiento de patrones saturados sin desbordamiento a carriles incompatibles, preservación de R8 ante días fuera de ventana, conservación de reparto de materias sin patrón, soporte de vacantes permitidas cuando una materia carece de pool docente y casos reales con fines de semana y festivos.

