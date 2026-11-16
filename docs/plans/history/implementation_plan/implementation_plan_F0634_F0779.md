<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# implementation_plan - Fases 634-779 (Rondas 44)

> Indice: [../../../implementation_plan.md](../../../implementation_plan.md) - Escision: 2026-11-16

---
### 157.63 Ronda 106 — Aislamiento y Resiliencia en Exportación Concurrente Asignaturas / Semestres (Fases 634 - 636)

#### 1. Fase 634 — Captura de Nombre de Libro Inicial y Restauración de Foco Libre de Fugas COM
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:** Al abrir la vista de vinculaciones, no se registraba la identidad del libro de trabajo activo del usuario. Al exportar por Asignaturas (`AsignaturasExcelExporter`), si se abría una copia de la plantilla `FV-FO-033`, Excel cambiaba el foco global a este nuevo libro, dejando la plantilla protegida activa y sin mecanismo para retornar al libro original del usuario.
- **Solución Implementada:** Se introdujo `_nombreLibroInicial` (cadena sin retener referencias COM) capturado en `OnLoad`, y se implementó `RestaurarFocoLibroInicial()` que busca el libro por nombre en `_excelApp.Workbooks` liberando todas las referencias COM intermedias mediante `Marshal.ReleaseComObject`.

#### 2. Fase 635 — Verificación Preventiva de Hoja Plantilla/Protegida en `btnExportar_Click`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:** Si el usuario ejecutaba "Exportar Asignaturas" y posteriormente "Exportar Semestres" sin cerrar el formulario, la exportación semestral se ejecutaba ciegamente sobre `_excelApp.ActiveSheet`, la cual era la hoja protegida `"PROGRAMACIÓN"` de la plantilla recién abierta, intentando limpiar celdas protegidas o sobreescribir la información presupuestal.
- **Solución Implementada:** En la rama `PorSemestres`, se verifica si la hoja activa pertenece a la plantilla (`PROGRAMACIÓN`, `NOTOCAR`) o está protegida (`ProtectContents`). Si es así, se restaura el foco al libro inicial y, si persiste en el libro de plantilla, se crea e inserta defensivamente una nueva hoja de cálculo limpia, protegiendo la plantilla de cualquier alteración.

#### 3. Fase 636 — Fallback Defensivo con Inserción de Hoja Limpia ante Libros Protegidos o Vacíos en `SemestresExcelExporter`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Diagnóstico:** En `SemestresExcelExporter.Exportar`, si no existía hoja activa o si la hoja activa estaba protegida, `allCells.Clear()` lanzaba `InvalidOperationException` irrecuperable con error COM 1004 ("La celda o el gráfico que intenta cambiar están en una hoja protegida").
- **Solución Implementada:** Se implementó verificación previa de `ws == null` (creando un nuevo libro con `Workbooks.Add()`) y de `ws.ProtectContents == true`, en cuyo caso inserta automáticamente una hoja limpia al final del libro activo (`wb.Worksheets.Add(Type.Missing, ultimoSheet)`), activa la nueva hoja y transfiere el ownership COM, garantizando que `allCells.Clear()` opere sin riesgo de colisión ni destrucción de datos previos.

### 157.64 Ronda 107 — Inserción Determinista No Destructiva de Hoja Semestres y Cero Fugas RCW (Fases 637 - 639)

#### 1. Fase 637 — Aislamiento Estricto de Plantilla FV-FO-033 en Exportación Semestral
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:** Si `RestaurarFocoLibroInicial()` no encontraba el libro inicial del usuario, la lógica anterior intentaba agregar una hoja dentro del libro activo, el cual correspondía a la plantilla `FV-FO-033-*` de Mis Documentos, contaminando la plantilla oficial de programación con hojas adicionales.
- **Solución Implementada:** Se verifica si el libro activo inicia con `FV-FO-033` al fallar la restauración de foco; de ser así, se crea un libro de trabajo independiente mediante `Workbooks.Add()` liberando su RCW, garantizando que la plantilla quede 100% inalterada.

#### 2. Fase 638 — Liberación Determinista de RCW de `ActiveWorkbook` en `OnLoad`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:** La evaluación directa `(string)_excelApp.ActiveWorkbook?.Name` en `OnLoad` instanciaba un RCW COM no administrado sin pasar por `Marshal.ReleaseComObject`, arriesgando la persistencia de procesos `Excel.exe` zombies al cerrar el complemento.
- **Solución Implementada:** Se aisló la captura en una referencia `dynamic wbInicial` envuelta en bloque `finally` con llamada defensiva a `Marshal.ReleaseComObject(wbInicial)`.

#### 3. Fase 639 — Creación Determinista de Hoja Nueva "Semestres", Validación Temprana y Liberación COM de Colecciones
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Diagnóstico:** `SemestresExcelExporter` ejecutaba `allCells.Clear()` sobre la hoja activa del usuario, borrando datos preexistentes si la hoja no estaba protegida. Además, si no había sesiones con fecha, se creaba una hoja vacía antes de abortar la exportación, y las llamadas repetidas a `wb.Worksheets` no liberaban los RCW intermedios.
- **Solución Implementada:** Se incorporó validación temprana de fechas (`!items.Any(s => s.FechaSesion.HasValue)`). Se implementó la inserción determinista de una nueva hoja dedicada `"Semestres"` (con desambiguación automática `"Semestres (n)"` ante colisión) al final del libro, garantizando la preservación íntegra de las hojas preexistentes del usuario y liberando estrictamente las referencias COM de `hojas`, `ultimoSheet` y `wb` en bloques `finally`.

### 157.65 Empaquetado Release v1.6.3.43 (WiX 4)

#### 1. Verificación y Sincronización Canónica
- **Script Ejecutado:** `.\scripts\build_installer.ps1 -NewVersion "1.6.3.43"`
- **Resultado:**
  - Versión sincronizada a `1.6.3.43` en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
  - Limpieza de instaladores MSI previos.
  - Ejecución de las 544 pruebas unitarias de `ConversorPDF.Tests.csproj`: **544 superadas, 0 errores, 0 omitidas (100% de éxito)**.
  - Compilación exitosa de ensamblados XLL en modo `Release|x64` (0 errores).
  - Empaquetado del instalador MSI mediante WiX 4 Toolset (`InstaladorConversorPDF.wixproj`).
  - Artefacto generado: `installer\InstaladorConversorPDFv1.6.3.43.msi`.

### 157.66 Ronda 108 — Validación del Diagnóstico y Plan Consolidado (Fases 640 - 644)

#### 1. Fase 640 — Filtro de Cohortes en Dos Pasadas (Exacta → Tolerante sólo para Cohortes sin Aula Exacta)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Sobre-inclusión en `ctx.Aulas`: `ExtraerCohorte` descartaba `idprograma` y la condición `progsSel.Contains(a.IdPrograma)` admitía aulas de programas homónimos no solicitados, cargando 38 aulas para 36 cohortes solicitadas. Entre ellas ingresaba el aula 263 (`3|1|2|I|A|PRESENCIAL`, huérfana sin asignaturas) y el aula 291.
- **Solución Implementada:** Se reestructuró el filtrado en dos pasadas: Pasada 1 realiza coincidencia estricta por `AulaRef` (que incluye `idprograma`). Pasada 2 aplica la tolerancia a homónimos (Fase 491) exclusivamente para las cohortes que no obtuvieron un match exacto en la Pasada 1.

#### 2. Fase 641 — Compuertas P07/P08 Acotadas a Cohortes Solicitadas
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Las compuertas de validación de asignaturas en P07 (`EjecutarDiagnosticoPreVueloAsync`) y P08 (`IniciarOptimizacionAsync`) evaluaban la totalidad de `_contextoActual.Aulas` contra `_contextoActual.AulaAsignaturas`, provocando bloqueos pre-vuelo ante aulas huérfanas residuales fuera de la selección del usuario (falso positivo sin vía de escape).
- **Solución Implementada:** Se acotó la verificación a las referencias solicitadas (`refsSolicitadas`). Las aulas huérfanas ajenas a la selección son excluidas de la compuerta de bloqueo y registradas como advertencia en la traza (`AulasHuerfanasNoSolicitadas`).

#### 3. Fase 642 — Traza Temprana Real de Aulas Huérfanas en P05
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** La métrica `AulasSinAsignaturas` se calculaba de forma binaria (`ctx.AulaAsignaturas.Count == 0 ? 1 : 0`), ocultando huérfanas parciales en el alcance hasta llegar a la compuerta P07/P08.
- **Solución Implementada:** Se evalúa y almacena el conteo real de `aulasHuerfanas` en el contexto, emitiendo una advertencia detallada `AulasHuerfanasAlcance` con el listado de referencias sin materias.

#### 4. Fase 643 — Homónimos Distinguibles en el Árbol (P03)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Programas con idéntico nombre y sin abreviatura se mostraban con el mismo texto en el árbol `treeContextoAcademico`, induciendo al operador a marcar ramas homónimas incorrectas.
- **Solución Implementada:** Al poblar el nivel de programas en el árbol, si existen múltiples programas con el mismo nombre normalizado, se anexa el identificador `$" (ID {prog.Id})"` al texto visible, preservando la selección técnica basada en `Tag` (`AgNodoArbolKey`).

#### 5. Fase 644 — Desbloqueo Operativo Inmediato (Dato, no código)
- **Archivos Modificados:** Base de datos SQLite del operador (`ag_aula`)
- **Diagnóstico:** El aula id 263 (`3|1|2|I|A|PRESENCIAL`) figuraba huérfana sin sesiones en la base de datos local del operador.
- **Solución Implementada:** Procedimiento documentado para inactivación (`UPDATE ag_aula SET estado = 0 WHERE aula_ref = '3|1|2|I|A|PRESENCIAL'`) o depuración previa validación de integridad referencial (`SELECT COUNT(*) FROM ag_sesiones WHERE id_aula = 263`).

### 157.67 Ronda 109 — Tolerancia a Homónimos en Validación de Anclaje y Consistencia Documental (Fases 645 - 646)

#### 1. Fase 645 — Validación de Anclaje P07/P08 con Fallback de Cohorte y Helper Centralizado `AgAulaFactory.ExtraerCohorte`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAulaFactory.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAulaFactoryTests.cs`
- **Diagnóstico:** Las compuertas de anclaje multi-combinación en P07 (`EjecutarDiagnosticoPreVueloAsync`) y P08 (`IniciarOptimizacionAsync`) evaluaban exclusivamente coincidencia exacta por `AulaRef`. Si un aula fue persistida con el ID del programa homónimo (admitida por la Pasada 2 de la Fase 640), el anclaje la reportaba erróneamente como faltante y bloqueaba la corrida. Además, la función `ExtraerCohorte` estaba duplicada como función local privada en `ConstruirContextoAsync`.
- **Solución Implementada:** Se promovió `ExtraerCohorte` a método público estático reutilizable en `AgAulaFactory` con cobertura de pruebas unitarias. Se actualizó la detección de combinaciones faltantes en P07 y P08 incorporando un fallback de cohorte que valida coincidencia con programas homónimos cuando no existe match exacto, preservando el anclaje estricto cuando existe.

#### 2. Fase 646 — Consistencia Documental de Artefactos de la Ronda 108
- **Archivos Modificados:** `implementation_plan.md`, `walkthrough.md`
- **Diagnóstico:** Los nombres de método descritos en la documentación de la Ronda 108 diferían de los identificadores reales en código (`EjecutarOptimizacionAsync` en vez de `IniciarOptimizacionAsync`, y `PrepararContextoAsync` en vez de `ConstruirContextoAsync`).
- **Solución Implementada:** Se normalizaron ambos identificadores en `implementation_plan.md` y `walkthrough.md` garantizando total paridad con el código fuente.

### 157.68 Ronda 110 — Anclaje Mono-Combinación con Tolerancia a Homónimos (Fase 647)

#### 1. Fase 647 — Anclaje de Combinación Única (P05) con Fallback de Cohorte Homónima
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al seleccionar una sola combinación cuya aula fue persistida bajo el ID de un programa homónimo, la búsqueda de coincidencias exactas por `AulaRef` en `ConstruirContextoAsync` resultaba vacía (`coincidencias.Count == 0`), asignando `IdAulaCorrida = 0` y provocando el bloqueo pre-vuelo en las ramas mono-combinación de P07 y P08 (`No existe Aula exacta para '...'`).
- **Solución Implementada:** Se incorporó en `ConstruirContextoAsync` un fallback por cohorte mediante `AgAulaFactory.ExtraerCohorte(aulaRefEsperada)` ante la ausencia de match exacto. Si se localiza una única aula coincidente por cohorte homónima, se ancla deterministamente `ctx.IdAulaCorrida = coincidencias[0].Id` con logging explícito, preservando el reporte de ambigüedad si existiesen múltiples.

### 157.69 Empaquetado Release v1.6.3.44 (WiX 4)

#### 1. Verificación y Sincronización Canónica
- **Script Ejecutado:** `.\scripts\build_installer.ps1 -NewVersion "1.6.3.44"`
- **Resultado:**
  - Versión sincronizada a `1.6.3.44` en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
  - Limpieza de instaladores MSI previos.
  - Ejecución de las 549 pruebas unitarias de `ConversorPDF.Tests.csproj`: **549 superadas, 0 errores, 0 omitidas (100% de éxito)**.
  - Compilación exitosa de ensamblados XLL en modo `Release|x64` (0 errores).
  - Empaquetado del instalador MSI mediante WiX 4 Toolset (`InstaladorConversorPDF.wixproj`).
  - Artefacto generado: `installer\InstaladorConversorPDFv1.6.3.44.msi`.

### 157.70 Ronda 111 — Plan Unificado v4: Decisiones β/γ/Horas y Unificación Militar de Horarios (Fases 648 – 656)

#### 1. Fase 648 — Contrato Puro: `AgFranjaJornada`, `AgDiasVentana`, `HorarioMilitarizador` y Ampliación de DTOs/Modelos
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgFranjaJornada.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgDiasVentana.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Application/Services/HorarioMilitarizador.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgNodoArbolKey.cs`
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgContextoSeleccionDto.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgAula.cs`
  - `ConversorPDF.Tests/Academic/Domain/AgFranjaJornadaTests.cs` (Nuevo)
  - `ConversorPDF.Tests/Academic/Domain/AgDiasVentanaTests.cs` (Nuevo)
  - `ConversorPDF.Tests/Academic/Application/Services/HorarioMilitarizadorTests.cs` (Nuevo)
- **Diagnóstico:** Desalineación entre opciones horarias en UI (12 h AM/PM) y persistencia del AG (militar 24 h); ausencia de un contrato canónico puro y testeable que mapee explícitamente jornada y horas/día a franjas militares, así como las 5 opciones de días (incluyendo JUE-SAB `{4,5,6}`).
- **Solución Implementada:**
  - Implementación de `AgFranjaJornada.Resolve(jornada, horasDia)`:
    - MAÑANA+4 -> `("08:00", "12:00", "MANANA", 4)`
    - MAÑANA+8 -> `("08:00", "17:00", "MANANA", 8)`
    - TARDE+4  -> `("13:00", "17:00", "TARDE", 4)`
    - TARDE+8  -> `("14:00", "22:00", "TARDE", 8)`
    - NOCHE+4  -> `("18:00", "22:00", "NOCHE", 4)`
    - NOCHE+8  -> `("14:00", "22:00", "NOCHE", 8)` (comparte TARDE)
  - Implementación de `AgDiasVentana`: 5 opciones canónicas (`LUN-JUE` "1,2,3,4", `MAR-VIE` "2,3,4,5", `JUE-SAB` "4,5,6", `VIE-SAB` "5,6", `VIE-DOM` "5,6,7").
  - Implementación de `HorarioMilitarizador`: helper puro e idempotente que convierte cadenas de horario (12 h o 24 h) a `"HH:mm A HH:mm"` usando `HorarioSesionParser.TryParse`.
  - Extensión de `AgNodoArbolKey`, `AgContextoSeleccionDto` y `AgAula` con las propiedades `Jornada`, `HorasDia`, `DiasVentana` manteniendo `AulaRef` inalterada.

#### 2. Fase 649 — Migración Aditiva de Schema y Semilla de 42 Bloques (`MigrationV62`)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV62_AddAgAulaVentanasYBloquesSemilla.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/Migrations/MigrationV62Tests.cs` (Nuevo)
- **Diagnóstico:** `ag_aula` y `ag_arbol_nodo` carecen de columnas para persistir las decisiones de jornada, horas por día y días de ventana; el catálogo `ag_bloque_horario` contiene únicamente 8 bloques y carece de las franjas completas de 4h y 8h para los 7 días de la semana.
- **Solución Implementada:**
  - Migración idempotente `MigrationV62` que añade las columnas `jornada` (TEXT), `horas_dia` (INTEGER) y `dias_ventana` (TEXT) a `ag_aula` y `ag_arbol_nodo`.
  - Inserción idempotente de los bloques faltantes en `ag_bloque_horario` hasta completar la cuadrícula de 6 bloques/día × 7 días = 42 bloques mediante `WHERE NOT EXISTS`.

#### 3. Fase 650 — Migración de Datos para Unificación a Formato Militar (`MigrationV63`)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV63_UnificarHorariosMilitar.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/Migrations/MigrationV63Tests.cs` (Nuevo)
- **Diagnóstico:** Los registros históricos en `vinculaciones.dia_horas_semanales` y `sesiones.horariosesion` contienen cadenas en formato 12 h AM/PM (`"1:00 PM A 5:00 PM"`), lo que genera fallos de coincidencia con las opciones militares del AG.
- **Solución Implementada:**
  - `MigrationV63`: migración de datos idempotente que normaliza `sesiones.horariosesion` y la tercera parte de `vinculaciones.dia_horas_semanales` a formato militar `"HH:mm A HH:mm"`, preservando intactos los valores de duración y fechas.

#### 4. Fase 651 — Sincronización Militar en `DayHoursDropdown` y Serialización 3-Partes en `AgMaterializadorSesiones`
- **Archivos Afectados:**
  - `ConversorPDF.XLL/UI/DayHoursDropdown.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs`
- **Diagnóstico:** `DayHoursDropdown` valida selecciones contra colecciones en 12 h AM/PM (`Items.Contains`), descartando horarios en formato militar y asignando `DBNull.Value`; asimismo, `AgMaterializadorSesiones` serializaba `dia_horas_semanales` en 2 partes sin horario y validaba la invariante I2 asumiendo sesiones fijas de 4.0 h.
- **Solución Implementada:**
  - Actualización de opciones de `DayHoursDropdown` a formato militar (4h: `"08:00 A 12:00"`, `"13:00 A 17:00"`, `"18:00 A 22:00"`; 8h: `"08:00 A 17:00"`, `"14:00 A 22:00"`).
  - Modificación de `AgMaterializadorSesiones` para formatear `dia_horas_semanales` en 3 partes `DIA|HORAS|HORARIO` (`bloque.ObtenerHorarioTexto()`), 1 bloque por día.
  - Actualización de la invariante I2: validación por sumatoria dinámica `Sum(s => s.Horas) == esperado`, admitiendo sesiones de 4 u 8 horas.

#### 5. Fase 652 — Respeto de Decisiones en Fábrica AG y Operadores Genéticos
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorMutacionTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorCruceTests.cs`
- **Diagnóstico:** Los operadores de mutación y cruce alteraban aleatoriamente el bloque horario sin respetar las decisiones explícitas de jornada/horas tomadas en el aula, provocando deriva temporal.
- **Solución Implementada:**
  - Incorporación de `AgGen.EsFijoTemporal` cuando la franja y jornada provienen de decisión de aula.
  - Guards en `AgOperadorMutacion` y `AgOperadorCruce` para omitir mutación/cruce de bloque horario en genes con `EsFijoTemporal == true`, permitiendo optimización exclusiva de docentes.
  - En `AgFabricaIndividuos`: resolución de bloques mediante `AgFranjaJornada.Resolve` cuando existen decisiones en `AgAula`.

#### 6. Fase 653 — OwnerDraw en Árbol Académico (Nivel 6: 5 Grupos de Radios y Cabeceras)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico:** La interfaz carecía de controles en el árbol para seleccionar las dimensiones de decisión (Jornada, Días y Horas/Día) por cohorte/modalidad en el Nivel 6.
- **Solución Implementada:**
  - Incorporación de cabeceras en `pnlArbolToolbar` (`lblColJornada@290`, `lblColDias@470`, `lblColHorasDia@850`).
  - Renderizado OwnerDraw en `treeContextoAcademico_DrawNode` para Nivel 6 usando `RadioButtonRenderer.DrawRadioButton`.
  - Hit-test extendido en `treeContextoAcademico_NodeMouseClick` con exclusividad por dimensión.

#### 7. Fase 654 — Persistencia de Perfiles en `ag_arbol_nodo` y Sincronización en `AgArbolMapper`
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgArbolMapper.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgArbolRepository.cs`
  - `ConversorPDF.Tests/Academic/Application/Mappers/AgArbolMapperTests.cs`
- **Diagnóstico:** Al guardar o cargar perfiles declarativos del árbol, las decisiones de jornada, horas_dia y dias_ventana no se propagaban a `ag_arbol_nodo`.
- **Solución Implementada:**
  - Mapeo bidireccional en `AgArbolMapper` de `Jornada`, `HorasDia` y `DiasVentana` entre `AgNodoArbolKey` y `AgArbolNodo`.
  - Extensión de sentencias de lectura y escritura en `AgArbolRepository`.

#### 8. Fase 655 — Pre-Vuelo: Validación de Ventana y Existencia de Bloque Horario
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs` (Nuevo)
- **Diagnóstico:** Ausencia de validación previa que alerte si los días del patrón de la asignatura no están contenidos en la ventana seleccionada (`dias_patron ⊆ ventana`) o si no existe un bloque horario activo para la tupla `(jornada, horas_dia)`.
- **Solución Implementada:**
  - Servicio puro `AgValidadorVentana` integrado en el diagnóstico pre-vuelo P07 y P08 de `AgFormGeneradorHorarios`.

#### 9. Fase 656 — Documentación Técnica de Arquitectura de Horarios
- **Archivos Afectados:**
  - `docs/articles/changes/2026-ag-horarios-academicos.md`
  - `implementation_plan.md`
  - `walkthrough.md`
- **Diagnóstico:** Desactualización de la documentación arquitectónica tras la consolidación del modelo v4.
- **Solución Implementada:**
  - Actualización exhaustiva de `2026-ag-horarios-academicos.md` reflejando el modelo de franjas, decisiones de ventana, formato militar unificado y trazabilidad de pruebas.

### 157.71 Ronda 112 — Plan v4: Gaps de Lógica en Congelamiento Temporal (Fases 657 – 659)

#### 1. Fase 657 — Preservación de `EsFijoTemporal` en `AgGen.Clonar()`
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`
  - `ConversorPDF.Tests/Academic/Domain/AgGenTests.cs`
- **Diagnóstico:** `AgGen.Clonar()` omite copiar la propiedad `EsFijoTemporal`. Al clonar progenitores en `AgOperadorCruce`, los descendientes pierden el flag de congelamiento temporal.
- **Solución Planificada:** Incorporar `EsFijoTemporal = this.EsFijoTemporal` en el inicializador de `AgGen.Clonar()`.

#### 2. Fase 658 — Cobertura Integral de Mutación para Genes con `EsFijoTemporal`
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorMutacionTests.cs`
- **Diagnóstico:** El guard de mutación solo se ubicaba sobre el bloque horario, permitiendo que la mutación de aula reasigne `IdAula` a otra cohorte con decisiones temporales distintas o nulas.
- **Solución Planificada:** Insertar la guarda temprana `if (gen.EsFijoTemporal) continue;` al inicio del bucle por gen en `AgOperadorMutacion` (tras L77), congelando aula y bloque de forma conjunta, mientras la mutación docente en L28-73 sigue activa.

#### 3. Fase 659 — Protección de Genes Congelados en Cascada de Cruce `HeredarOX`
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorCruceTests.cs`
- **Diagnóstico:** Los genes congelados tienen `Estado == Valido` y superan el filtro `Estado == Fijado` de L107 en `HeredarOX`, cayendo en la cascada de reasignación de bloque horario y patrón semanal que sobrescribe la decisión fijada por el usuario.
### 157.72 Ronda 114 — Fase 660: Colisión de UDF `Normalizar` por Auto-Registro Excel-DNA

#### 1. Fase 660 — Renombrado de Métodos `Normalizar` Nuevos a Nombres Calificados
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgFranjaJornada.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/HorarioMilitarizador.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV63_UnificarHorariosMilitar.cs`
  - `ConversorPDF.XLL/UI/DayHoursDropdown.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/HorarioMilitarizadorTests.cs`
- **Diagnóstico:** `ConversorPDF.XLL-AddIn64.dna` no especifica `ExplicitRegistration="true"`, provocando que Excel-DNA auto-registre todo método `public static` con firma soportada como función definida por el usuario (UDF). La existencia de tres métodos `Normalizar(string)` (`DiaSemanaNormalizador`, `AgFranjaJornada`, `HorarioMilitarizador`) provocaba colisiones en tiempo de carga con el mensaje `Repeated function name: 'Normalizar'`.
- **Solución Implementada:**
  - En `AgFranjaJornada.cs`, renombrado `Normalizar` a `NormalizarJornada(string jornada)`.
  - En `HorarioMilitarizador.cs`, renombrado `Normalizar` a `NormalizarHorario(string horario)`.
  - Actualización de todos los puntos de consumo en `MigrationV63_UnificarHorariosMilitar.cs`, `DayHoursDropdown.cs` y `HorarioMilitarizadorTests.cs`.
  - La UDF global `Normalizar` queda reservada exclusivamente a `DiaSemanaNormalizador`.

### 157.73 Ronda 115 — Fase 661: Corrección de UX en Nivel 6 de Árbol Académico (Etiquetas, Separadores, Alineación y Repintado)

#### 1. Fase 661 — Alineación Dinámica, Etiquetas Completas y Repintado de Fila en Nivel 6
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico:**
  - En el Nivel 6 del árbol, los radio buttons se renderizaban con etiquetas abreviadas corridas (`"M","T","N"`, `"L-J"..."V-D"`, `"4h","8h"`) sin separadores entre grupos de decisión (Jornada, Días, Horas/día).
  - Las cabeceras en `pnlArbolToolbar` tenían textos estáticos (`[M] [T] [N]`, etc.) en posiciones X fijas desalineadas respecto a la posición visual real de las columnas del nodo en el árbol (`bounds.X`).
  - El ancho del formulario (984 px) resultaba insuficiente para las dimensiones completas (~1060 px), provocando recortes y scroll horizontal.
  - Al hacer clic sobre los radios, la invalidación sólo cubría `e.Node.Bounds` (el texto de la modalidad), impidiendo el repintado visual inmediato de los radios.
  - El hit-test no tenía clamp superior en el grupo de horas (dx >= 690 activaba 8h de forma desmedida).
- **Solución Implementada:**
  - Etiquetas descriptivas completas: `"Mañana"`, `"Tarde"`, `"Noche"` en Jornada; `"Lun-Jue"`, `"Mar-Vie"`, `"Jue-Sab"`, `"Vie-Sab"`, `"Vie-Dom"` en Días; `"4h"`, `"8h"` en Horas/día.
  - Inserción de líneas separadoras verticales (`DibujarSeparadorVertical`) entre los tres grupos en `DrawNode` (X=+430 y X=+945).
  - Creación de separadores visuales correspondientes en la toolbar (`_pnlSepToolbar1`, `_pnlSepToolbar2`) y método dinámico `AlinearCabecerasNivel6()` para sincronizar los `Left` de cabeceras y separadores con el X real de los nodos de Nivel 6.
  - Redimensión del formulario (`ClientSize = 1120x661`), contenedores hijos a 1072 px y reducción de `Indent` a 14 px para asegurar espacio sin scroll horizontal.
  - Invalidación de fila completa en `NodeMouseClick`: `new Rectangle(0, r.Y, treeContextoAcademico.ClientSize.Width, r.Height)`.
  - Hit-test estricto con zonas delimitadas y no-op defensivo fuera de rango.
  - Detección de cursor `Cursors.Hand` en `MouseMove` al pasar sobre zonas interactivas de Nivel 6.

#### 2. Fase 662 — Compresión Compacta de Columnas de Nivel 6 (Prevención de Desborde a la Derecha)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico:**
  - Tras F661, con `Indent = 14`, el X base del nodo de Nivel 6 es aproximadamente `6 * 14 + 19 = 103`. La columna `8h` quedaba en `103 + 1015 = 1118` px y su etiqueta terminaba en ~1136 px, excediendo el ancho cliente utilizable del árbol (~1055 px con scrollbar vertical), provocando que el grupo Horas/día y el extremo de Vie-Dom quedaran recortados.
- **Solución Implementada:**
  - Compresión compacta de offsets horizontales sin pérdida de legibilidad:
    - Jornada: Mañana (+150), Tarde (+215), Noche (+275), Separador (+360).
    - Días: Lun-Jue (+375), Mar-Vie (+450), Jue-Sab (+525), Vie-Sab (+600), Vie-Dom (+675), Separador (+775).
    - Horas/día: 4h (+790), 8h (+855).
    - Límite derecho total: `103 + 855 + 25 = 983 px` (< 1055 px), dejando ~72 px de margen de seguridad antes del scroll vertical.
  - Sincronización exacta de `ProcesarClickDimensionNivel6` (hit-test), `AlinearCabecerasNivel6` (toolbar), `_pnlSepToolbar1/2.Left` y detección de cursor en `MouseMove`.

#### 3. Fase 663 — Render Condicionado al `Checked` de Modalidad y Selección Automática de Fila en Nivel 6
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  - Aunque la propagación del árbol solo marcaba la modalidad activa por defecto (`Presencial`), `DrawNode` renderizaba los radio buttons seleccionados (`Mañana`, `Mar-Vie`, `4h`) en todas las filas de Nivel 6 independientemente de su estado `Checked`. Esto generaba la ilusión visual de que múltiples modalidades estaban activas simultáneamente en un grupo.
- **Solución Implementada:**
  - En `TreeContextoAcademico_DrawNode`: condicionar el estado activo de los radios al flag `bool filaActiva = e.Node.Checked`. Las filas no marcadas (`Virtual`, `Semipresencial`) se dibujan con todos sus radios desmarcados.
  - En `ProcesarClickDimensionNivel6`: si el usuario hace clic sobre una opción de dimensión en una fila no seleccionada, la modalidad se marca automáticamente (`node.Checked = true`), se aplica exclusividad (`AplicarExclusividadModalidad`) y se disparan los eventos correspondientes.
  - En `NodeMouseClick` y `TreeContextoAcademico_AfterCheck`: invalidar el rectángulo completo de todas las modalidades hermanas para refrescar simultáneamente el encendido/apagado de sus radios.

### 157.74 Ronda 116 — Fase 664: Aislamiento Estricto de Asignaturas por Programa y Cohorte en AG (Prevención de Cruces Multi-Programa)

#### 1. Fase 664 — Aislamiento Curricular por Programa en Fábrica AG, Operadores Genéticos, Validador y Staging
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`
- **Diagnóstico:**
  - En escenarios de optimización multi-programa (por ejemplo, sedes con oferta simultánea de APT-52 y APT-53), `contexto.Asignaturas` contiene los catálogos combinados de todos los programas involucrados.
  - Al generar la población inicial en `AgFabricaIndividuos.CrearAleatorio`, el cruce entre `gruposAula` y `contexto.Asignaturas` se efectuaba exclusivamente validando la igualdad de semestre romano (`semAulaRom == semAsigRom`).
  - Como resultado, asignaturas de un programa con el mismo semestre se programaban en aulas de otro programa (ej. `FINANZAS PUBLICAS` de APT-52 Semestre VII en el aula de APT-53 Semestre VII B; y `Opción de Grado` de APT-53 Semestre X en el aula de APT-52 Semestre X A, desplazando a `TRABAJO DE GRADO` de APT-52).
  - Al abrir `FormSesionesAcademicas` desde `FormVinculaciones` mediante "Revisar", los selectores filtran por `idprograma` del aula y `semestre`, por lo cual la asignatura foránea no existe en el catálogo filtrado, quedando el combobox en blanco y no disponible.
  - La mutación (`AgOperadorMutacion`) y el cruce (`AgOperadorCruce.HeredarOX`) tampoco verificaban que las aulas o asignaturas recombinadas pertenecieran al mismo programa.
- **Solución Implementada:**
  1. En `AgFabricaIndividuos.cs`:
     - Helper `ProgramaDeAula(AgAula a)` para extraer de forma robusta `IdPrograma` de la entidad o de `AulaRef` (`partes[2]`).
     - Clave natural de cohorte en `ClaveCohorteAula` extendida para incluir `IdPrograma` (`$"{prog}|{sem}|{grp}"`).
     - Restricción estricta al conjunto de asignaturas `asigsValidasAula` configuradas en `ag_aula_asignatura` (`contexto.AulaAsignaturas`) para el aula específica.
     - Filtro incondicional de programa: `idProgAula > 0 && idProgAsig > 0 && idProgAula != idProgAsig => continue`.
     - Fijación del gen con `IdPrograma` procedente directamente de la asignatura.
  2. En `AgOperadorMutacion.cs`:
     - Restricción de aulas candidatas al mismo programa del gen: `a.IdPrograma <= 0 || gen.IdPrograma <= 0 || a.IdPrograma == gen.IdPrograma`.
  3. En `AgOperadorCruce.cs`:
     - Verificación en `HeredarOX`: `progAsig > 0 && progAula > 0 && progAsig != progAula => continue` antes de transferir `IdAsignatura`.
  4. En `AgValidadorReglas.cs`:
     - Penalización en regla dura R4 (`CompatibilidadAula`) si `aula.IdPrograma != gen.IdPrograma` o si `asigEval.IdPrograma != gen.IdPrograma`.
  5. En `AgGenToAgVinculacionMapper.cs`:
     - Resolución defensiva de `IdPrograma` en `AgVinculacion` vinculada al ID de programa real de la asignatura en `contexto.AsignaturasPorId`.
  6. En `AgMaterializadorSesiones.cs`:
     - Guarda fail-fast previa a la persistencia en staging que valida la correspondencia de `IdPrograma` entre la cohorte del aula y la vinculación.
  7. En `AgFabricaIndividuosTests.cs`:
     - Nueva prueba unitaria `CrearAleatorio_WhenMultiPrograma_ShouldNotAssignSubjectsFromOtherPrograms` verificando el aislamiento completo.

### 157.75 Ronda 117 — Fases 665 - 669: Soporte de Pernocta en Último Día (`isPernoctaUltimoDia`) y Ajuste de Duración y Rango en Bloques de Viáticos

#### 1. Fase 665 — Propagación de Columna y Vista (Plumbing de Persistencia y DTO)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Sql/AcademicSchema.sql`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV64_AddPernoctaUltimoDia.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/SesionListItem.cs`
- **Diagnóstico:**
  - Las comisiones de viáticos requieren sumar +1 día de regreso cuando el docente pernocta en el último día de clase. La vista `vw_SesionesDetalle` no exponía la columna `isPernoctaUltimoDia` y `SesionListItem` no la capturaba.
- **Solución Implementada:**
  - Proyección de `v.isPernoctaUltimoDia AS IsPernoctaUltimoDia` en `vw_SesionesDetalle` dentro de `AcademicSchema.sql`.
  - Migración `MigrationV64_AddPernoctaUltimoDia` idempotente con `ALTER TABLE vinculaciones ADD COLUMN isPernoctaUltimoDia INTEGER DEFAULT 0;` y recreación de `vw_SesionesDetalle`.
  - Registro de `MigrationV64` en `SqliteHelper.BuildMigrator()` después de `MigrationV63`.
  - Propiedad `public int IsPernoctaUltimoDia { get; set; }` en `SesionListItem.cs`.

#### 2. Fase 666 — Lógica de Negocio del Bloque de Viáticos (`ViaticoBloque.cs`)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/ViaticoBloque.cs`
- **Diagnóstico:**
  - La duración y texto del bloque se calculaban rígidamente con `FechaMaxima`, omitiendo el día extra de comisión por pernocta en el último día.
- **Solución Implementada:**
  - Propiedad `TienePernoctaUltimoDia => Sesiones != null && Sesiones.Any(s => s.IsPernoctaUltimoDia == 1);` (Regla A).
  - Propiedad `FechaFinComision => FechaMaxima?.AddDays(TienePernoctaUltimoDia ? 1 : 0);` preservando `FechaMaxima` intacta para el algoritmo `AgruparConsecutivos`.
  - Recálculo de `DuracionTotal = Math.Round((FechaFinComision.Value - FechaMinima.Value).Days + 0.5, 1)`.
  - Actualización de `DisplayText` para mostrar `FechaFinComision`.
  - Invalidación segura de caché en `Agregar` e `InvalidateCache`.

#### 3. Fase 667 — Consumidores del Rango y Duración de Viáticos
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
  - `ConversorPDF.XLL/VisTimeline/Application/Services/ViaticoTimelineMapper.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/ViaticoConflictAnalyzer.cs`
- **Diagnóstico:**
  - Múltiples pantallas y exportadores calculaban visualmente la fecha final sobre `FechaMaxima` y `Math.Round` con redondeo bancario (perdiendo el día extra en `2.5 → 2`).
- **Solución Implementada:**
  - En `FormVerVinculaciones.cs`: `fechaFinStr` proyectado desde `bloque.FechaFinComision ?? bloque.FechaMaxima`.
  - En `ViaticosIndividualExporter.cs`: `fFin = b.FechaFinComision ?? b.FechaMaxima ?? DateTime.Now` y `Math.Round(b.DuracionTotal, MidpointRounding.AwayFromZero)`.
  - En `ViaticoTimelineMapper.cs`: rango y tooltip proyectados con `bloque.FechaFinComision ?? bloque.FechaMaxima`.
  - En `ViaticoConflictAnalyzer.cs`: visualización y método de intersección `SeSolapan` extendidos para evaluar `FechaFinComision ?? FechaMaxima`.

#### 4. Fase 668 — Prerrequisito UI y Persistencia (`FormSesionesAcademicas`, `Vinculacion`, `AcademicRepository`)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.Designer.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Vinculacion.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
- **Diagnóstico:**
  - El formulario de sesiones académicas carecía de control interactivo para capturar `isPernoctaUltimoDia`, y el repositorio no persistía ni recuperaba el campo en las operaciones CRUD de vinculaciones.
- **Solución Implementada:**
  - En `Designer.cs`: control `chkPernoctaUltimoDia` ("Pernocta último día") en `Point(356, 143)` y ajuste armónico de `btnGenerar` a `Point(495, 137)` con tamaño `(145, 30)` sin solapamiento.
  - En `FormSesionesAcademicas.cs`: captura en `PrepararVinculacionDesdeForm` (`chkPernoctaUltimoDia.Checked ? 1 : 0`) y restauración en `CargarEstadoPrevio` (`vinc.IsPernoctaUltimoDia == 1`).
  - En `Vinculacion.cs`: propiedad `public int IsPernoctaUltimoDia { get; set; }`.
  - En `AcademicRepository.cs`: soporte completo en `GetVinculaciones()`, `InsertVinculacion()`, `GetVinculacionById()` y `UpdateVinculacion()`.

#### 5. Fase 669 — Pruebas Unitarias de ViaticoBloque y Migración V64
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Domain/ViaticoBloquePernoctaTests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV64Tests.cs`
- **Diagnóstico:**
  - Se requería blindar matemáticamente la regla de negocio y asegurar la idempotencia de esquema.
- **Solución Implementada:**
  - `ViaticoBloquePernoctaTests.cs`: 6 pruebas unitarias verificando casos sin flag (1.5), con flag (2.5), 1 día (1.5), dos bloques recurrentes (+1 ambos), bloque ajeno intacto y preservación de `FechaMaxima`.
  - `MigrationV64Tests.cs`: 4 pruebas unitarias verificando metadatos, migración sobre BD antigua con valor 0 por defecto, idempotencia y exposición de la columna en `vw_SesionesDetalle`.

### 157.76 Ronda 118 — Verificación de Fases 665-669 (Pernocta Último Día) y Bloqueante de Compilación

#### 1. Fase 670 — Migración V64 no registrada en el proyecto (BLOCKER de Compilación)
- **Archivo Afectado:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** `SqliteHelper.cs:365` invoca `new MigrationV64_AddPernoctaUltimoDia()`, pero el archivo no estaba registrado en `<Compile Include>`. Al no ser proyecto SDK-style (requiere inclusión explícita de archivos .cs), causaba fallo de compilación `CS0246`.
- **Solución Implementada:** Inserción de `<Compile Include="Academic\Infrastructure\Persistence\Migrations\MigrationV64_AddPernoctaUltimoDia.cs" />` inmediatamente tras `MigrationV63`.

#### 2. Fase 671 — `chkPernoctaUltimoDia` sin reset en la ruta "sin vinculación previa"
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** En `CargarEstadoPrevio`, la salida temprana por `vincResult.IsFailure` no restablecía los checkboxes a su estado por defecto, conservando valores de instancias previas.
- **Solución Implementada:** En el bloque `if (vincResult.IsFailure)`, se ejecuta `this.Invoke` para forzar `chkTerminalAereo.Checked = false; chkPernoctaUltimoDia.Checked = false; pnlTransporteAereo.Visible = false;`.

#### 3. Fase 672 — Caché de `ViaticoBloque` dependiente de `IsPernoctaUltimoDia`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/ViaticoBloque.cs`
- **Diagnóstico:** Si `IsPernoctaUltimoDia` mutaba in-place en una sesión existente del bloque, `_cachedDuracionTotal` no se invalidaba al no llamarse `Agregar()` ni `InvalidateCache()`.
- **Solución Implementada:** Incorporación del campo `_cachedPernocta`. En el getter de `DuracionTotal`, se verifica `_cachedPernocta == pernoctaActual`. Se limpia en `Agregar()` e `InvalidateCache()`. Adicionalmente, se añadió prueba unitaria `DuracionTotal_WhenSesionPernoctaMutatesInPlace_ShouldRecalculateWithoutCallingAgregar`.

#### 4. Fase 673 — Clarificación semántica en `ViaticosIndividualExporter`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** `Math.Round(b.DuracionTotal, MidpointRounding.AwayFromZero)` calcula los días de comisión para el objeto del viaje (1 día: 0.5 -> 1 día calendario; 2 días + pernocta: 2.5 -> 3 días calendario).
- **Solución Implementada:** Documentación en código explicitando la distinción entre días calendario de comisión (`dur`) y días de viático (`b.DuracionTotal`).

#### 5. Fase 674 — Acotación de `SeSolapan` en `ViaticoConflictAnalyzer`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/ViaticoConflictAnalyzer.cs`
- **Diagnóstico:** Evaluar solapes con `FechaFinComision` (+1 día de pernocta) generaba cruces con sesiones de otra sede en días sin sesión real en el bloque A, donde `EsCruceSalvado` no podía operar al no haber pares de `IdSesion`.
- **Solución Implementada:** `SeSolapan` se acota a `FechaMaxima` (días efectivos de clase), preservando `FechaFinComision` en visualizaciones de detalle y tooltip.

#### 6. Fase 675 — Margen de seguridad UI entre `chkPernoctaUltimoDia` y `btnGenerar`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.Designer.cs`
- **Diagnóstico:** Margen estrecho de 4-8px entre el texto del checkbox y `btnGenerar`.
- **Solución Implementada:** `btnGenerar` reubicado a `Point(500, 137)` con `Size(140, 30)`, garantizando margen de 9px sin alterar el borde exterior del contenedor (640px).

#### 7. Fase 676 — Documental: MODO AG sin soporte de pernocta
- **Archivos de Referencia:** `MigrationV64` / `MigrationV59`
- **Diagnóstico:** `ag_vinculaciones` no incluye `isPernoctaUltimoDia`. El modo AG opera con valor 0 por defecto sin error de runtime.
- **Acción:** Registrado como pendiente si se requiere paridad en MODO AG en iteraciones posteriores.

### 157.77 Ronda 119 — Verificación de Fases 670-676 (Correcciones Ronda 118)

#### 1. Fase 677 — `TienePernoctaUltimoDia` reevaluación (Evaluación)
- **Estado:** RETIRADA (Comportamiento deseado)
- **Diagnóstico:** Cachear `TienePernoctaUltimoDia` anularía la reactividad ante mutaciones in-place del flag en sesiones ya agregadas (verificada por `DuracionTotal_WhenSesionPernoctaMutatesInPlace_ShouldRecalculateWithoutCallingAgregar`), mientras que `Sesiones.Any` sobre un máximo de ≤8 elementos toma tiempo despreciable (sub-microsegundo).
- **Acción:** Se conserva la evaluación dinámica de la bandera, manteniendo la coherencia de la clave `_cachedPernocta`.

#### 2. Fase 678 — Endurecimiento de `this.Invoke` en reset de controles
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** Si `CargarEstadoPrevio` se invocara sin handle creado (`!IsHandleCreated`), `Control.Invoke` lanzaría `InvalidOperationException` y abortaría la asignación del `id_periodo` por defecto.
- **Solución Implementada:** Se trasladó la invocación dentro del bloque protegido `try-catch` y se condicionó con `if (IsHandleCreated)`.

### 157.78 Ronda 120 — Fases 679 a 684: "Pernocta Último Día" en Nivel 2 y Menú Contextual (FormVerVinculaciones)

#### 1. Fase 679 — Repositorio y Dominio Base
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IAcademicRepository.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/ViaticoBloque.cs`
- **Diagnóstico:**
  - Se requería actualizar la columna `isPernoctaUltimoDia` en la base de datos de manera atómica por vinculación, y exponer desde `ViaticoBloque` los IDs únicos de vinculación de sus sesiones agrupadas.
- **Solución Implementada:**
  - En `IAcademicRepository`: declarada la firma `SqliteResult<int> SetPernoctaUltimoDia(int idVinculacion, int valor)`.
  - En `AcademicRepository`: implementación ejecutando `UPDATE ... SET isPernoctaUltimoDia = @valor WHERE id = @id`, con habilitación de CR-SQLite y notificación `NotifyMarkDirty()`.
  - En `ViaticoBloque`: agregada propiedad `IdsVinculacion` que extrae los `IdVinculacion > 0` únicos de las sesiones.

#### 2. Fase 680 — Nivel 2: "Cierre" y Formato del Bloque
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/ViaticoBloqueFormatter.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:**
  - El nodo de Nivel 2 requería mostrar el horario de cierre (sesión con inicio más tardío del último día de clase) y la etiqueta `| Pernocta Ultimo Dia` condicional a `TienePernoctaUltimoDia`.
- **Solución Implementada:**
  - Creada la clase de servicio `ViaticoBloqueFormatter` con métodos `ObtenerTextoCierre(ViaticoBloque)` y `FormatearNodoNivel2(...)`.
  - Registrada la compilación en `ConversorPDF.XLL.csproj`.
  - En `FormVerVinculaciones.ConstruirArbolViaticos`, delegada la generación de `textoBloque` a `ViaticoBloqueFormatter.FormatearNodoNivel2(...)`.

#### 3. Fase 681 — Menú Contextual en `FormVerVinculaciones`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:**
  - Permitir al usuario alternar la pernocta directamente haciendo clic derecho sobre cualquier bloque de viático.
- **Solución Implementada:**
  - Añadido `itemPernocta` en `ConfigurarMenuContextual` entre `itemImprimirViatico` y el separador de pool.
  - Visibilidad configurada en `NodeMouseClick` exclusivamente para `_vistaActual == VistaVinculaciones.Viaticos && !_modo.EsModoAg`.
  - Texto dinámico: `bloque.TienePernoctaUltimoDia ? "🛏️ Quitar Pernocta Último Día" : "🛏️ Marcar Pernocta Último Día"`.

#### 4. Fase 682 — Handler `AlternarPernoctaUltimoDiaAsync`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:**
  - Ejecutar la mutación de pernocta de forma consistente para todas las vinculaciones del bloque (Regla A), con confirmación interactiva, transacción en base de datos, persistencia en vault y reconstrucción del árbol preservando la selección visual.
- **Solución Implementada:**
  - Implementado `AlternarPernoctaUltimoDiaAsync()` con confirmación de impacto (`1,5 ➔ 2,5 días`), transacción ACID con rollback ante fallos, llamada a `PersistirCambiosAlVault()`, actualización en memoria de `_items` y restauración de scroll/foco en el nodo correspondiente.

#### 5. Fase 683 — MODO AG (Aislamiento y Documental)
- **Archivo Afectado:** Documental (`implementation_plan.md` / `walkthrough.md`)
- **Diagnóstico:**
  - Evitar inconsistencias o excepciones en MODO AG dado que `ag_vinculaciones` no contiene la columna de pernocta.
- **Solución Implementada:**
  - Opción oculta en MODO AG (`!_modo.EsModoAg`). Registrada como deuda técnica controlada.

#### 6. Fase 684 — Pruebas Unitarias
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Application/Services/ViaticoBloqueFormatterTests.cs`
  - `ConversorPDF.Tests/Academic/Domain/ViaticoBloquePernoctaTests.cs`
- **Diagnóstico:**
  - Validar exhaustivamente el cálculo de la sesión de cierre, el formato del nodo con y sin pernocta, la extracción de IDs únicos de vinculación y el ciclo completo de alternancia (on/off).
- **Solución Implementada:**
  - `ViaticoBloqueFormatterTests.cs`: 6 pruebas unitarias (`ObtenerTextoCierre` con bloque nulo, vacío, un solo día, múltiples sesiones ordenadas por hora de inicio, formato 12h y formateo de Nivel 2).
  - `ViaticoBloquePernoctaTests.cs`: 2 pruebas añadidas (`IdsVinculacion_WhenSesionesConMultiplesVinculaciones_ShouldReturnDistinctIds` y `AlternarPernocta_WhenMutandoSesionesDelBloque_ShouldToggleEntre1Punto5Y2Punto5`).

### 157.79 Ronda 121 — Fases 685 a 689: Robustecimiento de Pernocta y Cobertura de Repositorio

#### 1. Fase 685 — Mensaje de Confirmación con Duración Dinámica
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:** `cambioDias` estaba hardcodeado a `1,5 ➔ 2,5 días`, resultando erróneo para bloques de otra extensión (ej. 1 día: `0,5 ➔ 1,5 días`; 3 días: `2,5 ➔ 3,5 días`).
- **Solución Implementada:** Derivado dinámicamente de `bloque.DuracionTotal`: `double durNueva = valorNuevo == 1 ? durActual + 1 : durActual - 1; string cambioDias = $"{durActual} ➔ {durNueva} días";`.

#### 2. Fase 686 — Aislamiento Cultural en Pruebas de Formatter
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Application/Services/ViaticoBloqueFormatterTests.cs`
- **Diagnóstico:** Las aserciones `"1,5 días"` y `"$120.000"` fallarían en runners con cultura `en-US` o invariante.
- **Solución Implementada:** En el constructor del test, se fija `CultureInfo.DefaultThreadCurrentCulture` y `CultureInfo.DefaultThreadCurrentUICulture` a `es-CO`.

#### 3. Fase 687 — Notificación `MarkDirty()` Post-Commit
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:** `AcademicRepository` notificaba `NotifyMarkDirty()` dentro de la transacción por cada UPDATE, pudiendo encolar sincronización antes del commit.
- **Solución Implementada:** Instanciado el repositorio con `suppressMarkDirty: true` y ejecutado `SyncContext.Scheduler?.MarkDirty()` de forma explícita tras `tx.Commit()`.

#### 4. Fase 688 — Restauración de Selección por Coincidencia Exacta
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:** `nDoc.Text.Contains(docenteSeleccionado)` fallaba ante nombres con prefijos compartidos (ej. "ADOLFO LEON GUERRERO" vs "ADOLFO LEON GUERRERO GARCIA").
- **Solución Implementada:** Búsqueda exacta evaluando `b.Sesiones.Any(s => string.Equals(s.DocenteNombre, docenteSeleccionado, StringComparison.Ordinal))` sobre el bloque correspondiente, con retorno anticipado una vez restaurado el nodo.

#### 5. Fase 689 — Cobertura Unitaria de `SetPernoctaUltimoDia`
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AcademicRepositorySetPernoctaTests.cs`
- **Diagnóstico:** Ausencia de pruebas sobre la persistencia SQL de `SetPernoctaUltimoDia`.
- **Solución Implementada:** Creada suite de pruebas unitarias SQLite en memoria verificando activación (0→1), desactivación (1→0), id inexistente (0 filas) y normalización de valores.

### 157.80 Ronda 122 — Fases 690 y 691 (Fase 666 y 666.b): Rango de Período Académico como Advertencia Confirmable y Ampliación de Ventana Maestra

#### 1. Fase 690 (Fase 666) — Validación de Rango de Fechas del Período (Advertencia Confirmable)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** El bloqueo rígido (`MessageBoxButtons.OK` con `return;`) impedía guardar vinculaciones válidas cuyas sesiones excedían ligeramente la ventana maestra configurada inicialmente en el formulario.
- **Solución Implementada:** Reemplazo por `MessageBoxButtons.YesNo` con advertencia confirmable. Si el usuario confirma continuar (`DialogResult.Yes`), se establece `ampliarRangoConfirmado = true;`.

#### 2. Fase 691 (Fase 666.b) — Ampliación Automática de la Ventana Maestra
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** Al confirmar sesiones fuera de rango, persistir `dtpInicio.Value` y `dtpFinal.Value` desincronizados generaba inconsistencias entre la vinculación y sus sesiones hijas.
- **Solución Implementada:** Antes de persistir, si `ampliarRangoConfirmado == true`, se desuscribe temporalmente `dtpInicio.ValueChanged` y se expande `dtpInicio.Value = minFecha` y `dtpFinal.Value = maxFecha` para envolver fielmente las fechas reales de las sesiones.

### 157.81 Ronda 123 — Verificación de Fases 690/691 y Robustecimiento de Edición de Fechas (Fases 692 a 695)

#### 1. Fase 692 — Instrucción de Remediación en Salida "No"
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** El diálogo informativo de fechas fuera de rango no indicaba al usuario la alternativa de corrección en caso de responder "No".
- **Solución Implementada:** Añadida la indicación explícita: `"Si selecciona 'No', ajuste las fechas del grid o modifique [Fecha Inicio]/[Fecha Final] antes de guardar.\n\n"`.

#### 2. Fase 693 — Tolerancia de Formatos y Objetos DateTime en `GridCellValueParser`
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Mappers/GridCellValueParser.cs`
  - `ConversorPDF.Tests/Unit/GridCellValueParserTests.cs`
- **Diagnóstico:** `TryParseFecha` rechazaba entradas con componentes de hora (`dd/MM/yyyy HH:mm:ss`), objetos `DateTime` nativos o formatos de 2 dígitos de año (`d/M/yy`, `dd/MM/yy`), provocando reversiones silenciosas de celdas en el grid.
- **Solución Implementada:** Manejo directo de `value is DateTime dt`, recorte de componente horario (`txt.IndexOf(' ')`), incorporación de formatos adicionales (`"d/M/yy", "dd/MM/yy", "yyyy/MM/dd", "yyyyMMdd"`) y ampliación de pruebas unitarias.

#### 3. Fase 694 — Evaluación de Semántica de Período vs Ventana de Generación
- **Estado:** RETIRADA (Diseño intencional)
- **Diagnóstico:** La ventana maestro-detalle de la vinculación se define intencionalmente por `[dtpInicio, dtpFinal]`, no por el semestre completo, para validar la programación específica de la asignatura.

#### 4. Fase 695 — Alineación de Trazabilidad Código vs Rondas
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** Los comentarios en código rotulaban "Fase 666 / Fase 666.b", mientras que el plan y walkthrough documentan Fases 690 / 691.
- **Solución Implementada:** Comentarios actualizados a `Fase 690 (plan 666)` y `Fase 691 (plan 666.b)`.

### 157.82 Ronda 124 — Robustecimiento de Fechas de 2 Dígitos, Separadores Whitespace y Terminología de Ventana (Fases 696 a 698)

#### 1. Fase 696 — Normalización de Años de 2 Dígitos
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Mappers/GridCellValueParser.cs`
  - `ConversorPDF.Tests/Unit/GridCellValueParserTests.cs`
- **Diagnóstico:** Al admitir `d/M/yy` y `dd/MM/yy` con `CultureInfo.InvariantCulture`, años a partir de 30 (ej. "15/10/30") se resolvían a 1930 por `TwoDigitYearMax = 2029`, permitiendo guardar fechas de siglo pasado.
- **Solución Implementada:** Añadida normalización si `fecha.Year < 2000 => fecha = fecha.AddYears(100)` y caso de prueba unitaria en `GridCellValueParserTests.cs`.

#### 2. Fase 697 — Terminología Precisa en Mensaje de Validación
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** El mensaje visible y el log rotulaban "fuera del período académico", perpetuando la ambigüedad tras confirmarse que la validación evalúa la ventana de la vinculación actual.
- **Solución Implementada:** Ajustado el texto del diálogo a "fuera de la ventana de la vinculación" y el título a "Fechas Fuera de la Ventana", así como el log correspondiente.

#### 3. Fase 698 — Recorte de Separadores Horarios Whitespace Múltiples (Espacio, Tab, NBSP)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Mappers/GridCellValueParser.cs`
  - `ConversorPDF.Tests/Unit/GridCellValueParserTests.cs`
- **Diagnóstico:** `IndexOf(' ')` fallaba al recortar componentes horarios cuando el separador era tabulador o espacio no separable (`\u00A0`), común al pegar desde navegadores o Excel.
- **Solución Implementada:** Utilizado `txt.Split` con separadores `{' ', '\t', '\r', '\n', '\u00A0'}` descartando entradas vacías y tomando el primer segmento. Pruebas unitarias extendidas.

### 157.83 Ronda 125 — Auto-ampliación Silenciosa de Ventana Maestra al Grid y Sincronización de Snapshot de Fechas (Fase 699)

#### 1. Fase 699 — Auto-ampliación de Ventana Maestra y Actualización FI/FF en Snapshot
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** Al modificar manualmente `Fecha Inicio` o `Fecha Final` en el formulario para acomodar sesiones extendidas, `ParametrosEstanDesincronizados()` reportaba falso positivo de desincronización ("Falta Generar Sesiones") porque `_ultimoEstadoGenerado` preservaba las fechas iniciales. Además, achicar la ventana por debajo del rango de sesiones del grid provocaba advertencias innecesarias al guardar.
- **Solución Implementada:**
  - Incorporada bandera anti-reentrancia `_ajustandoFechaMaestra`.
  - Creados helpers `TryObtenerRangoSesionesGrid`, `AmpliarVentanaMaestraAlGrid` y `SincronizarFechasEnEstadoGenerado`.
  - Actualizados handlers `DtpInicio_ValueChanged` y `DtpFinal_ValueChanged` suscritos simétricamente en el constructor.
  - Sincronización simétrica de supresión de eventos (`-= / +=`) para `dtpInicio` y `dtpFinal` en Generar, Fase 691, Post-commit y `CargarEstadoPrevio`.

### 157.84 Ronda 126 — Blindaje de Rango DateTimePicker, Filtro de Filas Marcadas y Desduplicación de Logs (Fases 700 a 703)

#### 1. Fase 700 — Acotamiento contra MinDate/MaxDate en Auto-ampliación
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** Asignar fechas anómalas aceptadas por el parser a `dtpInicio.Value` o `dtpFinal.Value` podía arrojar `ArgumentOutOfRangeException` dentro del handler `ValueChanged` en WinForms.
- **Solución Implementada:** Acotado preventivo `primera` contra `dtpInicio.MinDate` y `ultima` contra `dtpFinal.MaxDate` antes de asignar a los controles.

#### 2. Fase 701 — Criterio de Selección de Filas Marcadas en `TryObtenerRangoSesionesGrid`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** El rango de fechas se calculaba sobre todas las filas del grid, permitiendo que filas desmarcadas no persistidas ampliaran indebidamente la ventana maestra.
- **Solución Implementada:** Filtrado explícito con `if (!SafeCellBool(row.Cells["ColChecked"].Value)) continue;` para alinear con el criterio de guardado de `PrepararSesionesDesdeGrid`.

#### 3. Fase 702 — Actualización Documental del Snapshot `_ultimoEstadoGenerado`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** El comentario XML declaraba que el snapshot se congelaba exclusivamente en `[Generar]`, omitiendo el refresco quirúrgico de `FI/FF`.
- **Solución Implementada:** Documentación XML actualizada explicitando el refresco de fechas por edición manual preservando `T/H/D`.

#### 4. Fase 703 — Supresión de Logs Redundantes por Idempotencia de Snapshot
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** Re-entradas en la cadena de eventos de fechas registraban mensajes de log idénticos consecutivos sin cambios efectivos en el hash.
- **Solución Implementada:** Retorno anticipado `if (nuevo == _ultimoEstadoGenerado) return;` previo a la asignación y logging.

### 157.85 Ronda 127 — Sincronización Inmediata del Snapshot tras Ampliación Pre-Guardado (Fase 704)

#### 1. Fase 704 — Sincronización de Snapshot tras Mutación de Fechas en F691
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** F691 ampliaba `dtpInicio`/`dtpFinal` con eventos desuscritos antes de invocar `GuardarAsync`. Si el guardado fallaba por bloqueo de base de datos o se cancelaba, el snapshot `_ultimoEstadoGenerado` retenía los valores antiguos de `FI/FF`, provocando que el siguiente intento de guardado fuera bloqueado falsamente por `ParametrosEstanDesincronizados()` ("Falta Generar Sesiones").
- **Solución Implementada:** Invocado `SincronizarFechasEnEstadoGenerado()` inmediatamente después del bloque `finally` de F691, asegurando coherencia del snapshot independientemente del resultado de la persistencia asíncrona.

### 157.86 Ronda 128 — Consolidación Integral del Snapshot Pre-Persistencia (Fase 705)

#### 1. Fase 705 — Actualización de Snapshot tras Ajuste de Horas y Ampliación de Fechas
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** Cuando el usuario confirmaba la discrepancia de horas (F2.4), `txtTotalHoras.Text` se mutaba en UI pero `_ultimoEstadoGenerado` retenía el valor `H=` previo. Si la persistencia fallaba o se cancelaba, reintentar el guardado disparaba un falso positivo de "Falta Generar Sesiones".
- **Solución Implementada:** Asignado `_ultimoEstadoGenerado = ObtenerEstadoParametrosCriticos()` inmediatamente antes de inicializar `_guardadoCts`, blindando atómicamente todos los cambios pre-guardado (horas, fechas y parámetros de control).

### 157.87 Ronda 129 — Consistencia Terminológica en Aviso de Parámetros Desincronizados (Fase 706)

#### 1. Fase 706 — Exclusión de "Fechas" del Aviso "Falta Generar Sesiones"
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** El diálogo informativo advertía sobre la modificación de `(Docente, Horas, Distribución de días o Fechas)`. Dado que desde la Fase 699 la edición de fechas maestras se sincroniza quirúrgicamente en el snapshot y auto-amplía al grid, las fechas ya no constituyen causa de desincronización de generación.
- **Solución Implementada:** Removido el literal `"o Fechas"` del mensaje interactivo, restringiéndolo a `(Docente, Horas o Distribución de días)`.

### 157.88 Ronda 130 — Rango de Comisión de Viáticos con Pernocta en Último Día (Fases 707 a 710)

#### 1. Fase 707 — `AsignaturasExcelExporter.cs`: Rango Efectivo de Comisión
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:** La exportación a plantillas `FV-FO-033` escribía `FechaMaxima` directamente en la columna de fin de comisión sin considerar si el bloque finalizaba con pernocta (`isPernoctaUltimoDia = 1`). Esto provocaba que docentes con pernocta el último día tuvieran una fecha de fin igual al último día de clase, calculando 1,5 días en vez de 2,5 días en las fórmulas de la plantilla.
- **Solución Implementada:** Sustituido el uso directo de `bloquesParaEstaFila[i].FechaMaxima` por `ViaticoBloqueFormatter.ObtenerRangoComision(bloquesParaEstaFila[i])?.Fin`, preservando `FechaMinima` como inicio y asignando la fecha de regreso extendida calculada en el dominio.

#### 2. Fase 708 — `ViaticoBloqueFormatter.cs`: Método de Rango de Comisión y Pruebas Unitarias
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/ViaticoBloqueFormatter.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/ViaticoBloqueFormatterTests.cs`
- **Diagnóstico:** Se requería una función pura desacoplada que expusiera el rango efectivo `(Inicio, Fin)` de comisión para un bloque, aprovechando `bloque.FechaFinComision` y respetando el principio de responsabilidad única y *single-type-per-file*.
- **Solución Implementada:** Incorporado método `ObtenerRangoComision(ViaticoBloque bloque)` retornando `(DateTime Inicio, DateTime Fin)?` con fechas normalizadas `.Date`. Agregadas pruebas unitarias: con pernocta último día (fin extendido), sin pernocta (fin igual a `FechaMaxima`) y bloque vacío (`null`).

#### 3. Fase 709 — Documentación Técnica y Respaldo Snapshot
- **Archivos Afectados:**
  - `docs/articles/academic-viaticos.md`
  - `_FIXES/fix-20261016-182822/`
- **Diagnóstico:** La documentación técnica debía documentar explícitamente que la fecha de regreso en comisiones refleja `isPernoctaUltimoDia` y cómo la plantilla institucional deduce los días de viático.
- **Solución Implementada:** Actualizado `academic-viaticos.md` con la sección de rango de fechas de comisiones y pernocta en último día. Generado snapshot en `_FIXES/fix-20261016-182822/`.

#### 4. Fase 710 — Verificación Estática y Preservación de Invariantes
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/ViaticoBloque.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/ViaticoConflictAnalyzer.cs`
- **Diagnóstico:** Asegurar que `ViaticoBloque.FechaMaxima` permanezca intacta (base para algoritmos de consecutividad `AgruparConsecutivos` y solapamiento en `ViaticoConflictAnalyzer.SeSolapan`) y que no existan regresiones en plantillas no modificadas ni en modo AG.
- **Solución Implementada:** Verificación estática completada; invariantes preservados.

### 157.89 Ronda 131 — Verificación Runtime, Sincronización de DocFX y Estandarización de Auditoría (Fases 711 a 713)

#### 1. Fase 711 — Verificación de Compilación y Suite Completa de Pruebas Unitarias
- **Archivos Afectados:**
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.Tests/ConversorPDF.Tests.csproj`
- **Diagnóstico:** Fase 710 concluyó sin verificación por compilador de las tuplas `(DateTime Inicio, DateTime Fin)?` ni ejecución de la suite de pruebas tras la incorporación de las 3 pruebas unitarias de `ObtenerRangoComision`.
- **Solución Implementada:** Ejecutada compilación `dotnet build ConversorPDF.XLL\ConversorPDF.XLL.csproj -c Release -p:Platform=x64` (0 errores) y suite `dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release` (624 de 624 pruebas superadas, 0 errores).

#### 2. Fase 712 — Regeneración del Sitio Estático de Documentación (`docs/_site`)
- **Archivos Afectados:**
  - `docs/_site/articles/academic-viaticos.html`
- **Diagnóstico:** El HTML publicado en `docs/_site` permanecía desincronizado con la fecha previa al cambio de la documentación técnica en Markdown.
- **Solución Implementada:** Ejecutado `docfx build docs/docfx.json` regenerando completamente el sitio estático y sincronizando la sección de rango de comisiones con pernocta en `academic-viaticos.html`.

#### 3. Fase 713 — Corrección de Atribución de Plantillas y Estandarización de Snapshot de Auditoría
- **Archivos Afectados:**
  - `docs/articles/academic-viaticos.md`
  - `_FIXES/fix-20261016-1828.22.000/`
- **Diagnóstico:** La documentación técnica atribuía erróneamente columnas 1..8 a `GF-FO-023` mediante `AsignaturasExcelExporter` cuando corresponde exclusivamente a `FV-FO-033`. Asimismo, el snapshot previo carecía del encabezado de auditoría `# AUDITORÍA DE CAMBIO` y de la nomenclatura `fix-yyyyMMdd-HHmm.ss.fff` con archivos `.txt`.
- **Solución Implementada:** Clarificada la redacción en `academic-viaticos.md` distinguiendo el rol de `FV-FO-033` (`AsignaturasExcelExporter`) y `GF-FO-023` (`ViaticosIndividualExporter`). Creada la carpeta canónica `_FIXES/fix-20261016-1828.22.000` con encabezados de auditoría de cambio (versión 1.6.3.51) y archivos `.txt`, eliminando la copia previa.

### 157.90 Ronda 132 — Formato y Resaltado del Nivel 2 en la Vista Viáticos (Fases 714 a 720)

#### 1. Fase 714 — `ViaticoBloqueFormatterTests.cs`: Pruebas Unitarias de Formato Nivel 2 (TDD)
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Application/Services/ViaticoBloqueFormatterTests.cs`
- **Diagnóstico:** Las pruebas unitarias existentes exigían la presencia del código de trayecto (`VAL-GT045`), el prefijo `Cierre:` y la palabra completa `días`. Se requería adaptar las pruebas existentes a la nueva especificación y añadir cobertura para casos sin pernocta (verificando ausencia de `Cierre:` y `SIN-COD`), con pernocta y bloque nulo.
- **Solución Implementada:** Actualizadas las llamadas a `FormatearNodoNivel2` removiendo el parámetro `codRuta`, validando la abreviación `d`, la remoción de `Cierre: ` y la inclusión de `| Pernocta Ultimo Dia` solo cuando corresponde. Añadidos tests `FormatearNodoNivel2_WhenSinPernocta_ShouldNotIncludeCierrePrefixNorCodigoRuta` y `FormatearNodoNivel2_WhenBloqueNull_ShouldReturnEmpty`.

#### 2. Fase 715 — `ViaticoBloqueFormatter.cs`: Ajuste de Firma y Formateo de Cadena
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/ViaticoBloqueFormatter.cs`
- **Diagnóstico:** El método `FormatearNodoNivel2` concatenaba `codRuta`, la palabra `días` y el prefijo `Cierre: `, saturando la fila visual del árbol.
- **Solución Implementada:** Modificada la firma para eliminar el parámetro `codRuta`. Se actualizó la plantilla de texto a `📅 {fechaIni} - {fechaFin} | {b.DuracionTotal} d | {tipoTransporte} | {ruta} | ${valorTransporte:N0} | {ObtenerTextoCierre(b)}{marca}`, manteniendo intactos `ObtenerTextoCierre` y `ObtenerRangoComision`.

#### 3. Fase 716 — `FormVerVinculaciones.cs`: Integración UI y Resaltado en DarkOrange
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:** En la vista de viáticos, el Nivel 2 mostraba `codRutaDisplay` ("SIN-COD") y se dibujaba uniformemente en `Color.DarkBlue`, sin destacar viajes con pernocta en el último día.
- **Solución Implementada:** Preservada la invocación a `_trayectoEncoderService.GetOrCreateCodigoTrayecto` (garantizando pre-generación en base de datos), eliminada la variable local `codRutaDisplay`, adaptada la llamada a `FormatearNodoNivel2` con 6 argumentos y asignado `ForeColor = bloque.TienePernoctaUltimoDia ? Color.DarkOrange : Color.DarkBlue`.

#### 4. Fase 717 — Documentación Técnica de Formato y Resaltado Nivel 2
- **Archivos Afectados:**
  - `docs/articles/academic-viaticos.md`
  - `docs/articles/changes/2026-10-16-ajuste-formato-nivel2-viaticos.md`
- **Diagnóstico:** La documentación técnica debía reflejar la estructura exacta del nodo Nivel 2, la semántica de la abreviatura `d`, el retiro de `Cierre:` y código de ruta de la vista (aclarando la preservación de `COD_RUTA` en exportación) y el resaltado visual condicional por pernocta.
- **Solución Implementada:** Incorporada subsección detallada en `academic-viaticos.md` y creado artículo formal de cambio en `docs/articles/changes/2026-10-16-ajuste-formato-nivel2-viaticos.md`.

#### 5. Fase 718 — Sincronización de Sitio Estático DocFX (`docs/_site`)
- **Archivo Afectado:** `docs/_site/articles/academic-viaticos.html`
- **Diagnóstico:** El HTML publicado en `docs/_site` no contenía la nueva subsección de formato Nivel 2 ni el resaltado en `DarkOrange` al no haberse ejecutado DocFX tras la edición Markdown.
- **Solución Implementada:** Ejecutado `docfx build docfx.json` desde el directorio `docs/` (0 errores, 1 warning pre-existente en index.md); HTML estático sincronizado al 100%.

#### 6. Fase 719 — Respaldo de Seguridad y Estandarización de Snapshot `_FIXES`
- **Archivos Afectados:** `_FIXES/fix-20261016-1908.00.000/`, `scratch/fix_snapshot.ps1`
- **Diagnóstico:** Se requería preservar copia de seguridad histórica `.txt` con encabezado `# AUDITORÍA DE CAMBIO` conforme a la convención canónica del repositorio.
- **Solución Implementada:** Generada carpeta `_FIXES/fix-20261016-1908.00.000/` con respaldos de `ViaticoBloqueFormatter.txt`, `ViaticoBloqueFormatterTests.txt`, `FormVerVinculaciones.txt` y `academic-viaticos.txt` identificados con versión `1.6.3.52`.

#### 7. Fase 720 — Corrección de Nomenclatura de Test en Registro de Cambio
- **Archivo Afectado:** `docs/articles/changes/2026-10-16-ajuste-formato-nivel2-viaticos.md`
- **Diagnóstico:** En la sección Test Coverage se citaba `FormatearNodoNivel2_WhenSinPernocta_ShouldFormatExpectedString` (nombre transitorio del plan preliminar) en lugar de la prueba real `FormatearNodoNivel2_WhenSinPernocta_ShouldNotIncludePernoctaUltimoDiaSuffix`.
- **Solución Implementada:** Sincronizado el nombre exacto de la prueba en la documentación de cambio.

---

## 📦 Entrega de Release v1.6.3.54 (Instalador Oficial)

- **Versión Oficial:** `1.6.3.54`
- **Archivos de Versión Sincronizados:**
  - `ConversorPDF.XLL\Core\Helpers\App.cs` (`APP_VERSION = "1.6.3.54"`)
  - `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (`<Version>1.6.3.54</Version>`)
  - `installer\InstaladorConversorPDF.wxs` (`Version="1.6.3.54"`)
- **Documentación Técnica Generada (`/build-docs`):**
  - Generado sitio web estático DocFX (`docs\_site\`) sin errores.
  - Sincronizados XML comments, guías conceptuales en `docs\articles\academic-viaticos.md` y registro de cambios.
- **Verificación Integral de Pruebas Unitarias:**
  - Suite completa: **626 / 626 pruebas aprobadas (0 errores)** (`dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`).
- **Compilación de Add-In:**
  - `ConversorPDF.XLL` (Release|x64): **0 Errores** (`ConversorPDF.XLL-AddIn64.xll`).
- **Paquete Instalador Generado (`/build-release 1.6.3.54`):**
  - Archivo: `installer\InstaladorConversorPDFv1.6.3.54.msi`
  - Tamaño: **5,464,064 bytes** (5.46 MB)
  - Estado: Listo para distribución con auto-activación en Excel.

---

### 157.91 Ronda 133 — Sincronización de Plantillas Excel en el Instalador y Blindaje de Reemplazo (Fases 721 a 727)

#### 1. Fase 721 — Sincronización de Plantilla `FV-FO-033.xlsx` y Secuencia de Build
- **Archivo Afectado:** `ConversorPDF.XLL\ExcelTemplates\FV-FO-033.xlsx`, `installer\InstaladorConversorPDF.wxs`
- **Diagnóstico:** El instalador v1.6.3.54 empaquetó la versión obsoleta (736.198 B) en vez de la versión actualizada (777.061 B) debido a una inversión temporal en la secuencia de empaquetado/guardado de plantilla.
- **Solución Implementada:** Identificada la causa raíz y documentada la invariante de empaquetado en `build-release.md`. Preparada la versión 1.6.3.55 para distribución completa de la plantilla de 777.061 B.

#### 2. Fase 722 — Blindaje Integral de Reemplazo de Plantillas en `TemplateComponents`
- **Archivo Afectado:** `installer/InstaladorConversorPDF.wxs`
- **Diagnóstico:** Los archivos `.xlsx` carecen de versión PE en la tabla `File` del MSI. En un MajorUpgrade, si el archivo en destino fue tocado o tiene fecha posterior, Windows Installer (`CostFinalize`) asigna `FA_NoOverwrite` y omite copiarlo durante `InstallFiles`, aun si `RemoveFile` lo elimina durante `RemoveFiles`.
- **Solución Implementada:** 
  1. Conservados los mismos `File/@Id` (`filTemplate1`, `filTemplate2`) y Component GUIDs para mantener linaje e integridad MSI.
  2. Asignado `DefaultVersion="!(bind.FileVersion.filXllDll)"` a ambos archivos `<File>` para que Windows Installer reconozca una versión sintética válida ligada al assembly principal y fuerce la sobreescritura de cualquier archivo sin versión en disco.
  3. Declarados `RegistryValue` con `KeyPath="yes"` en `HKCU\Software\ESAP\ConversorPDF\Templates` con valor `[ProductVersion]` para forzar que los componentes se evalúen con estado de instalación activo en instalaciones per-user.
  4. Declarados elementos `<RemoveFile Id="PurgeOldTemplate1" Name="GF-FO-023.xlsx" On="install" />` y `<RemoveFile Id="PurgeOldTemplate2" Name="FV-FO-033.xlsx" On="install" />` para purga física en la secuencia `RemoveFiles` (3500) antes de `InstallFiles` (4000).

#### 3. Fase 723 — Eliminación de Carpeta `ExcelTemplates` en Desinstalación (`RemoveFolder`)
- **Archivo Afectado:** `installer/InstaladorConversorPDF.wxs`
- **Diagnóstico:** Al desinstalar el Add-In, la subcarpeta `ExcelTemplates` quedaba residente en `%APPDATA%\Microsoft\AddIns\ExcelTemplates` con sus archivos huérfanos (~3,7 MB), incumpliendo el estándar de desinstalación limpia.
- **Solución Implementada:** Declarado `<RemoveFolder Id="RemoveExcelTemplatesFolder" Directory="dir_ExcelTemplates" On="uninstall" />` en `cmpTemplate2` para asegurar la remoción de la carpeta del producto si queda vacía.

#### 4. Fase 724 — Corrección de Codificación (Mojibake) en Snapshots `_FIXES`
- **Archivos Afectados:** `.agents/workflows/contexto-previo.md`, `scratch/fix_snapshot.ps1`, `_FIXES/fix-20261016-1828.22.000/`, `_FIXES/fix-20261016-1908.00.000/`
- **Diagnóstico:** `Get-Content / Set-Content -Encoding UTF8` en scripts de PowerShell 5.1 sin BOM causaba lectura en ANSI (CP-1252) y re-escritura en UTF-8 con doble codificación persistida (mojibake `# AUDITORÃA DE CAMBIO` y caracteres acentuados).
- **Solución Implementada:** Reemplazado el patrón por `[System.IO.File]::WriteAllText($destino, $header + [System.IO.File]::ReadAllText(..., [System.Text.Encoding]::UTF8), $utf8)` en `contexto-previo.md` y `fix_snapshot.ps1`. Revertido y saneado el mojibake en los 7 archivos `.txt` de los snapshots de Ronda 131 y 132 a UTF-8 válido con BOM.

#### 5. Fase 725 — Sincronización de Encabezados y Registro de Invariante de Release
- **Archivos Afectados:** `docs/articles/changes/2026-10-16-ajuste-formato-nivel2-viaticos.md`, `.agents/workflows/build-release.md`, `implementation_plan.md`, `walkthrough.md`
- **Diagnóstico:** Encabezados de Ronda 132 desincronizados citando fases 714 a 717 en vez de 714 a 720; ausencia de advertencia en el workflow de release para verificación de tamaño de plantillas `.xlsx`.
- **Solución Implementada:** Actualizados los encabezados a `(Fases 714 a 720)` en plan, walkthrough y documento de cambios. Añadida la sección `⚠️ Lección Aprendida 2 — Plantillas ExcelTemplates (contenido, no código)` en `.agents/workflows/build-release.md`.

#### 6. Fase 726 (Opcional) — Evaluación de Referencias Irresolubles en `.dna`
- **Archivo Afectado:** `ConversorPDF.XLL/ConversorPDF.XLL-AddIn64.dna`
- **Diagnóstico:** Líneas 18-20 declaran `<Reference Path="System.Buffers.dll" />`, `System.Memory.dll`, `System.Numerics.Vectors.dll`, ausentes en `bin\x64\Release`.
- **Estado:** Pendiente de selección de alternativa (Opción A: retirar del `.dna`; Opción B: desplegar binarios en el instalador).

#### 7. Fase 727 (Opcional) — Limpieza de Código Muerto en `ViaticoBloque.cs` y Deconstrucción
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/ViaticoBloque.cs`, `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:** `ViaticoBloque.DisplayText` obsoleto sin consumidores y variables no usadas en deconstrucción de tupla (`origen`, `escala`, `destino`).
- **Estado:** Pendiente de aprobación para aplicación quirúrgica.

---

### 157.92 Ronda 134 — Habilitar "Usar Autofiltro" en Hoja Protegida de Exportar Asignaturas (Fase 728)

#### 1. Fase 728 — Habilitar "Usar Autofiltro" en Hoja Protegida de Exportar Asignaturas
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:** `ws.Protect(ConstantsHelper.FORMATO_PASSWORD)` en el bloque `finally` omite los argumentos opcionales booleanos de COM, estableciendo por defecto todos los permisos `Allow*` en `false`. Como resultado, los usuarios no pueden interactuar con los botones de autofiltro en la fila 13 de la hoja `PROGRAMACIÓN` tras la exportación.
- **Solución Implementada:** Se implementó el método auxiliar `ProtegerHojaPermitiendoAutofiltro(dynamic ws, string password)` utilizando `Type.InvokeMember` con 16 parámetros posicionales (`Missing.Value`), asignando la contraseña en `args[0]` y habilitando explícitamente `args[14] = true` (`AllowFiltering`). Se sustituyó la llamada directa en el bloque `finally`.
- **Documentación:** Creado registro en `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`.

---

### 157.93 Ronda 135 — Robustecimiento de Protección con Autofiltro/Ordenación y Trazabilidad de Rango (Fases 729 a 732)

#### 1. Fase 729 — `AllowSorting` Habilitado en Protección de Hoja
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:** Al habilitar únicamente `AllowFiltering` (`args[14] = true`), las opciones de ordenamiento ("Ordenar de A a Z / Z a A") del menú desplegable del autofiltro permanecían deshabilitadas debido a que Excel gestiona `AllowSorting` de forma independiente.
- **Solución Implementada:** Se configuró explícitamente `args[13] = true` (`AllowSorting`) en `ProtegerHojaPermitiendoAutofiltro`, permitiendo el ordenamiento directamente desde las cabeceras filtrables.

#### 2. Fase 730 — Protección Fallback ante Excepción en Invocación Reflexiva
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:** Si la llamada por reflexión `ProtegerHojaPermitiendoAutofiltro` fallaba (por desajuste de versión de Excel o fallo COM transitorio), el bloque `catch` silencioso dejaba el archivo completamente desprotegido.
- **Solución Implementada:** Se implementó fallback en el bloque `catch`: ante error en la invocación reflexiva se invoca `ws.Protect(ConstantsHelper.FORMATO_PASSWORD)` garantizando que la hoja nunca quede desprotegida.

#### 3. Fase 731 — Invariante de Rango de Datos de la Tabla `T_Principal8`
- **Archivo Afectado:** `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`
- **Diagnóstico:** La tabla nativa en plantilla `FV-FO-033` (`T_Principal8`) tiene rango fijo `A13:DF234` (221 filas de datos). Exportaciones con >221 filas sobrepasarían el rango de la tabla y del autofiltro.
- **Solución Implementada:** Se adoptó la Opción B (no destructiva): documentar el límite de 221 filas como invariante formal de diseño de la plantilla sin forzar mutaciones con `ListObject.Resize` que arriesguen la integridad de las fórmulas precalculadas de la plantilla.

#### 4. Fase 732 — Trazabilidad de Cobertura de Pruebas y Smoke Test
- **Archivo Afectado:** `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`
- **Diagnóstico:** Ausencia de especificación explícita sobre el entorno de prueba, la tabla involucrada (`T_Principal8`) y el comportamiento en los límites de fila.
- **Solución Implementada:** Se actualizó la sección de cobertura en el documento de cambios detallando el análisis estático COM y las condiciones del smoke test manual en Excel 365 x64.

---

### 157.94 Ronda 136 — Reversión de AllowSorting, Trazabilidad en Fallback y Blindaje de Reprotección (Fases 733 a 735)

#### 1. Fase 733 — Reversión de `AllowSorting` por Restricción de Celdas Bloqueadas
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`, `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`
- **Diagnóstico:** `AllowSorting = true` habilita el comando de menú pero Excel exige que todas las celdas del rango a ordenar estén desbloqueadas. Al tener celdas bloqueadas en la plantilla, el ordenamiento desde el desplegable arroja un cuadro de error de hoja protegida.
- **Solución Implementada:** Se revirtió la asignación de `args[13]`, dejándola en `Missing.Value` (deshabilitada) y preservando exclusivamente `args[14] = true` (`AllowFiltering`). Se ajustó la documentación técnica reflejando que el ordenamiento no está habilitado en esta plantilla.

#### 2. Fase 734 — Supresión de Reprotección sobre Hoja No Desprotegida
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:** Si la hoja tenía una contraseña distinta a la interna, la excepción ocurría pero en el bloque `finally` se intentaba invocar `Protect` sobre una hoja aún protegida, produciendo inconsistencias o excepciones COM silenciadas.
- **Solución Implementada:** Se introdujo la bandera booleana `hojaDesprotegida`, asignada con `!ws.ProtectContents` tras `ws.Unprotect(...)`. En el `finally`, la reprotección se ejecuta exclusivamente si `hojaDesprotegida == true`.

#### 3. Fase 735 — Observabilidad y Trazabilidad en Fallback de Protección
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:** El bloque `catch` del `finally` ejecutaba un fallback silencioso sin emitir logs de advertencia o error en caso de fallo reflexivo o del fallback.
- **Solución Implementada:** Se incorporó logging con `LoggerHelper.Warning` cuando la llamada reflexiva falla y `LoggerHelper.Error` con excepción si el fallback tampoco tiene éxito.

---

### 157.95 Ronda 137 — Blindaje de Bandera de Desprotección, Sincronización de Versión 1.6.3.57 y Saneamiento Documental (Fases 736 a 738)

#### 1. Fase 736 — Independencia de Bandera `hojaDesprotegida` respecto a Lectura COM
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:** Si la evaluación `hojaDesprotegida = !ws.ProtectContents` lanzaba excepción COM dentro del `try`, la bandera permanecía en `false` provocando que el `finally` omitiera la reprotección de una hoja ya desprotegida.
- **Solución Implementada:** Se asigna `hojaDesprotegida = true` inmediatamente tras la ejecución sin excepción de `ws.Unprotect(ConstantsHelper.FORMATO_PASSWORD)`. Posteriormente, si `ws.ProtectContents` confirma que la hoja sigue protegida (contraseña distinta), se anula `hojaDesprotegida = false`.

#### 2. Fase 737 — Sincronización de Versión Oficial a 1.6.3.57
- **Archivos Afectados:** `ConversorPDF.XLL/Core/Helpers/App.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `installer/InstaladorConversorPDF.wxs`
- **Diagnóstico:** Las fases 728-735 se incorporaron con posterioridad al empaquetado del instalador oficial 1.6.3.56; publicar un nuevo instalador con el mismo número de versión rompería la actualización por `MajorUpgrade`.
- **Solución Implementada:** Incrementada la versión a `1.6.3.57` en los tres archivos fuente del producto. La compilación y generación del MSI quedan preparadas para ejecución en el workflow `/build-release 1.6.3.57`.

#### 3. Fase 738 (Opcional) — Saneamiento Documental y Encabezados de Release
- **Archivos Afectados:** `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`, `implementation_plan.md`, `walkthrough.md`
- **Diagnóstico:** Presencia del literal de contraseña institucional en texto plano y encabezados de entrega de release desfasados.
- **Solución Implementada:** Sustituido el literal por la referencia a `ConstantsHelper.FORMATO_PASSWORD`, fechado el smoke test (2026-10-17 / Excel 365 x64 / build 1.6.3.57) y actualizados los encabezados a `1.6.3.57`.

---

### 157.96 Ronda 138 — Sincronización de DefaultVersion de Plantillas y Corrección de Estado de Release (Fases 739 y 740)

#### 1. Fase 739 — `DefaultVersion` de Plantillas en Instalador Sincronizado a 1.6.3.57
- **Archivo Afectado:** `installer/InstaladorConversorPDF.wxs`
- **Diagnóstico:** Los componentes `cmpTemplate1` y `cmpTemplate2` conservaban `DefaultVersion="1.6.3.56"`, lo que impedía que en una actualización 1.6.3.56 → 1.6.3.57 Windows Installer (`CostFinalize`) reconociera una versión superior sobre los archivos instalados en disco, arriesgando la anulación del blindaje de sobreescritura de plantillas.
- **Solución Implementada:** Se actualizaron los atributos `DefaultVersion` a `"1.6.3.57"` en `filTemplate1` (`GF-FO-023.xlsx`) y `filTemplate2` (`FV-FO-033.xlsx`), así como los comentarios explicativos correspondientes.

#### 2. Fase 740 — Corrección de Trazabilidad en Sección de Entrega de Release
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`
- **Diagnóstico:** La sección declaraba erróneamente resultados de pruebas unitarias, compilación y generación de MSI 1.6.3.57 como completados cuando en realidad se encuentran pendientes de ejecución bajo el workflow de empaquetado.
- **Solución Implementada:** Se corrigió la redacción marcando suite de pruebas, compilación de Add-In y artefacto MSI como pendientes de ejecución mediante `/build-release 1.6.3.57`.

---

### 157.97 Ronda 139 — Enlace Dinámico de Plantillas y Sincronización Integral de Release (Fases 741 a 743)

#### 1. Fase 741 — Invariante de Tamaños de Plantilla Sincronizado en `build-release.md`
- **Archivo Afectado:** `.agents/workflows/build-release.md`
- **Diagnóstico:** El checklist de release contenía tamaños desactualizados (`777061` y `2938523` bytes) en lugar de los tamaños vigentes de las plantillas en disco (`777082` y `2938592` bytes), induciendo a falsas discrepancias durante la verificación del instalador.
- **Solución Implementada:** Se corrigieron los valores en el checklist del workflow `build-release.md` a `777082` y `2938592`.

#### 2. Fase 742 — Enlace Dinámico de Versión de Plantillas con `!(bind.FileVersion.filXllDll)`
- **Archivo Afectado:** `installer/InstaladorConversorPDF.wxs`
- **Diagnóstico:** El uso de literales hardcodeados en `DefaultVersion` exigía mantenimiento manual repetitivo en cada bump y contradecía el estándar de enlace automático especificado en el workflow de release.
- **Solución Implementada:** Se asignó `DefaultVersion="!(bind.FileVersion.filXllDll)"` en `filTemplate1` y `filTemplate2`, vinculando de forma atómica y automática la versión sintética de las plantillas a la del ensamblado `filXllDll`.

#### 3. Fase 743 — Sincronización de Encabezados y Lista de Cambios en Documento Técnico
- **Archivo Afectado:** `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`
- **Diagnóstico:** El documento técnico omitía las rondas 137 y 138 en el encabezado y carecía del detalle de `hojaDesprotegida` (F736) y del enlace dinámico de plantillas en el instalador.
- **Solución Implementada:** Se sincronizó la cabecera cubriendo las Rondas 134 a 139 y se documentaron las mejoras en la sección Changes.

---

## 📦 Entrega de Release v1.6.3.57 (Instalador Oficial)

- **Versión Oficial:** `1.6.3.57`
- **Archivos de Versión Sincronizados:**
  - `ConversorPDF.XLL\Core\Helpers\App.cs` (`APP_VERSION = "1.6.3.57"`)
  - `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (`<Version>1.6.3.57</Version>`)
  - `installer\InstaladorConversorPDF.wxs` (`Version="1.6.3.57"`, `DefaultVersion="!(bind.FileVersion.filXllDll)"`)
- **Blindaje Integral de Plantillas Excel en Instalador:**
  - `File/@Id` conservado: `filTemplate1` (`GF-FO-023.xlsx`), `filTemplate2` (`FV-FO-033.xlsx`).
  - `FV-FO-033.xlsx`: **777.082 bytes**, `DefaultVersion="!(bind.FileVersion.filXllDll)"` en `InstaladorConversorPDF.wxs`.
  - `GF-FO-023.xlsx`: **2.938.592 bytes**, `DefaultVersion="!(bind.FileVersion.filXllDll)"` en `InstaladorConversorPDF.wxs`.
  - `DefaultVersion="!(bind.FileVersion.filXllDll)"`: Anula `FA_NoOverwrite` en `CostFinalize`, garantizando sobreescritura contra cualquier archivo sin versión en disco.
  - `RegistryValue KeyPath="yes"`: Claves en `HKCU\Software\ESAP\ConversorPDF\Templates` (`[ProductVersion]`) asegurando activación del componente.
  - `RemoveFile` integrado (`PurgeOldTemplate1`, `PurgeOldTemplate2`) para purga forzada en instalación.
  - `RemoveFolder` integrado (`RemoveExcelTemplatesFolder`) para desinstalación limpia de `dir_ExcelTemplates`.
- **Verificación Integral de Pruebas Unitarias:**
  - Suite completa: **626 / 626 pruebas aprobadas (0 errores, 46 s)** (`dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`).
- **Compilación de Add-In:**
  - `ConversorPDF.XLL` (Release|x64): **0 Errores** (`ConversorPDF.XLL-AddIn64.xll`).
- **Paquete Instalador Generado (`/build-release 1.6.3.57`):**
  - Archivo: `installer\InstaladorConversorPDFv1.6.3.57.msi`
  - Tamaño: **5,578,752 bytes** (5.57 MB)
  - Hash SHA-256: `D309E2CFABE5D110ACB081FFB6E3C15034F3C6BFB97556C0FDCD3AAC528EA49C`
  - Estado: Listo para distribución con auto-activación en Excel y reemplazo blindado dinámico de plantillas.

---

### 157.98 Ronda 140 — Reparación de Esquema CRDT en Periodos y Resolución de Hueco de Migración V50/V51 (Fases 744 a 751)

#### 1. Fase 744 — Helper `PeriodoVentanaCrdtRepair` para Soltar Triggers CRDT Obsoletos
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/PeriodoVentanaCrdtRepair.cs` (Nuevo)
- **Diagnóstico:**
  - `periodos` está registrada como tabla CRR en CR-SQLite (`CrdtTables`). Cuando se intentaba alterar o actualizar la tabla tras cambios de esquema, los triggers `periodos__crsql_itrig/utrig/dtrig` generados con la aridad previa (4 columnas) disparaban `SQL logic error expected 13 values, got 9` al ejecutar sentencias DML de backfill.
- **Solución Implementada:**
  - Creación del helper estático `PeriodoVentanaCrdtRepair` con el método `DropStaleCrsqlTriggers(connection, transaction)` que ejecuta de forma transaccional `DROP TRIGGER IF EXISTS` sobre los tres triggers stale.
  - Preserva intacta la tabla `periodos__crsql_clock` para mantener el historial de sincronización. Los triggers son recreados limpiamente por `SyncInitializer.EnableAllCrdts` con el esquema actualizado.

#### 2. Fase 745 — Migración V65: `MigrationV65_PeriodoVentanaCrdtRepair`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV65_PeriodoVentanaCrdtRepair.cs` (Nuevo)
- **Diagnóstico:**
  - Debido al fallo de V50 y V51 por mismatch de triggers CRDT, los vaults avanzaron hasta la versión 64 omitiendo dichas versiones aditivas tolerantes. Como resultado, la tabla `periodos` carecía físicamente de `periodo_inicio` y `periodo_fin`, impidiendo guardar o visualizar fechas de período y degradando la prioridad de resolución de ventana del motor AG.
- **Solución Implementada:**
  - Implementación de `MigrationV65_PeriodoVentanaCrdtRepair` (Versión 65, Aditiva, `IMigrationTolerante`).
  - Inspecciona `pragma_table_info('periodos')`, suelta triggers obsoletos mediante `PeriodoVentanaCrdtRepair`, añade físicamente `periodo_inicio TEXT` y `periodo_fin TEXT`, y ejecuta backfill canónico parametrizado:
    - `idperiodo = 1`: `2026-01-12` a `2026-06-19`
    - `idperiodo = 2`: `2026-08-11` a `2026-12-12` (incorpora regla V56)
    - `idperiodo = 3`: `2027-01-01` a `2027-06-30`
  - Filtra con `(periodo_inicio IS NULL OR periodo_inicio = '')` para garantizar idempotencia y no sobrescribir fechas previamente configuradas.

#### 3. Fase 746 — Registro de Migración V65 y Helper en Pipeline y Compilación
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:**
  - La nueva migración V65 requería ser registrada en el pipeline de ejecución de `SchemaMigrator` y los dos nuevos archivos debían incluirse explícitamente en el proyecto non-SDK `.csproj`.
- **Solución Implementada:**
  - En `SqliteHelper.BuildMigrator()`, añadido `new MigrationV65_PeriodoVentanaCrdtRepair()` inmediatamente tras `MigrationV64`.
  - En `ConversorPDF.XLL.csproj`, incorporados `<Compile Include="Academic\Infrastructure\Persistence\Migrations\PeriodoVentanaCrdtRepair.cs" />` y `<Compile Include="Academic\Infrastructure\Persistence\Migrations\MigrationV65_PeriodoVentanaCrdtRepair.cs" />`.

#### 4. Fase 747 — Blindaje Retroactivo de `MigrationV50` y `MigrationV51` para Vaults Legados (< 50)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV51_PeriodoVentanaRepair.cs`
- **Diagnóstico:**
  - Vaults antiguos que ejecuten migraciones desde versiones inferiores a 50 volverían a reproducir el error `expected 13 values, got 9` durante el backfill de fechas de V50 o V51 si los triggers CRDT estaban presentes.
- **Solución Implementada:**
  - Invocado `PeriodoVentanaCrdtRepair.DropStaleCrsqlTriggers` antes del DDL y antes del backfill DML en ambas migraciones, blindando la ruta para bases de datos legadas.

#### 5. Fase 748 — Observabilidad y Diagnóstico de Huecos de Versión en `SchemaMigrator`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`
- **Diagnóstico:**
  - Cuando una migración tolerante aditiva fallaba, el log registraba un simple `Warning` y la versión quedaba omitida en silencio. Posteriormente, `VerificarColumnasRequeridas` emitía un error sin indicar la acción correctiva.
- **Solución Implementada:**
  - En el `catch` de `IMigrationTolerante`, el nivel de registro se elevó a `_logger.Error` explicitando `HUECO DE VERSIÓN: v{migration.Version} ({migration.Name}) FALLÓ pero es ADITIVA. Se continúa el arranque sin registrar versión`.
  - En `VerificarColumnasRequeridas`, se añadió la indicación interactiva: `Acción requerida: aplique la migración de reparación correspondiente`.

#### 6. Fase 749 — Guard de Coherencia en `AcademicRepository.InsertPeriodo`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
- **Diagnóstico:**
  - Mientras `UpdatePeriodo` validaba y rechazaba con error descriptivo cuando el esquema no soportaba ventana temporal y venían fechas (`PeriodosTieneVentana() == false`), `InsertPeriodo` omitía silenciosamente las columnas de fecha e insertaba solo `(idperiodo, periodo, activo)`, produciendo inconsistencia de comportamiento.
- **Solución Implementada:**
  - Incorporado en `InsertPeriodo` el mismo guard defensivo: si `!tieneVentana && (!string.IsNullOrWhiteSpace(periodo.PeriodoInicio) || !string.IsNullOrWhiteSpace(periodo.PeriodoFin))`, retorna `SqliteResult<int>.Failure("El esquema no soporta ventana temporal (periodo_inicio/periodo_fin). Aplique las migraciones.")`.

#### 7. Fase 750 — Cobertura de Pruebas Unitarias Automatizadas
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV65Tests.cs` (Nuevo)
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV50Tests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV51Tests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AcademicRepositoryPeriodoTests.cs`
- **Diagnóstico:**
  - Se requería validar que V65 repare el esquema sobre tablas con triggers centinela simulando el error de aridad, que el backfill respete fechas existentes, que V50 y V51 no fallen ante triggers stale, y que `InsertPeriodo` aplique el guard ante esquemas sin ventana.
- **Solución Implementada:**
  - Creado `MigrationV65Tests.cs` con 5 pruebas unitarias (`MigrationV65_Metadata_ShouldBeCorrect`, `Up_WhenLegacySchemaAndStaleCrsqlTriggers_ShouldAddColumnsBackfillAndDropTriggers`, `Up_WhenColumnsExistWithNullDatesAndStaleTriggers_ShouldBackfillWithoutArityError`, `Up_WhenDatesAlreadyPresent_ShouldPreserveThem`, `Up_WhenPeriodosTableDoesNotExist_ShouldReturnGracefully`).
  - Añadida prueba `Up_WhenCrsqlTriggersAreStale_ShouldDropTriggersAndNotFailOnBackfill` en `MigrationV50Tests.cs` y `MigrationV51Tests.cs`.
  - Añadida prueba `InsertPeriodo_WhenEsquemaSinVentanaYConFechas_ShouldReturnFailure` en `AcademicRepositoryPeriodoTests.cs`.

#### 8. Fase 751 — Sincronización de Versión Oficial a 1.6.3.58
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Core/Helpers/App.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `installer/InstaladorConversorPDF.wxs`
- **Diagnóstico:**
  - Se requería incrementar la versión canónica para empaquetado de la solución en la siguiente entrega.
- **Solución Implementada:**
  - Sincronizada la versión `1.6.3.58` en `APP_VERSION`, `<Version>`, `<AssemblyVersion>`, `<FileVersion>` y `Package/@Version`.

---

### 157.99 Ronda 141 — Reparación de Esquema CRDT en Periodos — Revisión de Fases 744-751 (Fases 752 a 756)

#### 1. Fase 752 — Drop Condicionado de Triggers CRDT en V65 según Necesidad Real de Backfill
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV65_PeriodoVentanaCrdtRepair.cs`
- **Diagnóstico:**
  - `DropStaleCrsqlTriggers` se ejecutaba incondicionalmente dos veces en V65, incluso en bases de datos sanas (columnas y fechas completas). Si la posterior regeneración en `EnableAllCrdts` no ocurría (CR-SQLite no disponible o error en `crsql_alter_commit`), `periodos` quedaba sin triggers y en arranques posteriores el short-circuit de `SyncInitializer:295` impedía repararlos.
- **Solución Implementada:**
  - Se condicionó el drop de triggers exclusivamente a `requiereBackfill`. Si las columnas ya existen y todas las fechas están pobladas, se omite el drop de triggers y el backfill DML, preservando intacto el subsistema de sincronización.

#### 2. Fase 753 — Aplicación de Regla V56 y Relleno de Fechas Fin Nulas en Backfill de V65
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV65_PeriodoVentanaCrdtRepair.cs`
- **Diagnóstico:**
  - La condición previa evaluaba únicamente `periodo_inicio IS NULL OR periodo_inicio = ''`. Si una fila tenía inicio pero fin nulo, no se completaba. Asimismo, vaults donde V50 pobló `2026-08-15` pero V56 falló conservaban la fecha desactualizada.
- **Solución Implementada:**
  - Se actualizó el `UPDATE` utilizando `COALESCE(NULLIF(periodo_inicio, ''), @Ini)` y `COALESCE(NULLIF(periodo_fin, ''), @Fin)` para rellenar campos nulos individuales respetando datos existentes.
  - Se replicó la regla explícita de V56 para `2026-02`, actualizando `periodo_inicio = '2026-08-11'` donde fuera `2026-08-15`, NULL o vacío.

#### 3. Fase 754 — Blindaje de Drop Condicionado en V50 y V51
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV51_PeriodoVentanaRepair.cs`
- **Diagnóstico:**
  - Ambas migraciones contenían drops duplicados antes y después de `ALTER TABLE`, soltando triggers en vaults legados que ya contaban con las columnas y registros completos.
- **Solución Implementada:**
  - Se eliminó el drop previo al DDL y se condicionó el drop posterior a la evaluación de `requiereBackfill`, ejecutando el borrado de triggers y los updates solo si existen registros incompletos.

#### 4. Fase 755 — Detección y Restauración de Triggers CRDT Ausentes en `SyncInitializer`
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Sync/SyncInitializer.cs`
- **Diagnóstico:**
  - En arranques sin migraciones pendientes (`schemaMigrated == false`), si una tabla ya marcada como CRDT (`isAlreadyCrdt == true`) había perdido sus triggers por un drop defensivo no regenerado, la condición de short-circuit omitía la reconstrucción y dejaba la tabla permanentemente desincronizada.
- **Solución Implementada:**
  - Creado el helper `TableHasCrdtTriggers(connection, tableName)` que verifica la presencia de los tres triggers (`itrig`, `utrig`, `dtrig`).
  - Antes de evaluar el short-circuit, si `isAlreadyCrdt && !rebuilt && !TableHasCrdtTriggers(...)`, se invoca inmediatamente `SELECT crsql_as_crr('{safeTableName}');` para restaurar los triggers faltantes.

#### 5. Fase 756 — Sincronización de `DefaultVersion` en Instalador WiX
- **Archivo Afectado:** `installer/InstaladorConversorPDF.wxs`
- **Diagnóstico:**
  - `Package/@Version` se actualizó a `1.6.3.58`, pero los atributos `DefaultVersion` y comentarios de `filTemplate1` (`GF-FO-023.xlsx`) y `filTemplate2` (`FV-FO-033.xlsx`) conservaban `1.6.3.57`.
- **Solución Implementada:**
  - Sincronizados los atributos `DefaultVersion="1.6.3.58"` y sus comentarios en ambos componentes del instalador.

---

### 157.100 Ronda 142 — Verificación de Fases 752-756 (Reparación V65 / CRDT) (Fases 757 a 759)

#### 1. Fase 757 — Regla V56 Incondicional en Migration V65
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV65_PeriodoVentanaCrdtRepair.cs`
- **Diagnóstico:**
  - `requiereBackfill` evalúa si faltan columnas o si alguna fila contiene fechas NULL o vacías. Si ambas columnas ya existían físicamente y ninguna fecha estaba vacía, `requiereBackfill` resultaba `false`. Sin embargo, un vault que ejecutó V50 (`periodo_inicio = '2026-08-15'`) pero falló al aplicar V56 retenía el inicio incorrecto (`2026-08-15` en vez de `2026-08-11`). Al haber quedado el `UPDATE` de V56 anidado dentro de `if (requiereBackfill)`, no se ejecutaba y el test `Up_WhenInicioEs20260815_ShouldApplyV56Rule` fallaba.
- **Solución Implementada:**
  - Se desacopló la sentencia `UPDATE` de la regla V56 del bloque condicional `if (requiereBackfill)`, ejecutándola de manera incondicional. Con triggers intactos y esquema completo, actualizar registros existentes es seguro y corrige la fecha a `2026-08-11`.

#### 2. Fase 758 — Purga de Triggers Parciales Remanentes antes de `crsql_as_crr` en `SyncInitializer`
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Sync/SyncInitializer.cs`
- **Diagnóstico:**
  - Si una tabla experimentó una falla o interrupción y conservaba 1 o 2 de los 3 triggers CRDT (`dtrig`, `utrig`, `itrig`), la rama de restauración de la Fase 755 (`!TableHasCrdtTriggers`) invocaba directamente `crsql_as_crr`, fallando con `trigger ... already exists`. La tabla se marcaba fallida y quedaba permanentemente desincronizada.
- **Solución Implementada:**
  - Previo a invocar `SELECT crsql_as_crr('{safeTableName}');`, se añadieron sentencias `DROP TRIGGER IF EXISTS` para los 3 triggers CRDT (`dtrig`, `utrig`, `itrig`) sin alterar la tabla `_clock`, garantizando una inicialización limpia e idempotente.

#### 3. Fase 759 — Uso de `COALESCE` en Backfill de V50 y V51 para Relleno de Fechas Fin Nulas
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV51_PeriodoVentanaRepair.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV50Tests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV51Tests.cs`
- **Diagnóstico:**
  - En V50 y V51, `requiereBackfill` evaluaba si existían registros con `periodo_fin IS NULL`, pero la cláusula `WHERE` del `UPDATE` exigía `periodo_inicio IS NULL OR periodo_inicio = ''`. Si un registro tenía `periodo_inicio` poblado pero `periodo_fin` nulo, se soltaban triggers pero no se actualizaba ninguna fila, manteniendo el valor nulo.
- **Solución Implementada:**
  - Se unificó la sentencia `UPDATE` en ambas migraciones con `COALESCE(NULLIF(periodo_inicio, ''), @Ini)` y `COALESCE(NULLIF(periodo_fin, ''), @Fin)`, actualizando por `idperiodo = @Id` sin filtrar en el `WHERE`.
  - Se agregaron las pruebas unitarias `Up_WhenInicioPresenteYFinNulo_ShouldFillOnlyFin` en `MigrationV50Tests` y `MigrationV51Tests`.

#### 4. Fase 760 — Ajuste de Trazabilidad de Logs en V65
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV65_PeriodoVentanaCrdtRepair.cs`
- **Diagnóstico:**
  - El mensaje en el bloque `else` de `requiereBackfill` indicaba `"sin cambios ni drop de triggers"`. Tras Fase 757, la regla V56 se ejecuta incondicionalmente y puede modificar `periodo_inicio` a `2026-08-11`, haciendo inexacto el log de auditoría.
- **Solución Implementada:**
  - Se ajustó el mensaje del `else` a `"[V65] Columnas y fechas completas: sin drop de triggers."` y se añadió el log explícito `"[V65] Regla V56 verificada para 2026-02."` tras ejecutar la actualización.

---

### 157.101 Ronda 143 — Distribución de Viáticos por Asignatura/Sede y Prevención de Desbordes en Plantilla Excel (Fases 761 a 766)

#### 1. Fase 761 — Creación de `AsignaturaViaticoAllocator` y Modelo `AsignaturaViaticoFila` (Reglas R1 a R7)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/AsignaturaViaticoFila.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Application/Services/AsignaturaViaticoAllocator.cs` (Nuevo)
- **Diagnóstico:**
  - La asignación de bloques a filas en la exportación a Excel dependía de un recorrido plano secuencial sobre `bloquesGlobales.Skip(bloqueIndex).Take(8)`, asignando viáticos a cualquier asignatura disponible sin verificar qué vinculación originó el viaje. Esto provocaba que comisiones de una sede se escribieran en filas de otra sede y que bloques que cruzaban sedes se fragmentaran o desalinearan.
- **Solución Implementada:**
  - Implementado `AsignaturaViaticoAllocator` con el contrato `Distribuir(asignaturas, bloquesGlobales, advertencias)` y constante `MaxBloquesPorFila = 8`.
  - Aplicadas las reglas de negocio deterministas: R1 (asociación exacta por `IdVinculacion`), R2 (bloque multi-sede indivisible en vinculación predominante con desempate por U más temprana), R3 (cota máxima de 8 comisiones por fila), R4 (reasignación por saturación priorizando ventana [U, V], sede en ruta y cercanía), R5 (fila de continuación como último recurso inmediatamente tras la fila base), R6 (advertencia de auditoría si la fecha cae fuera de [U, V]) y R7 (invarianza de suma de bloques).
  - Parseo tolerante de fechas de vinculación multi-formato (`yyyy-MM-dd` y `dd/MM/yyyy`).

#### 2. Fase 762 — Integración en `AsignaturasExcelExporter.cs` y Corrección de Anulación de Tiquetes Aéreos
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - `AsignaturasExcelExporter` agrupaba asignaturas sin incluir la ventana temporal en la clave, impedía distinguir vinculaciones con fechas distintas, aplicaba el ciclo obsoleto `Skip/Take(8)`, omitía escribir U/V y Grupo en filas duplicadas, y en las líneas 477-483 forzaba `RequiereTiquetesAereos = "NO"` y `TerminalAerea = SIN_TERMINAL` en toda fila duplicada, anulando indebidamente el cálculo de terminal aérea (`BQ`) en filas de continuación.
- **Solución Implementada:**
  - Clave de agrupación enriquecida: `{ SedeCiudad, Subsede, AsignaturaNombre, Semestre, VinculacionFechaInicio, VinculacionFechaFin }`.
  - Reemplazo del algoritmo de particionado plano por la invocación a `AsignaturaViaticoAllocator.Distribuir`.
  - Emisión de `U`, `V` y `Grupo` en filas de continuación, preservando la omisión de horas (`L–O`) para evitar duplicar horas en fórmulas de Excel.
  - Eliminación definitiva del bloque de sobreescritura de tiquetes aéreos (líneas 477-483), permitiendo que cada fila liquide sus terminales aéreas según sus propios bloques asignados.
  - Guard defensivo de capacidad contra el límite de 234 filas de la tabla `T_Principal8`.
  - Exposición de `UltimosAvisos`, `UltimosBloquesReasignados` y `UltimasFilasContinuacion` con registro de logs `Warning`.

#### 3. Fase 763 — Actualización de UI en `FormVerVinculaciones.cs`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:**
  - El mensaje tras exportar asignaturas era genérico (`"Exportación finalizada."`), sin informar al usuario si hubo reasignaciones por saturación, generación de filas de continuación o advertencias de auditoría.
- **Solución Implementada:**
  - Mensaje interactivo actualizado a: `"Exportación finalizada. N bloques reasignados / M filas de continuación / K advertencias (ver log)."` con despliegue del primer aviso si existiesen advertencias e icono `MessageBoxIcon.Warning`.

#### 4. Fase 764 — Registro de Nuevos Servicios en `ConversorPDF.XLL.csproj`
- **Archivo Afectado:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:**
  - `AsignaturaViaticoFila.cs` y `AsignaturaViaticoAllocator.cs` requerían ser compilados en el proyecto non-SDK de la solución.
- **Solución Implementada:**
  - Agregadas las etiquetas `<Compile Include>` para ambos archivos bajo `Academic\Application\Services\`.

#### 5. Fase 765 — Suite de Pruebas Unitarias en `AsignaturaViaticoAllocatorTests.cs`
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Application/Services/AsignaturaViaticoAllocatorTests.cs` (Nuevo)
- **Diagnóstico:**
  - Se requería validar de forma automatizada y aislada las reglas R1 a R7, la no fragmentación de bloques multi-sede, la reasignación por saturación y la creación de filas de continuación.
- **Solución Implementada:**
  - Creación de 9 pruebas unitarias cubriendo: distribución canónica 15 bloques en 4 sedes (4/6/2/3), bloques multi-sede indivisibles, desborde R4, continuación R5, invariante <= 8 comisiones, validación R6 de ventana, parseo multi-formato y colecciones vacías.

#### 6. Fase 766 — Documentación Técnica del Cambio
- **Archivo Afectado:** `docs/articles/changes/2026-10-19-distribucion-viaticos-por-asignatura.md` (Nuevo)
- **Diagnóstico:**
  - Requisito de registro de arquitectura y trazabilidad de cambios para el subsistema de exportación de asignaturas.
- **Solución Implementada:**
  - Redacción del documento técnico detallando propósito, reglas R1 a R7, mitigación de riesgos y cobertura de pruebas.

---

### 157.102 Ronda 144 — Revisión de Fases 761-766 (Distribución de Viáticos por Asignatura/Sede) (Fases 767 a 773)

#### 1. Fase 767 — Propagación de Guard de Capacidad de Fila 234 y Prevención de Truncamiento en Cascada
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - El `break` ante `filaExcel > 234` salía únicamente del bucle interno `foreach (filaPlan)` del docente actual. Los docentes subsiguientes entraban con `filaExcel = 235`, generando un error repetitivo por cada fila omitida y saturando `_ultimosAvisos`.
- **Solución Implementada:**
  - Declarada la constante `FILA_MAXIMA_PLANTILLA = 234` y las banderas `capacidadExcedida` y `filasOmitidas`. Al superar la cota se computan las filas omitidas del lote actual, se interrumpe el bucle externo de docentes con `if (capacidadExcedida) break;` y se emite un único aviso consolidado con severidad `Error`.

#### 2. Fase 768 — Distribución en Dos Pases para Priorizar Asignación Natural (R1) sobre Desbordes (R4)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/AsignaturaViaticoAllocator.cs`
- **Diagnóstico:**
  - En un solo pase en orden cronológico, reasignaciones R4 por desborde de fechas tempranas consumían cupos en filas naturales que posteriormente requerían bloques cronológicamente más tardíos de esa misma fila, desplazándolos indebidamente o provocando filas de continuación innecesarias.
- **Solución Implementada:**
  - Reestructurado el algoritmo en dos pases: el Pase 1 reserva los cupos de las asignaciones unívocas naturales (R1) para cada fila con cupo; el Pase 2 resuelve los bloques pendientes aplicando las reglas R2 (multi-sede), R4 (reasignación por saturación) y R5 (continuación como último recurso).

#### 3. Fase 769 — Parseo Robusto de Fechas con Hora y Cultura `es-CO`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/AsignaturaViaticoAllocator.cs`
- **Diagnóstico:**
  - Cadenas con hora (ej. `dd/MM/yyyy HH:mm:ss`) no coincidían con `TryParseExact` y caían al fallback invariante `DateTime.TryParse`, el cual interpretaba `12/09/2026` como formato estadounidense `MM/dd/yyyy` (9 de diciembre en vez de 12 de septiembre), alterando silenciosamente la ventana temporal [U, V].
- **Solución Implementada:**
  - Incorporados formatos con hora en `FormatosFecha` y uso explícito de `CultureInfo.GetCultureInfo("es-CO")` ante separadores `/` en el fallback genérico.

#### 4. Fase 770 — Distinción de Advertencia para Bloques sin Vinculación Coincidente en R4
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/AsignaturaViaticoAllocator.cs`
- **Diagnóstico:**
  - Si un bloque no tenía vinculación coincidente (`candidatas.Count == 0`), se emitía `"reasignado por saturación"`, provocando que el contador `UltimosBloquesReasignados` registrara falsos positivos de saturación.
- **Solución Implementada:**
  - Diferenciado el mensaje a `"sin vinculación coincidente"` cuando `candidatas.Count == 0`, reservando `"reasignado por saturación"` exclusivamente para filas naturales llenas.

#### 5. Fase 771 — Replicación de Prestaciones y Puntaje Ocasionales en Filas de Continuación (LLAVE Y)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - Las filas de continuación omitían escribir las columnas `Prestaciones` (X) y `PuntajeOcasionales` (Z), dejando incompleta la fórmula institucional `Y = CONCAT(VINCULACIÓN, CATEGORÍA, PRESTACIONES)`.
- **Solución Implementada:**
  - Incorporada la escritura de `Prestaciones` y `PuntajeOcasionales` en la rama `if (esFilaDuplicada)`.

#### 6. Fase 772 — Priorización de Avisos Críticos en la Interfaz de Usuario
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Diagnóstico:**
  - El mensaje final desplegaba `UltimosAvisos[0]`, permitiendo que advertencias menores (como R6 fuera de ventana) ocultaran alertas críticas de capacidad, saturación o continuación.
- **Solución Implementada:**
  - Búsqueda preferente de avisos que contengan `"continuación"`, `"saturación"` o `"Capacidad"` antes de recurrir al primer aviso ordinario.

#### 7. Fase 773 — Cobertura Automatizada del Caso Real Multi-Sede y Fallback Huérfano
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Application/Services/AsignaturaViaticoAllocatorTests.cs`
- **Diagnóstico:**
  - La suite de pruebas no contemplaba sedes repetidas con ventanas solapadas en formato latino (Cali Fundamentos, Cali Historia, Cartago, El Águila) ni bloques con `IdVinculacion` no coincidente.
- **Solución Implementada:**
  - Incorporadas las pruebas `CasoRealDosCalisCartagoElAguila_15Bloques_RespetaVentanaYFila` (validación exacta 4/6/2/3 sin avisos) y `BloqueSinVinculacionCoincidente_AdvierteSinContarSaturacion`.

---

### 157.103 Ronda 145 — Revisión de Fases 767-773 (Distribución de Viáticos por Asignatura/Sede) (Fases 774 a 776)

#### 1. Fase 774 — Corrección de Conflicto de Ámbito de Variable CS0136 en Bucles Anidados de Filas
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - El nuevo bucle de exportación `for (int i = 0; i < filasPlan.Count; i++)` (F767) provocaba el error de compilación CS0136 debido a la colisión de ámbito con tres bucles internos `for (int i ...)` en las secciones 4 (comisiones) y 5 (ruta).
- **Solución Implementada:**
  - Renombrado el identificador del bucle externo a `for (int indiceFila = 0; indiceFila < filasPlan.Count; indiceFila++)` y acceso indexado `filasPlan[indiceFila]`, manteniendo intacto `i` en los bucles internos.

#### 2. Fase 775 — Métricas de Cierre Inconsistentes al Activar el Guard de Capacidad y Comentario Obsoleto
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - El contador `UltimasFilasContinuacion` se incrementaba anticipadamente antes de escribir en la hoja, contabilizando filas no escritas si el guard de capacidad truncaba el proceso. Asimismo, el aviso no informaba el docente donde ocurrió el corte, y existía un comentario obsoleto indicando que las fechas de vinculación no se escribían en filas duplicadas.
- **Solución Implementada:**
  - Se eliminó el conteo anticipado de continuaciones y se trasladó dentro del bucle tras el guard (`if (esFilaDuplicada) UltimasFilasContinuacion++;`). Se registró `docenteCorte = grupoDocente.Key` y se enriqueció el aviso consolidado: `$"Capacidad de T_Principal8 excedida (fila {filaMaximaPlantilla}). Filas omitidas desde el docente '{docenteCorte}': {filasOmitidas}."`. Se actualizó el comentario para aclarar que solo las horas L-O se omiten, mientras que U, V, Grupo, Prestaciones y Puntaje Ocasionales sí se replican.

#### 3. Fase 776 — Detección Dinámica del Límite de Capacidad Institucional `T_Principal8`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - El límite fijo de fila 234 (`FILA_MAXIMA_PLANTILLA`) impedía que el exportador aprovechara ampliaciones de la tabla institucional `T_Principal8` en la plantilla de Excel, truncando la exportación aunque existieran fórmulas configuradas.
- **Solución Implementada:**
  - Reemplazada la constante fija por `FILA_MAXIMA_PLANTILLA_DEFAULT = 234` e implementado el método `ObtenerFilaMaximaPlantilla(dynamic ws)` que calcula dinámicamente el límite a partir de `ws.ListObjects["T_Principal8"]` con fallback seguro al valor por defecto ante cualquier excepción.

---

### 157.104 Ronda 146 — Revisión de Fases 774-776 (Guard de Capacidad y Métricas de Exportación) (Fases 777 a 779)

#### 1. Fase 777 — Prevención de Fuga de Recursos COM en `ObtenerFilaMaximaPlantilla`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - Las referencias COM `ws.ListObjects["T_Principal8"]` y `tabla.Range` se obtenían sin `Marshal.ReleaseComObject`, dejando RCWs vivos que podían retener el proceso Excel en segundo plano.
- **Solución Implementada:**
  - Envolvimiento de la resolución de rango en bloque `try/finally`, liberando explícitamente `rango` y `tabla` mediante `Marshal.ReleaseComObject`.

#### 2. Fase 778 — Sincronización de Avisos y Reasignaciones Condicionada a `planCompleto`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - Los avisos y el conteo de bloques reasignados del docente se agregaban a las métricas globales antes de recorrer y escribir las filas en Excel, reportando avisos y reasignaciones correspondientes a filas omitidas cuando se activaba el guard de capacidad.
- **Solución Implementada:**
  - Introducida la bandera `planCompleto = true;` en el bucle de filas (asignando `planCompleto = false;` al activarse el guard) y trasladada la agregación a `_ultimosAvisos` y `UltimosBloquesReasignados` tras la escritura, registrando en el log institucional pero condicionando la inclusión en métricas de UI al volcado efectivo de las filas.

#### 3. Fase 779 — Sincronización de Documentación XML y Bitácora Histórica
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`, `walkthrough.md`
- **Diagnóstico:**
  - El XML doc del método `Exportar` mantenía una descripción obsoleta indicando que se duplicaba la última fila sin datos redundantes ante exceso de viáticos. Asimismo, en `walkthrough.md` L2054 figuraba `filasPlan.Count - i`, desincronizado del renombre a `indiceFila` de F774.
- **Solución Implementada:**
  - Actualizado el XML doc de `Exportar` reflejando las reglas R1–R7 y la condición de último recurso de las filas de continuación, y corregida la referencia en la bitácora `walkthrough.md` a `filasPlan.Count - indiceFila`.

---

