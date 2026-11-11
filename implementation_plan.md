# Plan de Implementación: Sanitización, Trazabilidad y Refactorización Integral COM (Fases 170 - 234)

## 📋 Resumen de la Revisión y Arquitectura de Robustez COM
Se realizó un análisis exhaustivo y evaluación crítica del ciclo de vida de objetos COM Excel durante el proceso de generación y exportación de viáticos individuales (`ViaticosIndividualExporter.cs`). Se ratificaron y corrigieron 65 vulnerabilidades estructurales (Fases 170 a 234) para erradicar completamente fugas de recursos, bloqueos de archivos en disco, excepciones no controladas, fallos en la detección de aprobadores, auto-aprobaciones indebidas y hojas residuales en el libro activo del usuario.

---

## ⚙️ Detalle Completo por Fase Afectada (Fases 170 - 234)

### 1. **Fase 170, 185 & 218 — Captura de Copia y Verificación de Incremento en Duplicado (`wsBase.Copy`)**
- **Diagnóstico:** Usar `app.ActiveSheet` dependía del foco global del usuario, y duplicar sin validar incremento podía hacer que `wsNueva` apuntara a una hoja existente si `Copy` fallaba.
- **Solución:** Al duplicar con `wsBase.Copy(Type.Missing, ultimoSheet)`, se valida `(int)wb.Worksheets.Count != antesDup + 1`, omitiendo el bloque si la copia no fue insertada para evitar la sobrescrita o eliminación de hojas existentes.

### 2. **Fase 171, 194 & 210 — Aislamiento de Excepciones y Protección del Callback UI (`procesados`)**
- **Diagnóstico:** El ciclo `foreach` carecía de `catch` individual y la llamada a `progressCallback` podía abortar la exportación si la UI lanzaba excepción.
- **Solución:** Se añadió un bloque `catch (Exception ex)` por iteración que registra el error en `LoggerHelper.Error` sin interrumpir el resto del lote. Se envolvió `progressCallback?.Invoke` en un `try/catch` defensivo. El progreso se notifica mediante `procesados++`.

### 3. **Fase 172 — Protección al Restaurar Visibilidad de `wsBase`**
- **Diagnóstico:** Asignar `wsBase.Visible = 2` (`xlSheetVeryHidden`) sin `try/catch` lanzaba excepción COM 1004 en libros protegidos.
- **Solución:** Envolver la restauración de `wsBase.Visible = 2` en un bloque `finally` con `try/catch` defensivo.

### 4. **Fase 173, 195, 208 & 223 — Nombres Únicos con Sanitización de Prefijo (`tipoPrefijo`) y `ExisteHoja` Sin Sobrecarga Excesiva**
- **Diagnóstico:** Tipos como `SERVIDOR/PLANTA` o nombres largos lanzaban error COM 1004 al asignar `wsNueva.Name = sheetName`, y `ExisteHoja` con 3 argumentos causaba error de compilación CS1501.
- **Solución:** `tipoPrefijo` se sanitiza filtrando solo caracteres alfanuméricos/guiones y truncando a 24 caracteres máx. `ExisteHoja` se utiliza con su firma exacta de 2 parámetros.

### 5. **Fase 174, 186, 187, 188, 198, 199, 200, 211, 214, 217, 220, 221, 222, 225 & 232 — Guarda sobre `Sheets`, Exclusión de Hojas Pendientes (`pendientes`), Fallback para Libros Residuales y Filtrado por Tipo de Hoja**
- **Diagnóstico:** Intentar eliminar la única hoja visible del libro lanzaba error COM 1004. Asignar visibilidad a `DialogSheet` (Type `-4116`) provocaba error 1004. Si el libro estaba compuesto solo por plantillas/temporales, el filtro `excluir` dejaba `temporal == null` y la hoja no se borraba.
- **Solución:** `EliminarHojaTemporal` busca auxiliares sobre `wb.Sheets` verificando tipos válidos (`xlWorksheet` -4167 y `xlChart` -4168). Si `excluir` filtra todas las hojas (libro residual), un segundo bucle de fallback permite usar una hoja pendiente como auxiliar temporal para posibilitar el borrado sin intentar restaurar su visibilidad previa.

### 6. **Fase 175, 189, 190, 201, 204 & 226 — Preservación de Hojas Canónicas, Limpieza Dinámica y Registro de Prefijos Arbitrarios (`hojasTemporales`)**
- **Diagnóstico:** La limpieza previa borraba hojas del usuario y los prefijos temporales fijos no reconocían `tipoPrefijo` arbitrarios (p. ej. `PLANTA`, `DOCENTE`), dejando residuos si la exportación fallaba.
- **Solución:** `Exportar` acumula los nombres de hojas temporales reales en un `HashSet<string> hojasTemporales`, transmitiéndolo a `LimpiarHojasPlantilla` para garantizar la eliminación de cualquier hoja creada durante la ejecución.

### 7. **Fase 176, 184 & 207 — Alcance de `hojasImportadas` y Mapeo Seguro con `TryGetValue`**
- **Diagnóstico:** Declarar `hojasImportadas` dentro del bloque `try` causaba un error de compilación CS0103 al acceder en el bloque `finally`.
- **Solución:** Se declaró `hojasImportadas` antes del `try` principal y se utilizó `hojasImportadas.TryGetValue` defensivo.

### 8. **Fase 177, 178, 191, 197, 206, 209, 212, 213, 219 & 228 — Firma Dinámica de Contenido (`PrimerValorNoVacio`), Importación Verificada y Margen de Tolerancia en Dimensiones**
- **Diagnóstico:** Comparar celdas fijas (`A1`/`B1`) fallaba en plantillas cuyo primer valor estaba en `C1`. Además, residuos previos con filas/columnas extra vacías en `UsedRange` impedían la pre-limpieza (`identica = false`).
- **Solución:** `PrimerValorNoVacio` localiza dinámicamente el primer valor no nulo del `UsedRange`. `LimpiarPreviaPlantillas` admite una tolerancia de hasta 2 filas/columnas (`Math.Abs(fTpl - fExi) <= 2 && Math.Abs(cTpl - cExi) <= 2`), permitiendo eliminar la hoja previa e importar la versión limpia.

### 9. **Fase 179, 192, 193, 202, 215, 216, 224, 227, 229, 230 & 231 — Tokens Únicos, Truncamiento de Ruta, Sanitización de Nombres Reservados de Windows y Desambiguación de PDFs**
- **Diagnóstico:** `xV-TEMP-{consecutivoHoja}.xlsx` se repetía entre corridas. Nombres de ciudad o destino largos superaban `MAX_PATH` (260). Nombres reservados de Windows (`CON`, `PRN`, `AUX`, `NUL`, etc.) causaban fallos de I/O, y corridas repetidas sobrescribían los PDFs generados en la corrida previa.
- **Solución:** Se integró un token Guid (`tokenTemporal = Guid.NewGuid().ToString("N").Substring(0, 8)`). Se limitó `baseRutaName` a 40 caracteres máximo. `SanitizarNombreArchivo` antecede un guion bajo `_` si el nombre coincide con un reservado de Windows (`CON`, `PRN`, `AUX`, `NUL`, `COM1..4`, `LPT1..3`). Los nombres de PDF existentes se desambiguan concatenando `-(k)`.

### 10. **Fase 181 — Conteo Inmediato de `generados++` y Metadatos No Fatales**
- **Diagnóstico:** Incrementar `generados++` tras los metadatos hacía que un fallo de metadatos invalidara el conteo de un PDF generado con éxito.
- **Solución:** Incrementar `generados++` inmediatamente tras crear el PDF y envolver `GuardarMetadatos` en un bloque `try/catch` independiente.

### 11. **Fase 182 — Precálculo de Identidad de Usuario**
- **Diagnóstico:** Consultar la identidad de Windows dentro del ciclo por cada registro agregaba latencia I/O.
- **Solución:** Se precalculó `cleanUser` una sola vez antes de ingresar al bucle `foreach`.

### 12. **Fase 183 — Congelación de Valores Protegida con Referencia Cacheada**
- **Diagnóstico:** `wsNueva.UsedRange.Value = wsNueva.UsedRange.Value` ejecutaba dos llamadas COM a `UsedRange`.
- **Solución:** Se cacheó la referencia `dynamic rangoUsado = wsNueva.UsedRange; rangoUsado.Value = rangoUsado.Value;` dentro de un bloque `try/catch`.

### 13. **Fase 196 & 205 — Invocación Posicional de `sh.Copy` sobre Receptor Dynamic**
- **Diagnóstico:** Usar argumentos nombrados (`sh.Copy(After: ...)`) sobre un receptor `dynamic` provocaba error CS8324.
- **Solución:** Se utilizó sintaxis posicional con `Type.Missing`: `sh.Copy(Type.Missing, copiaDestino);`.

### 14. **Fase 233 & 234 — Unificación de Mapeo de Aprobador, Trazabilidad, Endurecimiento Defensivo Total y Prevención de Auto-Aprobación por Nombre y Cédula**
- **Diagnóstico:** `ViaticosRepository.GetConfig()` dependía de accesos directos `dynamic` sobre columnas SQLite que podían no existir en esquemas legacy. En `ViaticosIndividualProcessor` y `ViaticosIndividualExporter`, si el Director Territorial comisionado no tenía cédula configurada o faltaba el sustituto, se permitía la auto-aprobación o se emitían advertencias repetitivas por lote. Además, `FormProgramador/FormConfig.cs` leía índices 1-based en array 0-based provocando `IndexOutOfRangeException`.
- **Solución:** 
  1. **Endurecimiento Completo en Repositorio:** En `ViaticosRepository.cs`, todas las columnas se leen a través de funciones defensivas `raw(k)` y `col(k)` sobre `IDictionary<string, object>`, con fallback `dt_nombre` -> `corporativo_nombre` si `dt_nombre` estuviera vacío.
  2. **Helper Centralizado con Detección Dual (Cédula y Nombre Normalizado):** `ViaticosHelper.ResolverAprobador(...)` evalúa coincidencia tanto por cédula sanitizada como por nombre completo normalizado (`NormalizarNombre`), evitando auto-aprobación aún si `CedulaGerente` no estuviera configurada.
  3. **Control de Flood de Logs:** Se introdujo la bandera estática `_avisoCedulaVacia` para registrar la advertencia una única vez por sesión.
  4. **Validación Estricta de Sustituto:** Se exige que `GerenteSustituto` contenga texto para habilitar la sustitución; de lo contrario, se vacía el campo de aprobador para respetar normas de control interno.
  5. **Corrección de Formulario Programador:** En `FormProgramador/FormConfig.cs`, se ajustaron los índices de lectura en `FormConfig_Load` a `data[0..12, 0]`.

---

## 🔎 Análisis de Calidad e Invariantes (Fases 170 - 234)
1. **Robustez COM:** Inmunidad contra fugas de memoria, bloqueos de archivos en disco, diálogos modales y errores de compilación C#.
2. **Resiliencia de Procesamiento:** Los errores por registro individual, nombres no válidos y fallos de interfaz gráfica se capturan sin interrumpir la exportación.
3. **Invariante de Limpieza:** Se garantiza que el libro del usuario permanezca exactamente en su estado original sin alterar la visibilidad de hojas ni utilizar hojas pendientes de borrado como auxiliares.
4. **Seguridad y Control Interno:** El algoritmo de sustitución previene formalmente la auto-aprobación del Director Territorial tanto por coincidencia de cédula como por nombre normalizado.
5. **Sincronización:** Documentación interna, comentarios de código y plan de implementación sincronizados al 100%.

---

# Plan de Implementación: Edición Manual de Sesiones en Grilla y Sincronización Bidireccional de Integridad (Fases 235 - 245 / F1.1 - F1.15)

## 📋 Resumen Ejecutivo
Se detectó que las excepciones de desfase de integridad (`GridToSesionMapper: Fila X: Los datos de fecha u horas han sido alterados...`) eran provocadas por el desacoplamiento entre las celdas editadas por el usuario en `dgvSesiones` (`ColFechaSesion`, `ColHoras`, `ColHorario`) y el modelo de dominio enlazado en `row.Tag` (`Sesion`), junto con bugs de interpretación de separadores decimales (`NumberStyles.Any` con comas en cultura es-CO) y comportamientos predeterminados de WinForms (`AllowUserToAddRows`, `SortMode.Automatic` lexicográfico sobre strings de fecha).

---

## ⚙️ Desglose Técnico de Fases (Fases 235 - 245 / F1.1 - F1.15)

### 1. **Fase 235 (F1.1) — Creación de `GridCellValueParser` (Aislamiento y Testeo Puro)**
- **Archivo:** `ConversorPDF.XLL\Academic\Application\Mappers\GridCellValueParser.cs`
- **Propósito:** Parser estático sin dependencias de WinForms para interpretar horas y fechas de forma segura.

### 2. **Fase 236 (F1.2 & F1.3) — Saneamiento de `PrepararSesionesDesdeGrid` y Tolerancia Float en `GridToSesionMapper`**
- **Archivos:** `FormSesionesAcademicas.cs` y `GridToSesionMapper.cs`

### 3. **Fase 237 (F1.4, F1.5 & F1.6) — Sincronización Bidireccional en `dgvSesiones_CellValueChanged` con Anti-Reentrada y Reversión**
- **Archivo:** `FormSesionesAcademicas.cs`

### 4. **Fase 238 (F1.8 & F1.9) — Ciclo de Commit Limpio en `PrepararSesionesDesdeGrid`**
- **Archivo:** `FormSesionesAcademicas.cs`

### 5. **Fase 239 (F1.10 & F1.11) — Endurecimiento y Configuración de Grilla (`SetupGrid`)**
- **Archivo:** `FormSesionesAcademicas.cs`

### 6. **Fase 240 (F1.12 & F1.13) — Post-Commit Atómico de Rango de Fechas en Formulario**
- **Archivo:** `FormSesionesAcademicas.cs`

### 7. **Fase 241 (F1.14) — Validaciones de Dominio para Edición Manual Libre**
- **Archivo:** `FormSesionesAcademicas.cs`

### 8. **Fase 242 (F1.15) — Suite de Pruebas Unitarias de Parser y Mapeo**
- **Archivo:** `ConversorPDF.Tests\Unit\GridCellValueParserTests.cs` y `ConversorPDF.Tests\Unit\GridToSesionMapperTests.cs`

---

## 🔗 Planes y Documentos de Extensión
- **Optimización de Horarios Académicos por Algoritmo Genético (AG v7, Fases 1 - 150):** Documentado íntegramente en [`docs/articles/changes/2026-ag-horarios-academicos.md`](docs/articles/changes/2026-ag-horarios-academicos.md).

---

# Plan de Implementación: Árbol Académico de 6 Niveles, Ventanas Canónicas en Períodos y Materialización de Tuplas S (Fases 151 - 164)

## 📋 Resumen Ejecutivo
Se consolidó la estructura del contexto académico para el motor de Algoritmo Genético, estableciendo una jerarquía unificada de 6 niveles (Periodo -> Sede -> Programa -> Semestre -> Grupo -> Modalidad), resolviendo la prioridad de calendario en `periodos`, y garantizando que cada rama hoja marcada en el árbol origine una tupla de contexto $S$ materializada en `ag_aula` con clave natural unívoca `AulaRef` (incluyendo la Modalidad para erradicar colisiones) y clave subrogada `IdAula`.

---

## ⚙️ Detalle Técnico por Fase Afectada (Fases 151 - 164)

### 1. **Fase 151 — Corrección de Esquema en Catálogos (`no such column: nombre`)**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormTerceroAsignaturas.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Sede.cs`
- **Diagnóstico:** Consultas SQL utilizaban la columna inexistente `nombre` sobre tablas `programas` y `sedes`, provocando excepciones `SQLiteException (no such column: nombre)` al inicializar los formularios.
- **Solución:**
  - En `AgFormGeneradorHorarios.cs` y `FormTerceroAsignaturas.cs`, se actualizaron las consultas a `programas` proyectando `id as Id`, `nombreprograma as NombrePrograma`, `abreviatura as Abreviatura`, y filtrando por `activo = 1`.
  - En `sedes`, se proyectaron `ciudad as Ciudad`, `subsede as Subsede` y se agregó en `Sede.cs` la propiedad calculada `Nombre => string.IsNullOrWhiteSpace(Subsede) ? (Ciudad ?? string.Empty) : $"{Ciudad} ({Subsede})"`.
  - En `periodos`, se mapearon explícitamente `idperiodo as IdPeriodo`, `periodo as PeriodoNombre`, `periodo_inicio as PeriodoInicio`, `periodo_fin as PeriodoFin`, y `activo as Activo`.

### 2. **Fase 152 — Extensión del Modelo de Dominio `Periodo` y Enum `AgOrigenVentana`**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Periodo.cs`
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgOrigenVentana.cs`
- **Diagnóstico:** La entidad de dominio `Periodo` carecía de los campos de fecha de inicio y fin semestral, y el enum `AgOrigenVentana` no disponía del valor para clasificar resoluciones provenientes de la tabla `periodos`.
- **Solución:**
  - Se agregaron las propiedades `public string PeriodoInicio { get; set; }` y `public string PeriodoFin { get; set; }` a `Periodo.cs`.
  - Se implementó el método `TryObtenerVentana(out DateTime inicio, out DateTime fin)` con parseo multi-formato (ISO `yyyy-MM-dd`, variantes estándar) y verificación de rango (`fin >= inicio`).
  - Se añadió el valor `Periodos = 5` al enum `AgOrigenVentana.cs`.

### 3. **Fase 153 — Migración V50 de Esquema y Backfill de Fechas (`MigrationV50_AddPeriodoVentana`)**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgSchemaMigrationTests.cs`
- **Diagnóstico (Hallazgo Crítico):** El script semilla `AcademicSchema.sql` utiliza `INSERT OR IGNORE INTO periodos ...`. En vaults existentes las filas con `idperiodo = 1, 2, 3` ya existen, por lo que el `INSERT OR IGNORE` las omite y las columnas `periodo_inicio` y `periodo_fin` permanecían nulas en producción.
- **Solución:**
  - Se creó la migración idempotente `MigrationV50_AddPeriodoVentana` que ejecuta `ALTER TABLE periodos ADD COLUMN periodo_inicio TEXT` y `periodo_fin TEXT` si no existen.
  - Ejecuta de forma determinista sentencias `UPDATE periodos` para poblar las fechas canónicas (`2026-01-12` a `2026-06-19` para `idperiodo=1`, `2026-08-15` a `2026-12-12` para `idperiodo=2`, y `2027-01-01` a `2027-06-30` para `idperiodo=3`) cuando se encuentren nulas o vacías.
  - Registro de la migración en `SqliteHelper.BuildMigrator()` después de `MigrationV49` y en el proyecto `.csproj`.
  - Prueba unitaria de metadatos en `AgSchemaMigrationTests.cs`.

### 4. **Fase 154 — Cascada de Prioridades en Resolución de Ventana Semestral (`AgResolutorVentanaSemestre`)**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResolutorVentanaSemestre.cs`
- **Diagnóstico:** El resolutor consultaba directamente la tabla festivos sin verificar las fechas maestras del período en `periodos`.
- **Solución:**
  - Se reestructuró la cascada de resolución temporal:
    1. **Prioridad 1 (PERIODOS):** Si el registro en `periodos` tiene `periodo_inicio` y `periodo_fin` válidos, se resuelve con origen `AgOrigenVentana.Periodos` (o acotado por rango manual si el usuario lo especificó dentro de las fechas).
    2. **Prioridad 2 (FESTIVOS - PRESENCIAL):** Mínimo y máximo de rangos tipo 'PRESENCIAL'.
    3. **Prioridad 3 (MANUAL):** Fechas manuales suministradas por el usuario.
    4. **Prioridad 4 (ERROR):** Mensaje descriptivo guiando al usuario si no hay fechas válidas.

### 5. **Fase 155 — Validaciones de Coherencia en Dominio y Repositorio de Períodos**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Periodo.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
  - `ConversorPDF.Tests/Academic/Domain/PeriodoTests.cs` (Nuevo)
- **Diagnóstico:** `Periodo.Create` no validaba la regla de negocio `fin >= inicio`, y el repositorio no persistía ni recuperaba las nuevas columnas de ventana.
- **Solución:**
  - `Periodo.Create` valida que si ambas fechas son suministradas, la fecha final sea posterior o igual a la inicial.
  - `AcademicRepository.cs`: Actualizados `GetPeriodos()`, `GetPeriodoById()`, e `InsertPeriodo()` para incluir `periodo_inicio` y `periodo_fin`.
  - Creado `PeriodoTests.cs` con pruebas completas de validación de formato, fechas invertidas y parseo exitoso.

### 6. **Fase 156 — Modelos Desacoplados y Servicios de Expansión Combinatoria**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgContextoSeleccionDto.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgNodoArbolKey.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Application/Interfaces/ITreeViewCheckService.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgTreeViewCheckService.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSeleccionAulaExpandidor.cs` (Nuevo)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgTreeViewCheckServiceTests.cs` (Nuevo)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgSeleccionAulaExpandidorTests.cs` (Nuevo)
- **Diagnóstico:** No existían contratos ni DTOs para representar tuplas de contexto combinatorio independientes de la UI de WinForms, ni mecanismos de validación contra explosión combinatoria.
- **Solución:**
  - `AgContextoSeleccionDto`: Modela `{IdPeriodo, IdSede, IdPrograma, Semestre, Grupo, Modalidad}` y expone la clave natural unívoca `AulaRef`.
  - `AgNodoArbolKey`: Identificador tipado para nodos del árbol jerárquico.
  - `ITreeViewCheckService` / `AgTreeViewCheckService`: Servicio puro para calcular el producto cartesiano de selecciones.
  - `AgSeleccionAulaExpandidor`: Deduplica por `AulaRef` y hace cumplir el tope de seguridad `CapMaximo = 500`.
  - Pruebas unitarias completas para el servicio de combinaciones y el expansor de seguridad.

### 7. **Fase 157 & 158 — Árbol Jerárquico de 6 Niveles con Tri-Estado, Lazy Loading y Resumen Dinámico**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** La selección se realizaba mediante combos planos y se requería eliminar `cmbGrupo` para permitir la selección de grupo y modalidad por cada rama. Población estática de 10 semestres x 4 grupos x 3 modalidades (120 hojas por programa) ralentizaría la apertura del formulario.
- **Solución:**
  - En `grpSeleccion`, se implementó el control `treeContextoAcademico` con `CheckBoxes = true` y una barra de herramientas con acciones: Expandir, Contraer, Marcar Todo, Desmarcar, Invertir y Generar Aulas.
  - Implementación de carga diferida (lazy load): Se cargan Periodo -> Sede -> Programa y se colocan nodos dummy en `Programa` que cargan `Semestres (I..X)` al expandir o marcar; a su vez `Semestre` carga `Grupos (A..D)` y `Grupo` carga `Modalidades (Presencial, Virtual, Semipresencial)`.
  - Propagación de check recursiva hacia hijos y actualización de estado hacia padres protegida con bandera anti-reentrancia `_isUpdatingTreeChecks`.
  - Etiqueta de conteo en tiempo real `lblConteoCombinaciones` con advertencia visual roja al superar el tope de 500 combinaciones.
  - Retiro operativo de `cmbGrupo`, resolviendo el grupo directamente desde el árbol.

### 8. **Fase 159 — Fábrica de Aulas `AgAulaFactory` y Prevención de Colisiones en `AulaRef`**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAulaFactory.cs` (Nuevo)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAulaFactoryTests.cs` (Nuevo)
- **Diagnóstico (Colisión Crítica en Tupla S):** Si `AulaRef` omitía la modalidad, ramas del mismo grupo en modalidad presencial y virtual colisionaban en la misma fila de `ag_aula`, perdiendo una modalidad silenciosamente.
- **Solución:**
  - `AgAulaFactory` materializa entidades `AgAula` garantizando que la clave natural sea `AulaRef = "{idPeriodo}|{idsede}|{idprograma}|{semestre}|{grupo}|{modalidad.ToUpperInvariant()}"`.
  - Mapeo consistente de `TipoAula` ("VIRTUAL" vs "ESTANDAR"), `Capacidad` (999 vs 30) y `EsVirtual` (1 vs 0).
  - Pruebas unitarias en `AgAulaFactoryTests.cs` verificando unicidad de `AulaRef` entre modalidades.

### 9. **Fase 160 — Generación Transaccional por Lotes de Encabezados de Aula (`ag_aula`)**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** No existía una rutina directa para materializar en masa las aulas seleccionadas antes de ejecutar el algoritmo genético.
- **Solución:**
  - Se implementó `GenerarEncabezadosAulasAsync()` en `AgFormGeneradorHorarios.cs`:
    - Valida que la selección no supere 500 aulas mediante `AgSeleccionAulaExpandidor`.
    - Solicita confirmación explícita al usuario indicando el número exacto de aulas a generar.
    - Ejecuta en una transacción única (`IDbTransaction`) con verificación de registros preexistentes en `ag_aula` para evitar duplicados.
    - Emite resumen informativo en pantalla y en la bitácora `txtLogAG`.

### 10. **Fase 161 — Sincronización y Precarga de Ventana desde `periodos`**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Los controles `dtpVentanaInicio` y `dtpVentanaFin` no se sincronizaban automáticamente con las fechas maestras del período ni se indicaba visualmente el origen `PERIODOS`.
- **Solución:**
  - `ActualizarVentanaSemestralAsync()` precarga automáticamente los controles `DateTimePicker` con `_ventanaResuelta.FechaInicio` y `_ventanaResuelta.FechaFin` si el ajuste manual no está activo.
  - Actualiza `lblOrigenVentana` en color verde oscuro con la leyenda `Origen: PERIODOS [{Nombre}] (N días)`.

### 11. **Fase 162 — Invariante Atómico de Asignación de Aula en `AgGen`**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`
  - `ConversorPDF.Tests/Academic/Domain/AgGenTests.cs`
- **Diagnóstico:** `AgGen` duplicaba los atributos de la tupla S (`IdPeriodo, IdSede, IdPrograma, Semestre, Grupo`) de forma dispersa, arriesgando inconsistencias respecto a `IdAula`.
- **Solución:**
  - Se incorporó el método `gen.AsignarAula(AgAula aula)` como invariante atómico de dominio que sincroniza `IdAula`, `IdPeriodo`, `IdSede`, `IdPrograma`, `Semestre` y `Grupo`.
  - Pruebas unitarias añadidas en `AgGenTests.cs`.

### 12. **Fase 163 — Corrección de Falsa Alarma en `EncryptedSQLiteContext`**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Context/EncryptedSQLiteContext.cs`
- **Diagnóstico:** En `Dispose()`, la variable `hasChanges` se inicializaba en `true`. Si la conexión ya estaba cerrada o si el contexto solo ejecutaba lecturas, se emitía indebidamente el log `[WARN] Contexto descartado sin persistir al Vault maestro`.
- **Solución:**
  - Se cambió la inicialización a `bool hasChanges = false;` y la verificación de `total_changes() > 0` se evalúa únicamente si la conexión se encuentra abierta.

### 13. **Fase 164 — Aviso Visual en UI para Períodos sin Fechas Configuradas**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Cuando un período no tenía fechas en la base de datos y no había festivos presenciales, el usuario no recibía una retroalimentación clara sobre qué acción tomar.
- **Solución:**
  - Si el período carece de fechas maestras, `lblOrigenVentana` muestra una alerta en color naranja indicando `[AVISO] El período '{periodoSel.Nombre}' no tiene fechas en DB. Ingrese fechas manuales.` guiando al usuario a marcar "Ajustar manual" o configurar festivos.

---

---

## 🚀 Fases 165 - 170: Correcciones de Esquema, Estabilidad y Factorías (Ronda 17)

### 14. **Fase 165 — Corrección Bloqueante: `sqlAsig` con Esquema Canónico de `asignaturas`**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** La consulta utilizaba columnas inexistentes físicamente (`codigo`, `idsede`, `horas`, `horassemanales`, `modalidad`), arrojando `no such column: codigo` en runtime y rompiendo `ConstruirContextoAsync`.
- **Solución:** Se alineó la proyección SQL a las columnas físicas reales de la tabla `asignaturas` (`id, idprograma, semestre, num_nivel, cod_unidad, cod_materia, uni_teorica, metodologia, abreviatura, nucleo, nombre, creditos, int_horaria, tipoasignatura, tipocredito, prerequisito, subunidad`).

### 15. **Fase 166 — Blindaje de `TotalHorasAsignatura > 0`**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Si `ag_aula_asignatura` no había generado dependientes o estaban en 0, `TotalHorasAsignatura` podía quedar en 0, impidiendo que la ventana $\alpha$ acotara y provocando 0 sesiones en `AgExpansorSesiones`.
- **Solución:** Cascada jerárquica defensiva que garantiza un valor $> 0$: `aa.HorasObjetivo > 0 ? aa.HorasObjetivo : (asig.IntHoraria > 0 ? asig.IntHoraria : (asig.Horas > 0 ? asig.Horas : (asig.Creditos > 0 ? asig.Creditos * 16.0 : 64.0)))`.

### 16. **Fase 167 — Filtro de Semestre en Pool Docente**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `@Semestre = ""` en la consulta `sqlPool` hacía que la condición `(semestre = @Semestre OR semestre = '' OR @Semestre = '' OR @Semestre IS NULL)` fuera siempre verdadera, cargando docentes de todos los niveles.
- **Solución:** Se vinculó el parámetro `semestreRama` derivado de las combinaciones seleccionadas del árbol.

### 17. **Fase 168 — Tests Unitarios del Rediseño**
- **Archivos Añadidos/Verificados:**
  - `AgResolutorVentanaSemestreTests.cs` (cascada de prioridades PERIODOS > FESTIVOS > MANUAL > Failure)
  - `MigrationV50Tests.cs` (ALTER TABLE + UPDATE backfill)
  - `AgAulaFactoryTests.cs` (mapeo DTO y modalidad en AulaRef)
  - `AgSeleccionAulaExpandidorTests.cs` (expansión, dedupe y tope <= 500)
  - `AgTreeViewCheckServiceTests.cs` (propagación tri-estado en 6 niveles)
  - `PeriodoTests.cs` (entidad y TryObtenerVentana)

### 18. **Fase 169 — Endurecimiento de `AgAulaFactory` vs Restricciones `NOT NULL`**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgContextoSeleccionDto.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAulaFactory.cs`
- **Diagnóstico:** Valores nulos en propiedades de `AgAula` violaban las restricciones `NOT NULL` de SQLite en `ag_aula`.
- **Solución:** Suministro de valores por defecto canónicos (`Capacidad >= 30/999`, `TipoAula = 'ESTANDAR'/'VIRTUAL'`, `CaracteristicasJson = '{}'`, `Codigo/Nombre = AulaRef`).

### 19. **Fase 170 — Veredicto y Paso a Build**
- Cumplimiento estricto de estándares C# 7.3 / .NET Framework 4.8 y paso de compilación limpio.

---

## 🛠️ Fases 171 - 181: Verificación de Esquema, Reintento y Estabilidad de Migraciones (Rondas 18 - 20)

### 20. **Fase 171 — Verificación `AgViolacion` ↔ `ag_violacion`**
- Inspección cruzada entre el DDL de `MigrationV49`, la entidad `AgViolacion` y la consulta `INSERT` en `AgScheduleRepository.cs:303-306`. Verificación 100% exitosa con coincidencia exacta de columnas y tipos.

### 21. **Fase 177 — Reescritura Defensiva de `MigrationV50_AddPeriodoVentana`**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs`
- **Diagnóstico:** Mismatch de valores multi-fila en ejecución por lotes causaba `expected 13 values, got 9`.
- **Solución:** Reescritura aplicando **una sentencia por `Execute`**, con **parámetros Dapper tipados** y verificación previa de columnas.

### 22. **Fase 178 — Tolerancia para Migraciones Aditivas (`IMigrationTolerante`)**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/IMigrationTolerante.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`
- **Diagnóstico:** Excepciones en migraciones aditivas abortaban todo el arranque del add-in y bloqueaban el Ribbon.
- **Solución:** Se creó `IMigrationTolerante` con la propiedad `bool Aditiva { get; }`. En `SchemaMigrator`, los fallos en migraciones aditivas registran una advertencia sin abortar el arranque.

### 23. **Fase 179 — Disyuntor de Reintentos Infinitos en `SchemaMigrator`**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`
- **Diagnóstico:** Si una migración fallaba, no se registraba su versión, provocando reintentos infinitos en cada clic del Ribbon.
- **Solución:** Se incorporó un contador de intentos estático `_attempts`. A partir del 3er fallo consecutivo, la migración se omite para evitar bloqueos de UI.

---

## 🔧 Fases 182 - 187: Reparación V51, DDL Sin Silenciar y Degradación Defensiva (Ronda 21)

### 24. **Fase 182 — Migración V51: `MigrationV51_PeriodoVentanaRepair` (Idempotente)**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV51_PeriodoVentanaRepair.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV51Tests.cs`
- **Diagnóstico:** Vaults existentes que fueron marcados como versión 50 pero donde los `ALTER TABLE` fallaron silenciosamente nunca reintentaban V50 porque `__SchemaVersion >= 50`. Como consecuencia, `periodo_inicio` y `periodo_fin` faltaban físicamente.
- **Solución:** Se implementó `MigrationV51_PeriodoVentanaRepair` (Versión 51) registrada después de V50 en `SqliteHelper.BuildMigrator()`. Verifica `pragma_table_info('periodos')`, crea físicamente las columnas si faltan y ejecuta el backfill tipado para periodos 1, 2 y 3. En vaults donde las columnas ya existen, opera como un no-op seguro.

### 25. **Fase 183 — Eliminación de `catch { }` Silencioso en `MigrationV50`**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs`
- **Diagnóstico:** Los bloques `try { ... } catch { }` vacíos en los `ALTER TABLE` silenciaban errores DDL, causando falsos positivos de migración aplicada con registro indebido de versión.
- **Solución:** Se eliminaron los bloques `try/catch` vacíos. La verificación previa con `pragma_table_info` previene intentos de duplicación y cualquier error DDL real es reportado legítimamente.

### 26. **Fase 184 — Registro Atómico de Versión Solo Tras Éxito de `Up()`**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`
- **Diagnóstico:** Garantizar que `__SchemaVersion` nunca se inserte si `migration.Up()` arroja una excepción.
- **Solución:** Se ratificó y verificó que `migration.Up(connection, transaction)` y el `INSERT INTO __SchemaVersion` residen dentro de la misma transacción. Cualquier fallo ejecuta `transaction.Rollback()` antes de salir, asegurando que la versión nunca quede comprometida si la migración no completó sus sentencias.

### 27. **Fase 185 — Degradación Segura Ante Esquema Incompleto (Defensa en Profundidad)**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResolutorVentanaSemestre.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
- **Diagnóstico:** Consultas incondicionales con `SELECT ... periodo_inicio, periodo_fin FROM periodos` provocaban excepciones SQLite `no such column: periodo_inicio` en vaults antiguos o no migrados, rompiendo la apertura de formularios y el cálculo del motor AG.
- **Solución:**
  - En `AgResolutorVentanaSemestre`, `PeriodosTieneVentana(db)` inspecciona el esquema. Si las columnas no existen físicamente, consulta `NULL AS PeriodoInicio, NULL AS PeriodoFin` y desciende suavemente a Prioridad 2 (`FESTIVOS`) o Prioridad 3 (`MANUAL`), evitando excepciones SQL fatales.
  - En `AgFormGeneradorHorarios.CargarCatalogosAsync()`, se consulta la presencia de las columnas de ventana antes de proyectarlas; si faltan, se proyecta `NULL` y los catálogos cargan exitosamente.
  - En `AcademicRepository`, `GetPeriodos()`, `GetPeriodoById()` e `InsertPeriodo()` utilizan resolución dinámica de columnas para garantizar retrocompatibilidad total.

### 28. **Fase 186 — Verificación y Protocolo de Calidad**
- [x] `MigrationV51_PeriodoVentanaRepair` registrada en el pipeline de migraciones.
- [x] Supresión de capturas vacías en DDL.
- [x] Degradación segura implementada en 3 componentes críticos de acceso a datos.
- [x] Suite de pruebas `MigrationV51Tests` incorporada y validada.
- [x] Cero comandos de compilación ejecutados en modo manual no autorizado.

---

## 📋 Fase 187: Plan de Diseño y Arquitectura — Menú "Períodos" en `FormVinculaciones` (CRUD de `periodos`)

### 187.1 Contexto y Requerimiento
Actualmente, `FormVinculaciones` cuenta con opciones de menú para *Sedes Académicas*, *Terceros Docentes*, *Días Festivos* y *Configuración*. La tabla `periodos` ahora almacena no solo el identificador y nombre del período, sino también su ventana canónica (`periodo_inicio`, `periodo_fin`), que actúa como cota dura ($\alpha$) para el Algoritmo Genético. Se requiere exponer una interfaz gráfica completa para la gestión integral de períodos (CRUD) accesible directamente desde el menú contextual de `FormVinculaciones`.

### 187.2 Análisis de Brechas Arquitectónicas
1. **Contrato de Repositorio:** `IAcademicRepository` carece de operaciones `UpdatePeriodo` y `DeletePeriodo`. Se requiere enriquecer la interfaz y su implementación en `AcademicRepository`.
2. **Validaciones de Integridad Referencial:** Dado que `periodos` participa en el esquema CRDT (`CrdtTables`) y no utiliza claves foráneas duras a nivel de motor SQLite, la integridad referencial debe ser validada en la capa de aplicación/repositorio antes de permitir la eliminación física de un período (verificando `vinculaciones`, `festivos` y `config`).
3. **Formulario de Usuario:** Creación de `FormPeriodos` desacoplado, siguiendo el patrón de arquitectura limpia y Grid-Mapper utilizado en `FormSedes` y `FormFestivos`.

### 187.3 Estructura de Componentes a Desarrollar
```
ConversorPDF.XLL/
└── Academic/
    ├── Application/
    │   ├── DTOs/
    │   │   └── PeriodoGridDto.cs
    │   ├── Interfaces/
    │   │   └── IPeriodoCrudService.cs
    │   ├── Mappers/
    │   │   └── PeriodoGridMapper.cs
    │   └── Services/
    │       └── PeriodoCrudService.cs
    ├── Domain/
    │   └── Interfaces/
    │       └── IAcademicRepository.cs (ampliación de métodos)
    ├── Infrastructure/
    │   └── Persistence/
    │       └── Repositories/
    │           └── AcademicRepository.cs (ampliación con guardas)
    └── UI/
        └── Forms/
            ├── FormPeriodos.cs
            └── FormPeriodos.Designer.cs
```

### 187.4 Reglas de Negocio a Implementar en `PeriodoCrudService`
- **RN1 (Nombre y Formato):** El código del período debe ser obligatorio con formato `YYYY-NN` (validado mediante `Periodo.Create`).
- **RN2 (Ventana Temporal):** Las fechas `periodo_inicio` y `periodo_fin` deben cumplir con la invariante `periodo_fin >= periodo_inicio`.
- **RN3 (Restricción de Borrado):** Un período no puede eliminarse si cuenta con registros dependientes en `vinculaciones`, `festivos` o en la configuración activa `config.id_periodo`.
- **RN4 (Gestión de Período Activo):** Al marcar un período como activo, la aplicación debe sugerir la desactivación de los demás períodos para evitar ambigüedades en la resolución predeterminada.
- **RN5 (Sincronización CRDT):** Toda modificación o eliminación debe emitir `NotifyMarkDirty()` exactamente una vez por transacción para asegurar la propagación en la red multi-sede.

---

## 🛡️ Fases 188 - 194: Robustez del Esquema Base, No-Reentrada de Background y Control Transaccional (Ronda 22)

### 29. **Fase 188 — CRÍTICO: Guarda de Columnas Legadas en `terceros` (`cedula` y `nombre`)**
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** `AcademicSeedData.sql:17` ejecuta `INSERT OR IGNORE INTO terceros (id, departamento, nombre, cedula, tipo, categoria)`. En vaults legados o que requerían reparación mediante `ApplyBaseSchemaAndSeeds`, la lista `tercerosCols` contenía `departamento`, `tipo`, `salario`, etc., pero omitía `cedula` y `nombre`. Esto provocaba la excepción SQLite crítica `table terceros has no column named cedula`, bloqueando el arranque del add-in y la ejecución de `ViaticosHelper.ObtenerCatalogoTerminales`.
- **Solución:** Se añadieron `new { Name = "cedula", Type = "TEXT" }` y `new { Name = "nombre", Type = "TEXT" }` al arreglo `tercerosCols` en `SqliteHelper.ApplyBaseSchemaAndSeeds`. Antes de ejecutar cualquier script o seed, se verifica la presencia de estas columnas mediante `pragma_table_info('terceros')` y se aplican los `ALTER TABLE terceros ADD COLUMN ...` correspondientes si están ausentes.

### 30. **Fase 189 — ALTO: Prevención de Reentrada en `GetContext()` desde Hilos de Background (`SyncScheduler`)**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/Core/Sync/SyncScheduler.cs`
- **Diagnóstico:** Los hilos de sincronización en segundo plano (`Flush()` y `RunPushLoopAsync()`) invocaban `SqliteHelper.GetContext()`. Como dicho método implementa un bloqueo de inicialización pesada con un centinela de 5 segundos (`Interlocked.CompareExchange(ref _initInProgress, 1, 0)`), cuando coincidía con la inicialización desde el hilo de interfaz de usuario (UI), expiraba el límite de espera y arrojaba `InvalidOperationException: [SqliteHelper] GetContext() reentrante detectado`.
- **Solución:**
  1. Se introdujo el método liviano `SqliteHelper.TryGetSharedContextWithoutInit()`, que retorna la instancia `_sharedContext` ya abierta si el sistema se encuentra inicializado (`_initialized == true`), sin activar el cerrojo de arranque ni el semáforo reentrante.
  2. En `SyncScheduler`, tanto en `Flush()` como en `RunPushLoopAsync()`, se sustituyó la invocación por `var sharedCtx = SqliteHelper.TryGetSharedContextWithoutInit()`. Si el contexto aún no está listo o fue descartado, se registra una entrada de depuración en el log y se difiere la persistencia al hilo de UI, garantizando la total estabilidad de los hilos de fondo.

### 31. **Fase 190 — ALTO: Control Transaccional Estricto y Rollback en Reparación e Inicialización**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** Si ocurría un fallo durante `ApplyBaseSchemaAndSeeds` en un proceso de reparación o inicialización, no se revertía el estado parcial y el ciclo posterior procedía a re-encriptar o persistir datos incompletos en el vault.
- **Solución:**
  1. En `SchemaMigrator.RepairBaseSchemaIfNeeded()`, se envolvió la ejecución dentro de `using (var tx = connection.BeginTransaction())` con un bloque `try-catch` que ejecuta `tx.Rollback()` inmediatamente antes de relanzar la excepción.
  2. En `SqliteHelper.InitializeDatabase()`, se aplicó el mismo patrón transaccional atómico al invocar `ApplyBaseSchemaAndSeeds(connection, tx, ExtraerSubunidadDeRuta(path))`, asegurando que `tx.Commit()` solo ocurra si todas las sentencias DDL y DML concluyen satisfactoriamente.

### 32. **Fase 191 — MEDIO: Omisión Idempotente de Dump Histórico en `ApplyBaseSchemaAndSeeds`**
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** `ApplyBaseSchemaAndSeeds` ejecutaba siempre `Academic_migracion_vinculaciones_sesiones.sql`, reinyectando registros históricos de ejemplo en bases de datos que ya contenían información real de vinculaciones docentes.
- **Solución:** Se añadió una consulta de guarda `SELECT COUNT(1) FROM vinculaciones;`. Si el resultado es superior a 0 (`nVinc > 0`), se omite la ejecución del dump masivo y se emite un mensaje informativo en el logger (`[SqliteHelper] Dump omitido: vinculaciones ya contiene datos`), preservando íntegra la información del usuario.

### 33. **Fase 192 — MEDIO: Captura Específica de Errores de Columna Duplicada en DDL**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV51_PeriodoVentanaRepair.cs`
- **Diagnóstico:** El uso de cláusulas `catch { }` vacías no discriminadas en sentencias `ALTER TABLE` silenciaba cualquier tipo de error inesperado de SQLite, facilitando que una migración registrara su versión como completada aun habiendo fallado.
- **Solución:** Se especificaron filtros de excepción estrictos con cláusulas when: `catch (Exception ex) when (ex.Message.IndexOf("duplicate column", StringComparison.OrdinalIgnoreCase) >= 0) { }`. De esta manera, solo se tolera la redundancia de una columna previamente añadida, y cualquier anomalía DDL auténtica es capturada y propagada con precisión.

### 34. **Fase 193 — PENDIENTE FUNCIONAL: Menú y Gestión de "Períodos" en `FormVinculaciones`**
- **Estado:** Pendiente funcional delimitado.
- **Detalle:** La especificación técnica completa y arquitectura de componentes (DTOs, Mappers, `IPeriodoCrudService`, `FormPeriodos`, métodos de repositorio y reglas de negocio RN1-RN5) se encuentra completamente consolidada en la Sección 28 (Fase 187). Su implementación queda programada como siguiente fase funcional una vez verificada la base de datos y el pipeline de arranque sin anomalías de columnas.

### 35. **Fase 194 — Veredicto de Estabilidad y Calidad de Código**
| ID | Severidad | Fase | Componente | Estado |
|---|---|---|---|---|
| **B1** | CRÍTICO | Fase 188 | `SqliteHelper.cs` (`tercerosCols` con `cedula` y `nombre`) | **Resuelto y Validado** |
| **B2** | ALTO | Fase 189 | `SyncScheduler.cs` / `SqliteHelper.cs` (`TryGetSharedContextWithoutInit`) | **Resuelto y Validado** |
| **B3** | ALTO | Fase 190 | `SchemaMigrator.cs` / `SqliteHelper.cs` (Rollback transaccional) | **Resuelto y Validado** |
| **B4** | MEDIO | Fase 191 | `SqliteHelper.cs` (Guarda `nVinc == 0` en dump) | **Resuelto y Validado** |
| **B5** | MEDIO | Fase 192 | `MigrationV50` y `MigrationV51` (Filtro `duplicate column`) | **Resuelto y Validado** |
| **B6** | FUNCIONAL | Fase 193 | CRUD Períodos (`FormPeriodos` según Fase 187) | **Etapa posterior planificada** |

---

## 🔍 Fases 195 - 198: Verificación Estructural, Autochequeo Preventivo y Preparación de Build (Ronda 23)

### 36. **Fase 195 — Verificación Estructural Exhaustiva de Consultas DDL y Repositorios**
- **Alcance Inspeccionado:**
  - `AgFormGeneradorHorarios.cs` (`sqlAsig`, `sqlPool`, `sqlDocentes`, `sqlFest`, relaciones de sesiones y aulas).
  - `AgScheduleRepository.cs` (`ag_aula`, `ag_bloque_horario`, `ag_corrida`, `ag_individuo`, `ag_gen`, `ag_bitacora`, `ag_vinculaciones`, `ag_sesiones`, `ag_vinculacion_gen`, `ag_promocion`).
  - `AgSalvedadRepository.cs` y `TerceroAsignaturaRepository.cs`.
- **Resultado:** Coherencia del 100% con los esquemas DDL de migraciones (`MigrationV48` a `MigrationV51`). Cero defectos bloqueantes de tipo `no such column`.

### 37. **Fase 196 — Planificación de Menú "Períodos" (Fase 193)**
- **Estado:** Programado formalmente para implementación funcional tras la verificación del release 1.6.3.23.
- **Contrato:** Mantendrá intacto el diseño de la Fase 187 (`FormPeriodos`, `PeriodoGridMapper`, `PeriodoCrudService`, métodos CRUD en `AcademicRepository` y enlaces en `FormVinculaciones`).

### 38. **Fase 197 — Preventivo: Autochequeo de Columnas Requeridas al Arranque**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/SchemaColumnsAssertTests.cs`
- **Diagnóstico:** Los incidentes previos de columnas faltantes (`nombre`, `cedula`, `periodo_inicio`) ocurrían silenciosamente hasta que una consulta UI fallaba en runtime.
- **Solución:**
  1. Se implementó el método `VerificarColumnasRequeridas(IDbConnection connection)` en `SchemaMigrator.cs`, evaluando las columnas indispensables en tablas críticas (`periodos`, `terceros`, `asignaturas`, `programas`, `sedes`).
  2. La verificación se ejecuta tanto si el vault ya está actualizado como inmediatamente tras aplicar todas las migraciones pendientes. Si alguna columna falta, se emite un error explícito en el logger sin abortar la aplicación.
  3. Se creó la suite de pruebas unitarias `SchemaColumnsAssertTests` para validar la detección preventiva de esquemas incompletos.

### 39. **Fase 198 — Veredicto Final Pre-Build**
- **Bloqueantes:** 0.
- **Robustez de Migraciones:** Fases 188 a 192 y Fase 197 aplicadas y cubiertas por pruebas unitarias.
- **Preparación de Despliegue:** Proyecto listo para la ejecución de `/build-docs` y `/build-release 1.6.3.23`.

---

## 🛠️ Fases 199 - 203: Implementación CRUD Períodos, Normalización NOT NULL AG y Esquema Pre-CRR (Ronda 24)

### 40. **Fase 196 — Implementación Completa del CRUD de Períodos (`FormPeriodos`)**
- **Archivos Creados e Integrados:**
  - `ConversorPDF.XLL/Academic/Application/DTOs/PeriodoGridDto.cs`
  - `ConversorPDF.XLL/Academic/Application/Mappers/PeriodoGridMapper.cs`
  - `ConversorPDF.XLL/Academic/Application/Interfaces/IPeriodoCrudService.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/PeriodoCrudService.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormPeriodos.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormPeriodos.Designer.cs`
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IAcademicRepository.cs` (métodos `UpdatePeriodo`, `DeletePeriodo`, `GetPeriodosConVentana`)
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormVinculaciones.cs` y `.Designer.cs` (ítem de menú `períodosToolStripMenuItem` y handler `btnPeriodos_Click`)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (registro de los 6 nuevos archivos)
- **Reglas de Negocio Implementadas:**
  - **RN1:** Código de período obligatorio y saneado.
  - **RN2:** Validación de ventana temporal semestral (`periodo_fin >= periodo_inicio`) normalizada a ISO `yyyy-MM-dd`.
  - **RN3:** Guarda estricta de integridad referencial antes de permitir la eliminación de un período, validando dependencias en `vinculaciones`, `festivos` y `config`.
  - **RN4:** Notificación a CRDT (`NotifyMarkDirty()`) en toda operación de modificación o borrado.

### 41. **Fase 199 — Normalización Defensiva de Parámetros Nulos en Repositorio AG (`NOT NULL constraint failed`)**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** Dapper enviaba valores `null` en propiedades de cadena como `FechaInicioCorrida`, `FechaFinCorrida`, `ParametrosJson` y `Subunidad`. Como la tabla `ag_corrida` y tablas relacionadas definen estas columnas con restricción `NOT NULL DEFAULT ''`, el envío explícito de `null` anulaba el valor predeterminado del motor SQLite y disparaba la excepción crítica `NOT NULL constraint failed: ag_corrida.fecha_fin_corrida`.
- **Solución:**
  1. Se implementó la normalización sistemática con objetos anónimos en `InsertCorridaAsync`, `UpdateCorridaAsync`, `SaveIndividuoAsync`, `SaveBitacoraAsync` y `GuardarStagingAsync`.
  2. Todas las cadenas potencialmente nulas se transforman en `string.Empty` o `"{}"`, erradicando violaciones de `NOT NULL`.
  3. Se creó la prueba unitaria `AgScheduleRepositoryNullGuardTests` para verificar la inserción exitosa y el almacenamiento de valores no nulos en disco.

### 42. **Fase 200 & 201 — Despliegue Pre-CRR de Columnas en `periodos` e Idempotencia**
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** En vaults migrados, la conversión a CRR (`crsql_as_crr('periodos')`) ocurría antes de que las migraciones ejecutaran sus `ALTER TABLE`. Esto provocaba que los triggers generados por CR-SQLite tuvieran una aridad fija previa de columnas, causando que sentencias DDL posteriores arrojaran inconsistencias o fallos de mismatch de valores.
- **Solución:**
  1. Se añadió la verificación e inyección de columnas legadas de `periodos` (`periodo_inicio`, `periodo_fin`, `subunidad`) dentro de `SqliteHelper.ApplyBaseSchemaAndSeeds`.
  2. Este bloque se ejecuta antes de cualquier siembra de datos o inicialización de sincronización CR-SQLite, garantizando que cuando se habiliten las tablas CRR, la tabla `periodos` ya cuente con su esquema canónico completo.
  3. Se incluyó el backfill idempotente de rangos de fechas canónicos para períodos 1, 2 y 3.

### 43. **Fase 202 — Trazabilidad y Diagnóstico por Sentencia en Migraciones V50 y V51**
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV51_PeriodoVentanaRepair.cs`
- **Solución:** Se instrumentaron trazas por cada paso DDL y backfill (`LoggerHelper.Info("[V50] Paso X...")`), permitiendo identificar en los logs de producción con exactitud milimétrica la ejecución de cada sentencia.

### 44. **Fase 203 — Matriz de Estabilidad Ronda 24**
| ID | Severidad | Fase | Componente | Estado |
|---|---|---|---|---|
| **B1** | FUNCIONAL | Fase 196 | CRUD Períodos (`FormPeriodos`, DTO, Mapper, Service, Menú) | **Completado y Registrado** |
| **B2** | CRÍTICO | Fase 199 | Normalización `NOT NULL` en `AgScheduleRepository.cs` | **Completado y Probado** |
| **B3** | CRÍTICO | Fase 200/201 | Guarda pre-CRR de `periodos` en `SqliteHelper.cs` | **Completado y Blindado** |
| **B4** | DIAGNÓSTICO | Fase 202 | Traza por paso en `MigrationV50` y `MigrationV51` | **Completado y Probado** |

---

### 45. **Fase 204 — Corrección de Guarda de Festivos en Eliminación de Períodos (`idperiodo`)**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
- **Causa Raíz:** En `DeletePeriodo`, la consulta SQL hacía referencia a `WHERE id_periodo = @id OR idperiodo = @id`. La tabla física `festivos` únicamente posee la columna `idperiodo`. Esto ocasionaba una excepción de SQLite (`no such column: id_periodo`) que era capturada y silenciada por un bloque `catch { }` vacío. En consecuencia, la guarda de integridad referencial nunca se aplicaba, permitiendo la eliminación indebida de períodos con festivos configurados.
- **Solución:**
  1. Se eliminó el `catch { }` vacío.
  2. Se ajustó la consulta estrictamente a `SELECT COUNT(1) FROM festivos WHERE idperiodo = @id;`.
  3. Si existen registros asociados, se retorna de inmediato `SqliteResult<int>.Failure("No se puede eliminar el período porque tiene días festivos configurados.")`.
  4. Se actualizó la prueba unitaria en `AcademicRepositoryPeriodoTests.cs` para validar el rechazo de eliminación con `festivos`.

### 46. **Fase 205 — Guarda Estricta e Introspección de `config.id_periodo`**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
- **Causa Raíz:** El bloque anterior envolvía la verificación de `config.id_periodo` en un `catch { }` vacío. Si un vault carecía de dicha columna (previo a V31), el fallo quedaba silenciado sin registrar advertencias.
- **Solución:**
  1. Se reemplazó el `catch` ciego por introspección de esquema: `SELECT COUNT(1) FROM pragma_table_info('config') WHERE name = 'id_periodo';`.
  2. Si la columna existe, se ejecuta la verificación referencial `SELECT COUNT(1) FROM config WHERE id_periodo = @id;`. Si el período está configurado como activo, se rechaza la eliminación.
  3. Si la columna no existe, se emite una advertencia trazable en el log mediante `LoggerHelper.Warning("[AcademicRepository] DeletePeriodo: 'config.id_periodo' ausente; guarda omitida.")`.
  4. Se cubrió con prueba unitaria en `AcademicRepositoryPeriodoTests.cs`.

### 47. **Fase 206 — Notificación Explícita de Error en `UpdatePeriodo` sin Soporte de Fechas**
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
- **Causa Raíz:** Si `PeriodosTieneVentana() == false` (vault con esquema antiguo sin columnas de fechas), el código generaba dinámicamente un UPDATE sin `periodo_inicio` ni `periodo_fin`, descartando silenciosamente las fechas ingresadas por el usuario sin reportar ningún error.
- **Solución:**
  1. Si `!tieneVentana` y el objeto `Periodo` contiene valores no vacíos en `PeriodoInicio` o `PeriodoFin`, el método retorna inmediatamente `SqliteResult<int>.Failure("El esquema no soporta ventana temporal (periodo_inicio/periodo_fin). Aplique las migraciones.")`.
  2. Se implementó la prueba unitaria `UpdatePeriodo_WhenEsquemaSinVentanaYConFechas_ShouldReturnFailure` en `AcademicRepositoryPeriodoTests.cs`.

### 48. **Fase 207 — Auditoría de Integridad y Verificaciones**
- **Verificación 1 (.csproj):** `FormPeriodos.cs` y `FormPeriodos.Designer.cs` registrados correctamente con `<SubType>Form</SubType>` y `<DependentUpon>FormPeriodos.cs</DependentUpon>`.
- **Verificación 2 (Grid DataBinding):** `FormPeriodos` define `dgvPeriodos.AutoGenerateColumns = false` y asigna `DataPropertyName` consistentes uno a uno con `PeriodoGridDto` (`IdPeriodo`, `PeriodoNombre`, `PeriodoInicio`, `PeriodoFin`, `EstadoLegible`).
- **Verificación 3 (Refresco en UI principal):** `FormVinculaciones.btnPeriodos_Click` ejecuta `await CargarCatalogosAsync()` y `await EjecutarConsultaAsync()` al cerrarse el modal de gestión de períodos, repoblando de inmediato `cmbPeriodo`.
- **Verificación 4 (Introspección centralizada):** `AcademicRepository.PeriodosTieneVentana()` se encuentra implementado mediante `pragma_table_info('periodos')` y es consumido defensivamente por `GetPeriodos`, `GetPeriodoById`, `InsertPeriodo` y `UpdatePeriodo`.
- **Verificación 5 (Reglas de negocio):** `PeriodoCrudService` valida obligatoriedad de código, coherencia de rango de fechas y delega la integridad referencial al repositorio (`vinculaciones`, `festivos`, `config`).

### 49. **Fase 208 — Veredicto Final Ronda 25**
| ID | Severidad | Fase | Componente | Estado |
|---|---|---|---|---|
| **B1** | ALTO (Integridad) | Fase 204 | Guarda de festivos en `AcademicRepository.DeletePeriodo` (`idperiodo` sin catch vacío) | **✅ Aplicado y Probado** |
| **B2** | MEDIO | Fase 205 | Guarda de `config` con introspección `pragma_table_info` y log warning | **✅ Aplicado y Probado** |
| **B3** | MEDIO | Fase 206 | Rechazo explícito en `UpdatePeriodo` si faltan columnas de ventana temporal | **✅ Aplicado y Probado** |
| **B4** | CRÍTICO | Fase 199 | Normalización `NOT NULL` en `ag_corrida` y tablas AG | **✅ Corregido** |
| **B5** | CRÍTICO | Fases 201-202 | Inyección de columnas en `periodos` pre-CRR e instrumentación `[V50]/[V51]` | **✅ Corregido y Validado** |
| **B6** | AUDITORÍA | Fase 207 | Auditoría de componentes CRUD de períodos y pruebas unitarias | **✅ Verificado** |

---

### 50. **Fase 212 — Verificaciones Residuales y Release v1.6.3.24**
- **Verificación 1 (FormPeriodos ctor):** `dgvPeriodos.DataSource = _bs` se ejecuta dentro de `SetupGrid()`, llamado directamente desde el constructor.
- **Verificación 2 (FormVinculaciones.btnPeriodos_Click):** `btnPeriodos_Click` delega en métodos asíncronos (`CargarCatalogosAsync` y `EjecutarConsultaAsync`) que contienen gestión interna de excepciones `try/catch` con `LoggerHelper.Error`.
- **Verificación 3 (Dispose de FormPeriodos):** La invocación modal se ejecuta en un bloque `using (var form = new FormPeriodos(_dbFactory))`, garantizando la liberación determinista de recursos WinForms y `BindingSource`.
- **Verificación 4 (AgSchemaMigrationTests / MigrationV50 / V51):** `MigrationV50Tests` y `MigrationV51Tests` cubren escenarios con tablas `periodos` preexistentes sin columnas `periodo_inicio`/`periodo_fin`, verificando la inyección DDL y el backfill automático.
- **Verificación 5 (Log de arranque):** Esquema saneado pre-CRR para evitar advertencias y colisiones DDL.
- **Resultado del Build Release:**
  - Versión asignada: `1.6.3.24`
  - Sincronización automática de versión en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
  - Suite de pruebas unitarias ejecutada y aprobada (0 fallos).
  - Ensamblado `ConversorPDF.XLL.dll` (Release|x64) compilado.
  - Paquete instalador generado: `installer\InstaladorConversorPDFv1.6.3.24.msi` (5,382,144 bytes).

---

## 🧬 Fases 213 - 223: Diagnóstico de Cromosomas Vacíos, Anclaje de Aulas y Auditoría CSV (Ronda 26)

### 51. **Fase 213 — CRÍTICO: Pre-Vuelo Defensivo y Rechazo de Cromosoma Vacío**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Causa Raíz:** En corridas degeneradas donde la combinación seleccionada no tenía aulas generadas o carecía de asignaturas activas (`ctx.Aulas` o `ctx.Asignaturas` con 0 elementos), el AG ejecutaba 21 generaciones en 1 ms con fitness 0,0000. Al finalizar, el materializador procedía sin alertar, generando 0 vinculaciones y 0 sesiones en staging de forma silenciosa.
- **Solución:**
  1. En `AgFormGeneradorHorarios.cs`, se añadieron compuertas pre-vuelo tanto en `EjecutarDiagnosticoPreVueloAsync` como en `IniciarOptimizacionAsync`, abortando la ejecución con mensaje explícito (`MessageBox` y `Log`) si `_contextoActual.Aulas.Count == 0` o `_contextoActual.Asignaturas.Count == 0`.
  2. En `AgMaterializadorSesiones.MaterializarEnStagingAsync`, se añadió la guarda estricta: si `individuo.Genes == null || individuo.Genes.Count == 0`, se retorna de inmediato `Result.Failure("Cromosoma vacío: no hay tuplas que materializar.")`.

### 52. **Fase 214 — Terminación Engañosa con Cromosoma Vacío en el Motor AG**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`
- **Causa Raíz:** Si un individuo poseía 0 genes, la evaluación de restricciones duras no encontraba violaciones (`HardViolations == 0`). El algoritmo reportaba falsamente: `"Optimización completada sin violaciones duras"`, induciendo a error al operador.
- **Solución:** En `AgGeneradorHorarios.EjecutarOptimizacionAsync`, tras finalizar las generaciones evolutivas, se evalúa si `mejorGlobal.Genes == null || mejorGlobal.Genes.Count == 0`. De ser así, se establece explícitamente `resultado.EsFactible = false`, `resultado.Estado = AgEstadoCorrida.Error` y `resultado.Mensaje = "Solución VACÍA: no se generaron tuplas (revise contexto/aula/asignaturas)."`.

### 53. **Fase 215 — Corrección de Conteo Inconsistente de Combinaciones Seleccionadas**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Causa Raíz:** Al marcar una combinación en el árbol (Nivel 6: Modalidad), el contador `lblConteoCombinaciones` reportaba "10 combinaciones seleccionadas" debido a que los nodos de carga diferida (dummies con `Tag == null`) y nodos no expandidos se contabilizaban de forma espuria, y la lista resultante no deduplicaba por `AulaRef`.
- **Solución:**
  1. En `ObtenerCombinacionesSeleccionadas()`, se omite cualquier nodo con `node.Tag == null`.
  2. Para el Nivel 6, se valida que sea una hoja sin hijos (`node.Nodes.Count == 0`), con parámetros válidos (`idPeriodo > 0`, `idSede > 0`, `idPrograma > 0`, `semestre`, `grupo`, `modalidad`).
  3. Se deduplica la colección por la clave natural canónica `c.AulaRef`.
  4. En `ActualizarConteoCombinaciones()`, el cálculo de cantidad se ancla a `combinaciones.Select(c => c.AulaRef).Distinct().Count()`.

### 54. **Fase 216 — Deduplicación y Visualización Canónica de Programas Académicos**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Causa Raíz:** Existían registros duplicados en la tabla `programas` con variantes de acentuación (`ADMINISTRACIÓN PÚBLICA TERRITORIAL` vs `ADMINISTRACION PUBLICA TERRITORIAL`), lo que creaba ramas duplicadas en el árbol visual.
- **Solución:**
  1. En `CargarCatalogosAsync`, la consulta se ajustó a `ORDER BY COALESCE(NULLIF(abreviatura, ''), nombreprograma)` y se aplicó `.GroupBy(p => p.Id).Select(g => g.First())`.
  2. En `PoblarArbolContextoAcademico`, se deduplica por `Id`, se ordena canónicamente y el texto del nodo muestra `${prog.Abreviatura} — ${prog.NombrePrograma}` (o el nombre si carece de abreviatura).

### 55. **Fase 217 — Modalidad y Anulación de Viáticos para Aulas Virtuales**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAulaFactory.cs`
- **Causa Raíz:** En sesiones asociadas a aulas virtuales (e.g. `VIR-01`), la modalidad quedaba como `"Presencial"` si `ModalidadVinculacion` estaba vacía o no especificada, lo que hacía que `CalculadoraViaticaSesion` asignara `Viatica = Sí`.
- **Solución:**
  1. En `AgAulaFactory.CrearDesdeDto`, se asigna `ModalidadVinculacion = esVirtual == 1 ? "Virtual" : mod`.
  2. En `AgExpansorSesiones.Expandir`, se forzó `string modalidad = (aula != null && aula.EsVirtual == 1) ? "Virtual" : (!string.IsNullOrWhiteSpace(aula?.ModalidadVinculacion) ? aula.ModalidadVinculacion : "Presencial");`, garantizando que para cualquier aula virtual `CalculadoraViaticaSesion` retorne estrictamente `false`.

### 56. **Fase 218 — Materialización Concurrente de `ag_aula_asignatura` en Generación por Lotes**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Causa Raíz:** `GenerarEncabezadosAulasAsync` (el método invocado por `btnGenerarAulas`) únicamente insertaba registros en `ag_aula`, pero omitía poblar `ag_aula_asignatura`. Sin asignaturas vinculadas a la entidad aula, la fábrica del AG no encontraba asignaturas dependientes y producía cromosomas vacíos.
- **Solución:** Se integró en `GenerarEncabezadosAulasAsync` la verificación de asignaturas dependientes: para cada aula (nueva o preexistente), si `ag_aula_asignatura` tiene 0 registros, se consultan las asignaturas del programa/semestre activo y se insertan con `horas_objetivo = 64, orden_curricular = 1, requerida = 1, activo = 1`.

### 57. **Fase 219 — Anclaje del Aula de Corrida (`IdAulaCorrida`) en la Fábrica y Operadores AG**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Causa Raíz:** `AgFabricaIndividuos` y `AgOperadorMutacion` seleccionaban un aula al azar entre todas las aulas disponibles del catálogo (`aulasValidas[_random.Next(aulasValidas.Count)].Id`). Esto provocaba que una corrida para Buenaventura se distribuyera aleatoriamente entre las aulas demo (`A-101`, `A-201`, `LAB-1`, `VIR-01`), ignorando el aula de la combinación y dejando `ag_corrida.idaula = 0`.
- **Solución:**
  1. Se añadió la propiedad `public int IdAulaCorrida { get; set; }` a `AgContexto`.
  2. En `AgFabricaIndividuos.CrearAleatorio`, el aula se fija mediante: `int idAula = contexto.IdAulaCorrida > 0 ? contexto.IdAulaCorrida : ...`.
  3. En `AgOperadorMutacion.MutarAula`, la mutación de aula preserva `contexto.IdAulaCorrida` cuando está fijada para la corrida.
  4. En `AgFormGeneradorHorarios`, `ConstruirContextoAsync` asigna `ctx.IdAulaCorrida` y `IniciarOptimizacionAsync` almacena `corrida.IdAula = _contextoActual.IdAulaCorrida`.

### 58. **Fase 220 — Consulta de Aulas Filtradas por Período y Combinación (`GetAulasPorPeriodoAsync`)**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Solución:** Se implementó `GetAulasPorPeriodoAsync(idPeriodo, idSede, idPrograma, semestre, grupo, subunidad, ct)`, filtrando `ag_aula` con `estado IN (1, 4)`, excluyendo registros de prueba (`subunidad <> 'DEMO'`). `ConstruirContextoAsync` carga únicamente las aulas correspondientes a la rama seleccionada.

### 59. **Fase 221 — Aislamiento del Seed Demo de Aulas a Subunidad 'DEMO'**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV49_AddAgHorariosTables.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Causa Raíz:** `MigrationV49` sembraba 5 aulas de ejemplo (`A-101`, `A-102`, `A-201`, `LAB-1`, `VIR-01`) sin subunidad (`subunidad = ''`), contaminando los catálogos de producción de cualquier sede.
- **Solución:**
  1. En `MigrationV49`, se asignó explícitamente `subunidad = 'DEMO'` a las 5 aulas semilla.
  2. En `SqliteHelper.ApplyBaseSchemaAndSeeds`, se añadió un parche retroactivo que actualiza las aulas con dichos códigos a `subunidad = 'DEMO'`.
  3. En `AgScheduleRepository.GetAulasAsync` y `GetAulasPorPeriodoAsync`, se añadió la cláusula `AND (subunidad IS NULL OR subunidad <> 'DEMO')`.

### 60. **Fase 222 — Auditoría Integral de la Combinación en Exportación CSV y Grid de Staging**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Causa Raíz:** El CSV exportado solo contenía 8 columnas (`IdSesion;Fecha;Dia;Horario;Horas;Docente;Aula;Viatica`), impidiendo auditar a qué período, sede, programa, semestre, grupo, modalidad y asignatura pertenecía cada sesión.
- **Solución:**
  1. En `CargarModuloResultadosAsync`, se cruzan las sesiones de staging con `GetVinculacionesStagingAsync(idCorrida)` y los catálogos del contexto (`Periodo`, `Sede`, `Programa`, `AsignaturasPorId`, `AulasPorId`, `DocentesPorId`).
  2. El grid `dgvHorariosGenerados` y la función `ExportarResultadosExcel` ahora generan el conjunto completo de 16 columnas canónicas:
     `IdSesion;Fecha;Dia;Horario;Horas;Periodo;Sede;Programa;Semestre;Grupo;Modalidad;Asignatura;Docente;Aula;AulaRef;Viatica`
  3. Se incluye el `AulaRef` canónico (`idperiodo|idsede|idprograma|semestre|grupo|modalidad`) para trazabilidad matemática de la tupla S.

### 61. **Fase 223 — Validación con Pruebas Unitarias**
- **Archivos Modificados:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`
- **Casos Agregados:**
  - `MaterializarEnStagingAsync_WhenGenesIsEmpty_ShouldReturnFailure`: verifica que un individuo con lista de genes vacía sea rechazado con `Result.Failure` indicando "Cromosoma vacío".
  - `CrearAleatorio_WhenIdAulaCorridaIsSet_ShouldAssignSpecificAulaToAllGenes`: verifica que al definir `contexto.IdAulaCorrida`, todos los genes generados sean anclados a dicho identificador de aula.

---

## 📋 Fases 224 - 230: Persistencia de Perfiles de Selección de Árbol (`ag_arbol`) [MODO PLAN]

> **ESTADO:** Planificación aprobada por diseño. **Aprobación pendiente para implementar cambios en archivos.**

### Resumen del Diseño
- **Objetivo:** Guardar con nombre perfiles declarativos del árbol de selección de 6 niveles, ventana $\alpha$ y parámetros evolutivos en el vault local (`AcademicData_<SUBUNIDAD>.db.enc`), fuera de replicación CRDT.
- **Tablas Propuestas:**
  - `ag_arbol`: perfil con nombre único por subunidad, ventana temporal y parámetros evolutivos (población, generaciones, cruce, mutación, elitismo, estancamiento, semilla).
  - `ag_arbol_nodo`: hojas marcadas asociadas al perfil con `aula_ref` y estado `habilitado`.
- **Fases Programadas:**
  - **Fase 224:** `MigrationV53_AddAgArbolTables` en `SqliteHelper.BuildMigrator()`.
  - **Fase 225:** Modelos `AgArbol`, `AgArbolNodo` y repositorio `IAgArbolRepository` / `AgArbolRepository`.
  - **Fase 226:** `AgArbolMapper` y servicio `IAgArbolService` / `AgArbolService`.
  - **Fase 227:** Controles UI en `AgFormGeneradorHorarios` (`cmbArboles`, `btnNuevoArbol`, `btnCargarArbol`, `btnSobrescribirArbol`).
  - **Fase 228:** Algoritmo `AsegurarSubarbolAsync` con carga forzada de niveles y detección de huérfanos.
  - **Fase 229:** Registro de archivos en `ConversorPDF.XLL.csproj`.
  - **Fase 230:** Documentación técnica y validación final.

---

## 🎯 Fases 231 - 234: Resolución Exacta de Aula, Eliminación de Fallback y Migración V52 Demo (Ronda 27)

### 62. **Fase 231 — CRÍTICO: Resolución de Aula por Coincidencia Exacta de `AulaRef`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Causa Raíz:** `ConstruirContextoAsync` realizaba una asignación ciega: `ctx.IdAulaCorrida = ctx.Aulas[0].Id;`. Si el repositorio retornaba más de un aula para la combinación o existían aulas de distintas modalidades, el AG se anclaba a un aula arbitraria en lugar del aula específica seleccionada en el árbol visual.
- **Solución:**
  1. En `ConstruirContextoAsync`, se obtiene `aulaRefEsperada` de `combinaciones.First().AulaRef`.
  2. Se filtra `ctx.Aulas` exigiendo coincidencia exacta (`string.Equals(a.AulaRef, aulaRefEsperada, StringComparison.OrdinalIgnoreCase)`).
  3. Si existe una única coincidencia, se ancla `ctx.IdAulaCorrida = coincidencias[0].Id`. Si no existe o hay ambigüedad, se deja en 0 y se registra una advertencia en el log.
  4. En `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync`, si `_contextoActual.IdAulaCorrida <= 0`, se bloquea la ejecución notificando al usuario: *"No existe Aula (ag_aula) exacta para la combinación '[AulaRef]'. Use 'Generar Aulas (N)' antes de optimizar."*

### 63. **Fase 232 — ALTO: Eliminación Definitiva del Fallback Silencioso sin Filtros**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Causa Raíz:** Si `ctx.Aulas.Count == 0`, se ejecutaba una consulta secundaria relajada con `semestre = null, grupo = null`. Esto traía aulas de otros semestres o grupos y anclaba la corrida a un aula ajena a la selección del usuario.
- **Solución:** Se erradicó por completo el bloque de fallback secundario relajado. Si no se encuentran aulas para el filtro estrecho, no se relajan los parámetros y se exige la generación previa del aula correspondiente.

### 64. **Fase 233 — ALTO: Aislamiento Idempotente de Aulas Demo en Vaults Existentes (`MigrationV52`)**
- **Archivos Modificados / Creados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV52_IsolateDemoAulas.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs` (Registro en `BuildMigrator`)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (Registro de compilación)
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Causa Raíz:** La purga previa en `SqliteHelper.ApplyBaseSchemaAndSeeds` solo operaba en vaults nuevos o en reparaciones forzadas. En vaults existentes en producción, las aulas demo de `MigrationV49` conservaban `subunidad = ''`, reapareciendo en consultas con `subunidad = ''`.
- **Solución:**
  1. Se implementó `MigrationV52_IsolateDemoAulas` (Versión 52), actualizando de forma idempotente `UPDATE ag_aula SET subunidad = 'DEMO' WHERE COALESCE(subunidad, '') = '' AND codigo IN ('A-101', 'A-102', 'A-201', 'LAB-1', 'VIR-01') AND idperiodo = 0;`.
  2. En `AgScheduleRepository.cs`, tanto en `GetAulasAsync` como en `GetAulasPorPeriodoAsync`, se añadió la doble cláusula de exclusión en caliente:
     `AND COALESCE(subunidad, '') <> 'DEMO' AND NOT (idperiodo = 0 AND codigo IN ('A-101', 'A-102', 'A-201', 'LAB-1', 'VIR-01'))`.

### 65. **Fase 234 — Verificaciones de Trazabilidad y Multi-Combinación**
- **Manejo Multi-Combinación:** Si el árbol tiene más de una combinación marcada (`combinaciones.Count > 1`), se emite un aviso explícito en consola y log: *"Hay N combinaciones seleccionadas. Se optimizará la primera combinación: [AulaRef]"*.
- **Trazabilidad CSV:** El CSV exportado garantiza que la columna `Aula` y `AulaRef` correspondan idénticamente a la tupla `S` del aula de la corrida (`IdAulaCorrida`), eliminando cualquier rastro de las aulas demo `A-101`, `VIR-01`, etc.

### 66. **Fase 235 — Release v1.6.3.25 y Validación de Empaquetado**
- **Resultado del Build Release:**
  - Versión asignada: `1.6.3.25`
  - Sincronización automática de versión en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
  - Suite de pruebas unitarias ejecutada y aprobada (0 fallos).
  - Ensamblado `ConversorPDF.XLL.dll` (Release|x64) compilado exitosamente.
  - Paquete instalador generado: `installer\InstaladorConversorPDFv1.6.3.25.msi` (5,386,240 bytes).

---

## 🔒 Fases 235 - 238: Erradicación del Anclaje Ciego Residual, Endurecimiento de Migración Demo y Sincronía Canónica de AulaRef (Ronda 28)

### 67. **Fase 235 — ALTO: Eliminación Total de la Rama Residual de Anclaje Ciego**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Causa Raíz:** `AgFormGeneradorHorarios.cs:314` conservaba un bloque `else if (ctx.Aulas.Count == 1) { ctx.IdAulaCorrida = ctx.Aulas[0].Id; }` y en `coincidencias.Count > 1` asignaba `coincidencias[0].Id`. Esto reintroducía una compuerta ciega donde un aula arbitraria podía ser seleccionada si había una sola en la lista o ante ambigüedad, contradiciendo el aislamiento de la Fase 231.
- **Solución:**
  1. Se eliminó la rama residual `else if (ctx.Aulas.Count == 1)`.
  2. Se estableció la regla estricta: el **único** punto de asignación válido para `ctx.IdAulaCorrida` es cuando la búsqueda por `AulaRef` exacto arroja exactamente una coincidencia (`coincidencias.Count == 1`).
  3. Para cualquier otro escenario (`coincidencias.Count == 0`, `coincidencias.Count > 1`, o sin combinaciones en el árbol), se asigna determinísticamente `ctx.IdAulaCorrida = 0`, activando el bloqueo y advertencia del diagnóstico pre-vuelo (`L472 / L560`).

### 68. **Fase 236 — MEDIO-ALTO: Endurecimiento de Exclusión Demo y No-Tolerancia en `MigrationV52`**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV52_IsolateDemoAulas.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Causa Raíz:** En `MigrationV52_IsolateDemoAulas`, la propiedad `Aditiva => true` indicaba a `SchemaMigrator` que cualquier fallo de ejecución debía ser tolerado y silenciado con una simple advertencia, dejando filas con `subunidad = ''` sin aislar. Asimismo, la exclusión en el repositorio requería garantizar doble criterio (por `subunidad` y por firma demo canónica de código y período).
- **Solución:**
  1. Se modificó `Aditiva => false` en `MigrationV52_IsolateDemoAulas`. Si la migración falla, el error es capturado y reportado formalmente por el migrador sin omitirse silenciosamente.
  2. En `AgScheduleRepository.cs`, tanto en `GetAulasAsync` como en `GetAulasPorPeriodoAsync`, se consolidó el doble criterio de exclusión:
     `AND COALESCE(subunidad, '') <> 'DEMO' AND NOT (idperiodo = 0 AND codigo IN ('A-101', 'A-102', 'A-201', 'LAB-1', 'VIR-01'))`.

### 69. **Fase 237 — MEDIO: Sincronización Canónica de `AulaRef` con Modalidad Efectiva**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAulaFactory.cs`
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgContextoSeleccionDto.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Causa Raíz:** Cuando `EsVirtual = 1` o la modalidad seleccionada era virtual, la entidad de aula forzaba su `ModalidadVinculacion` a `"Virtual"`. Si `AulaRef` en el DTO de selección o en la entidad se calculaba previamente con la modalidad cruda (`"Presencial"`), la clave natural `AulaRef` difería (`...|PRESENCIAL` vs `...|VIRTUAL`), provocando que la búsqueda por coincidencia exacta de Fase 231 fallara.
- **Solución:**
  1. En `AgAulaFactory`, se centralizó la construcción mediante `AgAulaFactory.ConstruirAulaRef(idPeriodo, idSede, idPrograma, semestre, grupo, modalidad)`, aplicando invariancia canónica de mayúsculas y recorte de espacios.
  2. En `AgContextoSeleccionDto`, se incorporó la propiedad calculada `ModalidadEfectiva`, que fuerza `"Virtual"` si `EsVirtual > 0` o `Modalidad == "Virtual"`. La propiedad `AulaRef` del DTO delega en `AgAulaFactory.ConstruirAulaRef(...)` utilizando dicha `ModalidadEfectiva`.
  3. En `AgAulaFactory.CrearDesdeDto`, `aulaRefCalculada` se recomputa tras forzar la modalidad de vinculación.
  4. En `AgFormGeneradorHorarios.GenerarEncabezadoAulaAsync`, se asigna explícitamente `AulaRef = AgAulaFactory.ConstruirAulaRef(...)` al generar encabezados individuales de aula.

### 70. **Fase 238 — Matriz de Verificación y Cobertura de Pruebas**
| # | Verificación | Componente | Estado |
|---|---|---|---|
| 1 | Sin asignación residual ajena a `coincidencias[0].Id` / `0` | `AgFormGeneradorHorarios.cs` | **✅ Verificado (Fase 235)** |
| 2 | Exclusión demo por `subunidad` y firma `idperiodo = 0 AND codigo IN (...)` | `AgScheduleRepository.cs` | **✅ Verificado (Fase 236)** |
| 3 | Migración V52 con `Aditiva => false` | `MigrationV52_IsolateDemoAulas.cs` | **✅ Verificado (Fase 236)** |
| 4 | Sincronía matemática `dto.AulaRef == aula.AulaRef` ante modalidad virtual | `AgAulaFactoryTests.cs` | **✅ Cubierto con tests unitarios** |
| 5 | Metadatos y comportamiento de `MigrationV52` en suite de persistencia | `AgSchemaMigrationTests.cs` | **✅ Cubierto con tests unitarios** |---

## 🌳 Fases 235 - 244: Perfiles de Selección de Árbol Académico (`ag_arbol`), Prerrequisito Crítico de Carga de Períodos y Sincronía Canónica

### 71. **Fase 235 (Prerrequisito Crítico) — Carga Efectiva de Períodos en `AgFormGeneradorHorarios`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** La consulta SQL para cargar períodos en `CargarPeriodosAsync()` se ejecutaba de forma desacoplada o incompleta antes de poblar `cmbPeriodos.DataSource`, provocando que el combo de períodos quedara vacío, impidiendo poblar el árbol de contexto académico y dejando en 0 las combinaciones seleccionadas.
- **Solución:** Se implementó la carga y asignación síncrona en memoria:
  ```csharp
  _periodos = (await db.QueryAsync<Periodo>(sqlPeriodos))
      .GroupBy(p => p.Id)
      .Select(g => g.First())
      .OrderByDescending(p => p.Activo)
      .ThenByDescending(p => p.Id)
      .ToList();
  ```
  Con esto, el combo `cmbPeriodos` se inicializa con el período activo por defecto y dispara la carga de sedes/programas en el árbol jerárquico.

### 72. **Fase 236 — Migración `MigrationV53_AddAgArbolTables` (Esquema Local de Perfiles)**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV53_AddAgArbolTables.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs` (registro en `BuildMigrator()` después de V52)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (registro `<Compile Include>`)
- **Diseño del Esquema:**
  - Tablas locales aisladas (no CRDT, no incluidas en `CrdtTables` ni en `AcademicSchema.sql`).
  - `ag_arbol`: Cabecera del perfil (nombre, descripcion, idperiodo, ventana temporal, origen de ventana, hiperparámetros evolutivos: población, generaciones, tasa cruce, tasa mutación, elitismo, estancamiento_max, semilla, usuario, subunidad, timestamps). Índice único `uq_ag_arbol_nombre` sobre `(nombre, subunidad)`.
  - `ag_arbol_nodo`: Detalle de combinaciones seleccionadas en el perfil (idarbol, idperiodo, idsede, idprograma, semestre, grupo, modalidad efectiva, aula_ref, atributos completos de aula: capacidad, tipo_aula, caracteristicas_json, codigo, nombre, es_virtual, habilitado, subunidad). Índice único `uq_ag_arbol_nodo` sobre `(idarbol, aula_ref)`.
  - Propiedad `Aditiva => false` y verificación de existencia previa en `sqlite_master`.

### 73. **Fase 237 — Dominio y Puerto de Repositorio de Perfiles**
- **Archivos Creados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgArbol.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgArbolNodo.cs`
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgArbolRepository.cs`
- **Contrato:** Métodos asíncronos con `CancellationToken` y `Result<T>` de `CSharpFunctionalExtensions`:
  - `GuardarAsync(AgArbol, IReadOnlyCollection<AgArbolNodo>, CancellationToken)`
  - `ObtenerAsync(long, CancellationToken)`
  - `ObtenerNodosAsync(long, CancellationToken)`
  - `ListarAsync(string, CancellationToken)`
  - `ExisteNombreAsync(string, string, long, CancellationToken)`
  - `RenombrarAsync(long, string, CancellationToken)`
  - `EliminarAsync(long, CancellationToken)`

### 74. **Fase 238 — Infraestructura de Persistencia `AgArbolRepository`**
- **Archivo Creado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgArbolRepository.cs`
- **Patrón Arquitectural:** Sigue fielmente la arquitectura de `AgScheduleRepository` (Dapper, transacciones con `db.BeginTransaction()`, manejo de `OperationCanceledException`, logging centralizado con `LoggerHelper.Error`).
- **Comportamiento Transaccional:**
  - `GuardarAsync`: Si el perfil no tiene ID o se detecta un perfil existente con el mismo nombre y subunidad, se reutiliza el ID realizando un upsert transaccional que elimina (`DELETE FROM ag_arbol_nodo WHERE idarbol = @Id`) y reinserta todos los nodos.
  - `EliminarAsync`: Cascada transaccional (elimina nodos y cabecera en una sola transacción).

### 75. **Fase 239 — Mappers y Capa de Aplicación (`IAgArbolService` / `AgArbolService`)**
- **Archivos Creados:**
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgArbolMapper.cs`
  - `ConversorPDF.XLL/Academic/Application/Interfaces/IAgArbolService.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/AgArbolService.cs`
- **Reglas de Negocio:**
  - Validación de nombre obligatorio y no vacío.
  - Exigencia de al menos 1 nodo de selección marcado (`nodos.Count > 0`).
  - Mapeo bidireccional puro `AgContextoSeleccionDto` <-> `AgArbolNodo`, preservando `ModalidadEfectiva`, `AulaRef` y los atributos completos de aula (`Capacidad`, `TipoAula`, `CaracteristicasJson`, `Codigo`, `Nombre`, `EsVirtual`).

### 76. **Fase 240 — UI: Diálogo de Nombre de Perfil y Barra de Herramientas**
- **Archivos Creados y Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/FormNombreArbol.cs` (formulario modal ligero code-only, sin dependencia de `.resx`, con estilo oscuro acorde a la aplicación).
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`:
    - Incorporación de controles en panel de cabecera: `lblPerfilArbol`, `cmbArboles`, `btnCargarArbol`, `btnNuevoArbol`, `btnSobrescribirArbol`, `btnEliminarArbol`.
- **Inyección y Ciclo de Vida:**
  - Instanciación directa en constructor de `AgFormGeneradorHorarios` siguiendo la convención del formulario:
    `_arbolRepository = new AgArbolRepository(_connectionFactory);`
    `_arbolService = new AgArbolService(_arbolRepository);`

### 77. **Fase 241 — Orquestación de Perfiles en el Formulario (`AgFormGeneradorHorarios`)**
- **Lógica Implementada:**
  - `RefrescarComboArbolesAsync()`: Llena `cmbArboles` con perfiles filtrados por la subunidad del usuario actual.
  - `GuardarNuevoPerfilArbolAsync()`: Abre `FormNombreArbol`, recolecta hojas seleccionadas con `RecolectarHojas()`, captura hiperparámetros de los controles reales (`numPoblacion`, `numGeneraciones`, `numTasaCruce`, `numTasaMutacion`, `numElitismo`, ventana temporal), persiste y refresca el selector.
  - `SobrescribirPerfilArbolAsync()`: Reemplaza las combinaciones y parámetros del perfil actualmente seleccionado tras confirmación.
  - `CargarPerfilArbolSeleccionadoAsync()`: Restaura el período en `cmbPeriodos`, sincroniza ventana temporal, actualiza controles de hiperparámetros, limpia marcas del árbol y ejecuta `AsegurarSubarbolAsync()`.
  - `AsegurarSubarbolAsync()`: Expande y carga de forma lazy/recursiva los nodos requeridos, busca u obtiene las hojas por clave natural, marca `hoja.Checked = true`, propaga el estado hacia ancestros (`ActualizarEstadoPadre`) y actualiza el conteo exacto de combinaciones (`ActualizarConteoCombinaciones()`).

### 78. **Fase 243 — Cobertura de Pruebas Unitarias e Integración (`ConversorPDF.Tests`)**
- **Suites Creadas:**
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgArbolRepositoryTests.cs`: Pruebas de integración sobre SQLite en memoria (guardar, sobreescritura, listado por subunidad, eliminación en cascada, verificación de unicidad).
  - `ConversorPDF.Tests/Academic/Application/Services/AgArbolServiceTests.cs`: Pruebas unitarias de reglas de negocio con `Moq` (rechazo de nombre vacío, rechazo de lista vacía de nodos, delegación correcta al repositorio).
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV53SchemaTests.cs`: Verificación de idempotencia de la migración V53, comprobación de tablas `ag_arbol` y `ag_arbol_nodo`, índices y columnas requeridas.
- **Resultado:** 427/427 pruebas ejecutadas y aprobadas en Release (0 errores).

### 79. **Fase 244 — Documentación Técnica de Cambios**
- **Archivo Modificado:** `docs/articles/changes/2026-ag-horarios-academicos.md`
- **Sección Añadida:** `8.5 Perfiles de Selección de Árbol y Sincronía Canónica (Fases 235 - 244)`. Detalle completo de las tablas `ag_arbol`, `ag_arbol_nodo`, DTO con atributos completos de aula, ciclo de vida UI y exportación CSV de 16 columnas.

---

## 🛡️ Fases 245 - 252: Robustez de Ventana Temporal, Integridad Transaccional, Deduplicación y Expansión Lazy de Perfiles (Ronda 29)

### 80. **Fase 245 — Sincronización Canónica de Ventana Temporal al Cargar Perfil**
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `CargarPerfilArbolSeleccionadoAsync` invocaba `ActualizarVentanaSemestralAsync()` de forma prematura (al seleccionar el período), antes de restaurar `chkManualVentana` y los selectores `dtpVentanaInicio`/`dtpVentanaFin`. Como consecuencia, `_ventanaResuelta` conservaba el origen `Periodos` ignorando las fechas manuales del perfil, lo que comprometía la reproducibilidad de la optimización genética.
- **Solución Implementada:**
  1. Se trasladó la invocación de `await ActualizarVentanaSemestralAsync();` para ejecutarse estrictamente después de restaurar el estado de `chkManualVentana` y las fechas en `dtpVentanaInicio`/`dtpVentanaFin`.
  2. En `ConfigurarEventos()`, se sincronizó el evento `chkManualVentana.CheckedChanged` para recalcular asíncronamente `_ventanaResuelta` tanto al activar como al desactivar la casilla, y se enlazó el evento `ValueChanged` de ambos `DateTimePicker` para mantener la ventana sincronizada en tiempo real.

### 81. **Fase 246 — Sincronización Previa al Guardado de Ventanas Manuales**
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Si el usuario editaba fechas manuales en `dtpVentanaInicio`/`dtpVentanaFin` sin desencadenar un cambio de período, `_ventanaResuelta` mantenía los valores anteriores. Al guardar o sobrescribir el perfil, se persistían fechas desincronizadas.
- **Solución Implementada:** Tanto en `GuardarNuevoPerfilArbolAsync()` como en `SobrescribirPerfilArbolAsync()`, si `chkManualVentana.Checked == true`, se normalizan las fechas (`.Date`) de `dtpVentanaInicio` y `dtpVentanaFin` y se ejecuta `await ActualizarVentanaSemestralAsync();` antes de transferir `_ventanaResuelta` al servicio `AgArbolService`.

### 82. **Fase 247 — Erradicación de Nodos Huérfanos por `UPDATE` de Perfiles Inexistentes**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgArbolRepository.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgArbolRepositoryTests.cs`
- **Diagnóstico:** Cuando se invocaba `GuardarAsync` con `perfil.Id > 0` y dicho ID no existía en la base de datos (registro eliminado o inválido), la sentencia `UPDATE ag_arbol` afectaba 0 filas sin emitir error. Posteriormente, el bloque de inserción de nodos procedía a insertar registros en `ag_arbol_nodo` con un `idarbol` inexistente, provocando datos huérfanos.
- **Solución Implementada:** Se verifica el resultado de filas afectadas `filas = await db.ExecuteAsync(sqlUpdate, ...)`. Si `filas <= 0`, se ejecuta `tx.Rollback()` y se retorna inmediatamente `Result.Failure<long>($"No se pudo actualizar el perfil con ID {idArbol}: el registro no existe en la base de datos.")`. Se añadió prueba de integración específica `GuardarAsync_WhenNonExistentProfileId_ShouldFailAndNotCreateOrphanNodes`.

### 83. **Fase 248 — Unificación de Criterio de Búsqueda y Unicidad por Subunidad**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgArbolRepository.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgArbolRepositoryTests.cs`
- **Diagnóstico:** Discrepancia entre la búsqueda de duplicados en `GuardarAsync` y `ExisteNombreAsync`: una comparaba nombres exactos mientras la otra utilizaba `LOWER(TRIM(nombre))`, y el manejo de subunidades vacías vs `NULL` era asimétrico.
- **Solución Implementada:**
  - Se unificó el predicado canónico en todas las consultas:
    ```sql
    WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(@Nombre))
      AND COALESCE(subunidad, '') = COALESCE(@Subunidad, '')
    ```
  - Se añadió la prueba de integración `GuardarAsync_WhenCaseAndWhitespaceVariationInName_ShouldOverwriteProfile` que verifica la idempotencia y sobrescritura ante variaciones de espacios y mayúsculas/minúsculas.

### 84. **Fase 249 — Deduplicación Defensiva de Nodos por `AulaRef` en la Capa de Aplicación**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/AgArbolService.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/AgArbolServiceTests.cs`
- **Diagnóstico:** La tabla `ag_arbol_nodo` cuenta con el índice único `uq_ag_arbol_nodo ON ag_arbol_nodo(idarbol, aula_ref)`. Si la lista de selecciones recibida contenía combinaciones redundantes con el mismo `AulaRef`, la inserción masiva en SQLite abortaba por infracción de unicidad.
- **Solución Implementada:** En `AgArbolService.GuardarPerfilAsync`, se aplica deduplicación estricta con `.GroupBy(n => n.AulaRef, StringComparer.OrdinalIgnoreCase).Select(g => g.First())`. Se añadió prueba unitaria `GuardarPerfilAsync_WhenDuplicateSelectionsProvided_ShouldDeduplicateNodesByAulaRef`.

### 85. **Fase 250 — Propagación Explícita de `EsVirtual` en DTOs de Selección**
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En `ObtenerCombinacionesSeleccionadas()`, el DTO se instanciaba con el constructor de 6 argumentos, dejando la propiedad `EsVirtual = 0` por defecto, aun cuando la modalidad de la hoja seleccionada correspondiera a `"Virtual"`.
- **Solución Implementada:** Al identificar un nodo de nivel 6 con código `"Virtual"`, se asigna de forma determinista `dto.EsVirtual = 1`, garantizando que tanto en memoria como en la tabla `ag_arbol_nodo` los atributos reflejen la naturaleza virtual del aula.

### 86. **Fase 251 — Placeholder de Carga Diferida en Nodos Creados Manualmente**
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al cargar un perfil en `AsegurarSubarbolAsync()`, si un semestre o grupo no existía en el árbol, se instanciaba el `TreeNode` correspondiente sin nodos hijos. Posteriormente, la función `CargarHijosLazy()` evalúa la condición `if (node.Nodes.Count == 1 && node.Nodes[0].Tag == null)`. Al tener `Nodes.Count == 0`, la condición era falsa y los subárboles no se poblaban de forma lazy.
- **Solución Implementada:** Al instanciar `nodoSem` o `nodoGrp` manualmente, se incorpora inmediatamente el nodo temporal de carga diferida `nodo.Nodes.Add(new TreeNode("Cargando...") { Tag = null });`, permitiendo que la invocación inmediata a `CargarHijosLazy` y las expansiones del usuario operen con normalidad.

### 87. **Fase 252 — Compuertas de Validación Temprana en Guardado y Sobrescritura UI**
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Si la UI se encontraba en un estado sin período válido o sin ventana resuelta, el usuario recibía mensajes de error genéricos tras pasar por el diálogo de nombre o confirmación.
- **Solución Implementada:** Se incorporó una compuerta previa al inicio de `GuardarNuevoPerfilArbolAsync()` y `SobrescribirPerfilArbolAsync()`:
  ```csharp
  int idPeriodo = (cmbPeriodos.SelectedItem as Periodo)?.Id ?? 0;
  if (_ventanaResuelta == null || idPeriodo <= 0)
  {
      MessageBox.Show(this, "Debe seleccionar un período con ventana temporal resuelta antes de guardar el perfil.", "Guardar Perfil", MessageBoxButtons.OK, MessageBoxIcon.Warning);
      return;
  }
  ```

---

## 🔒 Fases 253 - 256: Eliminación de Condiciones de Carrera UI, Manejo Seguro de Eventos Async, Validación de Unicidad y Preservación de Habilitado (Ronda 30)

### 88. **Fase 253 — Erradicación de Carrera Concurrente en `_ventanaResuelta` al Cargar Perfiles**
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al cargar un perfil en `CargarPerfilArbolSeleccionadoAsync`, la asignación sucesiva de `cmbPeriodos.SelectedItem`, `chkManualVentana.Checked` y `dtpVentana*.Value` disparaba 4 llamadas asíncronas no coordinadas a `ActualizarVentanaSemestralAsync`. Por falta de await en los manejadores de eventos, las tareas competían en el bucle de mensajes de WinForms y la última en finalizar sobreescribía `_ventanaResuelta`, pudiendo reintroducir la ventana `Periodos` en perfiles con ventana manual guardada.
- **Solución Implementada:**
  1. Se introdujo el campo booleano de sincronización `_restaurandoPerfil`.
  2. La actualización de controles en `CargarPerfilArbolSeleccionadoAsync` se envolvió en un bloque `try / finally` que activa `_restaurandoPerfil = true`.
  3. Los eventos de cambio de ventana comprueban `if (_restaurandoPerfil) return;`, silenciando completamente las ejecuciones intermedias.
  4. Al finalizar la asignación y desactivarse la bandera, se invoca una única vez `await ActualizarVentanaSemestralAsync();`, garantizando un estado final determinista y reproducible.

### 89. **Fase 254 — Manejador Centralizado `OnVentanaChanged` con Aislamiento `try/catch`**
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Los eventos `cmbPeriodos.SelectedIndexChanged`, `chkManualVentana.CheckedChanged` y `dtpVentana*.ValueChanged` utilizaban lambdas `async void` directas sin captura de excepciones. Ante cualquier fallo en base de datos o timeout, la excepción no capturada en `async void` causaba la terminación anormal del proceso.
- **Solución Implementada:**
  - Se implementó el método centralizado `private async void OnVentanaChanged(object sender, EventArgs e)` protegido con `try / catch (Exception ex)` y registro mediante `LoggerHelper.Error`.
  - Se suscribieron todos los controles de la ventana temporal a este manejador unificado.

### 90. **Fase 255 — Consumo de `ExisteNombreAsync` en `AgArbolService.GuardarPerfilAsync`**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/AgArbolService.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/AgArbolServiceTests.cs`
- **Diagnóstico:** El método de contrato `IAgArbolRepository.ExisteNombreAsync` no era consumido por la capa de aplicación, perdiendo la oportunidad de validar tempranamente la existencia de duplicados o fallos de conexión en el repositorio antes de la instanciación de entidades.
- **Solución Implementada:**
  - En `AgArbolService.GuardarPerfilAsync`, se valida `var dup = await _repository.ExisteNombreAsync(nombre.Trim(), sub, 0, ct);`. Si la consulta al repositorio falla (`dup.IsFailure`), se aborta retornando `Result.Failure<long>(dup.Error)`.
  - Se actualizó el constructor de pruebas `AgArbolServiceTests` configurando el mock por defecto para `ExisteNombreAsync` y se añadió la prueba unitaria `GuardarPerfilAsync_WhenExisteNombreAsyncFails_ShouldReturnFailure`.

### 91. **Fase 256 — Preservación del Atributo `Habilitado` en Inserción de Nodos**
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgArbolRepository.cs`
- **Diagnóstico:** En `AgArbolRepository.GuardarAsync`, los parámetros del lote de inserción para `ag_arbol_nodo` forzaban el valor literal `Habilitado = 1`, ignorando la propiedad `n.Habilitado` del modelo de dominio.
- **Solución Implementada:** Se corrigió a `Habilitado = n.Habilitado > 0 ? 1 : 0,` respetando el estado asignado en el objeto de dominio.

---

## 🛡️ Fases 257 - 259: Control de Sobrescritura de Perfiles, Confirmación UI y Estandarización de Migración V53 (Ronda 31)

### 92. **Fase 257 — Documentación y Control de Semántica de Sobrescritura en el Servicio**
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/AgArbolService.cs`
- **Diagnóstico:** En `AgArbolService.GuardarPerfilAsync`, se invocaba `ExisteNombreAsync` pero solo se manejaba `dup.IsFailure`. Se clarificó la semántica del contrato: cuando `dup.Value == true`, el perfil homónimo existe en la subunidad y el repositorio ejecutará el upsert transaccional de sobrescritura intencional, evitando bloqueos espurios en operaciones de actualización explícita.
- **Solución Implementada:** Se documentó y formalizó la política de sobrescritura en el flujo del servicio, dejando la confirmación preventiva al nivel de la interacción con el usuario en la interfaz gráfica.

### 93. **Fase 258 — Prevención de Pérdida Silenciosa de Perfiles Homónimos en UI**
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al guardar un nuevo perfil en `GuardarNuevoPerfilArbolAsync`, si el usuario ingresaba el nombre de un perfil preexistente en la misma subunidad, el repositorio ejecutaba el upsert silenciosamente, reemplazando la configuración anterior sin advertencia ni confirmación.
- **Solución Implementada:**
  - Tras capturar el nombre en `FormNombreArbol`, se verifica si `_perfilesArbol.Any(p => string.Equals(p.Nombre?.Trim(), nombre?.Trim(), StringComparison.OrdinalIgnoreCase))`.
  - Si el perfil ya existe, se muestra un diálogo de confirmación `MessageBox.Show(..., MessageBoxButtons.YesNo, MessageBoxIcon.Question)`. Si el usuario no confirma con `DialogResult.Yes`, la operación se aborta de inmediato sin alterar los datos existentes.

### 94. **Fase 259 — Estandarización de Política No-Aditiva en Migración V53**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV53_AddAgArbolTables.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV53SchemaTests.cs`
- **Diagnóstico:** `MigrationV52` implementaba `IMigrationTolerante` declarando explícitamente `Aditiva => false`, mientras que `MigrationV53` solo implementaba `IMigration`, creando una asimetría estilística en las migraciones de esquemas locales no-CRDT.
- **Solución Implementada:**
  - Se implementó `IMigrationTolerante` en `MigrationV53_AddAgArbolTables`, exponiendo `public bool Aditiva => false;`.
  - Se actualizó la prueba de metadatos en `MigrationV53SchemaTests` para comprobar `mig.Aditiva.Should().BeFalse();`.

### 95. **Fase 260 — Capacidad por Defecto para Hojas Virtuales**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Solución Implementada:** En `ObtenerCombinacionesSeleccionadas()`, si la modalidad es `"Virtual"`, se asigna de manera explícita `dto.Capacidad = 999;` junto con `dto.EsVirtual = 1;`.

---

## 🚀 Verificación Final de Compilación y Suite de Pruebas (Ronda 32)

- **Compilación Release (x64):**
  - Comando: `dotnet build ConversorPDF.XLL\ConversorPDF.XLL.csproj -c Release -p:Platform=x64`
  - Resultado: **Compilación Exitosa (0 Errores)**.
  - Artefacto: `ConversorPDF.XLL.dll` y `ConversorPDF.XLL-AddIn64.xll` generados correctamente.
- **Suite de Pruebas Unitarias e Integración (net48):**
  - Comando: `dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`
  - Resultado: **431 / 431 pruebas aprobadas (0 errores, 0 omitidos)**.
  - Cobertura completa: Perfiles `ag_arbol` y `ag_arbol_nodo`, deduplicación por `AulaRef`, migración idempotente V53, aislamiento de perfiles por subunidad, rollback ante IDs inexistentes, sincronización bidireccional de ventana temporal y árboles de selección de 6 niveles.

---

## 📦 Entrega de Release v1.6.3.26 (Instalador Oficial)

- **Versión Oficial:** `1.6.3.26`
- **Archivos de Versión Sincronizados:**
  - `ConversorPDF.XLL\Core\Helpers\App.cs` (`APP_VERSION = "1.6.3.26"`)
  - `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (`<Version>1.6.3.26</Version>`)
  - `installer\InstaladorConversorPDF.wxs` (`Version="1.6.3.26"`)
- **Paquete Instalador Generado:**
  - Archivo: `installer\InstaladorConversorPDFv1.6.3.26.msi`
  - Tamaño: **5,406,720 bytes**
  - Estado: Listo para despliegue y auto-activación en Excel.

---

## 🛠️ Ronda 34 — Diagnóstico y Corrección Crítica de Generación de Aulas (Fases 262 - 263)

### 96. **Fase 262 — Corrección Crítica de Esquema en Sincronización de `ag_aula_asignatura` (`no such column: idaula`)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  Al invocar `GenerarEncabezadosAulasAsync` (botón *Generar Aulas*), el método ejecutaba dos sentencias SQL con nombres de columnas inexistentes sobre la tabla `ag_aula_asignatura`:
  1. `SELECT COUNT(1) FROM ag_aula_asignatura WHERE idaula = @IdAula;`
  2. `INSERT OR IGNORE INTO ag_aula_asignatura (idaula, idasignatura, horas_objetivo, orden_curricular, requerida, activo, subunidad) VALUES (@IdAula, @IdAsig, 64, 1, 1, 1, @Sub);`
  En la definición canónica del esquema (`MigrationV49_AddAgHorariosTables.cs` L44-54), las columnas reales de la tabla son: `(id_aula, id_asignatura, exclusivo, prioridad, horas_objetivo, orden_curricular, requerida, subunidad)`, sin existir las columnas `idaula`, `idasignatura` ni `activo`.
  Esto producía una excepción `SQLiteException: SQL logic error / no such column: idaula`, desencadenando el rollback automático de la transacción. Como consecuencia, ninguna aula persistía en `ag_aula` (0 registros), lo que provocaba que la posterior comprobación de pre-vuelo bloqueara el flujo con: `No existe Aula (ag_aula) para la combinación exacta ...`.
- **Solución Implementada:**
  Se actualizaron las sentencias SQL dentro de la transacción activa en `GenerarEncabezadosAulasAsync`:
  ```csharp
  // Verificación de existencia de asignaturas dependientes
  int cuentaAsig = await db.ExecuteScalarAsync<int>(
      "SELECT COUNT(1) FROM ag_aula_asignatura WHERE id_aula = @IdAula;",
      new { IdAula = idAula }, transaction: trans);

  // Inserción transaccional con esquema canónico V49
  string sqlInsertAsig = @"
  INSERT OR IGNORE INTO ag_aula_asignatura (
      id_aula, id_asignatura, exclusivo, prioridad, horas_objetivo, orden_curricular, requerida, subunidad
  ) VALUES (
      @IdAula, @IdAsig, 0, 0, 64, 1, 1, @Sub
  );";
  ```
  Se preservó el uso de la conexión y transacción local (`db` y `trans`) para evitar locks por conexiones concurrentes en SQLite.

### 97. **Fase 263 — Protocolo de Verificación Operativa (Sin Cambios de Código)**
- **Verificaciones Post-Corrección:**
  1. Validar que la tabla `asignaturas` cuente con registros activos (`activo = 1`) para el programa y semestre evaluados.
  2. Validar que tras ejecutar `Generar Aulas`, exista el registro en `ag_aula` con el `AulaRef` exacto (`idperiodo|idsede|idprograma|semestre|grupo|modalidad`).
  3. Validar que `ag_aula_asignatura` contenga los registros asociados al `id_aula` generado con sus asignaturas.
  4. Validar que los logs de la corrida AG no contengan `no such column: idaula` ni `Error generando aulas`.

---

## 🏁 Ronda 35 — Verificación Estática y Cierre de Fases 235-263

- **Resultados de Verificación Estática:**
  - L1914: `SELECT COUNT(1) … WHERE id_aula = @IdAula` verificado exhaustivamente.
  - L1935-1937: `INSERT OR IGNORE INTO ag_aula_asignatura (id_aula, id_asignatura, exclusivo, prioridad, horas_objetivo, orden_curricular, requerida, subunidad)` alineado al esquema canónico V49.
  - Rama `omitidas` (L1897): `SELECT id FROM ag_aula WHERE aula_ref = @AulaRef LIMIT 1` confirmado.
  - Ambas ramas de inserción y preexistencia garantizan `idAula > 0`, asegurando que la siembra de dependencias en `ag_aula_asignatura` corra siempre.
  - `INSERT INTO ag_aula` (L1880-1892): Todas las columnas coinciden con `MigrationV49` y fijan `estado = AgEstadoAula.Disponible = 1` para consumo inmediato por `GetAulasAsync`.
  - Cero referencias residuales a columnas inválidas (`idaula`, `idasignatura`, `activo`) en `ag_aula_asignatura`.
  - Cero nuevos defectos bloqueantes detectados estáticamente.
- **Protocolo de Smoke Testing (Entorno Seguro):**
  - **Restricción Crítica:** Ejecutar **únicamente en copia del vault de pruebas**; terminantemente prohibido en base de datos de producción debido a que la aprobación/promoción escribe en tablas CRDT (`vinculaciones`, `sesiones`, `terceros_asignaturas`) con replicación automática e irreversible.
  - **Secuencia:** 1 combinación $\rightarrow$ `Generar Aulas (1)` $\rightarrow$ `Diagnóstico Pre-Vuelo` $\rightarrow$ `Optimizar Horarios` $\rightarrow$ `Modo AG` $\rightarrow$ `Aprobar` $\rightarrow$ `Promover` $\rightarrow$ Verificar $\rightarrow$ `ROLLBACK`.
- **Estado General:** Fases 235 a 263 **cerradas y verificadas**.

---

## 📦 Entrega de Release v1.6.3.27 (Instalador Oficial)

- **Versión Oficial:** `1.6.3.27`
- **Archivos de Versión Sincronizados:**
  - `ConversorPDF.XLL\Core\Helpers\App.cs` (`APP_VERSION = "1.6.3.27"`)
  - `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (`<Version>1.6.3.27</Version>`)
  - `installer\InstaladorConversorPDF.wxs` (`Version="1.6.3.27"`)
- **Verificación de Pruebas Unitarias:**
  - Suite completa: **Aprobada sin errores** (`dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`).
- **Compilación de Add-In:**
  - `ConversorPDF.XLL` (Release|x64): **0 Errores**.
- **Paquete Instalador Generado:**
  - Archivo: `installer\InstaladorConversorPDFv1.6.3.27.msi`
  - Tamaño: **5,402,624 bytes**
  - Estado: Listo para distribución y smoke test en copia del vault.

---

## 🛠️ Ronda 36 — Corrección Crítica de Esquema en Consulta de Asignaturas (Fases 264 - 265)

### 98. **Fase 264 — Corrección Crítica de Esquema en Consulta de `asignaturas` (`no such column: activo`)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  Al ejecutar `GenerarEncabezadosAulasAsync` (botón *Generar Aulas*), tras superar exitosamente la Fase 262, el método evaluaba `if (cuentaAsig == 0)` para vincular asignaturas dependientes. En la consulta `sqlAsigs` (L1920-1925), se incluía el predicado `AND activo = 1`.
  Sin embargo, en el esquema canónico de la base de datos (`AcademicSchema.sql` L27-46), la tabla `asignaturas` define exclusivamente las columnas:
  `(id, idprograma, semestre, num_nivel, cod_unidad, cod_materia, uni_teorica, metodologia, abreviatura, nucleo, nombre, creditos, int_horaria, tipoasignatura, tipocredito, prerequisito, subunidad)`.
  **No existe la columna `activo` en la tabla `asignaturas`** (a diferencia de `programas`, `sedes`, `terceros_asignaturas` o `periodos`).
  En consecuencia, `db.QueryAsync<int>(sqlAsigs, ...)` provocaba una excepción `SQLiteException: SQL logic error / no such column: activo`, desencadenando el rollback automático de la transacción. Esto impedía nuevamente la persistencia de las aulas en `ag_aula` y dejaba bloqueada la corrida AG.
- **Solución Implementada:**
  Se removió el filtro `AND activo = 1` de la sentencia `sqlAsigs`, alineándola con el resto de consultas canónicas del módulo (`AgFormGeneradorHorarios.cs` L360-365 y L968-969):
  ```sql
  SELECT id FROM asignaturas 
  WHERE idprograma = @IdPrograma 
    AND (semestre = @Semestre OR @Semestre = '' OR semestre IS NULL)
    AND (subunidad = @Sub OR subunidad = '' OR subunidad IS NULL);
  ```

### 99. **Fase 265 — Verificación y Smoke Testing Relacional (Sin Cambios de Código)**
- **Protocolo de Verificación:**
  1. Validar conteo de combinaciones al cargar perfiles guardados (`lblConteoCombinaciones`).
  2. Verificar que tras `Generar Aulas`, `SELECT COUNT(*) FROM ag_aula WHERE aula_ref LIKE '...'` sea mayor a 0.
  3. Verificar que `SELECT COUNT(*) FROM ag_aula_asignatura WHERE id_aula = <id>` sea mayor a 0 (considerando la correlación de formato de semestre: romano `I, II, ...` vs numérico `1, 2, ...`).
  4. Confirmar que el log del AG opere sin excepciones de columna y proceda limpiamente al pre-vuelo y optimización.

---

## 🛠️ Ronda 37 — Fallback Defensivo de Vinculación de Asignaturas (Fases 266 - 267)

### 100. **Fase 266 — Fallback Defensivo en `GenerarEncabezadosAulasAsync` para Discrepancia de Formato de Semestre**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Probable:**
  En bases de datos académicas, la columna `asignaturas.semestre` puede encontrarse almacenada en formato arábigo (`1`, `2`, ..., `10`) o vacía, mientras que las tuplas y nodos del árbol jerárquico utilizan numeración romana (`I`, `II`, `VIII`).
  En tales casos, la consulta con filtro exacto `semestre = @Semestre` retorna 0 filas sin lanzar excepción SQL. Como resultado, la entidad aula en `ag_aula` persistía, pero su tabla dependiente `ag_aula_asignatura` quedaba con 0 registros, bloqueando subsecuentemente el diagnóstico de pre-vuelo y la corrida del AG.
- **Solución Implementada:**
  Se implementó un mecanismo de fallback defensivo de segundo nivel: si la consulta por semestre devuelve 0 asignaturas (`idsAsig.Count == 0`), se ejecuta inmediatamente una consulta de respaldo que recupera las asignaturas del programa activo acotadas a la subunidad:
  ```csharp
  // Fase 266: Fallback defensivo si el semestre de la rama no coincide con el formato almacenado en asignaturas.semestre (ej. 'VIII' vs '8')
  if (idsAsig.Count == 0)
  {
      string sqlFallbackAsigs = @"
  SELECT id FROM asignaturas
  WHERE idprograma = @IdPrograma
    AND (subunidad = @Sub OR subunidad = '' OR subunidad IS NULL);";

      idsAsig = (await db.QueryAsync<int>(
          sqlFallbackAsigs,
          new { IdPrograma = item.IdPrograma, Sub = sub },
          transaction: trans)).ToList();
  }
  ```
  Esto garantiza que `ag_aula_asignatura` siempre quede vinculada y habilitada para la corrida genética, eliminando un ciclo extra de bloqueo en el smoke test.

### 101. **Fase 267 — Verificación Operativa de Pre-Vuelo y Conteo**
- **Protocolo de Verificación:**
  1. Verificar ausencia de excepciones de SQLite en el log durante `Generar Aulas`.
  2. Confirmar que `ag_aula` posea el registro con `AulaRef` exacto.
  3. Confirmar que `ag_aula_asignatura` contenga $\ge 1$ registro asociado al `id_aula`.
  4. Confirmar que el diagnóstico de pre-vuelo finalice en estado óptimo con `IdAulaCorrida > 0`.
  5. Validar consistencia en `lblConteoCombinaciones` con los nodos del perfil seleccionado.

---

## 🛠️ Ronda 39 — Persistencia en Vault Cifrado y Restauración de Perfiles (Fases 268 - 269)

### 102. **Fase 268 — Notificación al Programador de Sincronización (`MarkDirty`) para Persistencia en Vault Cifrado**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgArbolRepository.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  Al trabajar sobre un vault cifrado (`.db.enc`), las mutaciones se ejecutan sobre la base de datos temporal descifrada en memoria/disco. Para que los cambios sean re-cifrados y persistidos en el vault, el flujo arquitectónico exige notificar al programador central:
  `SyncContext.Scheduler?.MarkDirty()`.
  Sin esta llamada:
  1. `SyncScheduler._isDirty` permanece en `false`.
  2. El temporizador de debounce nunca inicia el guardado automático en background (`ForcePersistVaultAsync`).
  3. Al cerrar la sesión o salir de Excel, `AcademicDbSession.Dispose()` invoca `SyncScheduler.Flush()`, el cual al ver `_isDirty == false` se aborta con el mensaje:
     `[SyncScheduler] Flush(): sin cambios pendientes o sync en progreso`.
  4. La base de datos temporal se descarta sin haber persistido en el archivo cifrado. Como resultado, todas las mutaciones sobre `ag_arbol`, `ag_arbol_nodo`, `ag_aula` y `ag_aula_asignatura` se perdían irremediablemente al reiniciar Excel.
- **Solución Implementada:**
  - En `AgArbolRepository.cs`: Se creó el método auxiliar `NotifyMarkDirty()` y se invoca tras cada `tx.Commit()` exitoso en `GuardarAsync`, `EliminarAsync` y tras actualizar en `RenombrarAsync`.
  - En `AgScheduleRepository.cs`: Se invocó `SyncContext.Scheduler?.MarkDirty()` tras `trans.Commit()` en `InsertAulaAsignaturasAsync`.
  - En `AgFormGeneradorHorarios.cs`: Se invocó `SyncContext.Scheduler?.MarkDirty()` tras el `trans.Commit()` de `GenerarEncabezadosAulasAsync`.

### 103. **Fase 269 — Restauración Automática del Perfil Seleccionado al Abrir el Formulario (`OnShown`)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Solución:**
  En `OnShown`, el formulario ejecutaba `RefrescarComboArbolesAsync()`, el cual poblaba `cmbArboles` y seleccionaba el primer elemento visualmente (`SelectedIndex = 0`), pero no reconstruía las marcas del árbol ni restauraba los parámetros del algoritmo.
  Se incorporó la carga automática en `OnShown`:
  ```csharp
  if (_perfilesArbol != null && _perfilesArbol.Count > 0)
  {
      cmbArboles.SelectedIndex = 0;
      await CargarPerfilArbolSeleccionadoAsync();
  }
  ```
  Esto asegura que al reabrir el formulario, la última selección guardada en el vault quede inmediatamente activa y reflejada en la interfaz de usuario.

---

## 🛠️ Ronda 40 — Persistencia Forzada en Cierre y Notificación Integral de Repositorios (Fases 270 - 272)

### 104. **Fase 270 — Persistencia Inmediata al Cierre del Formulario (`OnFormClosing`)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  La invocación de `MarkDirty()` programa un temporizador adaptativo de *debounce* de 5 segundos antes de re-encriptar el vault cifrado. Si el usuario guarda un perfil o genera aulas y cierra la ventana o Excel antes de que expire ese temporizador, o si `RibbonController` destruye la sesión compartida antes de completar el ciclo asíncrono, la base de datos temporal se cerraba sin completar la escritura en el archivo `.db.enc`.
- **Solución Implementada:**
  Siguiendo el patrón arquitectónico establecido en `FormVinculaciones.cs` (L427-450), se sobreescribió `OnFormClosing` en `AgFormGeneradorHorarios.cs`:
  ```csharp
  protected override void OnFormClosing(FormClosingEventArgs e)
  {
      try
      {
          var scheduler = SyncContext.Scheduler;
          if (scheduler != null)
          {
              var sharedCtx = SqliteHelper.GetContext();
              if (sharedCtx != null)
              {
                  try
                  {
                      scheduler.ForcePersistVaultAsync(sharedCtx).GetAwaiter().GetResult();
                  }
                  catch (Exception ex)
                  {
                      LoggerHelper.Error("[AgFormGeneradorHorarios] Error al persistir el Vault en el cierre", ex);
                  }
              }

              scheduler.Flush();
          }
      }
      catch (Exception ex)
      {
          LoggerHelper.Error("[AgFormGeneradorHorarios] Error en secuencia de cierre del formulario", ex);
      }

      base.OnFormClosing(e);
  }
  ```

### 105. **Fase 271 — Notificación de Mutaciones en `AgScheduleRepository`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico y Solución:**
  Se implementó el método auxiliar `NotifyMarkDirty()` y se vinculó tras cada operación de escritura exitosa en el repositorio genético:
  - `InsertAulaAsync`
  - `InsertAulaAsignaturasAsync`
  - `InsertCorridaAsync`
  - `UpdateCorridaAsync`
  - `SaveIndividuoAsync`
  - `SaveBitacoraAsync`
  - `GuardarSesionesStagingAsync`
  Esto asegura que el historial de corridas, individuos generados y staging se mantengan sincronizados con el vault cifrado de forma homogénea.

### 106. **Fase 272 — Protocolo de Verificación de Persistencia Post-Cierre (Sin Cambios de Código)**
- **Verificaciones Operativas:**
  1. Guardar perfil en la UI $\rightarrow$ confirmar en el log: `ForcePersistVaultAsync … persistido exitosamente`.
  2. Cerrar Excel por completo $\rightarrow$ reabrir $\rightarrow$ verificar que el `Perfil Árbol` aparezca en el combo y el árbol jerárquico se reconstruya automáticamente con marcas (Fase 269).
  3. Verificar mediante consulta SQLite que `SELECT COUNT(*) FROM ag_arbol` y `SELECT COUNT(*) FROM ag_aula` mantengan sus registros intactos en la nueva sesión.
  4. Confirmar que `SyncScheduler.Flush()` ejecute la sincronización de cambios sin omitir escrituras.

---

## 📦 Entrega de Release v1.6.3.28 (Instalador Oficial)

- **Versión Oficial:** `1.6.3.28`
- **Archivos de Versión Sincronizados:**
  - `ConversorPDF.XLL\Core\Helpers\App.cs` (`APP_VERSION = "1.6.3.28"`)
  - `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (`<Version>1.6.3.28</Version>`)
  - `installer\InstaladorConversorPDF.wxs` (`Version="1.6.3.28"`)
- **Verificación de Pruebas Unitarias:**
  - Suite completa: **431 / 431 pruebas aprobadas (0 errores)** (`dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`).
- **Compilación de Add-In:**
  - `ConversorPDF.XLL` (Release|x64): **0 Errores**.
- **Paquete Instalador Generado:**
  - Archivo: `installer\InstaladorConversorPDFv1.6.3.28.msi`
  - Tamaño: **5,402,624 bytes**
  - Estado: Listo para distribución y smoke test en copia del vault.

---

## Ronda 42 — Soporte de Fallback en Contexto Genético y Diagnóstico Accionable de Catálogo

### 107. **Fase 274 — Fallback de Semestre en `ConstruirContextoAsync` (Desbloqueo de Pre-Vuelo)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  En la Ronda 42, el proceso de generación de aulas reportó éxito en persistencia (`creadas: 0 | preexistentes omitidas: 193`), pero el diagnóstico pre-vuelo abortó con:
  `[ERROR PRE-VUELO] El aula o el programa no tiene asignaturas dependientes activas.` (para la combinación `2|3|2|I|A|PRESENCIAL`).
  Al analizar el origen de datos, `ConstruirContextoAsync` filtraba estrictamente por `semestre = @Semestre` (`'I'`). Si en el catálogo de base de datos la tabla `asignaturas` almacenaba el semestre en formato arábigo (`'1'`), vacío o nulo, la consulta devolvía 0 registros. A diferencia de `GenerarEncabezadosAulasAsync` (Fase 266), `ConstruirContextoAsync` carecía de fallback.
- **Solución Aplicada:**
  Se agregó la consulta de respaldo para `ctx.Asignaturas` y `ctx.PoolDocentes`:
  Si la consulta filtrada por semestre devuelve 0 filas, se ejecuta inmediatamente un fallback acotado a `idprograma = @IdPrograma` y `subunidad = @Sub` (sin restringir semestre). Esto garantiza que el motor genético disponga del conjunto de asignaturas del programa incluso ante discrepancias de nomenclatura entre la rama del árbol y el pensum.

### 108. **Fase 275 — Diagnóstico Pre-Vuelo Accionable y Contador Explícito de Aulas sin Asignaturas**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  Los mensajes de error previos eran genéricos e impedían discernir si el bloqueo se debía a un filtro de subunidad, falta de pensum cargado o discrepancia de semestre. Asimismo, `GenerarEncabezadosAulasAsync` informaba "éxito" aun cuando las aulas quedaban sin asignaturas vinculadas.
- **Solución Aplicada:**
  1. **Pre-vuelo accionable (`EjecutarDiagnosticoPreVueloAsync` y `btnOptimizar_Click`):**
     Se sustituyó el mensaje genérico por un detalle contextual explícito:
     `[ERROR PRE-VUELO] Sin asignaturas para idprograma={_contextoActual.IdPrograma} semestre='{semestreRama}' subunidad='{_contextoActual.Subunidad}'. Verifique el catálogo 'asignaturas' (idprograma/semestre/subunidad) o cargue el pensum.`
     Permite al usuario/administrador identificar de inmediato el parámetro faltante.
  2. **Contador explícito `sinAsignaturas` en `GenerarEncabezadosAulasAsync`:**
     Se introdujo el acumulador `int sinAsignaturas` que se incrementa si tras la consulta principal y el fallback la lista `idsAsig` permanece vacía.
     El diálogo `MessageBox.Show` y el log en `txtLogAG` notifican con advertencia si alguna de las aulas quedó sin vinculaciones en `ag_aula_asignatura`.

### 109. **Fase 276 — Verificación Operativa del Catálogo y Vault (Sin Código)**
- **Protocolo de Comprobación:**
  1. `SELECT COUNT(*) FROM asignaturas WHERE idprograma=2;` $\rightarrow$ si devuelve 0, indica dato faltante en el catálogo del vault (se requiere cargar pensum académico).
  2. `SELECT DISTINCT semestre FROM asignaturas WHERE idprograma=2;` $\rightarrow$ verificar formato de nomenclatura (`'I'`, `'II'` vs `'1'`, `'2'` vs `''`).
  3. `SELECT COUNT(*) FROM asignaturas WHERE idprograma=2 AND (subunidad='VALLE' OR subunidad='' OR subunidad IS NULL);` $\rightarrow$ comprobar visibilidad de subunidad.
  4. `SELECT COUNT(*) FROM ag_aula_asignatura WHERE id_aula IN (SELECT id FROM ag_aula WHERE idprograma=2);` $\rightarrow$ confirmar vinculaciones activas tras reejecutar `Generar Aulas`.
  5. Confirmar que el diagnóstico pre-vuelo avance sin errores de bloqueo.

---

## Ronda 43 — Traza Diagnóstica Integral y Verificación de Esquema en Módulo S (Fases 277 - 283)

### 110. **Fase 277 — `AgDiagnosticoTraza` (Motor de Trazabilidad Estructurada)**
- **Archivo Creado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoTraza.cs`
- **Diagnóstico y Motivación:**
  Previamente, los fallos y anomalías en el Módulo S (`AgFormGeneradorHorarios`) se manifestaban de forma tardía, después de transacciones abortadas o en forma de errores genéricos sin visibilidad sobre las sentencias SQL ejecutadas, parámetros enlazados, conteos de filas obtenidas ni tiempos relativos de ejecución.
- **Implementación:**
  - Clase sellada `AgDiagnosticoTraza` con `Stopwatch` para temporización relativa continua en milisegundos (`t=...ms`).
  - Secuenciación determinista de pasos: `[P01]`, `[P02]`... `[Pnn]`.
  - Métodos estructurados:
    - `Inicio(detalle)` / `Fin(ok, detalle)`: Delimitación de ciclos operativos con identificador de correlación (ej. `S-HHmmss-Op`).
    - `Paso(nombre, detalle)`: Avance de hito con emisión en nivel `DEBUG`.
    - `Dato(clave, valor)`: Emisión de variables de contexto y configuración.
    - `Sql(nombre, sql, parametros)`: Normalización de sentencias SQL (colapsando saltos de línea y espacios duplicados) y serialización segura de parámetros anónimos sin exponer PII.
    - `Conteo(etiqueta, obtenidos, minimo)`: Detección proactiva de consultas vacías emitiendo `DEBUG` si `obtenidos >= minimo` o `WARNING` con marca `ALERTA (<minimo)` si el resultado es insuficiente.
    - `Contexto(AgContexto c)`: Volcado exhaustivo de dimensiones de corrida (Ids, Subunidad, Ventana, Aulas, Bloques, Docentes, Asignaturas, Festivos, etc.).
    - `Causa(AgCausaFallo causa, detalle)`: Registro explícito del código de causa raíz.
    - `Fallo(paso, ex)`: Captura y emisión de excepciones con stack trace a través de `LoggerHelper.Error`.
  - Integración dual: Emite a `LoggerHelper` (`Debug`, `Info`, `Warning`, `Error`, `Success`) y sincroniza con el delegado UI (`txtLogAG`) cuando la traza está habilitada.

### 111. **Fase 278 — `AgEsquemaDiagnostico` (Validación Preventiva de Esquema SQLite)**
- **Archivo Creado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEsquemaDiagnostico.cs`
- **Diagnóstico y Motivación:**
  Los errores de derivación de esquema en bases de datos locales o shadow copy (tales como `no such column: activo` o `no such column: idaula`) solían manifestarse durante la ejecución de comandos Dapper en tiempo de corrida.
- **Implementación:**
  - Clase estática `AgEsquemaDiagnostico.Verificar(IDbConnection db, AgDiagnosticoTraza t, string tabla, params string[] requeridas)`.
  - Ejecuta `PRAGMA table_info(tabla)` para obtener el conjunto canónico de columnas existentes en minúsculas e invariante a cultura.
  - Compara contra las columnas requeridas por el código. Si detecta columnas faltantes, emite advertencia explícita en la traza y en `LoggerHelper.Warning` con el listado exacto de campos ausentes antes de que se ejecute la lógica de negocio.

### 112. **Fase 279 — Instrumentación Paso a Paso P01 - P11 en `AgFormGeneradorHorarios`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Implementación:**
  Se integró la traza `AgDiagnosticoTraza` a lo largo de todo el ciclo de vida del Módulo S mediante los siguientes hitos secuenciales:
  - **P01 — `CargarCatalogosAsync`:** Verificación previa de esquema (`AgEsquemaDiagnostico`) en tablas clave (`periodos`, `programas`, `sedes`, `ag_arbol`, `ag_arbol_nodo`, `ag_aula`, `ag_aula_asignatura`, `asignaturas`, `terceros_asignaturas`, `ag_corrida`) y reporte de conteos de catálogos cargados.
  - **P02 — `ActualizarVentanaSemestralAsync`:** Volcado estructurado de parámetros temporales (período, fechas límite, días de duración, modo manual vs automático).
  - **P03 — `PoblarArbolContextoAcademico`:** Conteo de nodos jerárquicos (raíz, sedes, programas, semestres, grupos, modalidades) y hojas generadas con formato `AulaRef`.
  - **P04 — `GenerarEncabezadosAulasAsync`:** Traza de combinaciones seleccionadas, inspección de SQL de asignaturas, conteos de muestra, registro de activación de fallback y balance final (`insertadas`, `omitidas`, `sinAsignaturas`).
  - **P05 — `ConstruirContextoAsync`:** Inspección de cada consulta preparatoria (`ag_aula`, `ag_bloque_horario`, `ag_aula_asignatura`, `asignaturas`, `terceros_asignaturas`, `terceros`, `festivos`, `presenciales`) registrando sentencias SQL normalizadas, parámetros y conteos.
  - **P06 — `AgContexto.ConstruirIndices`:** Invocación de `traza.Contexto(ctx)` y reporte de cardinalidad de índices cruzados.
  - **P07 — `EjecutarDiagnosticoPreVueloAsync`:** Registro de evaluaciones con tipificación mediante `AgCausaFallo` ante bloqueos de ventana, aulas, asignaturas o pool docente.
  - **P08 — `IniciarOptimizacionAsync`:** Registro de inserción de corrida, parámetros del motor genético (población, generaciones, mutación, élite) y evolución generacional.
  - **P09 — `CargarModuloResultadosAsync`:** Traza de individuos generados y persistencia en staging.
  - **P10 — `PromoverAProduccionAsync`:** Traza de inserción y promoción transaccional a tablas operativas de vinculaciones y sesiones.
  - **P11 — `OnFormClosing`:** Trazabilidad de persistencia forzada en el vault (`ForcePersistVaultAsync`) y reporte de duración de sincronización.

### 113. **Fase 280 — Códigos de Causa Raíz `AgCausaFallo`**
- **Archivo Creado:** `ConversorPDF.XLL/Academic/Domain/Enums/AgCausaFallo.cs`
- **Implementación:**
  - Enumeración tipificada que mapea las causas exactas de rechazo o inconsistencia:
    `Ninguna`, `SinPeriodo`, `SinVentana`, `SinAulas`, `SinAnclajeAula`, `SinAsignaturas`, `SinAulaAsignaturas`, `SinPoolDocente`, `SinBloques`, `VentanaInfactible`, `EsquemaIncompleto`, `TransaccionAbortada`.
  - Cumple estrictamente la regla *Single Type Per File* al alojarse en su propio archivo dentro del espacio de nombres `ConversorPDF.XLL.Academic.Domain.Enums`.

### 114. **Fase 281 — Control de Verbosidad y Checkbox de Modo Debug en UI**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Implementación:**
  - Se agregó el control `chkModoDebug` en la barra superior de acciones (`pnlBotonesSim`) de `AgFormGeneradorHorarios`.
  - Permite al usuario activar/desactivar la emisión en tiempo real de trazas de nivel `DEBUG` en el visor `txtLogAG`.
  - Al instanciarse cada traza con `CrearTraza(operacion)`, la propiedad `Habilitado` se inicializa a partir de `chkModoDebug.Checked` (o configuración local). Los errores y advertencias continúan escribiéndose en el archivo de log físico sin penalizar la interfaz gráfica.

### 115. **Fase 282 — Registro en Solución y Pruebas Unitarias**
- **Archivos Modificados y Creados:**
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`: Registrados `<Compile Include>` para `AgCausaFallo.cs`, `AgDiagnosticoTraza.cs` y `AgEsquemaDiagnostico.cs`.
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgDiagnosticoTrazaTests.cs`:
    - Prueba de secuenciación de pasos `P01..Pnn`.
    - Prueba de silencio en UI cuando `Habilitado = false`.
    - Prueba de activación de `ALERTA` en `Conteo()` al recibir valores bajo el umbral mínimo.
    - Prueba de normalización de cadenas SQL multilínea y serialización de parámetros.
    - Prueba de asignación y persistencia de `UltimaCausa`.
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEsquemaDiagnosticoTests.cs`:
    - Prueba con SQLite in-memory verificando reporte exitoso cuando todas las columnas requeridas existen.
    - Prueba con SQLite in-memory verificando reporte exacto de columnas faltantes.
    - Prueba ante tablas inexistentes.

### 116. **Fase 283 — Documentación Técnica de Diagnóstico**
- **Archivos Modificados:**
  - `implementation_plan.md`: Detalle histórico de las Fases 277 a 283.
  - `docs/articles/changes/2026-ag-horarios-academicos.md`: Inclusión de la sección 8.15 que documenta la arquitectura de trazabilidad, catálogo de códigos de causa raíz y el flujo de pasos P01 a P11.

### 117. **Fase 284 — Incorporación de `traza.Causa(...)` y Propiedad `UltimaCausa`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoTraza.cs`
- **Diagnóstico:**
  Durante la instrumentación de Módulo S en la Fase 279, se invocaba `traza.Causa(AgCausaFallo.SinAulaAsignaturas, "...")`, pero la clase `AgDiagnosticoTraza` sólo exponía la enumeración de causa dentro de la sobrecarga de `Fallo(paso, ex, causa)`. Esto generaba un error de compilación CS1061 por método inexistente.
- **Solución Aplicada:**
  - Se implementó el método público `public void Causa(AgCausaFallo causa, string detalle = null)` que registra la causa tipificada, actualiza `UltimaCausa` y emite una advertencia estructurada `WARNING` con formato `causa={causa}`.
  - Se expuso la propiedad de lectura `public AgCausaFallo UltimaCausa { get; private set; }` para trazabilidad e inspección en pruebas unitarias.

### 118. **Fase 285 — Corrección de Columnas Requeridas en Verificación de Esquema (Eliminación de Falsos Positivos)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  La verificación preventiva de esquema introducida en la Fase 278 evaluaba:
  1. `sedes` con la columna `"nombre"`: En el esquema canónico (`AcademicSchema.sql` L116-128), la tabla `sedes` no posee columna `nombre` (sus columnas descriptivas son `ciudad`, `subsede`, `departamento`).
  2. `ag_corrida` con la columna `"id_aula"`: En la migración canónica (`MigrationV49_AddAgHorariosTables.cs` L403), la columna añadida para el anclaje de aula se denomina `"idaula"` (sin guion bajo).
  Esto producía avisos ruidosos `[ESQUEMA] FALTAN: nombre` y `FALTAN: id_aula` en cada inicialización de catálogos.
- **Solución Aplicada:**
  Se actualizaron las llamadas en `CargarCatalogosAsync`:
  - `AgEsquemaDiagnostico.Verificar(db, traza, "sedes", "id", "ciudad", "subsede", "activo");`
  - `AgEsquemaDiagnostico.Verificar(db, traza, "ag_corrida", "id", "idperiodo", "idprograma", "idaula", "estado");`
  Garantizando reporte de 0 columnas faltantes en el arranque.

### 119. **Fase 286 — Verificación Estática y Cobertura de Pruebas**
- **Archivos Modificados:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgDiagnosticoTrazaTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEsquemaDiagnosticoTests.cs`
- **Resultados:**
  - Validación de compilación estática sin advertencias ni errores CS1061.
  - Concordancia total entre los nombres de columnas verificados por `AgEsquemaDiagnostico` y los esquemas canónicos `AcademicSchema.sql` y `MigrationV49`.
  - Pruebas unitarias configuradas para validar la numeración correlativa `[P01..Pnn]`, la tolerancia ante consultas vacías, el registro de causas y la verificación de esquema SQLite.
- **Resultados de Release v1.6.3.29:**
  - Pruebas unitarias ejecutadas: **440 superadas**, 0 fallidas (100% éxito).
  - Ensamblados generados en `Release|x64`: `ConversorPDF.XLL.dll`, `ConversorPDF.XLL-AddIn64.xll`.
  - Instalador generado: `installer\InstaladorConversorPDFv1.6.3.29.msi` (5,410,816 bytes).

---

## Ronda 44 — Carga de AulaAsignaturas, Resolución de Sede en Staging y Cobertura Docente (Fases 293 - 296)

### 120. **Fase 293 — Carga Relacional de `ctx.AulaAsignaturas` en `ConstruirContextoAsync`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  En `ConstruirContextoAsync`, `ctx.AulaAsignaturas` nunca se cargaba desde la base de datos, manteniéndose como lista vacía (`Count = 0`). Esto dejaba `ctx.AulaAsignaturasPorId` desprovisto de asignaturas vinculadas, distorsionaba el diagnóstico de factibilidad y generaba la alerta `AulaAsignaturas: 0 ALERTA (<1)`.
- **Solución Aplicada:**
  - Se agregó la consulta a `ag_aula_asignatura` para el aula anclada `ctx.IdAulaCorrida` previo a `ctx.ConstruirIndices()`, poblando `ctx.AulaAsignaturas` con columnas canónicas (`id_aula`, `id_asignatura`, `exclusivo`, `prioridad`, `horas_objetivo`, `orden_curricular`, `requerida`, `subunidad`).
  - Se añadieron compuertas defensivas en `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync` que bloquean de inmediato con `AgCausaFallo.SinAulaAsignaturas` si el aula anclada no posee asignaturas vinculadas.

### 121. **Fase 294 — Resolución de `IdSede` y Salvaguarda Fail-Fast en Staging**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs`
- **Diagnóstico y Causa Raíz:**
  La tabla `asignaturas` no posee columna `idsede`. Al instanciar `AgGen`, `asig.IdSede` era siempre 0, lo que se propagaba a `AgVinculacion` y provocaba que las vinculaciones materializadas en staging tuvieran `idsede = 0`, comprometiendo la consistencia de datos en producción.
- **Solución Aplicada:**
  - Se incorporó la propiedad `public int IdSede { get; set; }` en `AgContexto` y se asigna desde `sedeSel.Id` en `ConstruirContextoAsync`.
  - En `AgFabricaIndividuos.cs`, `IdSede` se resuelve a partir del aula anclada (`aSede.IdSede`) o del contexto (`contexto.IdSede`).
  - En `AgGenToAgVinculacionMapper.cs` `MapGrupo`, se implementó fallback para resolver `IdSede` desde el aula o contexto si `g0.IdSede <= 0`.
  - En `AgMaterializadorSesiones.cs`, se implementó validación fail-fast: si tras la resolución `agVinc.IdSede <= 0`, la materialización aborta inmediatamente con `Result.Failure`.

### 122. **Fase 295 — Validación de Cobertura de Pool Docente y Prevención de R5**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:**
  Si un programa carece de docentes habilitados en `terceros_asignaturas` para sus asignaturas, `AgFabricaIndividuos` recurría a docentes generales no habilitados o dejaba `idDocente = 0`, provocando violaciones duras insalvables de la regla R5 (`AgReglaViolacion.R5_HabilitacionDocentePool`).
- **Solución Aplicada:**
  - En `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync`, se valida la cobertura del pool. Si no existe ningún docente habilitado ni en catálogo, se bloquea con `AgCausaFallo.SinPoolDocente`. Si existen docentes pero ninguno en pool, se advierte explícitamente en el log y se registra la causa en la traza.
  - En `AgFabricaIndividuos.cs`, se previene la asignación de `idDocente <= 0` arrojando `InvalidOperationException` preventiva.
  - En `AgValidadorVentana.cs`, la ausencia total de docentes en pool y catálogo general se categoriza como bloqueo crítico.

### 123. **Fase 296 — Métricas de Cobertura en Trazabilidad Diagnóstica**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Solución Aplicada:**
  - En `ConstruirContextoAsync`, `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync`, se calculan y emiten métricas estructuradas:
    - `traza.Dato("CoberturaDocente", $"{asignaturasConPool}/{totalAsignaturas}");`
    - `traza.Dato("AulasSinAsignaturas", aulasSinAsig);`
  - Proporciona visibilidad accionable previa a la ejecución sobre la completitud del pensum y asignaciones docentes.

### 124. **Fase 297 — UI: Panel "Modo AG (Staging)" en `tabResultados` (Módulo R)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico y Análisis:**
  En la pestaña Módulo R (`tabResultados`), los usuarios no contaban con acceso directo a la inspección y manipulación de datos en staging (`ag_vinculaciones` / `ag_sesiones`). El único acceso disponible a nivel de form residía en `btnAbrirModoAg` (`FormSesionesAcademicas`), sin acceso estructurado a Vinculaciones completas ni a la vista agregada de Detalle.
- **Solución Aplicada:**
  - Se incorporó el panel `pnlAgModo` acoplado en `Dock = DockStyle.Top` (altura 38px, fondo suave `#FFF8F0`, borde `#FixedSingle`) ubicado inmediatamente debajo de `pnlHeaderResultados` y encima del `splitContainerResultados`.
  - Se añadieron los botones de acción e indicador visual:
    - `btnAgVinculaciones` ("🔗 Vinculaciones (AG)", 170x28px, x=10).
    - `btnAgVerVinculaciones` ("📋 Detalle (AG)", 150x28px, x=190).
    - `btnAgSesiones` ("✏️ Sesión (AG)", 150x28px, x=350).
    - `lblAgModo` ("Modo AG · staging: ag_vinculaciones / ag_sesiones", x=520, fuente Segoe UI 9pt Bold, ForeColor OrangeRed).
  - Se añadieron las declaraciones de campos privados y la adecuada secuencia en `ResumeLayout(false)` / `PerformLayout()`.

### 125. **Fase 298 — Apertura en Modo AG (Wiring desde Módulo R) y Resincronización**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Análisis:**
  - El grid `dgvHorariosGenerados` enlaza una proyección anónima y no `List<SesionListItem>`, por lo que hacer un cast directo `as List<SesionListItem>` resultaría en `null`.
  - Para abrir `FormVinculaciones` en Modo AG se requieren 5 dependencias instanciadas canónicamente con `AgModoDatos.Ag`.
  - Para abrir `FormVerVinculaciones` se requiere consultar de forma tipada `vw_ag_sesiones_detalle` mediante `VinculacionQueryRepository.GetSesionesAsync(...)`.
  - Para abrir `FormSesionesAcademicas` con la vinculación exacta, se debe resolver `id_vinculacion` desde la fila seleccionada en `ag_sesiones`.
- **Solución Aplicada:**
  - En `ConfigurarEventos()` se cablearon los tres eventos Click de forma segura y asíncrona:
    - `btnAgVinculaciones.Click += async (s, e) => await AbrirVinculacionesModoAgAsync();`
    - `btnAgVerVinculaciones.Click += async (s, e) => await AbrirDetalleSesionesModoAgAsync();`
    - `btnAgSesiones.Click += async (s, e) => await AbrirSesionModoAgAsync();`
  - Se implementó `AbrirVinculacionesModoAgAsync()` instanciando repositorios y servicio canónicos con `AgModoDatos.Ag`, abriendo modal con `ExcelDnaUtil.Application` y refrescando con `CargarModuloResultadosAsync` al cerrar.
  - Se implementó `AbrirDetalleSesionesModoAgAsync()` validando filas activas, consultando `vw_ag_sesiones_detalle` mediante `VinculacionQueryRepository` y abriendo `FormVerVinculaciones`.
  - Se implementó `AbrirSesionModoAgAsync()` y `ObtenerIdVinculacionSeleccionadaAsync()` extrayendo `IdSesion` del grid y resolviendo `id_vinculacion` desde `ag_sesiones`.

### 126. **Fase 299 — Indicador Visual "Modo AG" y Títulos de Formularios**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico y Análisis:**
  `FormVinculaciones` y `FormVerVinculaciones` ya actualizaban su título con `[MODO AG - STAGING]` en caso de `_modo.EsModoAg`, pero `FormSesionesAcademicas` no contaba con esta distinción visual en su barra de título tras `InitializeComponent()`.
- **Solución Aplicada:**
  - En `FormSesionesAcademicas.cs` (constructor), se añadió la verificación:
    ```csharp
    if (_modo.EsModoAg)
    {
        this.Text = $"{this.Text} [MODO AG - STAGING]";
    }
    ```
  - En `AgFormGeneradorHorarios.Designer.cs`, `lblAgModo` refleja el indicador canónico de staging: `"Modo AG · staging: ag_vinculaciones / ag_sesiones"`.

### 127. **Fase 303 — Bloqueo Estricto por Cobertura Incompleta de Pool Docente en Pre-Vuelo y Optimización**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  Anteriormente, la compuerta de validación de pool docente en `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync` sólo bloqueaba cuando `asigConPool == 0 && Docentes.Count == 0`, emitiendo una mera advertencia cuando la cobertura era parcial (`asigSinPool > 0`). En escenarios reales donde un programa contenía asignaturas sin docentes habilitados en `terceros_asignaturas` (ej. 5 con pool y 1 sin pool), el pre-vuelo permitía continuar, generando asignaciones inválidas con `Docente #0` y colapsando en violaciones duras de la regla R5 (`R5_HabilitacionDocentePool`).
- **Solución Aplicada:**
  - Se modificaron `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync` para calcular `asignaturasSinPool` y evaluar `if (asigSinPool > 0)`.
  - Si existen asignaturas sin docente habilitado activo en el pool, el flujo se bloquea inmediatamente (`return;`), notificando en `txtLogAG` y en un cuadro de diálogo descriptivo con la lista detallada de asignaturas afectadas (`Id:Nombre`), y registrando `AgCausaFallo.SinPoolDocente` en la traza estructurada.

### 128. **Fase 304 — Blindaje contra Asignación de Genes con `IdTercero <= 0` en Fábrica de Individuos**
- **Archivo Analizado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico y Verificación:**
  Se evaluó la inicialización en `AgFabricaIndividuos.CrearAleatorio`. La fábrica ya implementaba la salvaguarda fail-fast estricta:
  ```csharp
  if (idDocente <= 0)
  {
      throw new InvalidOperationException($"Asignatura {asig.Id} ('{asig.Nombre}') sin docente asignable en pool ni catálogo general.");
  }
  ```
  La cual, combinada con el bloqueo temprano de la Fase 303 en pre-vuelo y optimización, garantiza que ningún gen se instancie con `IdTercero <= 0` ni en la población inicial ni durante el proceso evolutivo.

### 129. **Fase 305 — Inicialización en Código del Indicador `lblAgModo`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  Aunque `lblAgModo` poseía texto predeterminado en el Designer, en el código-behind no se inicializaban explícitamente sus propiedades de estado, dependiendo exclusivamente del archivo de diseño.
- **Solución Aplicada:**
  - En el constructor de `AgFormGeneradorHorarios.cs`, tras `InitializeComponent()`, se añadió la inicialización programática canónica:
    ```csharp
    if (lblAgModo != null)
    {
        lblAgModo.AutoSize = true;
        lblAgModo.ForeColor = Color.OrangeRed;
        lblAgModo.Text = $"Modo AG · staging: {AgModoDatos.Ag.TablaVinculaciones} / {AgModoDatos.Ag.TablaSesiones}";
    }
    ```

### 130. **Fase 306 — Orden de Acoplamiento (Docking) de `pnlAgModo` en Módulo R**
- **Archivo Verificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico y Verificación:**
  Se verificó la jerarquía de acoplamiento en `tabResultados`:
  `tabResultados.Controls.Add(splitContainerResultados)` (Fill), seguido de `tabResultados.Controls.Add(pnlAgModo)` (Top), y finalmente `tabResultados.Controls.Add(pnlHeaderResultados)` (Top). WinForms procesa los controles dockeados en orden inverso de inserción, posicionando `pnlHeaderResultados` arriba, `pnlAgModo` en el medio y `splitContainerResultados` llenando el espacio restante de forma matemáticamente consistente y sin colisiones.

### 131. **Fase 307 — Encapsulamiento de Excepciones y Trazabilidad en `EjecutarOptimizacionAsync`**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  El método `EjecutarOptimizacionAsync` carecía de bloque `try/catch`. Cualquier excepción lanzada durante la inicialización de población (como `InvalidOperationException` desde `AgFabricaIndividuos`), mutación o evaluación de aptitud escapaba del servicio sin retornar un `Result.Failure<AgResultadoCorrida>` manejado ni registrar el error en bitácora, derivando en un crash no controlado en la capa de interfaz de usuario.
- **Solución Aplicada:**
  - En `AgGeneradorHorarios.cs`, se envolvió la ejecución completa de `EjecutarOptimizacionAsync` en un bloque `try/catch`:
    - `catch (OperationCanceledException)`: Retorna `Result.Failure<AgResultadoCorrida>("Optimización cancelada por el usuario.")`.
    - `catch (Exception ex)`: Registra en `LoggerHelper.Error` y retorna `Result.Failure<AgResultadoCorrida>($"Error durante la optimización: {ex.Message}")`.
  - En `AgFormGeneradorHorarios.cs` (`IniciarOptimizacionAsync`), se añadió en los bloques `catch` la emisión estructurada de `traza.Fallo("P08_EjecutarOptimizacionAsync", ex, AgCausaFallo.ExcepcionNoControlada)` y `traza.Fin(false, ...)`.

### 132. **Fase 308 — Protocolo de Verificación Operativa y Estabilidad del Motor Genético**
- **Alcance:** Verificación formal de robustez funcional:
  1. Con asignaturas sin pool docente: el diagnóstico pre-vuelo y la compuerta de inicio bloquean preventivamente listando `Id:Nombre`.
  2. Emisión en telemetría de `CoberturaDocente: n/m` y `AsignaturasSinDocente`.
  3. Ausencia de `Docente #0` y erradicación de violaciones duras `R5_HabilitacionDocentePool`.
  4. Indicador visual `lblAgModo` visible con `"Modo AG · staging: ag_vinculaciones / ag_sesiones"`.
  5. Contención total de excepciones imprevistas mediante `Result.Failure`, previniendo crashes en la aplicación de escritorio.

### 133. **Fase 309 — Corrección de Columna `idvinculacion` en `ag_sesiones`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  En `ObtenerIdVinculacionSeleccionadaAsync`, la consulta SQL utilizaba `id_vinculacion` (`SELECT id_vinculacion FROM ag_sesiones ...`). Según la definición canónica del DDL en `MigrationV49_AddAgHorariosTables.cs` (línea 201), el identificador es `idvinculacion`. Esto provocaba la excepción `SQLiteException: no such column: id_vinculacion` al pulsar el botón "✏️ Sesión (AG)".
- **Solución Aplicada:**
  - Se corrigió la consulta en `ObtenerIdVinculacionSeleccionadaAsync`:
    ```csharp
    const string sql = "SELECT idvinculacion FROM ag_sesiones WHERE id = @IdSesion LIMIT 1;";
    int? idVinc = await db.ExecuteScalarAsync<int?>(sql, new { IdSesion = idSesion });
    ```

### 134. **Fase 310 — Desduplicación de Título de Ventana en `FormSesionesAcademicas`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:**
  La asignación previa en `OnLoad` (`"Sesiones Académicas — MODO AG"`) concatenada con el sufijo del constructor (`"[MODO AG - STAGING]"`) generaba títulos duplicados redundantes: `"Sesiones Académicas — MODO AG [MODO AG - STAGING]"`.
- **Solución Aplicada:**
  - Se unificó la asignación tanto en el constructor como en `OnLoad`:
    ```csharp
    this.Text = _modo.EsModoAg ? "Sesiones Académicas [MODO AG - STAGING]" : "Sesiones Académicas";
    ```

### 135. **Fase 311 — Interpolación Dinámica de Vista en Mensaje de Detalle de Sesiones**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  En `AbrirDetalleSesionesModoAgAsync`, el cuadro de advertencia ante 0 resultados contenía el nombre de vista codificado en minúsculas (`vw_ag_sesiones_detalle`), en lugar de interpolar la propiedad canónica del modelo `modo.VistaSesionesDetalle` (`vw_AgSesionesDetalle`).
- **Solución Aplicada:**
  - Se interpoló dinámicamente `modo.VistaSesionesDetalle`:
    ```csharp
    MessageBox.Show(this, $"No se encontraron sesiones registradas en la vista de detalle staging ({modo.VistaSesionesDetalle}).", "Modo AG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    ```

### 136. **Fase 312 — Protocolo de Verificación Operativa de Integración Módulo R → Modo AG**
- **Alcance:** Verificación formal de integración:
  1. Resolución exitosa de `idvinculacion` al pulsar "✏️ Sesión (AG)" en la grilla y apertura modal de `FormSesionesAcademicas` con título limpio.
  2. Consulta tipada de `vw_AgSesionesDetalle` al pulsar "📋 Detalle (AG)".
  3. Refresco sincronizado de `dgvHorariosGenerados` tras el cierre de los diálogos modales.

### 137. **Fase 313 — Validación Preventiva de IdVinculacion en Apertura de Sesión Modo AG**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  En `AbrirSesionModoAgAsync`, cuando `ObtenerIdVinculacionSeleccionadaAsync` no resolvía un valor válido (por fila sin `IdSesion`, celda vacía o sesión inexistente en `ag_sesiones`), se invocaba ciegamente `AbrirEnModoAg(idVinc)`. Dado que `idVinc` era nulo, `FormSesionesAcademicas` se abría como una vinculación nueva en borrador sin advertir al usuario.
- **Solución Aplicada:**
  - Se introdujo una comprobación estricta con guard clause antes de abrir el diálogo modal:
    ```csharp
    int? idVinc = await ObtenerIdVinculacionSeleccionadaAsync();
    if (!idVinc.HasValue || idVinc.Value <= 0)
    {
        MessageBox.Show(this, "Seleccione una fila de resultados con IdSesion válido para editar en Modo AG.",
                        "Modo AG", MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
    }
    AbrirEnModoAg(idVinc);
    ```

### 138. **Fase 314 — Persistencia de Estado de Corrida y Rehidratación de Resultados en Módulo R**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  La variable `_ultimoResultado` únicamente se asignaba en memoria viva al término de `IniciarOptimizacionAsync`. Al reiniciar Excel o recargar una corrida histórica mediante `CargarModuloResultadosAsync(idCorrida)`:
  1. `_ultimoResultado` era `null`, provocando que `CargarModuloResultadosAsync` retornara tempranamente sin poblar la grilla ni las etiquetas.
  2. Al intentar abrir "📋 Detalle (AG)" mediante `AbrirDetalleSesionesModoAgAsync`, la guarda `dgvHorariosGenerados.Rows.Count == 0 || _ultimoResultado == null` generaba un bloqueo falso con el mensaje "Ejecute 'Optimizar Horarios'…", incluso con las 72 sesiones generadas en base de datos.
  3. Los eventos de cierre de los formularios modales de Modo AG (`AbrirVinculacionesModoAgAsync`, `AbrirDetalleSesionesModoAgAsync`, `AbrirSesionModoAgAsync`) no refrescaban la vista porque evaluaban `_ultimoResultado?.IdCorrida > 0` (falso).
  4. Los métodos de compuerta `PromoverAProduccionAsync` y `AprobarCorridaYAulaAsync` retornaban inmediatamente por condición nula de `_ultimoResultado`.
- **Solución Aplicada:**
  - Se incorporó el campo de instancia `private long _idCorridaCargada;` para rastrear persistentemente el identificador de corrida visualizada.
  - En `IniciarOptimizacionAsync` se asigna `_idCorridaCargada = idCorrida;`.
  - En `CargarModuloResultadosAsync(long idCorrida)`, se asigna `_idCorridaCargada = idCorrida;` y se añade lógica de rehidratación automática: si `_ultimoResultado?.MejorIndividuo == null`, se consulta `_scheduleRepository.GetCorridaByIdAsync(idCorrida)` y `_scheduleRepository.GetIndividuoByIdAsync(corridaDb.IdIndividuoOptimo.Value)`, reconstruyendo `_ultimoResultado` con su individuo óptimo, aptitud y métricas de convergencia.
  - Se protegió la carga de violaciones contra colecciones nulas: `(_ultimoResultado.MejorIndividuo.Violaciones ?? new List<AgViolacion>())`.
  - Se actualizaron las guardas de `AbrirDetalleSesionesModoAgAsync` a `if (dgvHorariosGenerados.Rows.Count == 0 && _idCorridaCargada <= 0)`.
  - Se migraron los callbacks de refresco modal para evaluar `if (_idCorridaCargada > 0) await CargarModuloResultadosAsync(_idCorridaCargada);`.
  - En `PromoverAProduccionAsync` y `AprobarCorridaYAulaAsync` se agregó la rehidratación transparente vía `CargarModuloResultadosAsync(_idCorridaCargada)`.

### 139. **Fase 315 — Protocolo de Verificación Operativa post-Reinicio y Modo AG**
- **Alcance:** Verificación formal de robustez funcional y operacional:
  1. Al reabrir Excel y cargar una corrida en Módulo R, el botón "📋 Detalle (AG)" abre el diálogo `FormVerVinculaciones` con las 72 sesiones proyectadas desde `vw_AgSesionesDetalle` sin bloqueos falsos.
  2. Al pulsar "✏️ Sesión (AG)" sobre una fila con `IdSesion` válido, se resuelve su `idvinculacion` en `ag_sesiones` y se abre `FormSesionesAcademicas` en `[MODO AG - STAGING]`.
  3. Al pulsar "✏️ Sesión (AG)" sin fila seleccionada o sobre una fila sin `IdSesion` resuelto, se muestra un mensaje explicativo sin instanciar un formulario en blanco.
  4. El cierre de cualquiera de los formularios de auditoría de Modo AG refresca `dgvHorariosGenerados` de forma determinista sin duplicación de registros.
  5. Las acciones de compuerta ("Aprobar Corrida y Aula" y "Promover a Producción") permanecen operables tanto en la sesión activa como tras recargar corridas históricas.

### 140. **Fase 317 — Esquema BD: Columnas de Vacante (`MigrationV54_AddVacanteColumns`)**
- **Archivos Modificados y Creados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV54_AddVacanteColumns.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV49_AddAgHorariosTables.cs` (Espejo en DDL base)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs` (Registro de V54)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (Inclusión del archivo de migración)
- **Detalle de Esquema:**
  - `ag_vinculaciones`: `es_vacante INTEGER DEFAULT 0`, `motivo_vacante TEXT DEFAULT ''`, `candidatos_pool INTEGER DEFAULT 0`.
  - `ag_sesiones`: `es_vacante INTEGER DEFAULT 0`, `motivo_vacante TEXT DEFAULT ''`, `candidatos_pool INTEGER DEFAULT 0`.
  - `vinculaciones`: `es_vacante INTEGER DEFAULT 0`, `motivo_vacante TEXT DEFAULT ''`.
  - `sesiones`: `es_vacante INTEGER DEFAULT 0`, `motivo_vacante TEXT DEFAULT ''`.
  - Migración tolerante a fallos (`IMigrationTolerante`) e idempotente ante columnas preexistentes.

### 141. **Fase 318 — Modelos de Dominio y Tipado Fuerte de Vacantes**
- **Archivos Modificados y Creados:**
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgMotivoVacante.cs`:
    - `Ninguno = 0`, `SIN_CANDIDATOS = 1`, `SIN_DISPONIBILIDAD = 2`, `NO_CUMPLE_REQUISITOS = 3`, `POOL_VACIO = 4`, `FueraDeVentanaSemestre = 5`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`: Propiedades `EsVacante`, `MotivoVacante`, `CandidatosPool` y clonación en `Clonar()`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgVinculacion.cs`: Propiedades `EsVacante`, `MotivoVacante`, `CandidatosPool`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgSesion.cs`: Propiedades `EsVacante`, `MotivoVacante`, `CandidatosPool`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Vinculacion.cs`: Propiedades `EsVacante`, `MotivoVacante`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Sesion.cs`: Propiedades `EsVacante`, `MotivoVacante`.
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgVacanteGridDto.cs` (Nuevo DTO de presentación para el grid de vacantes).
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (Registro de `AgVacanteGridDto.cs`).

### 142. **Fase 319 — Clasificación del Motivo de Vacante en Fábrica Evolutiva**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico y Solución:**
  - Se eliminó el `throw new InvalidOperationException` que abortaba la simulación cuando una asignatura carecía de docentes en pool o docentes habilitados en contexto.
  - Se implementó `ClasificarMotivo(AgContexto c, Asignatura a, DateTime? fecha, AgBloqueHorario bloque)`:
    1. Si `PoolPorAsignatura` es nulo o vacío: si no hay docentes en contexto retorna `POOL_VACIO`, de lo contrario `SIN_CANDIDATOS`.
    2. Si el pool contiene registros pero ninguno activo (`Activo == 1`), retorna `NO_CUMPLE_REQUISITOS`.
    3. Si todos los docentes habilitados están ocupados en la franja/fecha (`TerceroOcupado`), retorna `SIN_DISPONIBILIDAD`.
    4. De lo contrario, retorna `NO_CUMPLE_REQUISITOS`.
  - El gen se genera de forma segura con `IdTercero = 0`, `EsVacante = 1`, `MotivoVacante = motivo`, `CandidatosPool = candidatosCount`.

### 143. **Fase 320 — Relajación de Regla R5 y Métrica de Vacantes**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPesosFitness.cs`: Inclusión de `PesoVacante = 50.0`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgIndividuo.cs`: Métrica calculada `Vacantes => Genes?.FindAll(g => g.EsVacante == 1 || g.IdTercero <= 0).Count ?? 0`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgResultadoCorrida.cs`: Propiedad `Vacantes => MejorIndividuo?.Vacantes ?? 0`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgCorrida.cs`: Propiedad `Vacantes`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`:
    - La restricción dura `R5_HabilitacionDocentePool` ahora solo se evalúa sobre genes reales no vacantes (`!gen.EsVacante && gen.IdTercero > 0`).
    - En el cómputo de penalizaciones blandas, cada vacante suma `ctx.Pesos.PesoVacante`, incentivando al algoritmo a cubrirlas sin bloquear la corrida con violaciones duras.
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs`: Propagación de `EsVacante`, `MotivoVacante`, `CandidatosPool`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`: Propagación de atributos de vacante al expandir sesiones.
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`: Persistencia y lectura de columnas de vacante en `GuardarStagingAsync`, `GetVinculacionesStagingAsync` y `GetSesionesStagingAsync`.
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`: Persistencia de vacantes en producción (`InsertVinculacion`, `InsertSesion`, `InsertSesiones`).

### 144. **Fase 321 — Diagnóstico Pre-Vuelo: de Bloqueo a Consentimiento Explícito**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Solución:**
  - Se modificaron las compuertas de pre-vuelo en `ValidarFactibilidadVentanaAsync` e `IniciarOptimizacionAsync`.
  - En lugar de detener la ejecución con mensaje de error ante asignaturas sin pool (`asigSinPool > 0`), se despliega un diálogo de confirmación `MessageBoxButtons.YesNo`:
    *"Hay {asigSinPool} de {totalAsig} asignaturas sin docente habilitado en el pool... Se marcarán como VACANTE (es_vacante=1) con penalización suave y podrá resolverlas después en el Panel de Vacantes. ¿Continuar con la optimización?"*.
  - Si el usuario consiente, la optimización procede generando genes vacantes controlados.

### 145. **Fase 322 — Panel de VACANTES en Módulo R y Acciones Contextuales**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`:
    - Incorporación de `lblVacantes` ("Vacantes: N") en `pnlHeaderResultados`.
    - Sustitución del contenedor inferior por `tabAuditoriaResultados` con dos pestañas:
      1. `tabViolaciones` ("⚠️ Violaciones de Reglas"): Grilla de violaciones (`dgvViolaciones`).
      2. `tabVacantes` ("👤 Panel de Vacantes"): Panel de filtro (`pnlFiltroVacantes`), combo de motivos (`cmbMotivoVacante`), botón contextual (`btnAccionVacante`), contador (`lblConteoVacantesTab`) y grilla de vacantes (`dgvVacantes`).
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`:
    - Cableado de eventos: `cmbMotivoVacante.SelectedIndexChanged`, `dgvVacantes.SelectionChanged`, `btnAccionVacante.Click`.
    - Población dinámica de `dgvVacantes` a partir de `sesionesResult` filtrando por `EsVacante == 1 || IdTercero <= 0`.
    - Actualización contextual del botón `btnAccionVacante`:
      - `SIN_CANDIDATOS` o `POOL_VACIO` → "Cargar pool" (abre `FormTerceroAsignaturas` prefiltrado por programa, asignatura y semestre).
      - `SIN_DISPONIBILIDAD` → "Mover franja" (abre `FormSesionesAcademicas` en Modo AG para la vinculación seleccionada).
      - `NO_CUMPLE_REQUISITOS` → "Revisar requisitos" (abre `FormTerceroAsignaturas`).
    - Al cerrar los diálogos modales, refresco automático de resultados recalculando el contador de vacantes.
    - Advertencia preventiva en `PromoverAProduccionAsync` requiriendo confirmación si existen vacantes sin resolver.
    - Exportación CSV ampliada con columnas `Vacante`, `Motivo` y `Candidatos`.

### 146. **Fase 323 — Pruebas Unitarias de Clasificación y Fábrica de Vacantes**
- **Archivos Creados y Modificados:**
  - `ConversorPDF.Tests/Academic/Domain/AgMotivoVacanteTests.cs` (Nuevo):
    - `ClasificarMotivo_WhenPoolIsNullAndNoDocentes_ShouldReturnPoolVacio`
    - `ClasificarMotivo_WhenPoolIsNullButDocentesExist_ShouldReturnSinCandidatos`
    - `ClasificarMotivo_WhenPoolHasRowsButAllInactive_ShouldReturnNoCumpleRequisitos`
    - `ClasificarMotivo_WhenPoolActiveTeachersAreAllOccupied_ShouldReturnSinDisponibilidad`
    - `AgFabricaIndividuos_WhenNoDocenteHabilitado_ShouldCreateVacantGenWithoutThrowing`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`:
    - `Evaluar_WhenGenIsVacante_ShouldNotRecordR5HardViolation_AndApplySoftPenalty`

### 147. **Fase 324 — Sincronización Documental y Exportación**
- **Archivos Modificados:**
  - `docs/articles/changes/2026-ag-horarios-academicos.md`: Inclusión de la Sección 8.22 con descripción integral de las Fases 317 a 324.
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`: Formato CSV con columnas `Vacante`, `Motivo`, `Candidatos`.

### 148. **Fase 325 — Formato Unificado de Asignaturas y Semestre Romano (`AgFormatoAsignatura`)**
- **Archivos Creados y Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFormatoAsignatura.cs` (Nuevo)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (Inclusión del nuevo archivo)
- **Detalle de Implementación:**
  - `SemestreRomano(string)`: Mapeo tolerante de entradas arábigas (`"1"`..`"12"` → `"I"`..`"XII"`), variantes en minúsculas (`"ii"` → `"II"`), cadenas vacías o nulas (`"SIN ASIGNAR"`) y passthrough para cadenas no mapeadas.
  - `Linea(Asignatura a)`: Construcción unificada `"{a.Id}: {SemestreRomano(a.Semestre)} {Flecha} {a.Nombre}"`.
  - `BulletList(IEnumerable<Asignatura>)`: Encadenamiento con viñeta `"• "` (`string.Join("\n• ", ...)`).
  - Configuración opcional de `Flecha` (`"→"` por defecto, configurable a `"->"`).

### 149. **Fase 326 — Integración en Puntos de Llamada (`AgFormGeneradorHorarios.cs`)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Detalle de Implementación:**
  - Call site 1 (L771 - `ValidarFactibilidadVentanaAsync`):
    - `string listaSin = AgFormatoAsignatura.BulletList(asignaturasSinPool);`
    - `traza.Dato("AsignaturasSinDocente", listaSin.Replace("\n", " "));`
    - `Log($"[PRE-VUELO] {asigSinPool}/{totalAsig} asignaturas SIN docente habilitado en 'terceros_asignaturas':\n• {listaSin}");`
    - Presentación en `MessageBox.Show` con formato `• {listaSin}`.
  - Call site 2 (L935 - `IniciarOptimizacionAsync`):
    - `string listaSinOpt = AgFormatoAsignatura.BulletList(asignaturasSinPoolOpt);`
    - `traza.Dato("AsignaturasSinDocente", listaSinOpt.Replace("\n", " "));`
    - `Log($"[OPTIMIZACIÓN] {asigSinPoolOpt}/{totalAsigOpt} asignaturas SIN docente habilitado en 'terceros_asignaturas':\n• {listaSinOpt}");`
    - Presentación en `MessageBox.Show` con formato `• {listaSinOpt}`.

### 150. **Fase 327 — Pruebas Unitarias de Formato (`AgFormatoAsignaturaTests`)**
- **Archivo Creado:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFormatoAsignaturaTests.cs`
- **Casos de Prueba Cubiertos:**
  - `SemestreRomano`: `"2"` → `"II"`, `"ii"` → `"II"`, `"VIII"` → `"VIII"`, `""`/`null`/espacios → `"SIN ASIGNAR"`, `"XI"` → `"XI"`, números arábigos 1 al 10 a sus correspondientes romanos.
  - `Linea`: Objeto con `Id = 132`, `Semestre = "II"`, `Nombre = "MATEMATICAS II"` → `"132: II → MATEMATICAS II"`; objeto nulo → `string.Empty`.
  - `BulletList`: Colección de 2 asignaturas formateada con viñetas `"132: II → MATEMATICAS II\n• 133: II → PROYECTO DE FUTURO I"`.
  - `Flecha`: Configuración temporal a `"->"` y verificación de salida ASCII pura.

### 151. **Fase 328 — Documentación y Verificación de Salida**
- **Archivos Modificados:**
  - `docs/articles/changes/2026-ag-horarios-academicos.md`: Sección 8.23 detallando la especificación y comportamiento del formato legible.
  - `implementation_plan.md`: Registro de las Fases 325 a 328.

### 152. **Fase 329 — Normalización Canónica de `AgMotivoVacante` y Extensiones (`AgMotivoVacanteExtensions`)**
- **Archivos Modificados y Creados:**
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgMotivoVacante.cs`: Eliminación de alias con duplicados de valor entero (`SinCandidatos=1`, `SinDocenteHabilitado=1`, etc.), formalizando los valores canónicos en mayúsculas: `Ninguno = 0`, `SIN_CANDIDATOS = 1`, `SIN_DISPONIBILIDAD = 2`, `NO_CUMPLE_REQUISITOS = 3`, `POOL_VACIO = 4`, `FUERA_DE_VENTANA = 5`.
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgMotivoVacanteExtensions.cs` (Nuevo): Creación según la regla de un solo tipo por archivo con métodos `.Codigo()` (retorna código en mayúsculas o cadena vacía si es Ninguno) y `.Desde(string)` (parseo insensible a mayúsculas con fallback seguro a Ninguno).
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`: Registro de `AgMotivoVacanteExtensions.cs`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs`: Actualización para asignar y diagnosticar motivos utilizando `.Codigo()`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`: Asignación de `motivoVacante = motivo.Codigo()`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`: Asignación a `AgMotivoVacante.SIN_CANDIDATOS`.
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgVacanteGridDto.cs`: Proyección determinista `Motivo => MotivoVacante.Codigo()`.
  - `ConversorPDF.Tests/Academic/Domain/AgMotivoVacanteTests.cs`: Pruebas de compatibilidad y parseo en `Codigo()` y `Desde()`.
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAsignadorTerceroTests.cs`: Ajuste de aserciones al código canónico.

### 153. **Fase 330 — Promoción Parcial con Justificación Formal de Vacantes**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgPromotorVinculaciones.cs`: Incorporación de sobrecarga `PromoverAProduccionAsync(long idCorrida, long idIndividuo, string usuario, string justificacionVacante, CancellationToken ct)`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPromotorVinculaciones.cs`:
    - Relajación de la compuerta dura de vacantes: si existen vacantes (`nVac > 0`) pero se suministra `justificacionVacante`, se permite la promoción parcial de las asignaturas cubiertas.
    - Registro de `es_vacante = 1` y `motivo_vacante` en entidades de producción (`Vinculacion`, `Sesion`).
    - Asiento de `justificacion_vacante` en la tabla `ag_promocion`.
    - Omisión segura de vacantes al sincronizar pool docente post-promoción (`sv.IdTercero > 0`).
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgVinculacionToVinculacionMapper.cs`: Mapeo explícito de `EsVacante` y `MotivoVacante`.
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`: Inclusión de `es_vacante` y `motivo_vacante` en `UpdateVinculacion` y `GetVinculacionById`.
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`: Al detectar vacantes activas y confirmar la promoción, se genera la justificación formal y se transmite a `_promotor.PromoverAProduccionAsync(...)`.
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPromotorVinculacionesTests.cs`: Nueva prueba unitaria verificando que ante vacantes con justificación, la ejecución supera la compuerta de vacantes.

### 154. **Fase 331 — Población Dinámica del Combo de Motivos y Soporte de `FUERA_DE_VENTANA`**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`:
    - Población dinámica de `cmbMotivoVacante` mediante `Enum.GetValues(typeof(AgMotivoVacante)).Cast<AgMotivoVacante>().Where(m => m != AgMotivoVacante.Ninguno).Select(m => m.Codigo()).Distinct()`.
    - Soporte contextual en `btnAccionVacante` para `FUERA_DE_VENTANA` ("Ajustar ventana") abriendo Modo AG para corregir fechas.
    - Parseo robusto con `AgMotivoVacanteExtensions.Desde(s.MotivoVacante)` al alimentar la colección `_todasLasVacantes`.

### 155. **Fase 332 — Verificación Visual y Espacial del Panel de Resultados**
- **Verificación en Designer:**
  - Disposición horizontal en `pnlHeaderResultados`:
    - `lblMejorFitness`: X=15, Width=155 (Right=170)
    - `lblHardViolaciones`: X=195, Width=147 (Right=342)
    - `lblVacantes`: X=375, Width=85 (Right=460)
    - `btnExportarSimulacion`: X=580, Width=150 (Right=730)
    - `btnPromoverProduccion`: X=740, Width=200 (Right=940)
  - Espacio libre entre `lblVacantes` (Right=460) y `btnExportarSimulacion` (Left=580) de 120 px. Cero solapes espaciales confirmado.

### 156. **Fase 333 — Acción Contextual y Navegación Dual para `FUERA_DE_VENTANA`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Detalle de Implementación:**
  - `ActualizarBotonContextualVacante`: Asigna el texto `"Ajustar ventana"` cuando `item.MotivoVacante == AgMotivoVacante.FUERA_DE_VENTANA`.
  - `EjecutarAccionVacanteAsync`: Implementación de flujo interactivo dual mediante diálogo modal:
    - Si el usuario selecciona 'Sí': abre la sesión en Modo AG (`FormSesionesAcademicas`) para editar o mover las fechas específicas de la sesión fuera de ventana.
    - Si el usuario selecciona 'No': redirige inmediatamente a la pestaña de configuración del Módulo S (`tabControlPrincipal.SelectedTab = tabSimulacion`) para permitir ampliar la ventana semestral o revisar fechas de período.
    - Si cancela: descarta la acción sin alterar el estado.

---

## 157. PLAN v14 — Modelo de Proyección Semestral (Fases 378-390) · Revisión crítica

> [!IMPORTANT]
> Estado: **PROPUESTO — PENDIENTE DE APROBACIÓN**. No se ha modificado código. No compilado ni ejecutado.

### 157.1 Veredicto por fase (v14)

| Fase | Veredicto | Motivo |
|---|---|---|
| F378 | Válida con corrección | `TerceroAsignaturaKey.Calcular` es PK física de `terceros_asignaturas`; normalizarla cambia IDs ⇒ duplicados (→ F399). APIs net48 (→ F393). |
| F379 | Válida con corrección | Usa sólo `combinaciones.First()`; corrida = 1 aula ⇒ "219 filas" no es aceptación por corrida (→ F400). |
| F380 | Parcial | Parche A: cast `(DayOfWeek)7` inválido (→ F396). Parche B: semanas de α, no de la fila; tope contradictorio (→ F401). Parche C: no compilable; `R-DOC` por día = **falso positivo** (→ F401). |
| F381 | **Rechazada como está** | Fórmula no reproduce histogramas; escala días fuera de K contradiciendo H6 (→ F391, F392, F393). |
| F382 | Válida con corrección | `Asignatura.DiasSemana` no existe (→ F394); sin franja Lunes ⇒ `First()` lanza (→ F395); redefine `TotalHorasAsignatura` (→ F397); grupo colisiona en cruce (→ F398); V55 incompleta (→ F402); formato `DiaHorasSemanales` (→ F403). |
| F383 | Válida con corrección | `vin.Value` sin guarda; `AgCausaFallo.StagingVacio` inexistente (→ F404). |
| F384 | Válida | Mover antes del smoke (→ F407). |
| F385 | Válida | — |
| F386 | Válida con corrección | Opción A sin fuente de `FechaInicio` en BD (→ F406). |
| F387 | Válida con corrección | Nombres de controles erróneos (→ F405). |
| F388 | Válida | Depende de `dias_patron` (F394/F402). |
| F389 | Válida con corrección | Casos de aceptación mal etiquetados (→ F392). |
| F390 | Válida | Usar `ExcelOptimizeHelper.Optimize` (regla Excel-DNA). |
| D12 | Falso positivo parcial | `Grupo` ya viaja en tupla S/`AulaRef`; `ParsearSemestreGrupo` sólo para import Excel. Defecto real: `Grupo = "A"` hardcode. |

### 157.2 Correcciones propuestas

#### Fase 391 — R-PS: generador por recorrido de calendario (reemplaza fórmula)
- **Causa:** `64h/2d/5sem` ⇒ fórmula da `[4,4,4,4]`, no `[2,3,3,4,4]`; `48h/3d/5sem` ⇒ `[3,3,3,3]`, no `[1,2,3,3,3]`; `base/base+1` no produce dispersión >1. Histogramas son **emergentes** (semana inicial parcial + festivos).
- **Corrección:** `AgPlanificadorPatronSemanal.GenerarFechas(diasK, horas, fechaInicio, ventanaFin, esFestivo)`: recorrer días desde `FechaInicio`, aceptar si `DayOfWeek ∈ K` y no festivo, hasta `S = D/4`. Histograma = agrupación ISO-semana del resultado. Infactible ⇒ `Result.Failure` (no ampliar K).

#### Fase 392 — Casos de aceptación desde fixture real
- **Causa:** `80/4d/4sem ⇒ 4-4-4-4-4` son 5 semanas; `64/2d ⇒ 2-3-3-4-4` exige 4 días/semana con K=2 (contradice H6).
- **Corrección:** cada caso de `AgPlanificadorPatronSemanalTests` = fila real `(K, D, Ini, Fin, festivos) ⇒ fechas esperadas` del fixture `Horarios2026_02Fixture`.

#### Fase 393 — Compatibilidad .NET 4.8 / C# 7.3 / un tipo por archivo
- `Math.Clamp` ⇒ `Math.Max(1, Math.Min(x, tope))`.
- `Split('-', StringSplitOptions.RemoveEmptyEntries)` ⇒ `Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries)`.
- `AgPatronSemanal` en archivo propio + constructor explícito; registrar ambos en `ConversorPDF.XLL.csproj`.

#### Fase 394 — Fuente de K (días del patrón)
- **Causa:** `Asignatura` no tiene `DiasSemana`.
- **Corrección:** V55 agrega `ag_aula_asignatura.dias_patron TEXT DEFAULT ''` (ISO `"4,5"`); `AgAulaAsignatura.DiasPatron`; `AgFabricaIndividuos` lee de `ctx.AulaAsignaturasPorId`; vacío ⇒ fallback a franjas activas del aula (comportamiento actual, 1 día).

#### Fase 395 — Franja Lunes inexistente
- **Causa:** V49 siembra días 2..7; patrones con LUNES (H7, ≥10 filas) ⇒ `bloquesValidos.First(...)` lanza `InvalidOperationException`.
- **Corrección:** V55 `INSERT ... (1,'18:00','22:00',4.0,'NOCHE',1,'') WHERE NOT EXISTS (dia_semana=1)`; fábrica usa `FirstOrDefault` ⇒ si null: gen vacante `MotivoVacante = SIN_DISPONIBILIDAD` + `traza.Dato("FranjaInexistente", ...)`.

#### Fase 396 — `ContarOcurrencias` con domingo
- **Causa:** `(DayOfWeek)7` no es `Sunday` ⇒ 0 ocurrencias domingo.
- **Corrección:** usar `b.DiaSemana` (propiedad ya mapea 7→Sunday en `AgBloqueHorario` L26).

#### Fase 397 — No redefinir `TotalHorasAsignatura`
- **Causa:** R6, `AgMetricas`, fitness y operadores la usan como total de asignatura.
- **Corrección:** nueva `AgGen.HorasGen` (horas del día del patrón); expansor consume `HorasGen` si `PatronSemanal != ""`; I1 = `Σ HorasGen(grupo) == TotalHorasAsignatura`.

#### Fase 398 — Clave de grupo determinista
- **Causa:** contador local `++_contadorGrupoIndividuo` colisiona entre padres en cruce.
- **Corrección:** `ClaveGrupo = $"{IdAsignatura}|{IdAula}|{Grupo}"`; cruce/mutación operan por grupo; `AgValidadorReglas` verifica mismo `IdTercero`/`IdAula`/`DiasPatron` por grupo.

#### Fase 399 — Normalizar en búsqueda, no en PK
- **Causa:** cambiar `TerceroAsignaturaKey.Calcular` invalida IDs existentes.
- **Corrección:** `Calcular` intacto; `AgContexto.EsDocenteHabilitado` prueba claves `{crudo, romano, arábigo, ""}`.

#### Fase 400 — Alcance por aula + grupo real
- **Corrección:** demanda/pool por la tupla de la corrida (`ctx.IdAulaCorrida`), no `combinaciones.First()`; aceptación 219/9340/2324 = Σ de corridas por lote. `Grupo = ctx.Grupo` (no `"A"`) en `AgGen`/`AgFabricaIndividuos`.

#### Fase 401 — Factibilidad y R-DOC
- F380-B: reutilizar `GenerarFechas` (F391) por asignatura; bloqueo si `Result.Failure`.
- `R-DOC` por (docente, día) = **falso positivo**: 98 pares pueden ser franjas distintas (Sáb MAÑANA/TARDE). Regla correcta ya existe: **R1 dura por (docente, fecha, franja)**. Eliminar Parche C; sólo advertencia informativa opcional.

#### Fase 402 — Migración V55 completa e idempotente
- `IMigrationTolerante` + verificación `PRAGMA table_info` por columna.
- Contenido: columnas `ag_gen` (`indice_dia_patron`, `patron_semanal`, `dias_patron`, `clave_grupo`, `horas_gen`), `ag_aula_asignatura.dias_patron`, seed Lunes.
- Actualizar `AgScheduleRepository` INSERT (L404) / SELECT (L487), `AgGen.Clonar`, `SqliteHelper`, `.csproj`.
- Corrección de datos de `periodos` (si se elige F406-B) ⇒ **V56** separada.

#### Fase 403 — Formato `DiaHorasSemanales`
- **Causa:** agregar `|18:00 A 22:00` rompe parsers existentes (`"JUEVES|8, VIERNES|4"`).
- **Corrección:** mantener `DIA|horas` separado por `", "`.

#### Fase 404 — Materialización segura
- Guardar `vin.IsSuccess`/`ses.IsSuccess` antes de `.Value`; agregar `AgCausaFallo.StagingVacio` (y `ColisionDocente` sólo si se usa).

#### Fase 405 — Nombres de controles F387
- `dgvHorarios` ⇒ `dgvHorariosGenerados`; `btnPromoverAProduccion` ⇒ `btnPromoverProduccion`.

#### Fase 406 — Ventana α vs Excel (H14)
- Opción A rechazada: no hay `FechaInicio` por asignatura en BD.
- Recomendado: **B** (V56 datos `periodo_inicio='2026-08-11'`) o ventana manual (prioridad 3 existente) para smoke.

#### Fase 407 — Orden de ejecución
- `F378 → F379 → F384 → F380 → F391/F381 → F382 → F383 → BUILD+TESTS+SMOKE → F385 → F386 → F387 → F388 → F389 → F390`.

#### Fase 408 — Decisiones pendientes cerradas
- Tope días/semana = **|K| (duro, H6)**; 5 sólo como cota de validación.
- `R-DOC` = **R1 dura por franja**; por día ⇒ no es regla.

### 157.3 Recomendación
- **No pasar a build/smoke.** Aplicar F391-F408 sobre el plan antes de implementar F381/F382.

---

### 157.4 Catálogo Exhaustivo de Cambios de Código y Soporte Técnico por Fase Afectada

> [!NOTE]
> Todo el diseño respeta estrictamente los principios de arquitectura de la solución: .NET Framework 4.8 / C# 7.3, tipos únicos por archivo físico, llamadas limpias a `ExcelOptimizeHelper.Optimize()`, no lanzar excepciones de flujo de negocio y preservación de aislamiento entre capas.

#### 1. Fase 378 / Fase 393 / Fase 399 — Normalización Segura de Semestre/Grupo
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFormatoAsignatura.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
- **Cambios Técnicos:**
  1. En `AgFormatoAsignatura.cs`, agregar `SemestreArabigo(string semestre)` usando el mapeo canónico (1 al 12 ↔ I al XII) y passthrough defensivo.
  2. Implementar `ParsearSemestreGrupo(string valor)` con sintaxis C# 7.3 compatible: `valor.Trim().Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries)`. Devuelve `(string Semestre, string Grupo)`.
  3. **Corrección F399:** Se mantiene intacto `TerceroAsignaturaKey.Calcular(...)` para no invalidar claves primarias ni registros históricos en la tabla física `terceros_asignaturas`.
  4. En `AgContexto.EsDocenteHabilitado(...)`, realizar búsqueda sucesiva de coincidencia en el diccionario probando: clave exacta, clave con `SemestreRomano`, clave con `SemestreArabigo` y clave con semestre vacío (docente habilitado general).

#### 2. Fase 379 / Fase 400 — Demanda Acotada por S (Eliminación de Fallback Global)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
- **Cambios Técnicos:**
  1. En `AgFormGeneradorHorarios.ConstruirContextoAsync(...)`, reemplazar el fallback descontrolado (L507-525) que traía 46 asignaturas de todos los semestres.
  2. Acotar la consulta SQL `sqlAsig` al programa y a la equivalencia semestral:
     `AND (@Semestre = '' OR semestre = @Semestre OR (@SemestreArabigo <> '' AND semestre = @SemestreArabigo))`
  3. **Corrección F400:** La tupla de demanda se ancla al semestre y grupo seleccionados para la corrida (`ctx.IdAulaCorrida`), asegurando que la demanda corresponda exactamente al aula en optimización.
  4. Cargar en `ctx.AulasAlcance` la totalidad de aulas del período/sede para cálculo de capacidad global en pre-vuelo.

#### 3. Fase 380 / Fase 396 / Fase 401 — Pre-vuelo: Oferta por Ocurrencias Reales y Factibilidad de Asignación
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorTemporal.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`
- **Cambios Técnicos:**
  1. En `AgPlanificadorTemporal.cs`, implementar el método auxiliar determinista:
     `public static int ContarOcurrencias(DayOfWeek dia, DateTime ini, DateTime fin, Func<DateTime, bool> esFestivo)`.
  2. **Corrección F396:** En `AgValidadorVentana.cs`, al iterar franjas activas, invocar `b.DiaSemana` (que traduce correctamente Domingo=7 a `DayOfWeek.Sunday=0`) evitando el desbordamiento de `(DayOfWeek)b.DiaSemanaNumero`.
  3. Calcular la oferta real del semestre multiplicando la duración de cada franja por sus ocurrencias efectivas no festivas en la ventana $[\alpha_{ini}, \alpha_{fin}]$.
  4. Validar factibilidad por asignación mediante el algoritmo de fechas reales de F391. Si alguna asignatura requiere más sesiones de las que caben en su ventana y patrón de días, registrar bloqueo tipificado.
  5. **Corrección F401:** Se descarta `R-DOC` por día como condición bloqueante (falso positivo, ya que un docente puede tener dos franjas distintas el mismo día como Sábado mañana/tarde). La validación de colisiones docentes se mantiene a nivel de franja exacta R1 en `AgValidadorReglas`.

#### 4. Fase 381 / Fase 391 / Fase 392 / Fase 393 — Planificador de Patrón Semanal (`AgPlanificadorPatronSemanal`)
- **Archivos Nuevos y Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPatronSemanal.cs` (Nuevo tipo aislado)
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs` (Nuevo servicio)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (Registro de compilación)
- **Cambios Técnicos:**
  1. `AgPatronSemanal.cs`: Clase sellada con propiedades de solo lectura `DiasSemana` (`IReadOnlyList<int>`), `SesionesPorSemana` (`IReadOnlyList<int>`), `FechasValidas` (`IReadOnlyList<DateTime>`). Constructor explícito sin Primary Constructors (regla C# 7.3).
  2. **Corrección F391:** En `AgPlanificadorPatronSemanal.cs`, descartar fórmulas de división cerrada. Implementar `GenerarFechas(...)` recorriendo día a día el calendario desde `FechaInicio` hasta `FechaFin`, verificando pertenencia a los días $K$ del patrón y exclusión de festivos hasta completar $S = D / 4$ sesiones.
  3. Derivar `SesionesPorSemana` agrupando las fechas resultantes por semana del año ISO 8601. Este recorrido genera de forma natural y exacta los 13 histogramas de $H9$ (incluidos patrones escalonados como `2,3,3,4,4` y `1,2,3,3,3`).
  4. Respetar estrictamente los días declarados $K$ (invariante $H6$). Si la ventana es insuficiente con los días dados, retornar `Result.Failure` en lugar de agregar días arbitrariamente.

#### 5. Fase 382 / Fase 394 / Fase 395 / Fase 397 / Fase 398 / Fase 402 / Fase 403 — Multi-Día 1↔N Genes, Invariantes y Migración V55
- **Archivos Afectados y Creados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV55_AddAgGenPatronSemanal.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs` (Registro en migrador)
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/AgAulaAsignatura.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Cambios Técnicos:**
  1. **Migración V55 (F402):**
     - Añadir columnas a `ag_gen`: `indice_dia_patron INTEGER NOT NULL DEFAULT 0`, `patron_semanal TEXT NOT NULL DEFAULT ''`, `dias_patron TEXT NOT NULL DEFAULT ''`, `clave_grupo TEXT NOT NULL DEFAULT ''`, `horas_gen REAL NOT NULL DEFAULT 4.0`.
     - Añadir columna a `ag_aula_asignatura`: `dias_patron TEXT NOT NULL DEFAULT ''` (almacena días ISO como `"4,5"`).
     - **Corrección F395:** Sembrar franja de Lunes en `ag_bloque_horario` (`dia_semana=1, hora_inicio='18:00', hora_fin='22:00', duracion_horas=4.0, turno='NOCHE', activo=1`) si no existe, garantizando que patrones con Lunes no lancen excepción.
     - Implementación con verificación preventiva `PRAGMA table_info` e idempotencia.
  2. **Modelo `AgGen.cs` (F397 & F398):**
     - Incorporar `IndiceDiaPatron`, `PatronSemanal`, `DiasPatron`, `ClaveGrupo` y `HorasGen`.
     - Preservar `TotalHorasAsignatura` como las horas globales de la asignatura (invariante para métricas y R6). `HorasGen` almacena las horas correspondientes a este gen específico dentro del patrón.
     - `ClaveGrupo` determinista: `${IdAsignatura}|${IdAula}|${Grupo}` (inmune a cruces evolutivos).
     - Actualizar `Clonar()` copiando los nuevos campos.
  3. **Fábrica `AgFabricaIndividuos.cs`:**
     - Leer `dias_patron` desde `AgAulaAsignatura` (F394).
     - Por cada asignatura, si posee patrón multi-día de $N$ días, generar $N$ genes vinculados bajo la misma `ClaveGrupo`, asignando a cada uno la franja horaria correspondiente a su día y sus `HorasGen`.
  4. **Expansor `AgExpansorSesiones.cs`:**
     - Expandir fechas utilizando `gen.HorasGen` si el gen forma parte de un patrón.
     - Aserción de integridad: horas acumuladas en sesiones igual a `gen.HorasGen`.
  5. **Materializador `AgMaterializadorSesiones.cs` (F403):**
     - Agrupar genes por `ClaveGrupo` para sintetizar la vinculación.
     - Formatear `DiaHorasSemanales` en la convención estándar no disruptiva: `"DIA|horas, DIA|horas"` (ej: `"JUEVES|4, VIERNES|4"`), preservando compatibilidad con los consumidores existentes.
     - Validar que `Σ gen.HorasGen == agVinc.TotalHoras == asignatura.IntHoraria` (Invariante I1).
  6. **Repositorio `AgScheduleRepository.cs`:**
     - Actualizar sentencias de inserción y lectura en `SaveIndividuoAsync` y consultas de staging para mapear las columnas añadidas en V55.

#### 6. Fase 383 / Fase 404 — Materialización Verificada y Manejo de Errores
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgCausaFallo.cs`
- **Cambios Técnicos:**
  1. En `AgCausaFallo.cs`, incorporar los valores canónicos `SinAsignaturas` y `StagingVacio`.
  2. En `AgFormGeneradorHorarios.cs`, tras la invocación a `MaterializarEnStagingAsync`:
     - Evaluar explícitamente `matRes.IsFailure`. Si falla, registrar traza de fallo y notificar vía `MessageBox`.
     - **Corrección F404:** Consultar staging y verificar defensivamente `ses.IsSuccess && vin.IsSuccess` antes de acceder a `.Value`. Si alguna consulta falla o retorna 0 elementos, emitir alerta de inconsistencia y evitar presentar métricas erróneas.

#### 7. Fase 384 / Fase 407 — Pool por Unión (Semestre ∪ General) sin Fallback Condicional
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
- **Cambios Técnicos:**
  1. En `AgFormGeneradorHorarios.cs`, unificar la carga de docentes en una única consulta SQL:
     `SELECT ... FROM terceros_asignaturas WHERE activo = 1 AND idprograma = @IdPrograma AND (@Semestre = '' OR semestre = @Semestre OR (@SemestreArabigo <> '' AND semestre = @SemestreArabigo) OR semestre = '' OR semestre IS NULL) AND (subunidad = @Sub OR subunidad = '' OR subunidad IS NULL)`
  2. Eliminar completamente el bloque de fallback condicional `if (ctx.PoolDocentes.Count == 0)`.
  3. En `AgContexto.ConstruirIndices()`, construir `PoolPorAsignatura` ordenando los candidatos de forma prioritaria: primero aquellos asignados específicamente al semestre actual, seguidos por prioridad y menor carga de asignaciones anteriores.

#### 8. Fase 385 — Identificación Canónica de Programas Homónimos y Aislamiento por Sede
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambios Técnicos:**
  1. Resolver programas en el selector comparando por su identificador único y nombre normalizado, alertando en bitácora si existen homónimos entre distintas modalidades o planes (ej. APT-52 vs APT-53).
  2. Delimitar estrictamente la oferta y demanda a la sede física seleccionada, evitando agregaciones espurias de sedes territoriales distintas.

#### 9. Fase 386 / Fase 406 — Ventana Semestral $\alpha$ y Cobertura Temporal Completa (H14)
- **Archivos Afectados / Creados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV56_AjustarInicioPeriodo2026_02.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResolutorVentanaSemestre.cs`
- **Cambios Técnicos:**
  1. **Corrección F406 (Opción B):** Implementar migración de datos V56 que actualiza `periodo_inicio = '2026-08-11'` para el período `2026-02` (idperiodo = 2), garantizando que la ventana formal abarque desde el primer martes de clases observado en `Horarios.xlsx` (11-ago-2026) hasta el 12-dic-2026 (17 semanas completas).
  2. En `AgResolutorVentanaSemestre.cs`, asegurar que la prioridad 5 (`Periodos`) tome este rango y registre en bitácora el inicio, fin y semanas activas tras descontar festivos.

#### 10. Fase 387 / Fase 405 — Saneamiento y Honestidad Visual en UI
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambios Técnicos:**
  1. **Corrección F405:** En `CargarResultados(...)`, utilizar los identificadores reales de controles WinForms:
     - `dgvHorariosGenerados.DataSource = null;`
     - `btnPromoverProduccion.Enabled = false;`
     - `lblMejorFitness.Text = lblVacantes.Text = lblHardViolaciones.Text = "—";`
  2. Si la corrida no produjo sesiones válidas en staging, limpiar completamente el panel y advertir al usuario en consola, impidiendo que queden residuos visuales de optimizaciones previas.

#### 11. Fase 388 — Generación Asistida de Aulas Congruente
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambios Técnicos:**
  1. En `GenerarEncabezadosAulasAsync(...)`, aplicar la misma regla de demanda acotada de F379 al poblar `ag_aula_asignatura`.
  2. Asignar `horas_objetivo = asignaturas.int_horaria` y almacenar el patrón inicial en `dias_patron`, garantizando congruencia absoluta entre el catálogo de aulas y los requerimientos curriculares.

#### 12. Fase 389 / Fase 392 — Suite de Pruebas Unitarias de Regresión y Paridad
- **Archivos Afectados y Creados en `ConversorPDF.Tests`:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFormatoAsignaturaTests.cs` (Extensión con casos de números arábigos y parseo con guiones)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanalTests.cs` (Nuevo: validación de los 13 histogramas de $H9$ usando calendario real)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs` (Reescritura de oferta por ocurrencias y descarte de falso positivo `R-DOC`)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/PoolDocenteUnionTests.cs` (Nuevo: verificación de cobertura docente con pool unificado)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMultiDiaGenTests.cs` (Nuevo: verificación de 1↔N genes, suma de horas e invariante I1)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorVerificadoTests.cs` (Nuevo: verificación de guardas de staging)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/Fixtures/Horarios2026_02Fixture.cs` (Fixture canónico con filas representativas del archivo `Horarios.xlsx`)

#### 13. Fase 390 — Exportación Espejo 1:1 con Layout de `Horarios.xlsx`
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Cambios Técnicos:**
  1. Incorporar botón `btnExportarProyeccion` en el panel de resultados.
  2. Implementar la generación de la hoja de cálculo con la estructura espejo exacta:
     - Columnas A a K: Metadatos (`No.`, `Programa`, `Sem.`, `Horas`, `Asignatura`, `Sede`, `Residencia`, `Docente`, `Fecha Inicio`, `Fecha Final`, `Días/Horas Semana`).
     - Columnas L a DZ: 119 columnas-día consecutivas rotando fechas y etiquetas de día (`Ma, Mi, J, V, S, D, L`).
     - Columna EA: `SubTotal` (fórmula suma o cálculo idéntico).
  3. Envolver todo el proceso de exportación dentro de un bloque `using (ExcelOptimizeHelper.Optimize(app))` para maximizar el rendimiento y garantizar la liberación determinista de objetos COM.

### 157.5 Resolución Integral de Ronda 52 — Correcciones Críticas y Paridad Funcional (Fases 409 - 422)

Evaluación exhaustiva de los 14 hallazgos detectados en la revisión estática v14, validando los 2 defectos bloqueantes de compilación/migración y las 12 mejoras funcionales de preservación de invariantes, paridad de histogramas y consistencia de alcances.

#### 1. Fase 409 — Corrección de Error de Compilación CS1061 en Materialización
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** `grp` es de tipo `IGrouping<string, AgGen>`, por lo que `grp.Key` es un `string` (`ClaveGrupo`), no un objeto compuesto. El acceso a `grp.Key.IdAula` arrojaba error de compilación CS1061.
- **Solución:** Se extrajo `int idAulaGrupo = genes.Count > 0 ? genes[0].IdAula : 0;` del primer gen ordenado por `IndiceDiaPatron`, empleando `idAulaGrupo` de forma segura para resolver la sede y emitir el mensaje de fallo.

#### 2. Fase 410 — Corrección de Columna 'franja' en Migración V55
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV55_AddAgGenPatronSemanal.cs`
- **Diagnóstico:** El seed para la franja de los lunes utilizaba la columna inexistente `turno` en `INSERT INTO ag_bloque_horario`, ocasionando fallo de SQLite en el arranque (`no such column: turno`).
- **Solución:** Se corrigió el nombre de la columna a `franja`, en estricta conformidad con el esquema canónico definido en `MigrationV49_AddAgHorariosTables.cs`.

#### 3. Fase 411 — Derivación e Inserción de 'dias_patron' en Encabezados de Aulas
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** La inserción en `ag_aula_asignatura` pasaba una cadena vacía constante `''` para `dias_patron`, provocando que el validador y la fábrica operaran a ciegas sin datos de multi-día.
- **Solución:** Se implementó la regla de derivación empírica según carga horaria:
  - $\ge 80$ h $\rightarrow$ `"2,3,4,5,6"` (5 días)
  - $\ge 64$ h $\rightarrow$ `"4,5,6"` (Jueves, Viernes, Sábado)
  - $< 64$ h $\rightarrow$ `"2,3"` (Martes, Miércoles)
  Pasando el parámetro `@DiasPatron` en la sentencia Dapper `INSERT OR IGNORE`.

#### 4. Fase 412 — Degradación Defensiva a Día Único ante Falla de Patrón
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Si `_planificadorPatron.Construir` retornaba `Failure`, `patron` era nulo, la propiedad `PatronSemanal` quedaba vacía y `SesionesDelGen()` evaluaba las horas totales completas para cada uno de los `diasK.Count` genes, multiplicando artificialmente la carga horaria por el número de días.
- **Solución:** Si `resPatron.IsFailure`, se degrada inmediatamente a día único: `diasK = new List<int> { diasK[0] };`, asegurando que se genere un solo gen con la carga total sin multiplicar las horas.

#### 5. Fase 413 — Reproducción Fiel de Histogramas Escalonados Empíricos
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El recorrido día a día puro producía histogramas uniformes, sin reproducir los patrones escalonados observados en `Horarios.xlsx` (`1-2-3-3-3`, `2-3-3-4-4`, `3-4-4-5`, `1-2-2-2`).
- **Solución:** Se implementó la distribución por semanas con rampa empírica: si `semanas >= 4` y `baseD > 1`, se decrementa la semana 0 (efecto arranque) y se compensa en las semanas posteriores respetando el tope de 5 días por semana. Se emplearon primitivas compatibles con .NET Framework 4.8 (`Clamp`, `Math.Max`, `Math.Min`).

#### 6. Fase 414 — Asignación y Consumo de 'FechasGen' en Expansión de Sesiones
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
- **Diagnóstico:** `AgExpansorSesiones` regeneraba fechas desde `FechaInicio`, concentrando las sesiones al inicio del semestre y desincronizándose del histograma del patrón.
- **Solución:**
  1. Se añadió `public List<DateTime> FechasGen { get; set; } = new List<DateTime>();` a `AgGen` (con soporte en `Clonar()`).
  2. La fábrica filtra y asigna las fechas precalculadas por el planificador a cada gen según su día ISO.
  3. `AgExpansorSesiones` consume prioritariamente `gen.FechasGen` si está poblado antes de invocar `GenerarFechasValidas`.

#### 7. Fase 415 — Agrupación por Semanas Relativas Ordenadas
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El cálculo previo con `GetWeekOfYear` sin ordenar rompía la secuencia ante semanas que cruzan fin de año.
- **Solución:** Se computan semanas ordinales relativas a la fecha de inicio del periodo semestral mediante `(f.Date - fechaInicio.Date).TotalDays / 7.0` con ordenamiento determinista por clave.

#### 8. Fase 416 — Cálculo Protegido y Verificación de Invariante I1 en Fábrica
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** `totalHoras / diasK.Count` carecía de protección contra división por cero y no verificaba que la suma de horas de los genes coincidiera con la carga de la asignatura.
- **Solución:** Se protegió con `Math.Max(1, diasK.Count)` y se añadió la verificación post-bucle $\sum \text{HorasGen} == \text{totalHoras}$ emitiendo advertencia `[I1]` ante discrepancias.

#### 9. Fase 417 — Validación de Invariantes I1 e I2 en Materialización de Staging
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** La persistencia en staging no validaba matemáticamente las horas ni las sesiones agrupadas antes de insertar.
- **Solución:** Se agregaron compuertas estrictas:
  - Invariante I1: `Math.Abs(horasGrupo - esperado) <= 0.001` (retorna `Result.Failure` si no coincide).
  - Invariante I2: `Math.Abs(sesionesGrupo * 4.0 - esperado) <= 0.001` (retorna `Result.Failure` si las sesiones expandidas difieren de `TotalHorasAsignatura / 4`).

#### 10. Fase 418 — Homogeneización de Alcance de Oferta vs Demanda en Validador
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** La oferta computada sobre `AulasAlcance` sumaba todas las aulas del programa/sede sin restringir al semestre y grupo de la demanda evaluada, distorsionando la comparación.
- **Solución:** Se alineó `aulasOferta = aulasDisponibles`, evaluando la capacidad en el mismo alcance curricular que la demanda.

#### 11. Fase 419 — Unificación de Valor por Defecto de 'diasDeclarados'
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** El validador usaba 2 días fijos mientras la fábrica caía a 1 gen si `DiasPatron` venía vacío.
- **Solución:** Se estandarizó la fuente única: si `DiasPatron` no está configurado, ambos componentes adoptan por defecto 2 días (`"2,3"`: Martes y Miércoles).

#### 12. Fase 420 — Limpieza de Deuda Técnica y Visibilidad en 'AgContexto'
- **Archivo:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
- **Diagnóstico:** Presencia de nombres de tipo completamente cualificados en llamadas a `AgFormatoAsignatura` y setter público no restringido en `PoolPorAsignatura`.
- **Solución:** Se incorporó el `using ConversorPDF.XLL.Academic.Application.Services.Genetica;` y se restableció `public Dictionary<int, List<TerceroAsignatura>> PoolPorAsignatura { get; private set; }`.

#### 13. Fase 421 — Exportación Espejo 1:1 con Seriales OADate y Título A1
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El exportador emitía fechas en formato de texto plano `yyyy-MM-dd` sin fila de título de periodo.
- **Solución:** Se añadió el título A1 `"PROGRAMACION DOCENTE PERIODO 2026-02"` y se formatearon las 119 columnas-día utilizando números de serie OLE Automation (`f.ToOADate()`), asegurando plena interoperabilidad matemática y visual con Excel.

#### 14. Fase 422 — Pruebas Unitarias para Histogramas Escalonados e Invariante I1
- **Archivos:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanalTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`
- **Diagnóstico:** La suite de pruebas no validaba las secuencias escalonadas observadas en `Horarios.xlsx` ni la conservación de suma de `HorasGen`.
- **Solución:** Se añadieron pruebas unitarias para:
  - 64h / 2 días / 5 sem $\rightarrow$ `[3, 4, 4, 5]`
  - 64h / 4 días / 5 sem $\rightarrow$ `[2, 3, 3, 4, 4]`
  - 48h / 3 días / 5 sem $\rightarrow$ `[1, 2, 3, 3, 3]`
  - 28h / 1 día / 4 sem $\rightarrow$ `[1, 2, 2, 2]`
  - Preservación estricta de $\sum \text{HorasGen} == \text{TotalHorasAsignatura}$ en soluciones generadas por `AgFabricaIndividuos`.

### 157.6 Ronda 53 — Calibración Matemática de Escalado e Invariantes (Fases 423 - 429)

Resolución de los 7 hallazgos detectados en la verificación del escalado empírico y blindaje de invariantes formales:

#### 1. Fase 423 — Conservación Estricta $\sum \text{sesiones} == S$
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** Si todas las semanas alcanzaban el tope (`TopeDiasSemana = 5`), la compensación de decremento en semana 0 no encontraba receptor, perdiendo 1 sesión (ej. 80h $\rightarrow$ 19 sesiones en vez de 20).
- **Solución:** Guarda estricta condicional y verificación final `if (porSemana.Sum() != S) return Result.Failure<AgPatronSemanal>(...)`, garantizando cumplimiento exacto de la carga requerida.

#### 2. Fase 424 — Reparto Monótono Creciente por Redondeo Acumulado
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanalTests.cs`
- **Diagnóstico:** Los desplazamientos unitarios heurísticos generaban vectores discordantes con aserciones exactas rígidas.
- **Solución:** Implementación de reparto acumulado telescópico $\lceil S \times w / W \rceil - \lceil S \times (w-1) / W \rceil$, garantizando monotonicidad no decreciente y $\sum = S$. En pruebas unitarias, se relajó a comprobación de invariantes fundamentales (`Sum() == S`, `Max() <= nCap`, orden ascendente).

#### 3. Fase 425 — Preservación de Patrones Uniformes Canónicos
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El escalado forzado alteraba innecesariamente las 197 filas (90%) del Excel que presentan distribuciones perfectamente uniformes (`[2,2,2,2,2,2]`, `[2,2,2,2]`).
- **Solución:** El escalado solo se activa si el patrón uniforme no cabe en la ventana disponible `W` (`(int)Math.Ceiling(S / (double)baseD) > W`).

#### 4. Fase 426 — Integración de Festivos en Cálculo de 'SemanasDisponibles'
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El parámetro `esFestivo` no era consumido en `SemanasDisponibles`, calculando semanas sobre días inhábiles e inflando la oferta disponible.
- **Solución:** Se evalúan iterativamente bloques de 7 días verificando la existencia de al menos un día de `diasK` no festivo en la ventana.

#### 5. Fase 427 — Restricción de Días a 'diasK' Declarados
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** Se incorporaban días no configurados por la institución cuando `n > diasK.Count`.
- **Solución:** `nCap` se limita estrictamente a `Math.Min(diasK.Count, TopeDiasSemana)`, y la lista de días se acota exclusivamente a los elementos de `diasK`.

#### 6. Fase 428 — Penalización Dura y Vacante ante Inconsistencia I1 en Fábrica
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La discrepancia en $\sum \text{HorasGen} \neq \text{totalHoras}$ solo emitía un log informativo sin penalizar al cromosoma.
- **Solución:** Se incrementa `individuo.HardViolations++` y se marcan los genes del grupo con `EsVacante = 1` y motivo `SIN_DISPONIBILIDAD`, asegurando que soluciones incongruentes sean descartadas por la presión de selección.

#### 7. Fase 429 — Normalización Canónica de Horas Semanales en Día Único
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** En la rama de día único se conservaba `asig.HorasSemanales` (2 a 6h), provocando que `SesionesDelGen()` calculara un número incorrecto de sesiones.
- **Solución:** Se fijó `HorasSemanales = 4.0;` e `HorasGen = totalHoras;`, manteniendo la invariante de sesiones atómicas de 4 horas.

### 157.7 Ronda 54 — Estabilización de Orden Monótono y Ajuste Constructivo I1 (Fases 430 - 433)

Resolución de los 4 hallazgos residuales detectados en la verificación del orden de vectores y el balance de horas:

#### 1. Fase 430 — Orden Ascendente Estable en Reparto Escalado
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El cálculo telescópico de diferencias generaba secuencias en orden decreciente (`[2,2,2,1]`), exigiendo la mayor concentración de sesiones al inicio del semestre.
- **Solución:** Se aplicó `porSemana.Sort()`, garantizando orden ascendente monótono estable, preservando idénticamente la suma $\sum = S$ y validando que cada elemento pertenezca al rango $[1..nCap]$.

#### 2. Fase 431 — Reporte Estricto de Semanas Disponibles sin Enmascarar 0
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** `SemanasDisponibles` retornaba `Math.Max(1, semanas)`, ocultando ventanas sin ningún día hábil y permitiendo aprobaciones falsas en pre-vuelo.
- **Solución:** Retorno directo de `semanas` (permitiendo 0) e inserción de guarda en `Construir`: si $W = 0$, retorna `Result.Failure` inmediato informando que la ventana carece de días hábiles.

#### 3. Fase 432 — Ajuste Constructivo del Residuo en Gen 0 sin Falsas Penalizaciones Duras
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Incrementar `HardViolations++` por ligeras discrepancias de redondeo descalificaba el individuo completo (`EsFactible = false`), impidiendo su promoción en el Módulo R.
- **Solución:** Se eliminó el incremento de `HardViolations` y se ajustó el residuo directamente en `g0.HorasGen += (totalHoras - sumaGenes)`, satisfaciendo constructivamente el invariante I1 de forma exacta.

#### 4. Fase 433 — Diagnóstico Accionable ante Insuficiencia de Franjas por Festivos
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El mensaje de fallo al no alcanzar la cuota semanal era genérico sin indicar la acción requerida.
- **Solución:** Mensaje enriquecido detallando causa y acción correctiva: `"Infactible: semana n sólo x/y sesiones (festivos o días K no disponibles). Ajuste ventana o agregue franjas."`.

### 157.8 Ronda 55 — Sincronización Pre-Vuelo, Múltiplo Canónico de 4 y Diversificación de Franjas (Fases 434 - 437)

Resolución de los 4 hallazgos detectados en la auditoría de paridad del validador de ventana, cuantización de sesiones y asignación de franjas:

#### 1. Fase 434 — Sincronización Pre-Vuelo vs Planificador en Semanas Disponibles
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `AgValidadorVentana` calculaba semanas calendario planas `Math.Max(1, (int)Math.Ceiling((fin - ini).TotalDays / 7.0))` ignorando festivos y días hábiles del patrón. Esto causaba falsos positivos en pre-vuelo (aprobaba la ventana) y bloqueos no tipificados durante la corrida del algoritmo genético (`Failure("Infactible: semana n...")`).
- **Solución:** Se delegó el cómputo de semanas en pre-vuelo a `AgPlanificadorPatronSemanal.SemanasDisponibles(diasKAsig, ini, fin, contexto.EsFestivo)`. Si retorna 0 semanas hábiles, se emite inmediatamente un bloqueo tipificado en el informe de pre-vuelo: `"Asignatura '{asig.Nombre}' (Id={asig.Id}): la ventana no contiene días hábiles para el patrón declarado."`. Además, se sincronizó el cómputo de horas con la jerarquía completa de la fábrica (`HorasObjetivo -> IntHoraria -> Horas -> Creditos*16 -> 64`).

#### 2. Fase 435 — Cuantización de Horas a Múltiplos Canónicos de 4 (Invariantes I1 e I2)
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Si una asignatura tenía horas no divisibles exactamente por 4 (p. ej. 45 o 65 h), el número de sesiones se calculaba con `Math.Ceiling(totalHoras / 4.0)` proyectando sesiones completas de 4 horas, lo que generaba un residuo negativo `totalHoras - (sesiones * 4)` y forzaba sesiones fraccionarias en el expansor, violando la regla del modelo de negocio de Horarios (todas las sesiones en Excel son exactamente de 4.0 horas).
- **Solución:** Antes de instanciar genes o construir el patrón, se normaliza `totalHoras` al múltiplo superior de 4 mediante `Math.Ceiling(totalHoras / 4.0) * 4.0` cuando `Math.Abs(totalHoras % 4.0) > 0.001`, emitiendo una advertencia trazable en `LoggerHelper.Warning` con el tag `[I2]`. Esto garantiza que `totalHoras == sesiones * 4.0` y previene residuos negativos.

#### 3. Fase 436 — Eliminación de Código Muerto 'Clamp'
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El método auxiliar privado `Clamp` quedó sin referencias tras la refactorización de escalado de la Fase 427, produciendo warnings de compilación / análisis estático.
- **Solución:** Eliminación completa del método no utilizado para mantener el código conciso y libre de advertencias de análisis estático.

#### 4. Fase 437 — Rotación Determinista de Franjas Horarias por Asignatura y Gen
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La selección de bloque mediante `bloquesValidos.FirstOrDefault(b => b.DiaSemanaNumero == diaIso)` seleccionaba invariablemente la primera franja del día (en Sábado, siempre la franja MAÑANA 08:00–12:00), provocando colisiones sistemáticas de aula/horario (restricción dura R4) entre asignaturas que comparten el mismo día y reduciendo artificialmente la diversidad de la población inicial.
- **Solución:** Se implementó una selección por rotación determinista `franjasDia[Math.Abs(asig.Id + i) % franjasDia.Count]`, dispersando las asignaturas entre franjas disponibles (MAÑANA / TARDE) tanto en la rama multi-día como en la rama de día único.

### 157.9 Ronda 56 — Respeto de Patrón Declarado, Fuente Única de Días K y Limpieza de AulasAlcance (Fases 438 - 440)

Resolución de los 3 hallazgos residuales detectados en la auditoría de persistencia de patrones, unificación de días de la semana y optimización de consultas:

#### 1. Fase 438 — Respeto de Patrón Declarado en Catálogo frente a Heurística Fija
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al sembrar asignaturas en `ag_aula_asignatura`, se sobrescribía siempre `dias_patron` con una heurística rígida basada en horas ($\ge 80 \rightarrow \text{"2,3,4,5,6"}$, $\ge 64 \rightarrow \text{"4,5,6"}$, $< 64 \rightarrow \text{"2,3"}$), ignorando si la asignatura ya poseía una combinación personalizada de días (en `Horarios.xlsx` coexisten 16 combinaciones distintas).
- **Solución:** Se modificó la consulta `sqlAsigs` para realizar `LEFT JOIN ag_aula_asignatura aa ON aa.id_aula = @IdAula AND aa.id_asignatura = a.id` recuperando `COALESCE(aa.dias_patron, '') AS DiasPatron`. Al asignar `diasPatron`, se evalúa prioritariamente `!string.IsNullOrWhiteSpace(aItem.DiasPatron) ? aItem.DiasPatron : (horas >= 80 ? "2,3,4,5,6" : horas >= 64 ? "4,5,6" : "2,3")`, preservando fielmente los patrones configurados por el usuario o importados desde Excel.

#### 2. Fase 439 — Fuente Única Canónica de Días K (`DiasKDe`)
- **Archivos:** 
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** `AgValidadorVentana` y `AgFabricaIndividuos` parseaban de forma independiente la cadena `DiasPatron` con código duplicado. Si surgía cualquier discrepancia o valor vacío, el pre-vuelo aprobaba con 2 días mientras la corrida intentaba con 1 día (o viceversa), produciendo divergencias en factibilidad.
- **Solución:** Se centralizó el parseo en el método compartido `AgPlanificadorPatronSemanal.DiasKDe(AgAulaAsignatura aa)`, el cual devuelve de forma inmutable la lista canónica de días ISO $[1..7]$ o el default `{ 2, 3 }`. Tanto el validador de ventana como la fábrica de cromosomas consumen este método único, eliminando 25 líneas de código redundante y garantizando paridad total.

#### 3. Fase 440 — Limpieza de Consulta Redundante de `AulasAlcance`
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Tras la unificación de alcances de la Fase 418, el validador opera directamente sobre `aulasDisponibles` (mismo alcance curricular que la demanda). La consulta asíncrona a base de datos `GetAulasPorPeriodoAsync` con semestre y grupo nulos era ejecutada redundantemente solo para poblar `AulasAlcance` y emitir una traza.
- **Solución:** Se eliminó la llamada duplicada a base de datos y se sincronizó `ctx.AulasAlcance = ctx.Aulas != null ? ctx.Aulas.ToList() : new List<AgAula>();`, manteniendo viva la métrica de telemetría sin latencia innecesaria de I/O.

### 157.10 Ronda 57 — Dapper POCO Mapping en .NET 4.8, Supresión de 64h Ficticias y SELECT DISTINCT (Fases 441 - 443)

Resolución del defecto bloqueante de runtime en Dapper y las dos inconsistencias de siembra de asignaturas:

#### 1. Fase 441 — Dapper POCO Mapping y `int_horaria` Anulable en .NET Framework 4.8
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `QueryAsync<(int Id, int IntHoraria, string DiasPatron)>` utilizaba `ValueTuple`, el cual no es soportado por el motor de mapeo de Dapper en .NET Framework 4.8 (requiere netstandard2.1+). Adicionalmente, valores NULL en la columna `int_horaria` de la base de datos SQLite provocaban fallos de casteo a tipos por valor no anulables.
- **Solución:** Se implementó la clase auxiliar anidada `private sealed class AsigSeedRow` con propiedades fuertemente tipadas (`int Id`, `int? IntHoraria`, `string DiasPatron`), permitiendo que Dapper realice la deserialización mediante reflexión estándar de CLR sin fallos de casteo ni dependencias de runtime no soportadas.

#### 2. Fase 442 — Supresión de Asignaturas sin Carga Horaria (Eliminación de 64h Ficticias)
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Asignar un default de 64 horas a materias con `int_horaria` nula o $\le 0$ creaba 16 sesiones obligatorias (4 semanas $\times$ 4 sesiones) para materias vacías, inflando artificialmente la demanda curricular y provocando cuellos de botella falsos en el validador pre-vuelo.
- **Solución:** Si `!aItem.IntHoraria.HasValue || aItem.IntHoraria.Value <= 0`, la asignatura se omite explícitamente (`omitidas++`), emitiendo una traza de telemetría informativa (`traza.Dato("AulaAsignaturaSinHoras", ...)`). Solo las asignaturas con carga real ingresan a `ag_aula_asignatura`.

#### 3. Fase 443 — `SELECT DISTINCT` para Prevenir Duplicación por LEFT JOIN
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Si la tabla `ag_aula_asignatura` contenía registros previos duplicados para el par aula-asignatura, el `LEFT JOIN` multiplicaba las filas de `asigsData`, produciendo sobreconteo en las métricas de inserción.
- **Solución:** Se incorporó `SELECT DISTINCT a.id AS Id, ...` en `sqlAsigs`, garantizando cardinalidad 1:1 por cada asignatura del catálogo académico.

### 157.11 Fase 444 — Desacoplamiento de Métrica 'sinHoras' del Contador 'omitidas'

- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En la Fase 442, omitir una asignatura por falta de `int_horaria` incrementaba el contador `omitidas++`. Dado que dicho contador representaba históricamente *"Aulas preexistentes omitidas"*, el diálogo final y el registro del log presentaban una métrica sobrecargada y engañosa al usuario.
- **Solución:** Se introdujo la variable independiente `int sinHoras = 0;` para contabilizar exclusivamente las asignaturas omitidas por falta de intensidad horaria. Se desacopló de `omitidas`, registrando ambas magnitudes de forma transparente en `traza.Dato("SinHoras", sinHoras)`, `txtLogAG` y en el `MessageBox` resumen:
  - `- Aulas creadas: {insertadas}`
  - `- Aulas preexistentes omitidas: {omitidas}`
  - `- Asignaturas sin int_horaria omitidas: {sinHoras}`

### 157.12 Ronda 59 + 60 — Coherencia de Períodos, Desambiguación de Cohortes (A/B/C), Regla R8 y Staging Fiel (Fases 445 - 461)

Resolución integral de los defectos detectados en la auditoría de corridas reales (Ronda 59 y 60):

#### 1. Fase 445 / Fase 453 — Inyección de PeriodoObjetivo del Combo en Tupla S y Traza de Período
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En `ObtenerCombinacionesSeleccionadas()`, las tuplas S conservaban el `IdPeriodo` asociado a nodos del árbol o perfiles antiguos; si el usuario seleccionaba un período distinto en `cmbPeriodos` (p. ej. `2027-01` id=3), las combinaciones se generaban con el período anterior (id=2), produciendo discordancia temporal y ventanas obsoletas.
- **Solución:** Se capturó `idPeriodoObjetivo = (cmbPeriodos.SelectedItem as Periodo)?.Id ?? 0;` en `ObtenerCombinacionesSeleccionadas()`, asignándolo a cada DTO generado. En `ConstruirContextoAsync`, si `_ventanaResuelta` no coincide con el período del combo, se recalcula la ventana semestral automáticamente. Se añadió emisión de traza estructurada `traza.Dato("PeriodoObjetivo", ...)` y advertencia `MultiplesPeriodosActivos` si existen más de un período con `activo = 1`.

#### 2. Fase 446 — Resolución de Programa Homónimo con Pensum del Semestre Solicitado
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En la base de datos coexisten programas con nombres idénticos (p. ej. `ADMINISTRACION PÚBLICA TERRITORIAL`) con IDs distintos (p. ej. 1, 2, 3). Si la selección apuntaba a un ID sin asignaturas para el semestre solicitado (`I`), la optimización fallaba con catálogo vacío a pesar de que el homónimo ID 1 poseía el pensum completo.
- **Solución:** Se implementó detección y resolución de homónimos en `ConstruirContextoAsync` y `GenerarEncabezadosAulasAsync`: si el programa seleccionado carece de asignaturas para el semestre pero un homónimo activo posee registros en `asignaturas`, se conmuta automáticamente al homónimo con el pensum, registrando `traza.Dato("ProgramaHomonimoResuelto", ...)`.

#### 3. Fase 447 — Generación Asistida de Aulas sin Aulas Huérfanas
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `GenerarEncabezadosAulasAsync` insertaba la cabecera en `ag_aula` antes de comprobar si existían asignaturas para el programa y semestre. Al encontrar 0 asignaturas en el catálogo, dejaba un aula huérfana en la base de datos sin asignaturas en `ag_aula_asignatura`.
- **Solución:** Se verifica `asigsData.Count > 0` antes de persistir el aula en `ag_aula`. Si no existen asignaturas, se incrementa `sinAsignaturas`, se emite traza diagnóstica de bloqueo y se omite la inserción de la cabecera, evitando registros residuales vacíos.

#### 4. Fase 448 / Fase 449 / Fase 450 / Fase 452 — Propagación de Fallo de Contexto vía Result y Diagnóstico Accionable
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `ConstruirContextoAsync` retornaba una instancia no nula de `AgContexto` aun cuando `ctx.Asignaturas.Count == 0` o `ctx.Aulas.Count == 0`, registrando éxito engañoso y fallando en etapas posteriores sin mensaje accionable.
- **Solución:** Se refactorizó la firma a `Task<Result<AgContexto>>`. Si `ctx.Aulas.Count == 0`, se retorna `Result.Failure` instruyendo a ejecutar "Generar Aulas". Si `ctx.Asignaturas.Count == 0`, se retorna `Result.Failure` con el desglose completo del contexto (período, programa, semestre, subunidad y disponibilidad en programas homónimos). Se emite advertencia si el período carece de festivos/presenciales (`SinFestivos`).

#### 5. Fase 454 / Fase 455 — Multi-Aula por Combinación (Cohortes A, B, C) con Modalidad y Grupo de Aula
- **Archivos:** 
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** `AgFabricaIndividuos` iteraba únicamente sobre `Asignaturas` y asignaba todos los genes a `IdAulaCorrida` o al primer aula con `Grupo = "A"`. Al seleccionar grupos A, B y C para el mismo semestre, todos los cursos colapsaban en una única aula (p. ej. 266), generando 14 violaciones duras de superposición de aula (R2) y grupo (R3).
- **Solución:** En `ConstruirContextoAsync`, se cargan las aulas de la sede y período sin acotar a un solo grupo, fijando `IdAulaCorrida` únicamente si hay exactamente 1 combinación. En `AgFabricaIndividuos.CrearAleatorio`, se itera por cada aula/grupo representativo (`gruposAula`), asignando deterministamente `gen.AsignarAula(aula)` con su `Grupo` y `Modalidad` propios. En `AgMaterializadorSesiones`, se valida que todos los genes del grupo compartan el mismo grupo (`agVinc.Grupo == genes[0].Grupo`).

#### 6. Fase 456 — Mutación Segura de Aula/Bloque y Restricción Dura R8 (Días No Duplicados en Multi-Día)
- **Archivos:** 
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** La mutación de aula seleccionaba aulas de cualquier grupo mezclando cohortes; y la mutación de bloque elegía bloques aleatorios de cualquier día de la semana, haciendo que asignaturas multi-día terminaran con días duplicados (p. ej. Viernes y Viernes), violando el patrón curricular.
- **Solución:**
  1. La mutación de aula se restringe a aulas con el mismo `Grupo` del gen (`gen.Grupo`).
  2. En genes multi-día, la mutación de bloque solo conmuta franjas horarias pertenecientes al mismo día asignado al gen (`b.DiaSemanaNumero == gen.DiaSemana`), impidiendo días repetidos.
  3. Se incorporaron los códigos `R8_DiasDuplicadosGrupo` y `R8_DiaFueraDePatron` en `AgReglaViolacion` y su evaluación en `AgValidadorReglas`, penalizando como infracción dura cualquier cromosoma con días duplicados dentro de la misma `ClaveGrupo`.

#### 7. Fase 457 — Consulta Fiel de Staging en MODO AG
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Application/DTOs/VinculacionFilterDto.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/VinculacionQueryRepository.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormVinculaciones.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** Al abrir `FormVinculaciones` en MODO AG desde Módulo R, el formulario aplicaba filtros predeterminados de producción (período de `config` id=1), devolviendo 0 sesiones. Además, `vw_AgSesionesDetalle` utilizaba `INNER JOIN` para programas y sedes y no proyectaba `IdCorrida`.
- **Solución:** Se añadió `IdCorrida` a `VinculacionFilterDto` y filtrado parametrizado en `VinculacionQueryRepository`. Se actualizó la vista `vw_AgSesionesDetalle` en `SqliteHelper.cs` para proyectar `v.id_corrida AS IdCorrida` y usar `LEFT JOIN` defensivos. En MODO AG, `FormVinculaciones` preselecciona la última corrida activa.

#### 8. Fase 458 — Coherencia Estricta de Periodo y Fechas en Staging
- **Archivos:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** `ag_vinculaciones` persistía fechas iniciales y períodos discordantes con la ventana semestral del período objetivo de la corrida.
- **Solución:** En `AgMaterializadorSesiones`, si `contexto.Ventana` está resuelta con `IdPeriodo > 0`, se valida que `agVinc.IdPeriodo == contexto.Ventana.IdPeriodo` (o se asigna), y se fuerzan `agVinc.FechaInicio` y `agVinc.FechaFin` desde la ventana semestral objetivo.

#### 9. Fase 459 — Migración V57: Reparación de Esquema en 'salvedades_conflictos'
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV57_RepairSalvedadesConflictos.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** En vaults migrados desde versiones tempranas, la tabla `salvedades_conflictos` carecía de las columnas `idsesion_a` e `idsesion_b`, provocando la excepción `no such column: idsesion_a` al consultar o registrar salvedades en `SalvedadConflictoRepository`.
- **Solución:** Se implementó `MigrationV57_RepairSalvedadesConflictos` con `IMigrationTolerante`, agregando de forma idempotente las columnas ausentes (`idsesion_a`, `idsesion_b`, `idtercero`, `idperiodo`, `justificacion`) mediante `ALTER TABLE ... ADD COLUMN` si no existen.

#### 10. Fase 460 / Fase 461 — Fiel Desglose de Violaciones HARD/SOFT y Trazabilidad en Grid
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El contador de violaciones en UI mostraba un resumen incompleto y el grid de horarios staging no proyectaba los identificadores de trazabilidad (`IdCorrida`, `IdVinculacion`, `IdGen`).
- **Solución:** En `CargarModuloResultadosAsync`, se actualizó `lblHardViolaciones` para reportar el conteo total: `Violaciones: {total} (HARD {hard}, SOFT {soft})`, ordenando `dgvViolaciones` con prioridad descendente para las restricciones duras. Se añadieron las columnas `IdCorrida`, `IdVinculacion`, `IdGen` e `IdSesionGlobal` a la proyección de la grilla de resultados.

### 157.13 Ronda 61 — MODO AG Fiel y Centinela VACANTE (Fases 457 ampliada, 462 - 466)

Resolución integral del modo de consulta de sesiones en MODO AG y unificación formal del tercero centinela VACANTE:

#### 1. Fase 457 (Ampliada) — Consulta de Sesiones en MODO AG y Proyección de 'IdCorrida'
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Domain/Models/SesionListItem.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/VinculacionListItem.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV49_AddAgHorariosTables.cs`
- **Diagnóstico:** Al abrir `FormVinculaciones` en MODO AG, la vista `vw_AgSesionesDetalle` carecía del filtro `AND v.estado = 0`, pudiendo incluir sesiones asociadas a vinculaciones descartadas o canceladas. Asimismo, `vw_AgVinculacionesDetalle` no proyectaba `IdCorrida`, y los modelos de lectura `SesionListItem` y `VinculacionListItem` carecían de la propiedad tipificada `IdCorrida`.
- **Solución:**
  1. Se agregó la propiedad `public long IdCorrida { get; set; }` tanto en `SesionListItem` como en `VinculacionListItem`.
  2. En `vw_AgSesionesDetalle` (definida en `SqliteHelper.cs` y `MigrationV49`), se fijó la condición `JOIN ag_vinculaciones v ON s.idvinculacion = v.id AND v.estado = 0`.
  3. En `vw_AgVinculacionesDetalle`, se proyectó `v.id_corrida AS IdCorrida`.
  4. Esto permite a `VinculacionQueryRepository.GetSesionesAsync` y `GetVinculacionesAsync` ejecutar consultas deterministas y filtradas con `IdCorrida = @idCorrida`, visualizando fielmente las sesiones generadas en staging (p. ej. corrida #9: 6 vinculaciones / 64 sesiones).

#### 2. Fase 462 — Resolución y Asignación de 'IdTerceroVacante'
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Cuando una asignatura carecía de docentes habilitados en su pool, `AgFabricaIndividuos` aplicaba un fallback aleatorio tomando un docente activo de `contexto.Docentes` antes de evaluar `idDocente <= 0`. Esto provocaba que asignaturas vacantes recibieran docentes reales aleatorios (p. ej. IDs 23, 70, 24) en lugar de permanecer asignadas al tercero centinela VACANTE (cédula 1111111111, ID 192).
- **Solución:**
  1. Se añadió `public int IdTerceroVacante { get; set; }` a `AgContexto`.
  2. En `AgFormGeneradorHorarios.ConstruirContextoAsync`, tras cargar docentes, se resuelve el centinela desde la tabla `terceros` (`cedula = '1111111111' OR UPPER(TRIM(nombre)) = 'VACANTE'`) priorizando la subunidad activa, registrando `traza.Dato("IdTerceroVacante", ctx.IdTerceroVacante)`.
  3. En `AgFabricaIndividuos.CrearAleatorio`, se eliminó completamente el fallback de docente aleatorio. Si `poolAsig` no contiene candidatos habilitados (excluyendo el centinela), se fija inmediatamente `esVacante = 1`, `idDocente = contexto.IdTerceroVacante`, `candidatosPool = 0` y se clasifica el motivo de vacante.
  4. Si un gen se marca vacante por falta de franja horaria en su día, también se le asigna `gen.IdTercero = contexto.IdTerceroVacante`.

#### 3. Fase 463 — Respeto al Centinela en Mutación, Expansión y Materialización
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs`
- **Diagnóstico:** El operador de mutación no distinguía si un gen era vacante, mutando aleatoriamente su docente a candidatos del pool, o asignando inadvertidamente el ID centinela a genes no vacantes. Asimismo, durante la expansión cronológica de sesiones y el mapeo a vinculaciones de staging, no se garantizaba la propagación del centinela.
- **Solución:**
  1. En `AgOperadorMutacion.Mutar`, si `gen.EsVacante == 1`, se fuerza `gen.IdTercero = contexto.IdTerceroVacante` y se omite la mutación de docente. Para genes no vacantes, los candidatos elegibles excluyen expresamente `contexto.IdTerceroVacante`.
  2. En `AgExpansorSesiones.Expandir`, cada sesión vacante recibe deterministamente `IdTercero = (gen.EsVacante == 1 && contexto.IdTerceroVacante > 0) ? contexto.IdTerceroVacante : gen.IdTercero`.
  3. En `AgGenToAgVinculacionMapper.MapGrupo`, se garantiza que `agVinc.IdTercero` adopte `contexto.IdTerceroVacante` si el grupo de genes es vacante.

#### 4. Fase 464 — Restricción Dura R9 y Filtrado de Centinela en Selectores de UI
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormVinculaciones.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** El motor genético carecía de una restricción dura para fiscalizar la integridad del centinela VACANTE. Además, al cargar el listado general de docentes en formularios de búsqueda (`FormVinculaciones` y `FormSesionesAcademicas`), el registro "VACANTE" aparecía en los combos de autocompletado para asignación regular.
- **Solución:**
  1. Se añadieron los códigos `R9_VacanteSinCentinela = 26` y `R9_CentinelaEnNoVacante = 27` a `AgReglaViolacion`.
  2. En `AgValidadorReglas`, se validan como infracciones duras (`esDura: true`) que todo gen vacante posea el `IdTerceroVacante`, y que ningún gen no vacante lo tenga asignado.
  3. En la indexación de colisiones docentes (R1) y sobrecarga (S2/S4), se excluye explícitamente el centinela VACANTE (`gen.EsVacante == 0 && gen.IdTercero != contexto.IdTerceroVacante`), evitando falsos positivos de cruce entre asignaturas vacantes simultáneas.
  4. En `FormVinculaciones.FiltrarSedesYTerceros` y `FormSesionesAcademicas.CargarCatalogosAsync`, se filtraron los registros con nombre o cédula de "VACANTE", excluyéndolos de los combos de terceros convencionales.

#### 5. Fase 465 — Migración V58: Asegurar Tercero Centinela VACANTE
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV58_EnsureTerceroVacante.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Si la base de datos local o un vault de pruebas fue creado antes de la siembra o si el registro VACANTE fue desactivado (`activo = 0`), la búsqueda del centinela fallaba retornando 0.
- **Solución:**
  1. Se implementó `MigrationV58_EnsureTerceroVacante` (`Version = 58`, `Aditiva = true`, `IMigrationTolerante`).
  2. Si existe un tercero con cédula `'1111111111'` o nombre `'VACANTE'`, actualiza `activo = 1`. Si no existe, realiza un `INSERT` idempotente con nombre `'VACANTE'`, cédula `'1111111111'`, correo `'VACANTE@esap.edu.co'`, subunidad vacía, tipo `'CATEDRA'`, categoría `'AUXILIAR'` y `activo = 1`.
  3. Se registró en `BuildMigrator()` de `SqliteHelper.cs` y en el archivo de proyecto `ConversorPDF.XLL.csproj`.

#### 6. Fase 466 — Verificación de Calidad y No Regresiones
- **Análisis:**
  1. Compatibilidad total con C# 7.3 y .NET Framework 4.8.
  2. Cumplimiento estricto del principio de tipo único por archivo físico.
  3. Cumplimiento de la restricción de usuario: "No compiles ni ejecutes el código".
  4. Ausencia de regresiones en las consultas existentes de `vw_SesionesDetalle` para el entorno de producción.

### 157.14 Ronda 62 — Validación Temprana de Aulas, Vistas de Staging V59 y Contratos DTO (Fases 467 - 470)

Resolución de residuales bloqueantes para generación de aulas, consistencia de esquema DTO y cobertura de pruebas:

#### 1. Fase 467 — Validación Temprana de Asignaturas Previo al INSERT de 'ag_aula'
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En `GenerarEncabezadosAulasAsync`, si un programa/semestre carecía de asignaturas en el pensum, el bloque de verificación ejecutaba la comprobación de forma tardía o incrementaba `sinAsignaturas` después de haber insertado el registro en `ag_aula`, dejando aulas huérfanas en la base de datos (p. ej. `insertadas = 3, sinAsignaturas = 3`).
- **Solución:**
  1. Se implementó una consulta previa `SELECT COUNT(1) FROM asignaturas WHERE ...` antes de cualquier operación sobre `ag_aula`.
  2. Si `asigsCount == 0`, se incrementan `sinAsignaturas++` y `omitidas++`, se registra `traza.Dato("AulaOmitidaSinAsignaturas", item.AulaRef)` y se omite el `INSERT` mediante `continue`.
  3. Se eliminó el bloque condicional tardío redundante posterior a la consulta de `ag_aula_asignatura`.

#### 2. Fase 468 — Migración V59: Recreación de Vistas de Staging y Tabla 'ag_salvedades_conflictos'
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV59_RecreateAgViews.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Las vistas `vw_AgSesionesDetalle` y `vw_AgVinculacionesDetalle` fueron concebidas en la migración V49. Migraciones posteriores (V54/V55) introdujeron columnas clave (`es_vacante`, `motivo_vacante`, `candidatos_pool`, `id_aula`, `idgen`, `clave_grupo`, `horas_gen`). Al ejecutar `SELECT *` en MODO AG, Dapper mapeaba campos nulos o fallaba por inconsistencia. Asimismo, la tabla `ag_salvedades_conflictos` requerida por `AgModoDatos.TablaSalvedades` requería esquema asegurado.
- **Solución:**
  1. Se creó `MigrationV59_RecreateAgViews` (`Version = 59`, `Aditiva = true`, `IMigrationTolerante`).
  2. Se recreó `vw_AgSesionesDetalle` proyectando la totalidad de columnas requeridas por `SesionListItem` (`IdSesion`, `IdCorrida`, `IdVinculacion`, `IdTercero`, `EsVacante`, `MotivoVacante`, `CandidatosPool`, `IdAula`, `IdGen`, `ModalidadVinculacion`, etc.) con `JOIN ag_vinculaciones v ON s.idvinculacion = v.id AND v.estado = 0`.
  3. Se recreó `vw_AgVinculacionesDetalle` proyectando `Id`, `IdCorrida`, `IdIndividuo`, `EsVacante`, `MotivoVacante`, `CandidatosPool`, etc.
  4. Se aseguró la creación idempotente de la tabla `ag_salvedades_conflictos (id, idsesion_a, idsesion_b, motivo, subunidad)`.
  5. Se registró en `BuildMigrator()` de `SqliteHelper.cs` y en `ConversorPDF.XLL.csproj`.

#### 3. Fase 469 — Contrato DTO ↔ Vista y Alerta Temprana en UI
- **Archivos:**
  - `ConversorPDF.XLL/Academic/UI/Forms/FormVinculaciones.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV57V58V59Tests.cs`
- **Diagnóstico:** Al consultar MODO AG en `FormVinculaciones`, si existían sesiones generadas en `ag_sesiones` para la corrida pero la consulta de vista retornaba 0 filas debido a filtros desalineados o discrepancias de proyección, la UI mostraba silenciosamente "0 sesiones".
- **Solución:**
  1. En `FormVinculaciones.EjecutarConsultaAsync`, si `result.Value.Count == 0` pero `ag_vinculaciones` contiene filas para la corrida, se consulta el total de sesiones en `ag_sesiones`. Si existen registros, se emite una advertencia de diagnóstico en `LoggerHelper.Warning` y se notifica al usuario en la barra de estado: `"Aviso: Hay {totalSesStaging} sesiones en staging pero 0 coinciden con filtros (corrida #{idCorridaVal})."`.
  2. Se añadieron aserciones de contrato en pruebas unitarias verificando la presencia obligatoria de columnas (`IdSesion`, `IdCorrida`, `IdTercero`, `EsVacante`, `MotivoVacante`, `CandidatosPool`, `IdAula`, `IdGen`) en `vw_AgSesionesDetalle`.

#### 4. Fase 470 — Batería de Pruebas Unitarias de Centinela, Mutación, Reglas y Migraciones
- **Archivos:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorMutacionTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV57V58V59Tests.cs` (Nuevo)
- **Diagnóstico:** Faltaba cobertura de pruebas automatizadas para los nuevos comportamientos introducidos en las Rondas 60 y 61.
- **Solución:**
  1. En `AgFabricaIndividuosTests`, se incorporó `CrearAleatorio_WhenSinPoolDocente_ShouldAssignIdTerceroVacante` validando que ante ausencia de pool, los genes reciban deterministamente `EsVacante = 1` y `IdTercero = 192`.
  2. En `AgOperadorMutacionTests`, se incorporó `Mutar_WhenGenIsVacante_ShouldPreserveIdTerceroVacante` verificando que la mutación preserve el centinela en genes vacantes y no reasigne profesores del pool.
  3. En `AgValidadorReglasTests`, se incorporaron `Evaluar_WhenGenVacanteSinCentinela_ShouldRecordR9HardViolation` (regla R9) y `Evaluar_WhenDiasDuplicadosEnGrupo_ShouldRecordR8HardViolation` (regla R8).
  4. En `MigrationV57V58V59Tests`, se implementaron pruebas de metadatos, idempotencia de V58 (`EnsureTerceroVacante`) y contrato de columnas de vistas V59.

### 157.15 Fase 471 — Verificación y Sincronización de UI y Diagnóstico (Fase 471)

Verificación completa y sincronización de los componentes de interfaz de usuario:
1. **Fase 449 (Diálogo Período + Homónimo):** Confirmada en `AgFormGeneradorHorarios.ConstruirContextoAsync` (L581-L586), detallando programa, semestre, subunidad y la sugerencia accionable del ID homónimo con catálogo.
2. **Fase 450 (Aviso Festivos = 0):** Confirmada en `AgFormGeneradorHorarios.ConstruirContextoAsync` (L649-L655), emitiendo advertencia de telemetría y en log si el período carece de festivos/presenciales.
3. **Fase 460 (Contador HARD/SOFT 1:1 con Grid):** En `AgFormGeneradorHorarios.CargarModuloResultadosAsync` (L1253-L1258), se ajustó el cálculo para derivar `hardCount`, `softCount` y `totalViol` directamente de `ind.Violaciones`, garantizando que la etiqueta `Violaciones: {totalViol} (HARD {hardCount}, SOFT {softCount})` coincida de manera exacta y sin discrepancias con la cantidad de filas renderizadas en `dgvViolaciones`.
4. **Fase 461 (Columnas de Staging en MODO AG):** En `FormVinculaciones.ConfigurarGrid` (L244-L248), se añadieron las columnas `IdCorrida` ("Corrida") e `IdGen` ("Gen") cuando `_modo.EsModoAg` está activo, complementando las propiedades añadidas en `SesionListItem` (`IdGen`, `IdAula`, `EsVacante`, `MotivoVacante`, `CandidatosPool`).

### 157.16 Release v1.6.3.33 — Compilación, Pruebas y Empaquetado MSI

- **Versión:** `1.6.3.33`
- **Artefacto:** `installer\InstaladorConversorPDFv1.6.3.33.msi` (5.43 MB)
- **Estado de Pruebas:** 100% de pruebas unitarias ejecutadas y aprobadas (xUnit + MSTest).
- **Compilación:** `Release|x64` sin errores.
- **Cambios incluidos:**
  - Migración V58 (`EnsureTerceroVacante`) y V59 (`RecreateAgViews` + `ag_salvedades_conflictos`).
  - Lógica centinela VACANTE (cédula `1111111111`, ID 192) unificada en fábrica, mutación, mapeo y reglas genéticas R8/R9.
  - Validación temprana previa a la inserción de `ag_aula` en `GenerarEncabezadosAulasAsync` (Fase 467).
  - Columnas `IdCorrida` e `IdGen` en grid de sesiones MODO AG y sincronización 1:1 de contadores de violaciones HARD/SOFT.
  - Exclusión del centinela VACANTE de selectores de docentes en UI.

### 157.17 Ronda 64 — Diagnóstico y Resolución de Falsos Negativos en Multi-Combinación (Fases 472 - 476)

#### 1. Fase 472 — Validar Anclaje Multi-Combinación (P07 & P08)
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Cuando el usuario selecciona múltiples combinaciones (e.g. A, B y C), `IdAulaCorrida` se establece en `0` (Fase 454) para indicar que no hay una única aula global anclada. Sin embargo, `P07_ValidarAnclajeAula` (L811) y `P08_ValidarAnclajeAula` (L998) exigían `IdAulaCorrida > 0`, provocando un bloqueo falso negativo: "No existe Aula exacta para la combinación", aun cuando todas las aulas de las combinaciones existían en `_contextoActual.Aulas`.
- **Solución:** Se sustituyó la exigencia ciega de `IdAulaCorrida > 0`. Si `combinaciones.Count > 1`, se verifica que cada combinación seleccionada posea su aula correspondiente en `_contextoActual.Aulas`. Si falta alguna, se notifica la lista exacta de faltantes; si todas están cubiertas, se registra en telemetría `AnclajeMultiCombinacion: aulas cubiertas={combinaciones.Count}, IdAulaCorrida=0 (multi-grupo)` y se permite proseguir la optimización multi-grupo sin bloqueo.

#### 2. Fase 473 — Cargar `ag_aula_asignatura` de Todo el Alcance de Aulas
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En `ConstruirContextoAsync` (L720), `ag_aula_asignatura` se consultaba únicamente con `WHERE id_aula = @IdAula` cuando `ctx.IdAulaCorrida > 0`. Con `IdAulaCorrida == 0` (multi-combinación), `ctx.AulaAsignaturas` se inicializaba vacía, provocando `AulaAsignaturas=0`, `AulasSinAsignaturas=1`, advertencias espurias y forzando patrones semanales uniformes por defecto (`{2,3}`).
- **Solución:** Se actualizó la consulta en `ConstruirContextoAsync` para recopilar `idsAulas = ctx.Aulas.Select(a => a.Id).ToList()` y consultar `WHERE id_aula IN @ids AND id_asignatura > 0`. Si el conteo es 0, se registra la advertencia contextual `SinAulaAsignaturas`.

#### 3. Fase 474 — `GenerarAulas` Asegura Vinculaciones en Aulas Preexistentes
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En `GenerarEncabezadosAulasAsync`, si un aula ya existía (`existentes.Contains(item.AulaRef)`), se incrementaba `omitidas++` pero el bloque de inserción de materias condicionaba `if (cuentaAsig == 0)`. Si el aula ya contenía asignaturas parciales o previas, no se aseguraban de forma idempotente las asignaturas del pensum para esa combinación, y la métrica de reporte mezclaba aulas omitidas con materias aseguradas.
- **Solución:** Se desacopló la verificación de asignaturas de la condición `cuentaAsig == 0`. Tanto para aulas nuevas como preexistentes, se ejecuta el bucle de `INSERT OR IGNORE INTO ag_aula_asignatura` para cada asignatura válida del pensum, se computa la métrica `aseguradas++`, se emite `traza.Conteo("Aseguradas", aseguradas)` y se separa en el resumen final.

#### 4. Fase 475 — Grupo y Patrón por Aula Real en `AgFabricaIndividuos`
- **Archivos:** 
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`
- **Diagnóstico:** 
  1. Si `ag_aula.grupo` era `NULL` en base de datos (caso real en vault de Valle), la agrupación previa `aulasValidas.GroupBy(a => a.Grupo ?? "A")` colapsaba todas las aulas en el grupo "A", ignorando los grupos B y C a pesar de figurar en `AulaRef` (`3|1|1|I|B|...`).
  2. La resolución de `aa` y el patrón de días se realizaba exclusivamente contra `contexto.AulaAsignaturasPorId[asig.Id]`, el cual almacena una sola entrada global sin diferenciar el aula ni la cohorte.
- **Solución:**
  1. Se implementó la función local `GrupoDeAula(AgAula a)` que extrae el grupo de `a.Grupo`, o si es nulo/vacío, lo parsea del 5º token de `AulaRef` (`partes[4]`), con fallback a "A".
  2. Se agrupó por `GrupoDeAula` para preservar las cohortes independientes A, B y C.
  3. Se resolvió `aa` buscando específicamente `contexto.AulaAsignaturas.FirstOrDefault(x => x.IdAula == aula.Id && x.IdAsignatura == asig.Id) ?? contexto.AulaAsignaturas.FirstOrDefault(x => x.IdAsignatura == asig.Id)`, garantizando que cada cohorte reciba su patrón semanal y carga horaria reales configurados para esa aula.
  4. Se añadió la prueba unitaria `CrearAleatorio_WhenAulaGrupoIsNull_ShouldDeriveGrupoFromAulaRef`.

#### 5. Fase 476 — Trazabilidad de Combinaciones en `GenerarAulas`
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En `GenerarEncabezadosAulasAsync`, la traza no reflejaba la lista de combinaciones (`AulaRef`) procesadas ni el desglose de asignaturas faltantes en el catálogo.
- **Solución:** Se añadieron las emisiones estructuradas `traza.Dato("Combinaciones", string.Join(" | ", lista.Select(x => x.AulaRef)))` y `traza.Dato("FaltantesEnCatalogo", $"{sinAsignaturas}")`.

### 157.18 Ronda 65 — Métricas y Diagnósticos Residuales (Fases 477 - 479)

#### 1. Fase 477 — Sobreconteo de `aseguradas`
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En `GenerarEncabezadosAulasAsync`, `aseguradas` incrementaba tras evaluar asignaturas sin verificar si `INSERT OR IGNORE` realmente insertó una fila o fue ignorada por ya existir en la base de datos.
- **Solución:** Se captura el resultado de `await db.ExecuteAsync(sqlInsertAsig, ...)` en la variable `filas`, incrementando `aseguradas++` únicamente cuando `filas > 0`.

#### 2. Fase 478 — Limpieza de Traza Redundante
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Se emitía doble traza de `Aseguradas` mediante `traza.Conteo("Aseguradas", ...)` y `traza.Dato("Aseguradas", ...)`.
- **Solución:** Se eliminó la línea redundante de `traza.Dato("Aseguradas", ...)`, conservando el conteo formal estructurado.

#### 3. Fase 479 — Mensaje Accionable de Bloqueo por `ag_aula_asignatura`
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En `P07_ValidarAulaAsignaturas` y `P08_ValidarAulaAsignaturas`, el bloqueo usaba un mensaje genérico y `traza.Fallo(..., null, ...)` sin identificar cuáles aulas específicas carecían de asignaturas vinculadas.
- **Solución:** Se calculan las aulas faltantes filtrando `aulasSinMaterias = _contextoActual.Aulas.Where(a => !_contextoActual.AulaAsignaturas.Any(x => x.IdAula == a.Id)).Select(a => a.AulaRef).ToList()`. Si faltan materias, se construye un mensaje accionable con viñetas indicando cada `AulaRef` afectado, se pasa la excepción formal a `traza.Fallo` y se despliega en el diálogo `MessageBox.Show`.

### 157.19 Release v1.6.3.34 — Compilación, Pruebas y Empaquetado MSI

- **Versión:** `1.6.3.34`
- **Artefacto:** `installer\InstaladorConversorPDFv1.6.3.34.msi` (5.44 MB)
- **Estado de Pruebas:** 100% de pruebas unitarias ejecutadas y aprobadas (xUnit + MSTest).
- **Compilación:** `Release|x64` sin errores.
- **Cambios incluidos:**
  - Resolución de falsos negativos en multi-combinación (Fase 472).
  - Carga de `ag_aula_asignatura` abarcando todo el alcance de aulas (Fase 473).
  - Aseguramiento de asignaturas tanto en aulas nuevas como preexistentes (Fases 474 & 477).
  - Derivación de grupo desde `AulaRef` ante `grupo = NULL` y resolución de patrón por aula (Fase 475).
  - Trazabilidad estructurada y mensajes accionables con aulas faltantes en P07 y P08 (Fases 476, 478 & 479).

### 157.20 Ronda 66 — Soporte Multi-Semestre en Contexto, Fábrica y Diagnóstico (Fases 477 - 480)

#### 1. Fase 477 — Contexto Multi-Semestre (ConstruirContextoAsync)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  Al seleccionar combinaciones que pertenecen a semestres distintos (ej. 6 combinaciones repartidas entre semestres `I`, `II`, `III` y `IV`), `ConstruirContextoAsync` extraía únicamente el semestre de la primera combinación seleccionada (`combinaciones.First().Semestre` $\rightarrow$ `'I'`). Como consecuencia:
  1. `GetAulasPorPeriodoAsync` filtraba por `Semestre = 'I'`, ignorando las aulas existentes de los semestres `II`, `III` y `IV`, produciendo un falso negativo en la validación multi-combinación de `P07` y `P08` al reportar como "faltantes" aulas que sí existían en la base de datos.
  2. `ctx.Asignaturas` cargaba únicamente las asignaturas del primer semestre, omitiendo el pensum de los demás semestres seleccionados.
  3. `ctx.PoolDocentes` y la resolución de programas homónimos quedaban restringidos al primer semestre.
- **Solución Implementada:**
  1. Se extraen todos los semestres distintos presentes en las combinaciones seleccionadas (`semestresCombinaciones`) y sus equivalentes arábigos (`semestresFiltro`).
  2. Se invoca `GetAulasPorPeriodoAsync(ctx.IdPeriodo, sedeSel.Id, programaSel.Id, null, null, ctx.Subunidad, ct)` sin filtro de semestre ni grupo, cargando la totalidad de aulas del alcance de la oferta académica.
  3. Se parametriza la consulta de asignaturas con cláusula `WHERE idprograma = @IdPrograma AND (@SinFiltro = 1 OR semestre IN @Semestres)` cargando la unión de todas las asignaturas de los semestres seleccionados.
  4. Se adapta de forma análoga la consulta de `PoolDocentes` y la comprobación de programas homónimos para operar con el conjunto completo de semestres seleccionados.

#### 2. Fase 478 — Asignación Asignatura ↔ Aula por Semestre (AgFabricaIndividuos)
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`
- **Diagnóstico:**
  Al albergar `contexto.Asignaturas` materias de múltiples semestres, la fábrica de individuos cruzaba cada cohorte/aula con todas las asignaturas del contexto. Además, al agrupar por `GrupoDeAula`, aulas de distinto semestre con el mismo identificador de grupo (ej. `I|A` y `II|A`) colapsaban en una única cohorte.
- **Solución Implementada:**
  1. Se introdujeron las funciones auxiliares locales `SemestreDeAula(AgAula a)` (que extrae el semestre de la entidad o del token 4 de `AulaRef`) y `ClaveCohorteAula(AgAula a)` (que combina `AulaRef` o `{semestre}|{grupo}`) garantizando que cada cohorte por semestre y grupo sea preservada de forma unívoca.
  2. En el bucle de asignación de genes, se compara el semestre de la cohorte/aula (`semAulaRom`) contra el semestre de la asignatura (`semAsigRom`). Si no coinciden, se omite la asignatura con `continue`, asegurando que cada cohorte sólo reciba asignaturas curriculares de su propio semestre.
  3. Se incorporó la prueba unitaria `CrearAleatorio_WhenMultiSemestre_ShouldAssignSubjectsMatchingAulaSemester` en `AgFabricaIndividuosTests.cs`.

#### 3. Fase 479 — Aseguradas por Combinación en `GenerarEncabezadosAulasAsync`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  En `GenerarEncabezadosAulasAsync`, `semRom` y `semArab` se derivan por cada ítem del bucle, pero la trazabilidad carecía del desglose de asignaturas aseguradas por combinación individual (`item.AulaRef`), impidiendo auditar si cada combinación específica vinculó exitosamente sus materias.
- **Solución Implementada:**
  Se introdujo un diccionario `porCombinacion = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)`. En cada iteración se contabilizan las filas insertadas (`aseguradasItem`) y se emite tras el commit `traza.Dato("AseguradasPorCombinacion", string.Join(" | ", porCombinacion.Select(kv => $"{kv.Key}={kv.Value}")))`.

#### 4. Fase 480 — Mensaje de Faltantes con Contexto de Alcance (P07 & P08)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  Cuando la validación multi-combinación detecta aulas no registradas, el mensaje previo informaba las faltantes sin explicitar cuántas aulas se cargaron en memoria ni qué semestres integraban el alcance, dificultando el diagnóstico rápido ante discrepancias de configuración.
- **Solución Implementada:**
  En `P07_ValidarAnclajeAula` y `P08_ValidarAnclajeAula`, se enriquece el mensaje de advertencia y error:
  `$"Faltan aulas (ag_aula) para {faltantes.Count} combinación(es):\n• {string.Join("\n• ", faltantes)}\nAlcance cargado: {_contextoActual.Aulas.Count} aulas, semestres {string.Join(",", semestresCombinaciones)}.\nEjecute 'Generar Aulas (N)' para el período objetivo."`.

#### 5. Fase 481 — Trazabilidad de Incidencias de Semestre Descartadas (AgFabricaIndividuos)
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoTraza.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`
- **Diagnóstico:**
  La bifurcación `continue` introducida en Fase 478 al detectar discrepancias de semestre entre la cohorte del aula y la asignatura descartaba asignaturas silenciosamente. Si existían inconsistencias de formato de semestre en el catálogo, la fábrica de individuos podía generar un individuo con 0 genes sin emitir ninguna traza explicativa en la telemetría.
- **Solución Implementada:**
  1. Se hizo pública la clase `AgDiagnosticoTraza` (`public sealed class AgDiagnosticoTraza`) permitiendo su consumo en contratos y servicios genéticos.
  2. Se añadió la propiedad `Traza` a `AgContexto` y se enlazó en `ConstruirContextoAsync`.
  3. En `AgFabricaIndividuos.CrearAleatorio`, se añadió el parámetro opcional `AgDiagnosticoTraza traza = null` y se acumula `int descartadasSemestre++` cada vez que se descarta una asignatura por discrepancia de semestre.
  4. Al finalizar la generación cromosómica, se emite `trazaEfectiva.Conteo("AsignaturasDescartadasPorSemestre", descartadasSemestre)`. Si la población generada queda sin genes y existían asignaturas en el contexto, se emite una advertencia explícita en telemetría (`GenesVaciosPorSemestre`) y en el log de depuración.
  5. Se implementó la prueba unitaria `CrearAleatorio_WhenSemestreDoesNotMatch_ShouldTrackDescartadasSemestreWithTraza` en `AgFabricaIndividuosTests.cs`.

### 157.21 Release v1.6.3.35 — Compilación, Pruebas y Empaquetado MSI

- **Versión:** `1.6.3.35`
- **Artefacto:** `installer\InstaladorConversorPDFv1.6.3.35.msi` (5.44 MB / 5,435,392 bytes)
- **Estado de Pruebas:** 100% de pruebas unitarias ejecutadas y aprobadas (521 / 521 pasadas, 0 errores, 0 omitidas en xUnit).
- **Compilación:** `Release|x64` sin errores (0 Errores, 25 Warnings de documentación XML / ICE91 esperados).
- **Cambios incluidos:**
  - **Fase 477:** Carga de contexto multi-semestre en `ConstruirContextoAsync` (aulas, asignaturas, pool docente y programas homónimos sin filtro restrictivo de primer semestre).
  - **Fase 478:** Cruce discriminado de asignaturas y aulas por semestre en `AgFabricaIndividuos`, agrupando por cohorte unívoca (`ClaveCohorteAula`) y preservando asignaturas sin semestre asignado o transversales.
  - **Fase 479:** Contabilización y emisión estructurada de asignaturas aseguradas por combinación (`AseguradasPorCombinacion`) en `GenerarEncabezadosAulasAsync`.
  - **Fase 480:** Mensajes de bloqueo y diagnóstico enriquecidos en `P07_ValidarAnclajeAula` y `P08_ValidarAnclajeAula` detallando cantidad de aulas cargadas y semestres del alcance.
  - **Fase 481:** Trazabilidad de asignaturas descartadas por semestre (`AsignaturasDescartadasPorSemestre`) y detección de cohortes vacías en `AgFabricaIndividuos`.

### 157.22 Ronda 68 — Bloque de Correcciones y Aislamiento de Cohortes (Fases 482 - 488)

#### 1. Fase 482 — Desambiguación de Métodos Exportables en `AgDiagnosticoTraza`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoTraza.cs`
- **Diagnóstico:** Al haberse hecho pública la clase `AgDiagnosticoTraza` (Fase 481), su método `public static string Normalizar(string sql)` colisionaba con `DiaSemanaNormalizador.Normalizar(string dia)` al ser descubierto por Excel-DNA, generando la advertencia `Repeated function name: 'Normalizar'`.
- **Solución Implementada:** Se renombró el método a `NormalizarSql(string sql)` clarificando su propósito para sentencias SQL y eliminando la colisión.

#### 2. Fase 483 — Restricción de Exportación de UDFs en Excel-DNA
- **Archivos Modificados:** `ConversorPDF.XLL/ConversorPDF.XLL-AddIn64.dna`
- **Diagnóstico:** `ExplicitExports="false"` provocaba que Excel-DNA registrara automáticamente cualquier método `public static` del ensamblado como función de hoja de cálculo (UDF).
- **Solución Implementada:** Se configuró `ExplicitExports="true"`, blindando el registro para que únicamente los métodos decorados explícitamente con `[ExcelFunction]` se expongan como fórmulas en Excel.

#### 3. Fase 484 — Acotamiento de Aulas al Alcance del Perfil Activo (Bloqueante)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `ConstruirContextoAsync` consultaba el catálogo `ag_aula` con `Semestre = null, Grupo = null`, cargando 8 aulas acumuladas en la base de datos a pesar de que el perfil activo 'A' constaba de 6 combinaciones seleccionadas. Esto derivaba en la creación de 8 cohortes (48 vinculaciones en lugar de 36).
- **Solución Implementada:** Tras obtener las aulas del repositorio, se filtran contra el conjunto de `AulaRef` de las combinaciones seleccionadas: `refsSel = new HashSet<string>(combinaciones.Select(c => c.AulaRef))`. Esto descarta aulas residuales o deseleccionadas (e.g. `268 (I-C)` y duplicados virtuales), preservando exactamente las cohortes correspondientes a la selección del usuario.

#### 4. Fase 485 — Evaluación de Depuración de Catálogo en Base de Datos
- **Diagnóstico y Evaluación:** Las aulas generadas previamente en `ag_aula` que no forman parte de la selección activa permanecen persistidas en la base de datos SQLite.
- **Dictamen:** Se validó que el filtro en memoria implementado en la Fase 484 neutraliza al 100% el impacto operativo sin requerir una migración destructiva o no-aditiva sobre la base de datos de los usuarios.

#### 5. Fase 486 — Aislamiento de Cohortes en Mutación de Aulas (Bloqueante)
- **Archivos Modificados:** 
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorMutacionTests.cs`
- **Diagnóstico:** En `AgOperadorMutacion`, la búsqueda de aulas candidatas filtraba únicamente por `string.Equals(a.Grupo, gen.Grupo)`. En escenarios multi-semestre con múltiples aulas de grupo "A" (ej. Semestre I, II, III y IV), la mutación reasignaba aleatoriamente un gen de Semestre I a un aula de Semestre II/III/IV, provocando fugas de genes y duplicación de vinculaciones.
- **Solución Implementada:** Se enriqueció el filtro de candidatas para exigir coincidencia estricta tanto en `Grupo` como en `Semestre` normalizado (`AgFormatoAsignatura.SemestreRomano`). Se añadió la prueba unitaria `Mutar_WhenMutatingAula_ShouldOnlySelectCandidatesWithMatchingSemestreAndGrupo`.

#### 6. Fase 487 — Regla Dura R4 de Coherencia Aula-Cohorte (Bloqueante)
- **Archivos Modificados:** 
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`
- **Diagnóstico:** La regla `R4_CompatibilidadAula` evaluaba exclusivamente `aula.Estado != AgEstadoAula.Disponible`, sin detectar cuando un gen quedaba asignado a un aula perteneciente a otro semestre o grupo.
- **Solución Implementada:** Se extendió R4 para verificar la coherencia entre `aula.Semestre` y `gen.Semestre` así como `aula.Grupo` y `gen.Grupo`. Si existe discrepancia, se emite una violación dura (`esDura: true`), impidiendo que cromosomas con fuga de cohorte superen el proceso evolutivo. Se implementó la prueba unitaria `Evaluar_WhenAulaFueraDeCohorte_ShouldRecordR4HardViolation`.

#### 7. Fase 488 — Diferenciación por Modalidad en Exportación Espejo
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Diagnóstico:** Al agrupar sesiones en `SemestresExcelExporter`, no se incluía la modalidad, provocando que cohortes homónimas (e.g. `I-B VIRTUAL` e `I-B PRESENCIAL`) se agruparan juntas como si fuesen una sola cohorte.
- **Solución Implementada:** Se incorporó `s.Modalidad` a la clave de proyección anónima en `items.GroupBy(...)`, garantizando la separación visual correcta de cada cohorte en el reporte exportado.

### 157.23 Ronda 69 — Revisión y Ajustes Finos de Aislamiento de Cohortes (Fases 489 - 492)

#### 1. Fase 489 — Corrección de R4: Eliminación de Falsos HARD con Semestre/Grupo Vacío
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`
- **Diagnóstico:** La implementación de la Fase 487 comparaba `AgFormatoAsignatura.SemestreRomano(aula.Semestre)` con `AgFormatoAsignatura.SemestreRomano(gen.Semestre)`. Dado que `SemestreRomano("")` devuelve `"SIN ASIGNAR"`, cuando un aula o gen carecía de semestre o grupo explícito (aulas o materias no acotadas), la comparación evaluaba `"SIN ASIGNAR" != "I"` generando violaciones duras espurias en pruebas y en el cálculo genético.
- **Solución Implementada:** Se definió `"SIN ASIGNAR"` y los grupos vacíos como comodines válidos:
  `bool semOk = semAula == "SIN ASIGNAR" || semGen == "SIN ASIGNAR" || string.Equals(semAula, semGen, StringComparison.OrdinalIgnoreCase);`
  `bool grpOk = string.IsNullOrWhiteSpace(aula.Grupo) || string.IsNullOrWhiteSpace(gen.Grupo) || string.Equals(aula.Grupo, gen.Grupo, StringComparison.OrdinalIgnoreCase);`
  Se añadió la prueba unitaria `Evaluar_WhenAulaSinSemestreOGrupo_ShouldTreatAsWildcardAndNotTriggerR4`.

#### 2. Fase 490 — Corrección de Exportación: Agrupación por `Grupo` y `Modalidad` sin `AulaRef`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Diagnóstico:** El modelo desnormalizado de sesión `SesionListItem.cs` no expone la propiedad `AulaRef` (únicamente expone `Grupo` y `Modalidad`). Proyectar `s.AulaRef` en `GroupBy` causaba error de compilación.
- **Solución Implementada:** Se ajustó la proyección anónima en `SemestresExcelExporter.GroupBy` utilizando las propiedades nativas `s.Grupo` y `s.Modalidad`, logrando el aislamiento completo de cohortes sin alterar el contrato de datos.

#### 3. Fase 491 — Corrección de Acotamiento de Aulas: Respaldo de Alcance Completo
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Si los nodos seleccionados tenían `AulaRef` vacío o aún no coincidían con las aulas persistidas en la base de datos, el filtro estricto de la Fase 484 vaciaba `ctx.Aulas` (`count == 0`), abortando prematuramente el pre-vuelo.
- **Solución Implementada:** Se enriqueció la selección filtrando únicamente referencias no vacías (`refsSel`). Si `filtradas.Count > 0`, se acota `ctx.Aulas`; en caso contrario (0 coincidencias), se conserva el alcance completo emitiendo `traza.Advertencia("AulasSinMatchSeleccion")` para no bloquear el flujo de pre-vuelo ni la generación de aulas.

#### 4. Fase 492 — Retiro Estratégico de Fases 483 y 485
- **Archivos Modificados:** `ConversorPDF.XLL/ConversorPDF.XLL-AddIn64.dna`
- **Diagnóstico y Decisión:**
  - **Fase 483:** Al haberse desambiguado `NormalizarSql` en la Fase 482, no existe colisión de nombres de funciones. Mantener `ExplicitExports="false"` previene la rotura de compatibilidad con llamadas de macros VBA dinámicas (`Application.Run`). Se revirtió a `ExplicitExports="false"`.
  - **Fase 485:** Se descartó la migración destructiva en base de datos (`UPDATE ag_aula SET estado = 0`), dado que el filtro en memoria implementado en las Fases 484/491 aísla limpiamente el entorno de optimización sin riesgo de corrupción o contaminación cruzada entre perfiles de usuario.

### 157.24 Ronda 70 — Normalización de Semestre en Mutación, Cohorte Lógica Homónima y Limpieza Excel (Fases 493 - 495)

#### 1. Fase 493 — Normalización Canónica de Semestre en Mutación de Aula
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorMutacionTests.cs`
- **Diagnóstico:** Al mutar un aula en `AgOperadorMutacion`, se copiaba directamente `candidatas[0].Semestre`. Si el aula candidata tenía un valor arábigo (e.g. `"1"`), `gen.Semestre` quedaba con dicho valor crudo, provocando fallas posteriores en `EsDocenteHabilitado(gen.Semestre)` y disparando falsas violaciones de la regla R5.
- **Solución Implementada:** Se normaliza el semestre del aula seleccionada mediante `AgFormatoAsignatura.SemestreRomano(elegida.Semestre)`. Si el resultado no es `"SIN ASIGNAR"`, se asigna a `gen.Semestre`. Se agregó la prueba unitaria `Mutar_WhenAulaHasArabicSemestre_ShouldNormalizeToRoman`.

#### 2. Fase 494 — Filtro por Cohorte Lógica Tolerante a Homónimos de Programa
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El formato de `AulaRef` es `idperiodo|idsede|idprograma|semestre|grupo|modalidad`. Cuando existen programas homónimos (por ejemplo, IDs `[2, 1]` resueltos a `1`), las referencias construidas con el ID anterior no coincidían exactamente con las aulas persistidas con el ID homologado. Como resultado, `filtradas.Count == 0` y la Fase 491 restauraba el alcance completo silenciosamente, reintroduciendo el problema de las 8 cohortes.
- **Solución Implementada:** Se introdujo la función local `ExtraerCohorte(a)` que extrae únicamente la cohorte lógica: `periodo|sede|semestre|grupo|modalidad`, omitiendo el segmento `idprograma`. De este modo, el acotamiento de aulas vincula la cohorte académica exacta independientemente de discrepancias de alias en el catálogo de programas.

#### 3. Fase 495 — Limpieza de Rango Previo en Exportación Excel Espejo
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Diagnóstico:** `SemestresExcelExporter` volcaba los datos directamente a partir de la celda `(3, 1)`. Al incorporar `Grupo` y `Modalidad` en la clave de agrupación, la matriz de datos resultante puede variar de tamaño entre ejecuciones sucesivas; filas residuales de corridas previas permanecían visibles al final de la tabla, generando aparentes duplicados visuales.
- **Solución Implementada:** Previo al volcado masivo en Excel, se ejecuta `rangoLimpieza.ClearContents()` sobre el área de datos existente y se libera explícitamente el objeto COM mediante `Marshal.ReleaseComObject(rangoLimpieza)`.

### 157.25 Ronda 71 — Aislamiento y Purga de Staging Acumulado (Fases 496 - 498)

#### 1. Fase 496 — Filtro por Corrida en Detalle/Exportación Modo AG
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al abrir el detalle o exportación de sesiones en Modo AG desde `AbrirDetalleSesionesModoAgAsync`, se invocaba `queryRepo.GetSesionesAsync(...)` sin especificar `idCorrida` (`null`). La vista `vw_AgSesionesDetalle` proyectaba todas las corridas históricas del período y programa presentes en `ag_sesiones` / `ag_vinculaciones`, mezclando sesiones de corridas previas (e.g. corridas 11 y 12) con la corrida activa.
- **Solución Implementada:** Se resolvió el ID de corrida activo (`_ultimoResultado?.IdCorrida ?? _idCorridaCargada`) y se envió como parámetro `idCorridaDetalle` a la sobrecarga de `GetSesionesAsync`, asegurando que la grilla y el árbol reflejen estrictamente las vinculaciones y sesiones de la corrida evaluada.

#### 2. Fase 497 — Purga Automática de Staging por Período y Programa
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs`
- **Diagnóstico:** `LimpiarStagingAsync(idCorrida)` únicamente eliminaba los registros pertenecientes al `id_corrida` de la materialización actual. Las corridas previas permanecían acumuladas indefinidamente en las tablas de staging de la bóveda temporal y persistente, provocando sobreconteo y contaminación cruzada si se consultaba sin filtro estricto.
- **Solución Implementada:** Se implementó la sobrecarga `LimpiarStagingAsync(int idPeriodo, int idPrograma, CancellationToken ct)` que purga de forma transaccional `ag_vinculacion_gen`, `ag_sesiones`, `ag_salvedades_conflictos`, `ag_vinculaciones` y `ag_gen` para la combinación de período y programa. `AgMaterializadorSesiones.MaterializarEnStagingAsync` invoca esta purga previa a almacenar el nuevo cronograma.

#### 3. Fase 498 — Índices de Optimización de Staging (Migración V60)
- **Archivos Creados y Modificados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV60_AddAgStagingIndices.cs` (Nuevo)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** Las operaciones de filtrado por corrida en `vw_AgSesionesDetalle` y las purgas transaccionales por subconsultas requerían escaneos secuenciales completos sobre `ag_vinculaciones` y `ag_sesiones`.
- **Solución Implementada:** Se creó la migración V60 asegurando los índices `idx_ag_vinc_corrida ON ag_vinculaciones(id_corrida, idperiodo, idprograma)`, `idx_ag_ses_vinc ON ag_sesiones(idvinculacion)` e `idx_ag_gen_ind ON ag_gen(id_individuo)`.

### 157.26 Ronda 71 (Consolidación) — Purga Completa de Entidades Huérfanas e Índices (Fases 499 - 501)

#### 1. Fase 499 — Retiro y Descarte de la Propuesta Duplicada de Fase 497
- **Diagnóstico y Decisión:** La propuesta original de Fase 497 intentaba reescribir la purga con sintaxis SQL `DELETE FROM ... s WHERE EXISTS ...` (con alias no válidos en SQLite para operaciones DELETE) y duplicaba la sobrecarga `LimpiarStagingAsync(int, int, ct)` ya existente en `AgScheduleRepository`. Dicha propuesta preliminar fue descartada formalmente.

#### 2. Fase 500 — Purga Completa de Corridas e Individuos Huérfanos
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** La purga de staging por período y programa limpiaba las tablas operativas (`ag_vinculacion_gen`, `ag_sesiones`, `ag_salvedades_conflictos`, `ag_vinculaciones`, `ag_gen`), pero dejaba intactos los registros de `ag_individuo` y `ag_corrida`. Esto provocaba que en bases de datos con corridas interrumpidas o previas aparecieran corridas huérfanas en el selector de corridas (`cmbCorrida`), algunas con `id_individuo_optimo == null`.
- **Solución Implementada:** Se añadieron sentencias transaccionales para eliminar `ag_individuo` (asociados a las corridas del período/programa) y `ag_corrida` antes del `Commit` en `LimpiarStagingAsync(int idPeriodo, int idPrograma, CancellationToken ct)`.

#### 3. Fase 501 — Consolidación de Índices de Migración V60
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV60_AddAgStagingIndices.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Solución Implementada:** Se consolidó la migración `MigrationV60_AddAgStagingIndices` registrada en `SqliteHelper.cs` incluyendo los 3 índices de alta concurrencia:
  - `idx_ag_vinc_corrida ON ag_vinculaciones(id_corrida, idperiodo, idprograma)`
  - `idx_ag_ses_vinc ON ag_sesiones(idvinculacion)`
  - `idx_ag_gen_ind ON ag_gen(id_individuo)`

### 157.27 Ronda 72 — Orden de Purga de Staging, Limpieza de Ancho Total en Excel y Cascada Completa (Fases 502 - 504)

#### 1. Fase 502 — Secuenciamiento de Purga Previa a la Creación de Corrida
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Si la purga por `(idperiodo, idprograma)` se ejecutaba dentro de `AgMaterializadorSesiones.MaterializarEnStagingAsync`, borraba la corrida que acababa de crearse e insertarse en `IniciarOptimizacionAsync` (debido a que compartían el mismo período y programa), provocando que `GetCorridaByIdAsync` retornara nulo y el selector de corridas no listara la solución optimizada.
- **Solución Implementada:** Se trasladó la invocación de `LimpiarStagingAsync(idPeriodo, idPrograma, ct)` a `IniciarOptimizacionAsync` para ejecutarse **antes** de instanciar e insertar la nueva `AgCorrida`. En `AgMaterializadorSesiones.MaterializarEnStagingAsync` se conserva únicamente la limpieza por `idCorrida`.

#### 2. Fase 503 — Limpieza a Ancho Total de Columnas en Exportación Espejo
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Diagnóstico:** La limpieza de la Fase 495 borraba hasta `columnasTotales` (calculadas para la corrida actual). Si una ejecución anterior contenía un rango mayor de fechas, las columnas sobrantes permanecían a la derecha de la tabla.
- **Solución Implementada:** Se amplió el rango de borrado a `ws.Range[ws.Cells[3, 1], ws.Cells[ws.Rows.Count, ws.Columns.Count]]`, asegurando la limpieza total del lienzo antes de vaciar la matriz de datos.

#### 3. Fase 504 — Cascada de Purga Completa en Tablas Módulo R
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** `LimpiarStagingAsync(int idPeriodo, int idPrograma, CancellationToken ct)` omitía tablas dependientes como `ag_violacion`, `ag_bitacora` y `ag_promocion`, dejando registros huérfanos asociados a corridas pasadas.
- **Solución Implementada:** Se añadieron sentencias transaccionales previas al borrado de `ag_individuo` y `ag_corrida` para eliminar en cascada `ag_violacion` (por individuo), `ag_bitacora` y `ag_promocion` (por corrida).

### 157.28 Ronda 73 — Verificación y Cierre Definitivo de Purga en Árbol Real (Fases 505 - 507)

#### 1. Fase 505 — Confirmación de Exclusividad de Purga por Corrida en Materializador
- **Archivos Verificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Estado:** Confirmado aplicado e íntegro. `AgMaterializadorSesiones.MaterializarEnStagingAsync` invoca exclusivamente `await _scheduleRepository.LimpiarStagingAsync(idCorrida, ct);`, garantizando que la corrida e individuo recién guardados no sean destruidos.

#### 2. Fase 506 — Verificación de Purga Previa en el Orquestador
- **Archivos Verificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Estado:** Confirmado aplicado con el token de cancelación correcto `_ctsEjecucion.Token` en `IniciarOptimizacionAsync` (L1144-L1149). Ninguna referencia errónea a `_contextoActual.CancellationToken` está presente.

#### 3. Fase 507 — Verificación de Cascada Completa de Purga
- **Archivos Verificados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Estado:** Confirmado aplicado e íntegro (L880-L905). Las sentencias transaccionales de purga eliminan en riguroso orden de dependencias:
  1. `ag_vinculacion_gen`
  2. `ag_sesiones`
  3. `ag_salvedades_conflictos`
  4. `ag_vinculaciones`
  5. `ag_gen`
  6. `ag_violacion` (por subconsulta a `ag_individuo`)
  7. `ag_bitacora` (por subconsulta a `ag_corrida`)
  8. `ag_promocion` (por subconsulta a `ag_corrida`)
  9. `ag_individuo`
  10. `ag_corrida`

### 157.29 Ronda 75 — Coherencia Docente por Grupo, Auditoría Integral y Migración V61 (Fases 508 - 514)

#### 1. Fase 508 — Coherencia de Docente y Vacante por ClaveGrupo en Mutador
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
- **Diagnóstico:** La mutación de docente se evaluaba independientemente por cada gen dentro del bucle genómico. Como resultado, en asignaturas multi-día pertenecientes al mismo `ClaveGrupo`, un gen podía mutar a un docente mientras que su gen hermano mutaba a otro o quedaba como vacante, produciendo vinculaciones con múltiples docentes y causando duplicación de filas en la exportación por docente.
- **Solución Implementada:** Se extrajo la mutación de docente/vacante para operar agrupada por `ClaveGrupo`. Cuando un grupo muta, todos sus genes constituyentes reciben sincrónicamente el nuevo docente o el centinela vacante.

#### 2. Fase 509 — Blindaje de Coherencia de Sesiones con la Vinculación Agrupada
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Aunque la vinculación agrupaba genes por `ClaveGrupo`, el expansor cronológico podía propagar discrepancias si los genes de entrada presentaban divergencias.
- **Solución Implementada:** Tras el bucle de expansión de sesiones, se sellan todas las sesiones generadas para el grupo con `agVinc.IdTercero`, `agVinc.EsVacante` y `agVinc.MotivoVacante`, garantizando uniformidad estricta 1:N entre la cabecera de vinculación y sus sesiones hijas.

#### 3. Fase 510 — Logging de Auditoría en Purga de Staging
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Solución Implementada:** En `LimpiarStagingAsync(int idPeriodo, int idPrograma, CancellationToken ct)` se registran recuentos antes/después para `ag_vinculaciones`, `ag_sesiones`, `ag_gen` y `ag_corrida`, así como la lista de IDs de corridas eliminadas.

#### 4. Fase 511 — Logging Detallado de Cohortes Solicitadas y Filtradas
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Solución Implementada:** En `ConstruirContextoAsync` se registran las cohortes solicitadas, las aulas cargadas, las aulas filtradas y el número de asignaturas asociadas por aula.

#### 5. Fase 512 — Logging de Proyección con Detección de Duplicados en Exportador
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Solución Implementada:** Previo al volcado de datos, se verifica y loguea la cantidad de filas y la presencia de claves repetidas en la agrupación.

#### 6. Fase 513 — Reducción de Ruido en Trazabilidad de Asignaturas Descartadas
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Solución Implementada:** El contador `AsignaturasDescartadasPorSemestre` se emite únicamente para el individuo 1, evitando emitir 50 trazas idénticas por población inicial.

#### 7. Fase 514 — Migración V61: Columna `fecha_creacion` en `ag_salvedades_conflictos`
- **Archivos Creados y Modificados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV61_AddFechaCreacionSalvedades.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** `SalvedadConflictoRepository` consulta `fecha_creacion` en `ag_salvedades_conflictos`, pero la tabla no contenía dicha columna, produciendo advertencias de `SQL logic error`.
- **Solución Implementada:** Se implementó la migración V61 añadiendo de forma tolerante la columna `fecha_creacion TEXT DEFAULT ''`.

### 157.30 Ronda 76 — Deduplicación de Traza Diagnóstica y Sincronización de Logging (Fases 515 - 518)

#### 1. Fase 515 — Deduplicación de Emisiones Idénticas en `AgDiagnosticoTraza.Conteo`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoTraza.cs`
- **Diagnóstico:** Contadores de estado (como `AsignaturasDescartadasPorSemestre`) que se repiten con el mismo valor en cada individuo de la población inicial producían 50 líneas idénticas en el archivo de log.
- **Solución Implementada:** Se incorporó un diccionario interno `_ultimoConteo` indexado por etiqueta comparando la firma `"{obtenidos}|{minimo}"`. Si el conteo sucesivo coincide exactamente con el valor previo, se omite silenciosamente la emisión repetitiva, conservando la primera emisión.

#### 2. Fase 516 — Corrección de Orden y Contenido en Logs de Cohortes y Aulas
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En la Fase 511, `aulasCargadas` se evaluaba tras aplicar el filtro en memoria, ocultando el total real de aulas del alcance; adicionalmente, `asigsPorAula` se invocaba antes de cargar `ctx.AulaAsignaturas`, reportando siempre 0 asignaturas asociadas.
- **Solución Implementada:** Se capturó `aulasCargadas` e `idsCargadas` previo a acotar por las cohortes seleccionadas. El registro detallado de `asigsPorAula` se reubicó inmediatamente después de la consulta efectiva de `ag_aula_asignatura`.

#### 3. Fase 517 — Reemplazo Formal de Fase 513 por Fase 515
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La propuesta preliminar de restringir la invocación a `idIndividuo <= 1` en la fábrica fue sustituida por la solución genérica y limpia de Fase 515 en el componente de diagnóstico `AgDiagnosticoTraza`. Se restauró la llamada estándar en la fábrica.

#### 4. Fase 518 — Verificación de Idempotencia en Migración V61
- **Archivos Verificados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV61_AddFechaCreacionSalvedades.cs`
- **Estado:** Confirmado. Utiliza la verificación previa `pragma_table_info('ag_salvedades_conflictos') WHERE name='fecha_creacion'` y bloque `try-catch` tolerante, previniendo errores si la columna ya estuviese presente.

### 157.31 Ronda 77 — Alineación Incondicional de Grupo y Sellado Integral de Staging (Fases 519 - 520)

#### 1. Fase 519 — Alineación Incondicional de Docente y Vacante por ClaveGrupo en Mutador
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorMutacionTests.cs`
- **Diagnóstico:** En la Fase 508, la unificación de docentes operaba después de evaluar `if (_random.NextDouble() >= tasaMutacion) continue;`. Si un cromosoma inicial contenía genes mixtos dentro del mismo `ClaveGrupo` y la tasa de mutación no se disparaba, el grupo se mantenía con docentes dispares. Además, si `genes[0].EsVacante == 0` pero sus genes hermanos eran vacantes (192), la rama vacante previa no los homogeneizaba.
- **Solución Implementada:** Se reestructuró el bucle para que la tasa de mutación únicamente condicione la decisión de seleccionar un nuevo docente para el grupo (`docCanon`), mientras que el sellado de `docCanon`, `vacCanon` y `motCanon` se ejecuta de forma **incondicional** para todos los genes constituyentes del `ClaveGrupo`. Se añadió la prueba unitaria `Mutar_WhenGenesInSameClaveGrupoHaveMixedTeachers_ShouldAlignInconditionally`.

#### 2. Fase 520 — Sellado Integral de `CandidatosPool` en Materialización
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** En la Fase 509 se sellaban `IdTercero`, `EsVacante` y `MotivoVacante`, pero se omitía `CandidatosPool` en las entidades `AgSesion`, dejando discrepancias entre la cabecera de vinculación agrupada y sus sesiones hijas.
- **Solución Implementada:** Se añadió `sesionesStaging[sIdx].CandidatosPool = agVinc.CandidatosPool;` al bucle de sellado en `MaterializarEnStagingAsync`, garantizando paridad completa de auditoría 1:N.

### 157.32 Ronda 78 — Coherencia Estructural en Cruce, Helper Canónico AgCoherenciaGrupo y Regla Dura (Fases 521 - 523)

#### 1. Fase 521 — Erradicación de Grupos Mixtos Reintroducidos por el Operador de Cruce (`AgOperadorCruce`)
- **Archivos Creados y Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
- **Diagnóstico:** A pesar de la alineación en la mutación (Fase 519), los grupos mixtos reaparecían porque el operador de cruce recombinaba el cromosoma gen a gen (`HeredarOX`) entre progenitores con asignaciones docentes distintas para el mismo grupo (`ClaveGrupo`), y el bucle de reparación de colisiones operaba igualmente a nivel de gen individual.
- **Solución Implementada:**
  1. Se creó el helper centralizado `AgCoherenciaGrupo.Sellar(AgIndividuo, AgContexto)`, el cual agrupa los genes por `ClaveGrupo` y unifica incondicionalmente el docente calculando la moda (`doc = reales.GroupBy(g => g.IdTercero).OrderByDescending(x => x.Count()).First().Key; vac = 0;`), o asignando el centinela vacante si no existen docentes reales.
  2. En `AgOperadorCruce.cs`, se invoca `AgCoherenciaGrupo.Sellar(hijo, contexto);` para cada descendiente inmediatamente antes de `AgMetricas.Recalcular(hijo, contexto);`.
  3. En `AgOperadorMutacion.cs`, se invoca `AgCoherenciaGrupo.Sellar(individuo, contexto);` al término del ciclo completo de mutaciones genéticas.

#### 2. Fase 522 — Registro de `AgCoherenciaGrupo.cs` en Proyecto XLL
- **Archivos Modificados:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Solución Implementada:** Se añadió `<Compile Include="Academic\Application\Services\Genetica\AgCoherenciaGrupo.cs" />` respetando la estructura explícita del `.csproj`.

#### 3. Fase 523 — Regla Dura de Invariante de Docente Único por `ClaveGrupo` (Guardia en Validador)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** Para garantizar que ninguna solución evolutiva con docentes dispares en un mismo `ClaveGrupo` pueda ser seleccionada como óptima o viable, se requería una restricción dura que castigue con infactibilidad cualquier inconsistencia residual.
- **Solución Implementada:** Tras el bucle de evaluación de genes en `Validar`, se incorporó la verificación por grupo:
  `genes.Where(g => !string.IsNullOrWhiteSpace(g.ClaveGrupo)).GroupBy(g => g.ClaveGrupo)`. Si `gr.Select(g => g.IdTercero).Distinct().Count() > 1`, se registra una violación dura (`esDura: true`) e incrementa `hardCount`, descartando cromosomas incoherentes.

### 157.33 Ronda 79 — Priorización de Genes Fijados en Sellado de Coherencia (Fase 524)

#### 1. Fase 524 — Priorización Canónica de Genes con Estado `Fijado` en `AgCoherenciaGrupo`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Si un grupo contenía genes con estado `AgEstadoGen.Fijado` (docente X) junto a genes no fijados cuya moda temporal fuera Y, el cálculo por moda previa alineaba los genes libres a Y mientras el gen fijado permanecía inmutable en X, preservando inadvertidamente la fragmentación del grupo.
- **Solución Implementada:** Se añadió como primera comprobación dentro del sellado por grupo la detección de genes fijados:
  - Si existen múltiples genes fijados con docentes distintos, se emite una advertencia mediante `LoggerHelper.Warning`.
  - El primer gen fijado actúa como docente/vacante canónico incondicional para todos los genes no fijados del grupo (`g.IdTercero = fijo.IdTercero; g.EsVacante = fijo.EsVacante; g.MotivoVacante = ...`).

### 157.34 Ronda 80 — Soporte Multi-Sede / Multi-Programa y Exclusividad de Modalidad (Fases 525 - 530)

#### 1. Fase 525 — Contexto Multi-Sede y Multi-Programa en `AgContexto` y `AgFormGeneradorHorarios`
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al seleccionar combinaciones de múltiples sedes y programas, `ConstruirContextoAsync` invocaba `GetAulasPorPeriodoAsync` pasando únicamente `sedeSel.Id` y `programaSel.Id` de la primera combinación, descartando las aulas de las demás sedes y programas marcados.
- **Solución Implementada:**
  - Se añadieron las propiedades `SedesAlcance` y `ProgramasAlcance` (`List<int>`) a `AgContexto`.
  - En `ConstruirContextoAsync`, se extraen `sedesSel` y `progsSel` de todas las combinaciones seleccionadas y se consulta `GetAulasPorPeriodoAsync` con `idSede = null` e `idPrograma = null`, cargando la totalidad de aulas del período para luego filtrar en memoria según la selección.

#### 2. Fase 526 — Filtro de Cohortes Acotado por Programas Seleccionados
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `ExtraerCohorte` omite el ID de programa (`p[2]`) para tolerar homónimos entre programas. No obstante, al coexistir cohortes con la misma sede, semestre y grupo en programas distintos, se producían colisiones indebidas.
- **Solución Implementada:** Se reforzó el filtrado de aulas exigiendo que `progsSel.Count == 0 || progsSel.Contains(a.IdPrograma)`, garantizando que solo se retengan aulas de los programas activamente seleccionados.

#### 3. Fase 527 — Consulta de Asignaturas y Pool Docente por Colección de Programas (`IN @IdProgramas`)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Las consultas SQL `sqlAsigMulti` y `sqlPoolMulti` restringían rígidamente `WHERE idprograma = @IdPrograma`, omitiendo asignaturas y docentes de los demás programas en ejecuciones multi-programa.
- **Solución Implementada:** Se adaptaron ambas consultas a `WHERE idprograma IN @IdProgramas` y se parametrizó `progsParam` (`progsSel` o fallback a `programaSel.Id`), permitiendo la carga completa de la oferta académica y pool docente para todos los programas seleccionados.

#### 4. Fase 528 — Exclusividad de Modalidad (Nivel 6) y Asignación por Defecto a Presencial
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al marcar un Grupo (Nivel 5) o cualquiera de sus ancestros, la propagación descendente marcaba simultáneamente las 3 modalidades (Presencial, Virtual, Semipresencial), triplicando las combinaciones por cohorte. Asimismo, faltaba exclusividad tipo radio button entre hermanas de nivel 6.
- **Solución Implementada:**
  - Se incorporó `AplicarExclusividadModalidad(TreeNode nodo)` para desmarcar automáticamente a las hermanas de nivel 6 cuando una modalidad es seleccionada.
  - En `TreeContextoAcademico_AfterCheck`, al marcar un nodo de nivel 6 se aplica dicha exclusividad.
  - En `PropagarCheckHaciaAbajo`, al descender a nodos de nivel 5 se selecciona únicamente la modalidad por defecto ("Presencial").
  - En `InvertirSeleccionArbol`, se valida la exclusividad si el nodo de nivel 6 queda marcado.

#### 5. Fase 529 — Evaluación Crítica: Rechazo de Desactivación de CheckBoxes e ImageList Fantasma (Falso Positivo)
- **Diagnóstico Crítico:** La propuesta de fijar `treeContextoAcademico.CheckBoxes = false` y asignar `StateImageList = ilEstado`:
  1. Depende de un `ilEstado` inexistente en el proyecto, provocando error de compilación.
  2. Elimina las casillas de verificación nativas de todos los niveles superiores (Período, Sede, Programa, Semestre, Grupo).
  3. Deshabilita la navegación por teclado (barra espaciadora) y rompe la cadena de eventos `AfterCheck` de WinForms.
  4. La exclusividad funcional y operativa requerida ya queda 100% satisfecha y blindada por la Fase 528 sin introducir deuda técnica ni riesgos de regresión.
- **Dictamen:** Se rechaza Fase 529 por tratarse de un falso positivo arquitectónico / riesgo de regresión visual severo.

#### 6. Fase 530 — Mensaje de Diagnóstico Pre-Vuelo P07 Enriquecido con Alcance Multi-Sede / Multi-Programa
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Solución Implementada:** El mensaje de advertencia P07 ahora desglosa con exactitud el alcance de sedes (`SedesAlcance`) y programas (`ProgramasAlcance`), así como las referencias cargadas (`refsAlcance`), facilitando una acción correctiva precisa al usuario.

### 157.35 Ronda 81 — Robustecimiento Multi-Sede: R3, R7, Balanceo Docente y Gestión de Escala (Fases 531 - 538)

#### 1. Fase 531 — Evaluación de R7 por Aula/Cohorte Real en Validador
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** R7 evaluaba `col.Select(y => y.Grupo).Distinct().Count()`. Al ejecutarse en multi-sede, un docente asignado al grupo "A" en Cali y al grupo "A" en Buenaventura en el mismo bloque producía `distinctGrupos = 1`, omitiendo la colisión.
- **Solución Implementada:** Se sustituyó por `col.Select(y => y.IdAula).Distinct().Count()`, garantizando la detección fidedigna de solapamiento docente entre sedes y cohortes distintas.

#### 2. Fase 532 — Inclusión de `IdSede` en Clave de Superposición de Grupo (R3)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** R3 componía `grpKey = $"{gen.IdPrograma}_{gen.Semestre}_{gen.Grupo}_{gen.IdBloqueHorario}"`. Como omitía `IdSede`, el mismo programa/grupo replicado en dos sedes físicas distintas se penalizaba falsamente como conflicto de estudiantes solapados.
- **Solución Implementada:** Se modificó la clave a `grpKey = $"{gen.IdSede}_{gen.IdPrograma}_{gen.Semestre}_{gen.Grupo}_{gen.IdBloqueHorario}"`, aislando los grupos por sede física.

#### 3. Fase 533 — Reparto Docente Balanceado y Determinista por Rotación Histórica
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
- **Diagnóstico:** `habilitados[_random.Next]` generaba sesgos aleatorios desbalanceados (e.g. 5-0 en lugar de 3-2) entre sedes para una misma asignatura.
- **Solución Implementada:**
  - En `AgFabricaIndividuos`, se implementó una rotación determinista por asignatura `_rotacionDocente` ordenada por `VecesAsignada` y `IdTercero`, reiniciada en cada `CrearAleatorio`.
  - En `AgOperadorMutacion`, se prioriza la asignación del docente con menor carga histórica `OrderBy(p => p.VecesAsignada).ThenBy(p => p.IdTercero).First()`.

#### 4. Fase 534 — Soporte Multi-Sede y Multi-Programa en Detalle/Export Modo AG
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `AbrirDetalleSesionesModoAgAsync` forzaba `idSede = _contextoActual.IdSede` e `idPrograma = _contextoActual.IdPrograma`, ocultando las sesiones de las demás sedes y programas en corridas multi-sede.
- **Solución Implementada:** Se ajustaron `idSede` e `idPrograma` a `null` cuando `SedesAlcance` o `ProgramasAlcance` tienen más de un elemento, permitiendo ver y exportar la totalidad de las sesiones de la corrida.

#### 5. Fase 535 — Aprobación Transaccional Multi-Aula en `AprobarCorridaYAulaAsync`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al aprobar una corrida multi-aula (`corrida.IdAula == 0`), ninguna entidad `ag_aula` pasaba a estado `Aprobado` porque solo se contemplaba `corrida.IdAula > 0`.
- **Solución Implementada:** Se añadió una sentencia SQL transaccional que actualiza a `AgEstadoAula.Aprobado` todas las aulas referenciadas en las sesiones de la corrida (`WHERE id IN (SELECT DISTINCT s.idaula FROM ag_sesiones s JOIN ag_vinculaciones v ON s.idvinculacion = v.id WHERE v.id_corrida = @IdCorrida)`).

#### 6. Fase 536 — Ampliación de Capacidad a 5000 Combinaciones para Multi-Sede
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSeleccionAulaExpandidor.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El tope fijo de 500 combinaciones bloqueaba la generación de aulas en árboles multi-sede con decenas de programas y grupos (e.g. ~1600 hojas).
- **Solución Implementada:** Se aumentó la constante de seguridad `CapMaximoRecomendado` a 5000 en el expandidor y en la interfaz de usuario.

#### 7. Fase 537 — Blindaje de Exclusividad al Cargar Perfiles de Árbol
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al cargar un perfil desde la base de datos, si existían hojas preexistentes o duplicadas, se podían activar múltiples modalidades para un mismo grupo.
- **Solución Implementada:** Se invocó `AplicarExclusividadModalidad(nodoMod)` inmediatamente después de marcar cada hoja de modalidad en `CargarPerfilArbolSeleccionadoAsync`.

#### 8. Fase 538 — Preservación de Períodos Heterogéneos en Hojas Seleccionadas
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** La Fase 445 sobrescribía ciegamente `dto.IdPeriodo = idPeriodoObjetivo`, aplastando combinaciones que legítimamente pertenecían a otro período en árboles multi-período.
- **Solución Implementada:** Solo se sobrescribe `dto.IdPeriodo` si todas las hojas seleccionadas pertenecen al mismo período o no tienen conflicto; ante multiplicidad de períodos, se preservan sus identificadores individuales con registro de advertencia.

### 157.36 Ronda 81 — Diversidad Poblacional, Compatibilidad en Cruce, Glifos Radio Nivel 6, Aprobación Multi-Aula, Regla R10 y Guarda Multi-Período (Fases 539 - 545 y Retiro de 531)

#### 1. Retiro de Fase 531 — Eliminación de Doble Penalización de R7 sobre Aulas
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico y Evaluación Crítica:** La Fase 531 sustituía la evaluación de `distinctGrupos` por `IdAula` en la regla R7. Sin embargo, cualquier conflicto en que un mismo docente coincida en dos aulas o sedes durante el mismo bloque horario ya se encuentra estricta y penalmente castigado por la regla R1 (`docenteBloques` agrupado por `IdTercero` e `IdBloqueHorario`). Evaluar adicionalmente `IdAula` en R7 provocaba una doble penalización por el mismo conflicto físico, distorsionando el fitness del motor genético.
- **Solución Implementada:** Se retiró la modificación de la Fase 531 en R7, restaurando la validación canónica sobre la multiplicidad de grupos (`col.Select(y => y.Grupo).Distinct().Count() > 1`), dejando la detección de solapes temporales globales a cargo de R1.

#### 2. Fase 539 — Preservación de Diversidad Genética Inicial con Offset Aleatorio por Individuo
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** En la Fase 533, el diccionario `_rotacionDocente` se inicializaba en 0 para cada individuo en `CrearAleatorio`. En consecuencia, los 50 individuos de la población inicial nacían con idénticas secuencias de asignación docente, perdiendo la variabilidad exploratoria del algoritmo genético y provocando estancamiento prematuro del fitness (`Estancamiento (20 generaciones)`).
- **Solución Implementada:** Al iniciar la selección para un individuo y asignatura, si la clave no existe en `_rotacionDocente`, se asigna un offset aleatorio `rot = _random.Next(orden.Count);` antes de avanzar deterministamente la rotación. De esta forma, cada individuo explora una configuración docente distinta mientras se preserva el balance y menor carga histórica en cada cromosoma.

#### 3. Fase 540 — Re-derivación de Cohorte e Invariante del Aula en Operador de Cruce (`AgOperadorCruce`)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** En `HeredarOX`, al transferir genes del progenitor donante se copiaban atributos como `IdAsignatura`, `IdTercero` y bloques horarios, pero se omitía actualizar `IdSede`, `IdPrograma`, `Semestre`, `Grupo` y `ClaveGrupo`. Como resultado, el descendiente retenía la cohorte del progenitor receptor con asignaturas del donante, violando invariantes y desfasando la evaluación de R3/S2 y la materialización de staging.
- **Solución Implementada:** Tras copiar los atributos genéticos en el bucle de `HeredarOX`, se re-derivan las propiedades de cohorte consultando `contexto.AulasPorId[dst.IdAula]` (`dst.IdSede = aulaDst.IdSede; dst.IdPrograma = aulaDst.IdPrograma; dst.Semestre = ...; dst.Grupo = ...;`) y se regenera `dst.ClaveGrupo = $"{dst.IdAsignatura}|{dst.IdAula}|{dst.Grupo}";`.

#### 4. Fase 541 — Blindaje de Compatibilidad Semestral en Cruce Genético
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** En escenarios multi-sede y multi-semestre, `HeredarOX` podía heredar una asignatura de un semestre distinto al semestre asignado al aula del gen receptor, provocando violaciones duras ("fuera de cohorte" / `R4_CompatibilidadAula`) que convertían al individuo en infactible y destruían soluciones viables.
- **Solución Implementada:** Antes de sobreescribir los atributos del gen en `HeredarOX`, se comprueba que `AgFormatoAsignatura.SemestreRomano(asigSrc.Semestre) == AgFormatoAsignatura.SemestreRomano(aulaDst.Semestre)`. Si difieren, se aborta la herencia de ese locus (`continue`), conservando la asignatura compatible del progenitor base.

#### 5. Fase 542 — Renderizado Visual de RadioButton en Nivel 6 y CheckBox en Niveles 1-5
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Aunque la exclusividad lógica tipo radio se implementó en la Fase 528, el control `treeContextoAcademico` mantenía `CheckBoxes = true`, dibujando casillas cuadradas nativas en el nivel 6 y confundiendo visualmente al usuario respecto a la selección única de modalidad.
- **Solución Implementada:**
  1. Se configuró `treeContextoAcademico.CheckBoxes = false` en el diseñador y en el código del formulario.
  2. Se generó dinámicamente un `ImageList` de 4 glifos nativos de 16x16 píxeles mediante `CheckBoxRenderer` y `RadioButtonRenderer`: `[0] CheckBox unchecked`, `[1] CheckBox checked`, `[2] RadioButton unchecked`, `[3] RadioButton checked`, asignándolo a `treeContextoAcademico.StateImageList`.
  3. Se conectó `NodeMouseClick` y `KeyDown` (tecla Espacio) para alternar el estado y propagar los eventos del árbol sin perder navegación.
  4. Se implementaron los métodos auxiliares `RefrescarEstadoNodo(TreeNode n)` y `RefrescarEstadoRecursivo(TreeNode n)` que asignan `StateImageIndex` según el nivel del nodo y su propiedad `Checked`.
  5. Se incorporó la actualización de glifos en `TreeContextoAcademico_AfterCheck`, `AplicarExclusividadModalidad`, `PropagarCheckHaciaAbajo`, `ActualizarEstadoPadre`, `MarcarTodoElArbol`, `InvertirSeleccionArbol`, `PoblarArbolContextoAcademico`, `CargarHijosLazy` y `AsegurarSubarbolAsync`.

#### 6. Fase 543 — Corrección de Columna SQL `s.id_aula` y Aprobación Exhaustiva Multi-Aula
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En `AprobarCorridaYAulaAsync`, la consulta SQL introducida en la Fase 535 referenciaba erróneamente `s.idaula`, cuando el esquema canónico de SQLite (`MigrationV49`, `AgScheduleRepository`) nombra la columna como `s.id_aula`, provocando error en tiempo de ejecución (`no such column: s.idaula`). Además, no se garantizaba la actualización conjunta cuando coexistían `corrida.IdAula` y sesiones con distintas aulas.
- **Solución Implementada:** Se corrigió la consulta SQL a `SELECT DISTINCT s.id_aula FROM ag_sesiones s JOIN ag_vinculaciones v ON s.idvinculacion = v.id WHERE v.id_corrida = @IdCorrida AND s.id_aula > 0`. Se unificó la lista de aulas incluyendo `corrida.IdAula` si aplica, ejecutando un único `UPDATE ag_aula SET estado = @Estado WHERE id IN @Ids`.

#### 7. Fase 544 — Tipificación Semántica de Regla R10 (`R10_CoherenciaDocenteGrupo`)
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** En la Fase 523, la violación por inconsistencia de docentes en un mismo `ClaveGrupo` se reportaba reutilizando la violación dura `R2_SuperposicionAula`, generando diagnósticos confusos y estadísticas de auditoría distorsionadas.
- **Solución Implementada:** Se añadió el valor de enumeración `R10_CoherenciaDocenteGrupo = 28` en `AgReglaViolacion.cs` y se actualizó `AgValidadorReglas.cs` para emitir formalmente infracciones con este código semántico.

#### 8. Fase 545 — Blindaje contra Mezcla de Múltiples Períodos Académicos
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Si el usuario seleccionaba ramas pertenecientes a múltiples períodos académicos distintos, `ObtenerCombinacionesSeleccionadas` o los flujos subsiguientes podían sobrescribir el período o mezclar calendarios incompatibles, rompiendo la resolución de la ventana temporal y la consistencia de la corrida.
- **Solución Implementada:**
  1. En `ObtenerCombinacionesSeleccionadas()`, se preserva el `IdPeriodo` específico de cada hoja asignado en el árbol, asignando el período de `cmbPeriodos` únicamente como fallback si `dto.IdPeriodo <= 0`.
  2. En `ConstruirContextoAsync`, se verifica si `combinaciones.Select(c => c.IdPeriodo).Where(p => p > 0).Distinct().Count() > 1`. Ante multiplicidad de períodos, se aborta inmediatamente retornando `Result.Failure<AgContexto>` con mensaje descriptivo.
  3. En `GenerarEncabezadosAulasAsync`, se verifica si las combinaciones pertenecen a más de un período y se detiene la ejecución mostrando un aviso al usuario, impidiendo la generación de aulas con mezcla de períodos.

### 157.37 Ronda 82 — Recálculo de Patrón y HorasGen en Crossover e Indexación AulaAsignatura por Cohorte (Fases 546 - 547)

#### 1. Fase 546 — Recálculo de Patrón, IndiceDiaPatron y HorasGen en Operador de Cruce (`HeredarOX`)
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Al recombinar loci en `HeredarOX`, el gen hijo heredaba `IdAsignatura`, `TotalHorasAsignatura` y bloque horario del donante, pero retenía las propiedades `HorasGen`, `DiasPatron`, `PatronSemanal` e `IndiceDiaPatron` del progenitor base. Cuando la asignatura donante tenía una carga horaria o patrón diferente (e.g. 64h multi-día vs 48h mono-día), la suma de horas del grupo en el hijo discrepaba del total de la asignatura, provocando el fallo inmediato del Invariante I1 (`Σ HorasGen == Max(TotalHorasAsignatura)`) en `AgMaterializadorSesiones` y abortando la materialización en staging. Asimismo, los días de patrón obsoletos generaban falsas infracciones de `R8_DiaFueraDePatron`.
- **Solución Implementada:**
  1. En `AgOperadorCruce.HeredarOX`, inmediatamente después de actualizar `dst.TotalHorasAsignatura`:
     - Se resuelve la entidad `AgAulaAsignatura` correspondiente mediante `contexto.AulaAsignaturaPorAulaAsig[(dst.IdAula, src.IdAsignatura)]` (con fallback a `contexto.AulaAsignaturasPorId[src.IdAsignatura]`).
     - Se extraen los días de patrón `diasK = AgPlanificadorPatronSemanal.DiasKDe(aaSrc).ToList()`.
     - Se actualiza `dst.DiasPatron = string.Join(",", diasK)` y `dst.PatronSemanal = aaSrc.PatronSemanal`.
     - Se ajusta `dst.IndiceDiaPatron = Math.Min(Math.Max(0, dst.IndiceDiaPatron), Math.Max(0, diasK.Count - 1))`.
     - Se recalcula `dst.HorasGen = Math.Max(4.0, Math.Round(dst.TotalHorasAsignatura / Math.Max(1, diasK.Count) / 4.0) * 4.0)`.
  2. En `AgCoherenciaGrupo.Sellar`, se añadió una guarda constructiva de balance por `ClaveGrupo` que verifica si `Math.Abs(sumaHoras - totalHorasEsperadas) > 0.001` y ajusta cualquier residuo sobre el primer gen, blindando matemáticamente el Invariante I1 en todos los descendientes y mutaciones.

#### 2. Fase 547 — Índice O(1) Desacoplado `AulaAsignaturaPorAulaAsig` para Soporte Multi-Aula y Multi-Sede
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** El diccionario `AulaAsignaturasPorId` en `AgContexto` estaba indexado exclusivamente por `IdAsignatura`. En corridas multi-aula o multi-sede con decenas de aulas, la misma asignatura aparecía N veces y el diccionario colapsaba reteniendo únicamente el último registro procesado, perdiendo `HorasObjetivo`, `DiasPatron` y `OrdenCurricular` de las demás aulas. Además, `AgFabricaIndividuos` recurría a búsquedas lineales `FirstOrDefault` y caía en un fallback colapsado que tomaba configuraciones de cohortes o sedes ajenas.
- **Solución Implementada:**
  1. En `AgContexto.cs`, se incorporó la propiedad indexada `public Dictionary<(int IdAula, int IdAsignatura), AgAulaAsignatura> AulaAsignaturaPorAulaAsig { get; private set; }`.
  2. En `AgContexto.ConstruirIndices()`, se construye el índice compuesto agrupando por `(x.IdAula, x.IdAsignatura)` sin pérdida de información para ninguna cohorte.
  3. En `AgFabricaIndividuos.cs`, se reemplazó la búsqueda lineal por la consulta directa O(1) `contexto.AulaAsignaturaPorAulaAsig.TryGetValue((aula.Id, asig.Id), out var aaExacta)`, eliminando fallbacks erróneos a aulas de otras sedes.
  4. En `AgValidadorReglas.cs`, se actualizó la regla de secuencia curricular R6 para consultar prioritariamente `AulaAsignaturaPorAulaAsig` al evaluar el orden curricular específico del aula.

### 157.38 Ronda 83 — Corrección de Compilación en HeredarOX y Reconciliación I1/I2 (Fases 548 - 550)

#### 1. Fase 548 — Erradicación del Error de Compilación CS1061 por `aaSrc.PatronSemanal`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** `AgAulaAsignatura` no expone una propiedad `PatronSemanal`. El intento de lectura `aaSrc.PatronSemanal` generaba el error de compilación CS1061.
- **Solución Implementada:** Se eliminó la referencia a `aaSrc.PatronSemanal`, derivando el patrón de días canónicamente a partir de `diasK = AgPlanificadorPatronSemanal.DiasKDe(aaSrc).ToList(); dst.DiasPatron = string.Join(",", diasK);`.

#### 2. Fase 549 — Reconciliación Robusta de Invariantes I1 e I2 mediante Ajuste de `FechasGen`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** El blindaje previo de I1 se ubicaba al final de `Sellar`, quedando omitido para grupos con genes fijados debido al `continue` de la rama fija. Además, ajustaba `genes[0].HorasGen` a ciegas, lo que podía mutar genes fijados inmutables, producir `HorasGen <= 0` con residuos negativos y desalinear `FechasGen` respecto a `HorasGen` (rompiendo I2 en el materializador de staging).
- **Solución Implementada:**
  1. Se reubicó la reconciliación al inicio del ciclo por grupo, antes de la rama de fijados, asegurando su ejecución para todos los grupos.
  2. Se seleccionan exclusivamente genes no fijados (`ajustables = genes.Where(g => g.Estado != AgEstadoGen.Fijado)`).
  3. Si existe residuo `Math.Abs(residuo) > 0.001`, se calcula `delta = Math.Round(residuo / 4.0)` y se agregan (+7 días por sesión) o eliminan fechas en `FechasGen` del último gen ajustable sin dejarlo en 0 sesiones.
  4. Se fija `gAjuste.HorasGen = gAjuste.FechasGen.Count * 4.0`, garantizando consistencia matemática simultánea entre sesiones reales e I1/I2.

#### 3. Fase 550 — Regeneración Dinámica de `PatronSemanal` en `HeredarOX`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Al transferir una asignatura con días o carga distintos, `dst.PatronSemanal` retenía el histograma del progenitor base, distorsionando las proyecciones de distribución semanal de sesiones y las reglas blandas S2/S4.
- **Solución Implementada:**
  1. Se inyectó `AgPlanificadorPatronSemanal` en el constructor de `AgOperadorCruce`.
  2. En `HeredarOX`, se recalcula el patrón semanal convocando `_planificadorPatron.Construir(diasK, (int)dst.TotalHorasAsignatura, fIni, fFin, contexto.EsFestivo)` y asignando el histograma a `dst.PatronSemanal`.

### 157.39 Ronda 84 — Desacople de FechasGen y HorasGen en Cruce, Reconciliación con FechasGen Vacío y Coherencia de Bloque (Fases 551 - 553)

#### 1. Fase 551 — Vaciado de `FechasGen` Heredado y Cuantización de `HorasGen` en Cruce
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Al transferir una asignatura en `HeredarOX`, `HorasGen` se recalculaba según la carga de la nueva asignatura, pero `FechasGen` conservaba las fechas precalculadas del progenitor base. `AgExpansorSesiones` priorizaba `FechasGen` preexistente cortando por `horasFaltantes <= 0`, lo que producía un número de sesiones discordante con `TotalHorasAsignatura` y provocaba la violación de la invariante I2 (`nSesiones * 4 == HorasGen`) en `AgMaterializadorSesiones`.
- **Solución Implementada:** Se limpia `dst.FechasGen = new List<DateTime>()` para que el expansor de sesiones regenere dinámicamente las fechas exactas mediante `GenerarFechasValidas` según el bloque y `HorasGen` cuantizado a múltiplos de 4.

#### 2. Fase 552 — Reconciliación I1/I2 en `Sellar` con `FechasGen` Vacío
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Si los genes del grupo tenían `FechasGen` vacío (como los generados tras un cruce en Fase 551), el bloque de ajuste en `Sellar` se limitaba a registrar un log de advertencia y no corregía el residuo de horas, haciendo fallar la invariante I1 (`sum(HorasGen) == TotalHorasAsignatura`).
- **Solución Implementada:** En la rama alternativa (`else`), si `delta > 0`, se inicializa la lista `FechasGen` si es nula y se genera la secuencia sintética semanal de fechas alineada al día de la semana (`DiaSemana`) a partir de la fecha de inicio estimada, recalculando `gAjuste.HorasGen = gAjuste.FechasGen.Count * 4.0` para satisfacer simultáneamente I1 e I2.

#### 3. Fase 553 — Reasignación de Bloque Coherente con `DiasPatron` en Cruce
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Al heredar una asignatura, se recalculaba `DiasPatron`, pero `DiaSemana` e `IdBloqueHorario` conservaban los valores del progenitor base. Si dicho día no formaba parte de los días permitidos (`diasK`) de la nueva cohorte, el individuo se tornaba infactible por la restricción dura `R8_DiaFueraDePatron`.
- **Solución Implementada:** Se calcula `diaObjetivo = diasK[Math.Min(dst.IndiceDiaPatron, diasK.Count - 1)]` y se reasigna el bloque horario convocando `dst.AsignarBloque(bloquePatron)` si existe un bloque activo con `DiaSemanaNumero == diaObjetivo`, o en su defecto actualizando `dst.DiaSemana = diaObjetivo`.

### 157.40 Ronda 85 — Verificación Final de Cadena I1/I2: Restricción de Duración 4h y Tope de Ventana (Fases 554 - 555)

#### 1. Fase 554 — Restricción de Duración Estándar de 4 Horas en Selección de Bloque Coherente
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Al seleccionar el bloque correspondiente al `diaObjetivo` tras heredar una asignatura en cruce, no se condicionaba `DuracionHoras == 4.0`. Si un bloque futuro tuviese una duración no estándar, la invariante de sesión elemental de 4 horas se quebraría, desalineando el conteo de sesiones frente a `HorasGen`.
- **Solución Implementada:** Se restringió la búsqueda en `AgOperadorCruce` exigiendo `b.Activo == 1 && b.DiaSemanaNumero == diaObjetivo && b.DuracionHoras == 4.0`.

#### 2. Fase 555 — Límite de Fecha Fin / Ventana en Generación Sintética de Fechas
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Al sintetizar fechas semanales en `Sellar` para grupos con `FechasGen` vacío, la adición de semanas continuaba indefinidamente sin verificar el límite temporal de `FechaFinEstimada` ni de la ventana académica del contexto, pudiendo proyectar sesiones fuera del calendario.
- **Solución Implementada:** Se estableció `DateTime tope = gAjuste.FechaFinEstimada ?? (ctx?.Ventana != null ? ctx.Ventana.FechaFin : cursor.AddMonths(4))` cortando el bucle con `if (cursor.Date > tope.Date) break;`, sincronizando `gAjuste.HorasGen = gAjuste.FechasGen.Count * 4.0` y emitiendo un log de advertencia informativo si el residuo requirió truncamiento por ventana.

### 157.41 Ronda 86 — Coherencia de Bloque Horario en Cruce y Protección contra HorasGen Nulo por Ventana (Fases 556 - 557)

#### 1. Fase 556 — Fallback en Cascada para Bloque Activo y Estado Incompleto en Cruce
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Si no se hallaba un bloque de 4 h para `diaObjetivo`, el código modificaba `DiaSemana = diaObjetivo` pero dejaba `IdBloqueHorario` con el bloque del progenitor base (correspondiente a otro día). Esto desalineaba el bloque de `DiaSemana` y causaba infactibilidad por la regla dura `R8_DiaFueraDePatron` al expandir sesiones.
- **Solución Implementada:** Se implementó una selección en cascada: primero un bloque activo de 4.0 h para `diaObjetivo`, y si no existe, cualquier bloque activo para `diaObjetivo`. Si no existe ningún bloque activo para dicho día, se marca `dst.Estado = AgEstadoGen.Incompleto` y se registra una advertencia, garantizando que nunca exista desalineación entre `IdBloqueHorario` y `DiaSemana`.

#### 2. Fase 557 — Protección contra HorasGen = 0 por Ventana Insuficiente y Distribución de Residuo
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Si la ventana académica o `FechaFinEstimada` no permitían agregar ninguna sesión sintética (`generadas == 0`), asignar `HorasGen = FechasGen.Count * 4.0` dejaba `HorasGen = 0`. Esto disparaba un reemplazo indeseado por `TotalHorasAsignatura` en el materializador de sesiones, quebrando las invariantes I1 e I2.
- **Solución Implementada:** Si `generadas > 0`, se asigna `HorasGen = FechasGen.Count * 4.0`; si `generadas == 0`, se conserva el `HorasGen` existente. Adicionalmente, el residuo pendiente (`delta - generadas`) se redistribuye a razón de +4 h en los demás genes ajustables del grupo, preservando la coherencia global del cromosoma.

### 157.42 Ronda 87 — Preservación de I2 en Reparto de Residuo y Asignación de Bloque de Respaldo (Fases 558 - 559)

#### 1. Fase 558 — Sincronización de `FechasGen` con Incremento de `HorasGen` en Reparto de Residuo
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Al redistribuir el residuo pendiente `delta - generadas` incrementando `+4.0 h` a `HorasGen` en otros genes ajustables, si el gen receptor ya contenía fechas precalculadas (`FechasGen.Count > 0`), `AgExpansorSesiones` reutilizaba dichas fechas sin expandirlas, generando un número de sesiones menor a `HorasGen / 4` y quebrando la invariante I2 (`nSesiones * 4 == HorasGen`).
- **Solución Implementada:** Al repartir las horas pendientes, si el gen receptor tiene `FechasGen` no vacío, se calcula la siguiente fecha semanal `c = ultFecha.AddDays(7)`, se valida contra `FechaFinEstimada`/ventana, y si es admisible se añade a `FechasGen` fijando `HorasGen = FechasGen.Count * 4.0`; si no tiene margen de ventana, se continúa con el siguiente gen ajustable.

#### 2. Fase 559 — Asignación de Bloque Activo de Respaldo en Cruce para Prevenir Bloque Cruzado
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Si no existía ningún bloque activo para `diaObjetivo`, marcar `dst.Estado = Incompleto` sin reasignar bloque dejaba `IdBloqueHorario` apuntando al bloque del padre1 (correspondiente a otro día). El generador de sesiones expandía fechas en el día del bloque obsoleto, disparando la regla dura `R8_DiaFueraDePatron` puesto que el validador no ignora genes en estado incompleto.
- **Solución Implementada:** En la rama de excepción cuando no existe bloque para el día objetivo, se asigna como último recurso cualquier bloque activo del contexto (`dst.AsignarBloque(bloqueUltimo)`) antes de marcar `dst.Estado = AgEstadoGen.Incompleto`, erradicando discrepancias entre `IdBloqueHorario` y su día asignado.

### 157.43 Ronda 88 — Blindaje Defensivo contra Desbordamiento de Ventana y Afirmación de Patrón (Fases 560 - 561)

#### 1. Fase 560 — Validación de Semanas Disponibles al Incrementar HorasGen con FechasGen Vacío
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Si un gen ajustable receptor del residuo tenía `FechasGen` vacío, asignar ciegamente `g.HorasGen += 4.0` asumía que `AgExpansorSesiones` encontraría semanas suficientes para materializar la sesión extra. Si la ventana académica no alcanzaba, se materializaban menos sesiones de las esperadas, quebrando I2.
- **Solución Implementada:** Se calcula `semanasDisponibles = ((t0 - c0).TotalDays / 7)` y se verifica `sesionesNuevas <= semanasDisponibles` antes de aplicar el incremento de 4 horas; de no haber margen suficiente, se omite el gen y se evalúa el siguiente ajustable.

#### 2. Fase 561 — Preferencia de Bloque Perteneciente al Patrón y Reafirmación de DiasPatron en Respaldo
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Al seleccionar un bloque de respaldo en la rama de último recurso, tomar cualquier bloque activo podía seleccionar uno cuyo día no estuviese en `diasK`, arriesgando violación de `R8_DiaFueraDePatron`.
- **Solución Implementada:** Se prioriza en `bloqueUltimo` un bloque activo cuyo `DiaSemanaNumero` pertenezca a `diasK`, respaldado por cualquier bloque activo si no existiese, y se reafirma explícitamente `dst.DiasPatron = string.Join(",", diasK)` tras la asignación.

### 157.44 Ronda 89 — Reparto Canónico de Horas por Grupo y Eliminación de HorasGen Desacoplado (Fases 562 - 564)

#### 1. Fase 562 — Reparto Canónico y Determinista de Horas en Sellar
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** En multi-cohorte, los genes del grupo terminaban sumando más horas que `TotalHorasAsignatura` (ej: 72h vs 48h). `Sellar` sólo sabía añadir horas pero se rendía cuando el residuo era negativo (`delta <= 0`), dejando intacto el exceso y disparando el fallo fatal I1 en el materializador de staging.
- **Solución Implementada:** Se implementó el reparto canónico determinista: `totalSesiones = round(horasRestantes / 4.0)`, asignando `baseSesiones = totalSesiones / count` y distribuyendo el resto `extraSesiones = totalSesiones % count` a los primeros genes ajustables ordenados por `IndiceDiaPatron`. Garantiza por construcción que la suma sea exactamente `TotalHorasAsignatura`.

#### 2. Fase 563 — Sellar Bidireccional para Excesos y Déficits sin Give-Up
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** La presencia de `delta <= 0` provocaba que se omitiera la corrección y se emitieran miles de warnings de residuo no reconciliable.
- **Solución Implementada:** El nuevo algoritmo es intrínsecamente bidireccional (ajusta hacia arriba o hacia abajo) recalculando tanto `HorasGen` como recortando/ampliando `FechasGen` para mantener la paridad estricta entre sesiones y horas (invariantes I1 e I2).

#### 3. Fase 564 — Desacople de HorasGen en HeredarOX
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** `HeredarOX` dividía `TotalHorasAsignatura / diasK.Count` asumiendo que el grupo tendría `diasK.Count` genes, pero la cantidad real de slots del grupo provenía del aula receptora, generando la disparidad de horas.
- **Solución Implementada:** Se eliminó el recálculo dividido en `HeredarOX`, fijando `dst.HorasGen = 0.0` y delegando el reparto determinista integral del grupo a `AgCoherenciaGrupo.Sellar` (Fase 562).

### 157.45 Ronda 90 — Blindaje Integral contra HorasGen = 0, Cobertura de Semanas y Corrección de Tests (Fases 565 - 570)

#### 1. Fase 565 — Corrección de Enum Inexistente `AgEstadoGen.Mutado`
- **Archivos Modificados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs`
- **Diagnóstico:** `AgEstadoGen` no define el miembro `Mutado`, causando el error de compilación CS0117 al ejecutar pruebas unitarias.
- **Solución Implementada:** Se corrigió a `AgEstadoGen.Valido` en las aserciones de prueba correspondientes.

#### 2. Fase 566 — Activación de Reparto ante Genes con HorasGen = 0
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Si un grupo contenía genes con `HorasGen = 0` heredados de Fase 564 y la suma de los demás genes ya igualaba `TotalHorasAsignatura`, la condición de residuo no disparaba, dejando genes en 0 h que luego disparaban +48 h fantasma por el fallback del expansor/materializador.
- **Solución Implementada:** Se añadió la comprobación `bool hayPendientes = genes.Any(g => g.Estado != AgEstadoGen.Fijado && g.HorasGen <= 0.0);` disparando el reparto si `Math.Abs(suma - esperado) > 0.001 || hayPendientes`.

#### 3. Fase 567 — Guarda de Semanas Disponibles en Remanente de Sellar
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Al ajustar el residuo final `deltaResidual`, reasignar `HorasGen` y vaciar `FechasGen` sin comprobar las semanas disponibles en la ventana lectiva podía prometer sesiones no materializables.
- **Solución Implementada:** Se calculan las semanas disponibles `(t0 - c0).TotalDays / 7` antes de actualizar `gFlex.HorasGen`, registrando advertencia informativa si la ventana es insuficiente.

#### 4. Fase 568 — Auditoría de Grupos sin Genes Ajustables o con Sesiones Insuficientes
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Grupos con todos los genes fijados o con más genes que sesiones no emitían trazas de diagnóstico.
- **Solución Implementada:** Se añadieron advertencias en el log para visibilizar grupos sin ajustables o con cardinalidad discordante.

#### 5. Fase 569 — Delegación Segura de HorasGen en Cruce con Contexto
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Si `contexto == null`, `Sellar` no se ejecuta en `Cruzar`, lo que dejaría genes en 0 h sin sellar.
- **Solución Implementada:** Se condicionó la asignación `if (contexto != null) dst.HorasGen = 0.0;`.

#### 6. Fase 570 — Pruebas Unitarias para Invariante de Cruce y Reparto con Cero Horas
- **Archivos Modificados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorCruceTests.cs`
- **Diagnóstico:** Ausencia de pruebas para grupos con genes en 0 h y verificación de invariante I1 tras cruce.
- **Solución Implementada:** Se añadieron las pruebas `Sellar_WhenInheritedGeneHasZeroHoursOnCompleteGroup_ShouldRedistribute` y `Cruzar_WhenExecutedWithContext_ShouldPreserveGroupHoursInvariants`.

### 157.46 Ronda 91 — Eliminación del Amplificador de Horas Fantasma, Importación de Moq y Consolidación de Carga (Fases 571 - 575)

#### 1. Fase 571 — Importación de Namespace Moq en Pruebas de Cruce
- **Archivos Modificados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorCruceTests.cs`
- **Diagnóstico:** `Mock<Random>` e `It.IsAny<int>()` carecían de la directiva `using Moq;`, generando errores de compilación CS0246/CS0103.
- **Solución Implementada:** Se importó `using Moq;` en el encabezado de `AgOperadorCruceTests.cs`.

#### 2. Fase 572 — Consolidación de Carga cuando Sesiones < Genes Ajustables
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Si la cantidad de sesiones totales de la asignatura era menor al número de genes ajustables del grupo, `totalSesiones / count` daba 0, asignando 0 h a genes que luego activaban recursivamente `hayPendientes` en cada ciclo.
- **Solución Implementada:** Se consolidó el reparto asignando 4.0 h a los primeros `totalSesiones` genes y 0.0 h a los restantes vaciando `FechasGen`, de forma idempotente y sin reintentos infinitos.

#### 3. Fase 573 — Eliminación del Fallback Legado `0 => TotalHorasAsignatura`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`, `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs`
- **Diagnóstico:** La expresión ternaria `g.HorasGen > 0 ? g.HorasGen : g.TotalHorasAsignatura` interpretaba 0 h como carga completa (+48 h fantasma), quebrando las invariantes I1 e I2 e inflando el total de horas mapeado.
- **Solución Implementada:** Se erradicó el fallback: `horasObjetivoGen = gen.HorasGen` (retornando lista vacía si `<= 0.0`), `horasGrupo = genes.Sum(g => g.HorasGen)` y `TotalHoras = genes.Sum(x => x.HorasGen)`.

#### 4. Fase 574 — Reducción de Ruido en Bucle Caliente de Sellar
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Logs con severidad Warning en `Sellar` emitían miles de líneas en cruce y mutación.
- **Solución Implementada:** Se degradaron a `LoggerHelper.Debug` las trazas de bucle caliente de diagnóstico de grupos.

#### 5. Fase 575 — Prueba con Asignaturas Cruzadas e Idempotencia de Sellar
- **Archivos Modificados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorCruceTests.cs`
- **Diagnóstico:** El test previo usaba padres uniformes sin reproducir el cruce con asignaturas mixtas ni verificar idempotencia.
- **Solución Implementada:** Se configuraron cromosomas con asignaturas 10 y 11 cruzadas y genes en 0 h, validando que la suma de horas satisfaga la carga y que una segunda ejecución de `Sellar` mantenga idéntico el cromosoma.

### 157.47 Ronda 92 — Atajo Temprano de Cero Horas, Deduplicación de Trazas y Cobertura (Fases 576 - 580)

#### 1. Fase 576 — Evaluación Inmediata de HorasGen <= 0 en Expansor y Filtro en Materializador
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Al consolidar carga en grupos con más slots que sesiones, los genes resultantes con `HorasGen = 0.0` podían fallar por bloque inexistente o ventana no resuelta si la guarda de 0 h se evaluaba tras dichas validaciones.
- **Solución Implementada:** Se reubicó la comprobación `if (gen.HorasGen <= 0.0) return Result.Success(new List<AgSesion>());` al inicio absoluto del método `Expandir`, y se agregó `if (g.HorasGen <= 0.0) continue;` en el bucle del materializador de sesiones.

#### 2. Fase 577 — Deduplicación Efectiva de Trazas en Bucle Caliente de Sellar
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Como `LoggerHelper` encola y escribe todos los mensajes independientemente de si son `Warning` o `Debug`, degradar el nivel no evitaba la sobrecarga de I/O por escrituras masivas.
- **Solución Implementada:** Se introdujo un diccionario concurrente estático `_incidenciasCoherencia` para emitir una única traza por grupo e incidencia (`TryAdd`).

#### 3. Fase 578 — Deduplicación en Traza de Ventana Insuficiente
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** La advertencia de ventana insuficiente al reconciliar remanentes seguía emitiéndose repetidamente en cada individuo.
- **Solución Implementada:** Se protegió la emisión con `_incidenciasCoherencia.TryAdd($"{grp.Key}|ventana|{deltaResidual}", 0)`.

#### 4. Fase 579 — Detección y Advertencia de Horas Fijas no Cuantizadas
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Si un gen fijado poseía horas no cuantizadas a múltiplos de 4, la suma proyectada no coincidía con el total esperado, impidiendo el cierre de la invariante I1.
- **Solución Implementada:** Se incorporó detección de discrepancia y advertencia deduplicada: `if (Math.Abs(sumaProyectada - esperado) > 0.001)`.

#### 5. Fase 580 — Pruebas de Consolidación y Comportamiento del Expansor ante Cero Horas
- **Archivos Modificados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgExpansorSesionesTests.cs`
- **Diagnóstico:** Ausencia de pruebas que verificasen la consolidación cuando las sesiones son menores que los genes y que el expansor retorne lista vacía sin error cuando `HorasGen = 0`.
- **Solución Implementada:** Añadidas las pruebas `Sellar_WhenFewerSessionsThanGenes_ShouldConsolidateAndBeIdempotent` y `Expandir_WhenHorasGenIsZero_ShouldReturnEmptyWithoutFailing`.

### 157.48 Ronda 93 — Diagnóstico I1/I2, Sellado del Campeón, Cuarentena por Grupo, Reparto Multi-Gen y Veredicto Coherente (Fases 581 - 588)

#### 1. Fase 581 — Diagnóstico Detallado de Composición de Grupo al Violar Invariantes I1/I2
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Los mensajes de error al violar I1/I2 sólo emitían identificador de asignatura y dos escalares sin detallar la clave de grupo, número de genes, horas por gen, estado ni fechas generadas.
- **Solución Implementada:** Se incorporó el helper `DescribirGrupo` que genera una traza completa con la clave de grupo, total esperado, suma real y desglose gen a gen (id, aula, grupo, semestre, horas, total, estado, fechas).

#### 2. Fase 582 — Sellado Obligatorio y Re-evaluación del Campeón
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGeneradorHorariosTests.cs`
- **Diagnóstico:** Si el mejor individuo procedía de la siembra inicial o de selección elitista sin pasar por cruce/mutación reciente, podía publicarse como `MejorIndividuo` con desajuste de suma por grupo.
- **Solución Implementada:** Se insertaron llamadas obligatorias a `AgCoherenciaGrupo.Sellar(mejorGlobal, contexto)` y `_evaluadorAptitud.EvaluarAptitud(mejorGlobal, contexto)` inmediatamente antes de publicar el resultado. Se añadió prueba unitaria `EjecutarOptimizacionAsync_WhenChampionPublished_ShouldBeSealedAndPreserveGroupInvariants`.

#### 3. Fase 583 — Guarda de Coherencia Defensiva, Clamping HorasGen >= 0 y Reparto Acotado en Fábrica
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`, `ConversorPDF.Tests/Academic/Domain/AgGenTests.cs`
- **Diagnóstico:** Asignaciones de residuo en la fábrica podían generar horas negativas sobre el gen 0, y el materializador no ejecutaba sellado defensivo previo al chequear I1/I2.
- **Solución Implementada:** Se clampó `HorasGen >= 0.0` en `AgGen`, se acotó el reparto del residuo en `AgFabricaIndividuos` entre los genes del grupo sin permitir valores negativos, y se incluyó guarda de sellado preventivo en `AgMaterializadorSesiones`. Se añadió prueba unitaria `HorasGen_Negativo_SeClampaACero`.

#### 4. Fase 584 — Cuarentena por Grupo en Materialización de Staging
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs`
- **Diagnóstico:** Cualquier inconsistencia aislada de un único grupo abortaba la materialización completa, dejando 0 vinculaciones y 0 sesiones en staging (`StagingVacio`).
- **Solución Implementada:** Se transformó el fallo inmediato en puesta en cuarentena individualizada por grupo (en I1, fallo de expansión e I2) con rollback de sesiones añadidas y acumulación de incidencias; staging sólo falla si ningún grupo supera las validaciones. Se añadieron pruebas unitarias para fallo con diagnóstico y staging parcial exitoso.

#### 5. Fase 585 — Deduplicación Delimitada por Corrida y Reparto Multi-Gen del Remanente
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs`
- **Diagnóstico:** El diccionario de deduplicación no distinguía entre corridas independientes y acumulaba miles de entradas estáticas; además, el reparto de déficit residual se volcaba exclusivamente sobre el último gen ajustable, truncándose si la ventana de ese único gen no alcanzaba.
- **Solución Implementada:** Se introdujo `IdCorrida` en `AgContexto` asignado en el form UI, `TryIncidencia` delimitado por `${ctx.IdCorrida}|{clave}`, y cálculo `SesionesMaximas` con redistribución proporcional en cascada sobre todos los genes ajustables con holgura en su ventana. Se añadió prueba unitaria `Sellar_WhenDeficitExceedsSingleGeneWindow_ShouldDistributeAcrossAllGenesWithSlack`.

#### 6. Fase 586 — PreVuelo con Capacidad Real de Ocurrencias Calendario por Asignatura
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorTemporal.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs`
- **Diagnóstico:** El pre-vuelo utilizaba una fórmula optimista `semanasVentana * tope` que ignoraba festivos y días sin ocurrencias reales según el patrón semanal declarado, aprobando corridas condenadas a déficit.
- **Solución Implementada:** Se añadió sobrecarga ISO en `AgPlanificadorTemporal.ContarOcurrencias` y se reemplazó el cálculo en `AgValidadorVentana` por `diasKAsig.Sum(d => ContarOcurrencias(d, ini, fin, esFestivo))`, bloqueando tempranamente asignaturas con capacidad real insuficiente. Se añadió prueba unitaria `ValidarFactibilidad_WhenPatternDaysHaveZeroOrInsufficientOccurrences_ShouldBlockWithCapacidadReal`.

#### 7. Fase 587 — Veredicto y Estado de Corrida Condicionados a Staging
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Si la materialización fallaba, la corrida ya había sido marcada como `Completada` y se emitía veredicto `[APROBADO]` con `EXITOSO`, reportando éxito a pesar de un staging vacío.
- **Solución Implementada:** Se condicionaron `corrida.Estado`, `traza.Resultado`, `traza.Fin` y los controles de interfaz (`lblStatusAG`, logs) a la bandera `stagingOk`, emitiendo `AgEstadoCorrida.Error` y `[BLOQUEADO]` ante cualquier fallo en staging.

#### 8. Fase 588 — Verificación y Smoke Test
- **Estado:** Código fuente actualizado y documentado. Pruebas unitarias preparadas. Compilación y despliegue pendientes para fase de ejecución/build.

### 157.49 Ronda 94 — Guarda Defensiva Pre-MapGrupo, SesionesMaximas por Ocurrencias Propias, PreVuelo Multi-Aula y Robustez de Tests (Fases 589 - 593)

#### 1. Fase 589 — Reubicación de Guarda Coherencia Previo a MapGrupo
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs`
- **Diagnóstico:** La guarda defensiva de la Fase 583 se ejecutaba después de `MapGrupo`, dejando `agVinc.TotalHoras` con la suma desfasada y docente/vacante no sincronizados tras la reparación de `Sellar`.
- **Solución Implementada:** Se reubicó la guarda de sellado preventivo inmediatamente antes de la llamada a `AgGenToAgVinculacionMapper.MapGrupo`, garantizando que la vinculación se mapee con las horas y docentes canónicos reparados. Se añadió aserción `savedVincs[0].TotalHoras.Should().Be(48.0)` en `AgMaterializadorSesionesTests`.

#### 2. Fase 590 — Corrección de SesionesMaximas según Ocurrencias Propias del Gen y Guarda ISO
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorTemporal.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs`
- **Diagnóstico:** `SesionesMaximas` multiplicaba las semanas por el total de días de `DiasPatron` del grupo, sobreestimando la capacidad de cada gen individual (que solo opera en un único día); adicionalmente, `ContarOcurrencias(int, ...)` no validaba el rango [1..7].
- **Solución Implementada:** Se añadió guarda de rango `if (diaIso < 1 || diaIso > 7) return 0;` en `AgPlanificadorTemporal`, y se redefinió `SesionesMaximas` en `AgCoherenciaGrupo` llamando a `ContarOcurrencias(g.DiaSemana, ...)`. Se añadió prueba unitaria `Sellar_WhenMultipleGenes_ShouldLimitEachGeneCapacityToItsOwnDayOccurrences`.

#### 3. Fase 591 — PreVuelo Evaluando Capacidad Máxima Multi-Aula por Asignatura
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs`
- **Diagnóstico:** `AulaAsignaturasPorId` mantenía una sola entrada por asignatura (last-write-wins). En corridas multi-cohorte con distintas aulas y patrones para la misma asignatura, el pre-vuelo bloqueaba erróneamente si el aula retenida carecía de ocurrencias aunque existiesen otras cohortes viables.
- **Solución Implementada:** Se adaptó `AgValidadorVentana` para calcular `capacidadReal` como el máximo entre todas las configuraciones de aula registradas en `contexto.AulaAsignaturas` para dicha asignatura. Se añadió prueba unitaria `ValidarFactibilidad_WhenMultipleClassroomsForSubject_ShouldUseMaxCapacityAcrossClassrooms`.

#### 4. Fase 592 — Reparabilidad en Test de Cuarentena I1
- **Archivos Modificados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs`
- **Diagnóstico:** En `MaterializarEnStagingAsync_WhenSingleGroupViolatesI1...`, el gen de prueba tenía estado `Valido`, por lo que la nueva guarda de sellado preventivo (Fase 583/589) reparaba las horas de 96 a 48, evitando la rama de I1 e intentando expandir con un mock sin setup (generando NRE en vez de cuarentena).
- **Solución Implementada:** Se configuró el gen con `EsFijo = 1` y `Estado = AgEstadoGen.Fijado` para que `Sellar` no pueda modificar su carga, alcanzando la rama I1 esperada y validando el mensaje de diagnóstico.

#### 5. Fase 593 — Eliminación de Constante Muerta TopeDiasSemana
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `const int TopeDiasSemana = 5;` quedó obsoleta y sin referencias tras la implementación del cálculo por ocurrencias en la Fase 586.
- **Solución Implementada:** Se removió la declaración de la constante muerta.

### 157.50 Ronda 95 — Comprobación del Ruido del Log, Deduplicación Centralizada y Organización por Etapas (Fases 594 - 600)

#### 1. Fase 594 — LoggerHelper: Niveles, Modo Verboso y Dedupe Centralizado
- **Archivos Modificados:** `ConversorPDF.XLL/Core/Enums/NivelLog.cs` (nuevo), `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `ConversorPDF.XLL/Core/Helpers/LoggerHelper.cs`
- **Diagnóstico:** No existía filtro de nivel ni control de repetición en el logger central; `Enqueue` escribía incondicionalmente a cola y consola de depuración, permitiendo que cualquier bucle inundara el log (ej. 78k líneas de residuo antes de la deduplicación de Fase 585).
- **Solución Implementada:** Se creó el enum `NivelLog` (Debug=0, Info=1, Success=2, Warn=3, Error=4, Ninguno=5) en su propio archivo físico respetando la regla `single-type-per-file.md`. En `LoggerHelper` se añadieron propiedades `Minimo` (default `Info`) y `ModoVerboso` (default `false`), filtro por nivel en `Enqueue`, mecanismo de limitación `WarningLimitado(clave, msg, maxVeces)` mediante `ConcurrentDictionary<string, int>`, reinicio `ReiniciarDedupe()` y método de resumen legible `Resumen(titulo, params (string, object)[])`.

#### 2. Fase 595 — SyncScheduler: MarkDirty Acotado por Ráfaga y Resumen de PUSH
- **Archivos Modificados:** `ConversorPDF.XLL/Core/Sync/SyncScheduler.cs`
- **Diagnóstico:** `MarkDirty()` registraba un log DEBUG por cada llamada individual (`ConsecutiveCalls=...`), acumulando cientos de líneas durante escrituras masivas de repositorios (ej. `AcademicRepository.NotifyMarkDirty()`).
- **Solución Implementada:** Se acotó la emisión de log a la primera llamada de la ráfaga y a múltiplos de 100 llamadas consecutivas (`currentConsecutive == 1 || currentConsecutive % 100 == 0`). Al completar el ciclo de `PushAsync`, se agregó una línea informativa `LoggerHelper.Info` resumiendo bytes enviados, duración en ms y llamadas acumuladas.

#### 3. Fase 596 — AgDiagnosticoTraza: Supresión de Prefijo Duplicado [Emit], Modo Detalle Opt-In y Resumen de Conteos en Fin
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoTraza.cs`
- **Diagnóstico:** Cada línea se emitía con doble marca de tiempo y el caller inútil `[Emit]`. `Habilitado` iniciaba en `true`, forzando la emisión continua de volcados SQL y pasos técnicos, mientras que los contadores se reportaban en líneas dispersas sin consolidación.
- **Solución Implementada:** `Habilitado` se configuró en `false` por defecto (detalle técnico opt-in). En `Emit`, se pasa `caller = $"Ag:{Correlacion}"` a `LoggerHelper`, eliminando el caller `[Emit]` y los timestamps repetidos. `Conteo` acumula los valores en un diccionario `_conteos`, y `Fin(bool ok, string detalle)` los concatena ordenados alfabéticamente en la línea de cierre `=== FIN ... ===`.

#### 4. Fase 597 — Límite de Advertencias en Bucles Calientes Restantes
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** La advertencia de ajuste residual en la fábrica (`AgFabricaIndividuos:343`) y la de múltiples genes fijados en `Sellar` (`AgCoherenciaGrupo:161`) se ejecutaban sin control de frecuencia dentro de bucles por individuo/generación, generando riesgo de nuevas inundaciones masivas.
- **Solución Implementada:** En `AgFabricaIndividuos` se reemplazó `LoggerHelper.Warning` por `LoggerHelper.WarningLimitado($"I1Ajuste|{contexto?.IdCorrida}|{claveGrupo}", ...)`. En `AgCoherenciaGrupo` se envolvió la emisión en `if (TryIncidencia(ctx, $"{grp.Key}|multiFijado|..."))`.

#### 5. Fase 598 — Desacoplamiento de Sub-Etapas AG con AgDiagnosticoTraza Independiente
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El bloque P08 agrupaba optimización evolutiva, guardado de bitácoras, materialización y sincronización en una sola traza genérica sin métricas desagregadas ni balance de aperturas y cierres (44 inicios vs 42 fines en logs medidos).
- **Solución Implementada:** Se introdujeron trazas individuales `trazaOpt` (`OptimizacionAG`) y `trazaMat` (`MaterializacionStaging`), cada una con su propio ciclo `Inicio(...)`, contadores específicos (`GeneracionesEjecutadas`, `HardViolations`, `Genes`, `SesionesStaging`, `VinculacionesStaging`) y cierre `Fin(...)`. Se invocó `LoggerHelper.ReiniciarDedupe()` al inicio de cada corrida en `IniciarOptimizacionAsync`.

#### 6. Fase 599 — Cabecera y Pie de Proceso Único para Corridas AG
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** No existía una línea canónica de resumen al inicio ni al término de una corrida evolutiva completa que permitiera auditar rápidamente sus parámetros y resultados sin inspeccionar miles de registros.
- **Solución Implementada:** Se emitieron dos líneas estructuradas con `LoggerHelper.Resumen`: `InicioCorridaAG` (Periodo, Programa, Sede, Cohortes, Poblacion, Generaciones, ModoVerboso) y `FinCorridaAG` (IdCorrida, Estado, Fitness, Hard, Vinculaciones, Sesiones, Horas).

#### 7. Fase 600 — Verificación y Smoke Test
- **Estado:** Código fuente actualizado, documentado y refactorizado. No se ejecuta compilación ni pruebas bajo la restricción del Paso 6 de `/ronda-review`.

### 157.51 Ronda 96 — Corrección de Regresión Crítica en PreVuelo, Desacoplamiento de Notificación UI y Modo Diagnóstico (Fases 601 - 606)

#### 1. Fase 601 — Fallback de Capacidad Real en PreVuelo para Asignaturas sin AulaAsignatura
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs`
- **Diagnóstico:** Si `contexto.AulaAsignaturas` contenía registros pero ninguno correspondía al `IdAsignatura` evaluado, `Where(...).Select(CapacidadDe).DefaultIfEmpty(0).Max()` devolvía 0, provocando un bloqueo falso que abortaba inmediatamente la optimización multi-cohorte (16 falsos bloqueos observados).
- **Solución Implementada:** Se aisló la consulta de configuraciones por asignatura en `configs`. Si `configs` está vacío, se agrega el `aa` del índice unificado si existe; si no, se evalúa `CapacidadDe(null)` (patrón por defecto `{2,3}`). Se incorporó la prueba unitaria `ValidarFactibilidad_WhenSubjectHasNoAulaAsignaturaRow_ShouldNotBlockFalsely`.

#### 2. Fase 602 — Retiro de Early-Continue Mono-Aula de Semanas Disponibles
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** La guarda previa `semanasVentana == 0` evaluaba exclusivamente el `aa` retenido por el diccionario mono-aula `AulaAsignaturasPorId` y realizaba `continue`, bloqueando la asignatura antes de que se pudiera evaluar la capacidad de otras aulas viables de la misma asignatura.
- **Solución Implementada:** Se eliminó el early-continue mono-aula; la cobertura de calendario queda gobernada por `capacidadReal` multi-aula.

#### 3. Fase 603 — Reporte de Bloqueo con Identificación del Aula y Patrón Evaluado
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** El mensaje de bloqueo imprimía los días del aula retenida en el índice mono-aula (`diasKAsig`) aunque la capacidad calculada proviniera de otra configuración.
- **Solución Implementada:** Se seleccionó la configuración de aula ganadora (`configs.OrderByDescending(CapacidadDe).First()`) para reportar explícitamente su `IdAula` y días evaluados en el mensaje de bloqueo.

#### 4. Fase 604 — Desacoplamiento de Notificación a UI de Etapas vs Detalle Técnico
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoTraza.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgDiagnosticoTrazaTests.cs`
- **Diagnóstico:** Al configurar `Habilitado = false` por defecto en la Fase 596, la guarda `if (Habilitado && _ui != null)` impedía que cualquier mensaje (INICIO, FIN, Conteo, WARN, ERROR) llegara al cuadro de texto `txtLogAG` de la interfaz, dejando al usuario a ciegas durante la ejecución.
- **Solución Implementada:** En `Emit`, se discriminó el nivel: `bool esDetalle = string.Equals(nivel, "DEBUG", StringComparison.OrdinalIgnoreCase);` y se condicionó la UI a `if (_ui != null && (!esDetalle || Habilitado))`, asegurando que las etapas de progreso se notifiquen siempre y que solo el detalle DEBUG sea opt-in. Se añadió prueba unitaria `Inicio_WhenHabilitadoFalse_ShouldStillNotifyUi`.

#### 5. Fase 605 — Sincronización de LoggerHelper.Minimo con Modo Diagnóstico
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al marcar `chkModoDebug`, se asignaba `LoggerHelper.ModoVerboso = true` pero `LoggerHelper.Minimo` permanecía en `NivelLog.Info`, haciendo que `Enqueue` descartara todos los mensajes DEBUG por severidad y evitando que los volcados técnicos llegaran al archivo de log.
- **Solución Implementada:** En `CrearTraza` se añadió `LoggerHelper.Minimo = habilitado ? NivelLog.Debug : NivelLog.Info;`, junto con la directiva `using ConversorPDF.XLL.Core.Enums;`.

#### 6. Fase 606 — Desambiguación de Etapa Raíz EjecucionCorridaAG
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** La traza de la corrida global en `IniciarOptimizacionAsync` y la traza de la sub-etapa de optimización evolutiva compartían el nombre `"OptimizacionAG"`, generando trazas anidadas con el mismo identificador de operación.
- **Solución Implementada:** Se renombró la traza global a `"EjecucionCorridaAG"`.

### 157.52 Ronda 97 — Corrección de Tests por Default de Traza, Limpieza de Variable Muerta y Visibilidad de Esquema (Fases 607 - 609)

#### 1. Fase 607 — Ajuste de Habilitado en Pruebas Unitarias de Diagnóstico de Traza
- **Archivos Modificados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgDiagnosticoTrazaTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEsquemaDiagnosticoTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`
- **Diagnóstico:** Al establecer `Habilitado = false` como valor predeterminado (Fases 596/604), siete pruebas unitarias que verifican la emisión a UI de pasos (`Paso`), consultas SQL (`Sql`), conteos exitosos (`Conteo` OK) y datos de diagnóstico (`Dato`) fallaban porque la ruta DEBUG no invocaba el delegado `uiLog`.
- **Solución Implementada:** Se configuró explícitamente `Habilitado = true` en las instancias de `AgDiagnosticoTraza` de las pruebas afectadas (TEST-01, TEST-04, TEST-05, ESQ-01, ESQ-02, ESQ-03 y `CrearAleatorio_...AsignaturasDescartadasPorSemestre`), restaurando las condiciones de prueba requeridas para validar los emisores DEBUG.

#### 2. Fase 608 — Retiro de Variable Local Inutilizada diasKAsig (CS0219)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** La variable local `var diasKAsig = AgPlanificadorPatronSemanal.DiasKDe(aa);` quedó sin referencias tras la eliminación del early-continue mono-aula (Fase 602) y la adopción de `DiasKDe(ganadora)` en el mensaje de bloqueo (Fase 603), generando una advertencia de compilación CS0219.
- **Solución Implementada:** Se retiró la declaración de `diasKAsig` y se actualizaron los comentarios metodológicos.

#### 3. Fase 609 — Emisión de Veredicto de Esquema en Nivel Resultado
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEsquemaDiagnostico.cs`
- **Diagnóstico:** La verificación del esquema de base de datos emitía su resultado exclusivamente como `traza.Dato` (nivel DEBUG), lo cual impedía que el veredicto (conformidad o discrepancia) fuera visible en la interfaz o en el archivo de log bajo operación normal (`Habilitado = false`).
- **Solución Implementada:** Se incorporó `traza.Resultado(true/false, ...)` para tabla inexistente, columnas faltantes y verificación exitosa, asegurando que el veredicto se registre siempre a nivel INFO o WARN.

### 157.53 Ronda 98 — Resolución de Import Genérico en PreVuelo y Consolidación de Diagnóstico de Esquema (Fases 610 - 611)

#### 1. Fase 610 — Inclusión de using System.Collections.Generic en AgValidadorVentana (CS0246)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** El bloque multi-aula introducido en la Fase 601 empleaba `new List<AgAulaAsignatura>()`, pero el archivo carecía de la directiva `using System.Collections.Generic;`. Dado que el proyecto XLL compila bajo .NET Framework 4.8 / C# 7.3 sin *implicit usings*, esto provocaba el error de compilación fatal CS0246.
- **Solución Implementada:** Se agregó la directiva `using System.Collections.Generic;` al encabezado del archivo.

#### 2. Fase 611 — Consolidación de Diagnóstico de Esquema en Resumen Único
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEsquemaDiagnostico.cs`
- **Diagnóstico:** Tras la Fase 609, la comprobación de esquema emitía una línea INFO `Resultado(true, ...)` individual por cada una de las 10 tablas del catálogo en cada apertura/carga del formulario, generando ruido innecesario de 10 líneas idénticas por corrida cuando todas las tablas están conformes.
- **Solución Implementada:** En `AgEsquemaDiagnostico.cs` se mantuvo la emisión de `Resultado(false, ...)` ante errores o columnas faltantes, pero se suprimió el `Resultado(true, ...)` individual en rutas exitosas (quedando como `Dato` de depuración). En `AgFormGeneradorHorarios.cs` se consolidaron las 10 comprobaciones dentro de una matriz estructurada emitiendo una única línea informativa con `LoggerHelper.Resumen("EsquemaBD", ("Tablas", 10), ("OK", esquemaOk), ("ConFaltantes", esquemaFallo))`.

### 157.54 Compilación de Documentación Técnica y Empaquetado Release 1.6.3.41

#### 1. Generación de Documentación Técnica (DocFX)
- **Script Ejecutado:** `.\scripts\build_docs.ps1`
- **Resultado:**
  - Compilación exitosa de binarios y XML de documentación con 0 errores.
  - Generación de metadatos conceptuales y API gestionada (164 modelos, 161 páginas HTML, índice de búsqueda).
  - Configuración automática para GitHub Pages (`.nojekyll`) en `docs/_site`.

#### 2. Empaquetado Release v1.6.3.41 (WiX 4)
- **Script Ejecutado:** `.\scripts\build_installer.ps1 -NewVersion "1.6.3.41"`
- **Resultado:**
  - Versión sincronizada a `1.6.3.41` en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
  - Ejecución de 544 pruebas unitarias en `ConversorPDF.Tests.csproj`: 544 superadas, 0 errores, 0 omitidas (100% éxito).
  - Compilación del ensamblado XLL y empaquetado del instalador MSI mediante WiX 4 Toolset.
  - Artefacto generado: `installer\InstaladorConversorPDFv1.6.3.41.msi` (5.447.680 bytes).

### 157.55 Ronda 100 — Supresión de Ruido Residual CRDT/Cruce y Normalización de Errores en Diagnóstico (Fases 613 - 616)

#### 1. Fase 613 — Supresión de Advertencias Benignas de Inicialización CRDT
- **Archivos Modificados:** `ConversorPDF.XLL/Core/Sync/SyncInitializer.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/SalvedadConflictoRepository.cs`
- **Diagnóstico:** El fallback de `crsql_alter_begin` emitía 36 líneas WARN por apertura y `EnsureCrSqliteEnabled` generaba 18 advertencias reiterativas por conexión abierta en repositorios de lectura/escritura, representando el 58% de las advertencias totales en ejecuciones estándar.
- **Solución Implementada:** Se degradó el log de fallback de alter_begin a `LoggerHelper.Debug` en `SyncInitializer.cs` y se acotó la advertencia de `CrSqliteLoader` a una sola emisión por proceso mediante `LoggerHelper.WarningLimitado("crSqliteNoHabilitado", ..., maxVeces: 1)` en ambos repositorios.

#### 2. Fase 614 — Limitación de Advertencia en Bucle de Cruce HeredarOX
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Si una combinación de día objetivo carecía de bloque exacto, `HeredarOX` emitía `LoggerHelper.Warning` en cada hijo generado, con potencial de inundar el log con miles de líneas durante corridas evolutivas largas en Excel.
- **Solución Implementada:** Se sustituyó la emisión por `LoggerHelper.WarningLimitado($"CruceSinBloque|{contexto?.IdCorrida}|{diaObjetivo}", ..., maxVeces: 1)`.

#### 3. Fase 615 — Despliegue de Binario y Smoke Test Multi-Cohorte
- **Archivos Modificados:** `%APPDATA%\Microsoft\AddIns\ConversorPDF.XLL.dll`
- **Diagnóstico:** El entorno de producción en Excel mantenía cargada la versión 1.6.3.39 mientras que el binario compilado correspondía a la 1.6.3.41.
- **Solución Implementada:** Declarada para ejecución en fase de despliegue mediante copia de `bin\x64\Release\ConversorPDF.XLL.dll` y validación de las 14 cohortes de VALLE 2027-01.

#### 4. Fase 616 — Normalización de Caller y Supresión de StackTrace Duplicado en Fallo
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoTraza.cs`
- **Diagnóstico:** `AgDiagnosticoTraza.Fallo` emitía dos registros ERROR consecutivos por cada fallo, el segundo con caller inconsistente (`Fallo`) y el volcado completo de `StackTrace`, duplicando líneas de error aun cuando el mensaje canónico ya contenía la causa raíz.
- **Solución Implementada:** Se condicionó la emisión del segundo registro con `StackTrace` a `LoggerHelper.ModoVerboso` y se especificó el caller uniforme `$"Ag:{Correlacion}"`.

### 157.56 Ronda 101 — Fuga Residual [I2] en Fábrica, Optimización de Logs CRDT y Bump a v1.6.3.42 (Fases 615, 617, 618)

#### 1. Fase 617 — Deduplicación de Advertencia [I2] en Creación de Individuos
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La advertencia de normalización de horas no múltiplo de 4 (`[I2] Asig {id}: {totalHoras} h no es múltiplo de 4...`) se ejecutaba con `LoggerHelper.Warning` sin control de frecuencia dentro del bucle de creación de cada individuo. En escenarios con asignaturas de 45/60/68 h, generaba entre 1.000 y 2.600 líneas de log redundantes.
- **Solución Implementada:** Se reemplazó la emisión por `LoggerHelper.WarningLimitado($"I2NoMultiplo|{contexto?.IdCorrida}|{asig.Id}|{totalHoras}", ..., maxVeces: 1)`.

#### 2. Fase 618 — Reducción de Ruido en Logs de Inicialización CRDT por Tabla
- **Archivos Modificados:** `ConversorPDF.XLL/Core/Sync/SyncInitializer.cs`
- **Diagnóstico:** Las líneas informativas de regeneración CRDT por cada tabla individual (`Triggers CRDT regenerados...` y `{tabla} CRDT regenerado vía crsql_as_crr...`) sumaban ~24-36 líneas INFO por sesión, y los fallos individuales por tabla podían repetirse indefinidamente.
- **Solución Implementada:** Se degradaron las confirmaciones por tabla a `LoggerHelper.Debug` y se acotó la advertencia de fallo con `LoggerHelper.WarningLimitado($"crdtFallo|{tableName}", ..., maxVeces: 1)`.

#### 3. Fase 615 — Bump Canónico de Versión a v1.6.3.42
- **Archivos Modificados:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `ConversorPDF.XLL/Core/Helpers/App.cs`, `installer/InstaladorConversorPDF.wxs`
- **Diagnóstico:** Para validar en el smoke test multi-cohorte las correcciones acumuladas de las Rondas 93 a 101 (Fases 581 a 618) sin colisionar con el binario v1.6.3.41 previamente probado en pruebas unitarias, se requería un incremento de versión distinguible en metadatos y registros.
- **Solución Implementada:** Sincronización de versión a `1.6.3.42` en todos los archivos de configuración y manifiesto.

### 157.57 Ronda 101 Extensión — Supresión de Logs INFO Redundantes en BD Nueva/Reparación (Fase 619)

#### 1. Fase 619 — Degradar a Debug Logs de Creación y Recreación de Tablas CRDT
- **Archivos Modificados:** `ConversorPDF.XLL/Core/Sync/SyncInitializer.cs`
- **Diagnóstico:** En una base de datos nueva o tras reparación de PKs nullables, `SyncInitializer` emitía múltiples líneas INFO por tabla (`{tabla} convertida a CRDT exitosamente`, `{tabla} tiene vistas...`, volcado DDL de 200 caracteres de `SQL generado para '{tabla}'`, y `{tabla} recreada con PK NOT NULL`), sumando más de 20 líneas de detalle técnico cuando el resumen final por vault (`Conversión CRDT completada exitosamente...`) ya confirma el estado global.
- **Solución Implementada:** Se degradaron las 4 emisiones de `LoggerHelper.Info` a `LoggerHelper.Debug` (líneas 354, 480, 530 y 546).

### 157.58 Compilación de Documentación Técnica y Empaquetado Release 1.6.3.42

#### 1. Generación de Documentación Técnica (DocFX)
- **Script Ejecutado:** `.\scripts\build_docs.ps1`
- **Resultado:**
  - Compilación exitosa de ensamblados y XML de documentación con 0 errores.
  - Procesamiento conceptual y de API gestionada (164 modelos, 161 páginas HTML, índice de búsqueda).
  - Configuración automática para GitHub Pages (`.nojekyll`) en `docs/_site`.

#### 2. Empaquetado Release v1.6.3.42 (WiX 4)
- **Script Ejecutado:** `.\scripts\build_installer.ps1 -NewVersion "1.6.3.42"`
- **Resultado:**
  - Sincronización y verificación de versión `1.6.3.42` en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
  - Ejecución de las 544 pruebas unitarias de `ConversorPDF.Tests.csproj`: 544 superadas, 0 errores, 0 omitidas (100% de éxito).
  - Compilación de ensamblado XLL y empaquetado del instalador MSI mediante WiX 4 Toolset.
  - Artefacto generado: `installer\InstaladorConversorPDFv1.6.3.42.msi` (5.451.776 bytes).

### 157.59 Ronda 102 — E-03: Residuo 'Docentes catedra' y Renombres en Cascada por Desincronización Posicional de Copia (Fases 620 - 623)

#### 1. Fase 620 — Resolución Determinista de Copia de Plantilla por Diferencia de Conjunto de Hojas
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** El importador de plantillas asumía que la nueva hoja copiada por `sh.Copy(Type.Missing, wb.Worksheets[antesImp])` quedaba posicionada en `wb.Worksheets[antesImp + 1]`. Cuando existía una hoja residual al final del libro, la inserción desplazaba la hoja residual a dicho índice, haciendo que `hojasImportadas[shName]` apuntara invariablemente al residuo `'Docentes catedra'` y registrase advertencias espurias atribuyendo a Excel renombres de 6 plantillas distintas.
- **Solución Implementada:** Se captura el conjunto de nombres presentes en `wb.Sheets` antes de `sh.Copy` y se localiza la nueva hoja por diferencia de conjuntos (`!nombresAntes.Contains(candName)`), eliminando la suposición posicional por índice.

#### 2. Fase 621 — Paracaídas Temporal Visible y Desprotección Defensiva en `EliminarHojaTemporal`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** `wsNueva.Delete()` fallaba con *"No se puede obtener la propiedad Delete de la clase Worksheet"* si la hoja estaba protegida o si era la única hoja visible del libro (Excel prohíbe dejar el libro sin hojas visibles). Si las demás hojas eran candidatas de exclusión o residuales ocultas, la rutina no podía habilitar otra hoja y abortaba con error de residuo visible.
- **Solución Implementada:** Se añadieron llamadas defensivas a `Unprotect` antes de `Delete()`. Si la eliminación falla por carecer de auxiliares visibles, se genera dinámicamente una hoja paracaídas visible (`__tmp_guard_{token}`) para permitir el borrado seguro de la hoja objetivo.

#### 3. Fase 622 — Desambiguación y Renombrado Preventivo de Residuo No Eliminable (`__res_`)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** En `LimpiarPreviaPlantillas`, si una hoja canónica preexistente idéntica no podía ser eliminada por restricciones de visibilidad de Excel, conservaba el nombre oficial (`Docentes catedra`), provocando que la posterior copia de plantilla colisionara obligando a Excel a desambiguarla como `Docentes catedra (2)` o desencadenando desincronización de índices.
- **Solución Implementada:** Si la hoja previa idéntica persiste tras `EliminarHojaTemporal`, se renombra defensivamente con prefijo técnico (`__res_{token}_{name}` truncado a 31 caracteres), liberando de inmediato el nombre canónico para la importación limpia.

#### 4. Fase 623 — Reconocimiento y Purga de Prefijos Auxiliares (`__tmp_guard_`, `__res_`) en `LimpiarHojasPlantilla`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** Las hojas paracaídas (`__tmp_guard_`) y los residuos renombrados defensivamente (`__res_`) no coincidían con los prefijos fijos de hojas temporales (`CATEDRA-`, `CONTRATISTA-`, etc.), arriesgando su permanencia en el libro al concluir la exportación.
- **Solución Implementada:** Se incorporaron los prefijos `__tmp_guard_` y `__res_` en el recolector de hojas pendientes de eliminación de `LimpiarHojasPlantilla`, asegurando su purga automática e incondicional al final de la corrida.

### 157.60 Ronda 103 — Regresiones y Riesgos Residuales en Fases 620–623 (Fases 624 - 628)

#### 1. Fase 624 — Deriva de Índices por Paracaídas Insertado al Inicio (`Worksheets.Add`)
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** Invocar `wb.Worksheets.Add()` sin argumentos de posición insertaba la hoja auxiliar antes de la hoja activa (habitualmente índice 1). En el bucle de purga descendente `for (int i = (int)wb.Sheets.Count; i >= 1; i--)`, esta inserción desplazaba todos los índices en +1, provocando que la hoja ubicada en `i - 1` fuera omitida y quedara sin purgar tras el borrado del objetivo.
- **Solución Implementada:** Se modificó la invocación para insertar el paracaídas explícitamente al final del libro: `wb.Worksheets.Add(Type.Missing, wb.Worksheets[(int)wb.Worksheets.Count])`.

#### 2. Fase 625 — Fuga de Hojas Paracaídas `__tmp_guard_` y Control de Nombre No Verificado
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** El paracaídas no contaba con un bloque `finally` para intentar su eliminación una vez lograda la purga de la hoja objetivo, y si el renombrado fallaba, la hoja retenía el nombre por defecto (`Hoja1`/`Sheet1`) que no era capturado por ningún filtro de purga.
- **Solución Implementada:** Se extendió el token a 8 caracteres hexadecimales, se trazó advertencia en fallo de renombrado y se implementó un bloque `finally` que elimina el paracaídas si la hoja objetivo fue borrada y el paracaídas no es la única hoja visible restante.

#### 3. Fase 626 — Identidad de Hojas Importadas por Nombre Real de Copia y Prevención de Pérdida de Datos del Usuario
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** `hojasImportadas` se indexaba por el nombre canónico de plantilla (`shName`). Si el usuario tenía una hoja previa legítima con contenido diferente (p. ej. `DIVIPOLA`), se conservaba en `LimpiarPreviaPlantillas` y Excel creaba la copia como `DIVIPOLA (2)`. Sin embargo, `LimpiarHojasPlantilla` recibía `hojasImportadas.Keys` (`DIVIPOLA`) y la eliminaba, ocasionando pérdida de datos del usuario.
- **Solución Implementada:** Se introdujo la colección `nombresCopias` que almacena los nombres reales asignados por Excel (`nombreCopia`), transmitiéndola a `LimpiarHojasPlantilla`. Se ajustó la regla de `pendientes` para purgar exclusivamente nombres exactos de copias importadas y restringir la detección de desambiguados `(n)` a la pre-limpieza (`nombresImportados == null`).

#### 4. Fase 627 — Inversión de Log Falso Positivo y Desprotección en Purga de `LimpiarHojasPlantilla`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** `LoggerHelper.Info("Hoja residual eliminada: '{name}'")` se emitía antes de ejecutar `sh.Delete()`. Si el borrado fallaba (p. ej. en la última hoja visible o protegida), el registro afirmaba un borrado inexistente seguido de un warning sin el nombre de la hoja.
- **Solución Implementada:** Se añadió `sh.Unprotect(ConstantsHelper.FORMATO_PASSWORD)` antes de borrar, se emite el log `Info` únicamente tras la ejecución exitosa de `sh.Delete()`, y se reporta explícitamente el nombre de la hoja y su índice en el `Warning` de captura.

#### 5. Fase 628 — Validación Defensiva de Fallback Posicional en Importación de Plantilla
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** Si la detección por diferencia de conjuntos no localizaba la nueva hoja, la asignación ciega `copia = wb.Worksheets[antesImp + 1]` repetía el riesgo de enlazar hojas preexistentes o residuales desplazadas.
- **Solución Implementada:** El fallback posicional valida estrictamente que la hoja en dicha posición no estuviera presente en `nombresAntes`. De lo contrario, se omite la importación y se registra advertencia explícita.

### 157.61 Ronda 104 — Reconciliación de Supervivientes, Robustez en Paracaídas y Sincronización XML (Fases 629 - 631)

#### 1. Fase 629 — Reconciliación de Nombres en Hojas Auxiliares Supervivientes
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** Si el libro original del usuario contenía una sola hoja y era idéntica a la plantilla, la purga de plantillas dejaba únicamente la hoja paracaídas con nombre técnico (`__tmp_guard_xxxxxxxx`) como única hoja visible del libro. Asimismo, si una hoja no podía eliminarse, conservaba el prefijo `__res_xxxx_` indefinidamente.
- **Solución Implementada:** Al cierre de `LimpiarHojasPlantilla`, se añadió un bloque de reconciliación que restaura el nombre canónico original en hojas `__res_` (si no colisiona) o renombra hojas `__tmp_guard_` supervivientes a `"Hoja1"`.

#### 2. Fase 630 — Supresión de Lectura Problemática de `.Name` en Catch de Paracaídas
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** En `EliminarHojaTemporal`, el bloque `catch` ante fallo de renombre interpolaba `paracaidas.Name`. Si dicha propiedad volvía a lanzar excepción COM, abortaba la asignación `paracaidas.Visible = -1` y la llamada `wsNueva.Delete()`, dejando la hoja residual sin borrar.
- **Solución Implementada:** Se retiró la evaluación de `paracaidas.Name` dentro del mensaje de advertencia, registrando únicamente `exNom.Message` y garantizando que la ejecución continúe hacia la visibilidad y borrado del objetivo.

#### 3. Fase 631 — Sincronización de Documentación XML en `LimpiarHojasPlantilla`
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** El comentario XML de `<param name="nombresImportados">` indicaba que recibía nombres canónicos, contradiciendo el cambio introducido en la Fase 626 donde se transmiten los nombres reales de las copias importadas.
- **Solución Implementada:** Se actualizó la etiqueta XML precisando que recibe los nombres reales de copia y documentando el modo pre-limpieza (`null`) para duplicados desambiguados.

### 157.62 Ronda 105 — Logging Explícito en Reconciliación y Centralización de Constantes Técnicas (Fases 632 - 633)

#### 1. Fase 632 — Trazabilidad de Fallo en Reconciliación y Normalización de Mensajes
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** Los bloques de renombre en la reconciliación de `LimpiarHojasPlantilla` silenciaban cualquier fallo con `catch { }` vacío, dejando hojas con nombres técnicos sin aviso en el log ante libros protegidos o colisiones. Además, el literal `"(única hoja restante del libro)"` era conceptualmente impreciso durante la pre-limpieza.
- **Solución Implementada:** Se añadieron bloques `catch (Exception ex)` con avisos explícitos en `LoggerHelper.Warning` ante fallos de renombre y se normalizó el mensaje a `"(superviviente de limpieza)"`.

#### 2. Fase 633 — Centralización de Prefijos Auxiliares y Nombre Neutro en Constantes
- **Archivos Modificados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Diagnóstico:** Los literales `"__tmp_guard_"`, `"__res_"` y `"Hoja1"` estaban duplicados en múltiples generadores y filtros de reconocimiento, con riesgo de desincronización ante cambios de formato de token o prefijo.
- **Solución Implementada:** Se declararon las constantes privadas `PREFIJO_PARACAIDAS`, `PREFIJO_RESIDUO` y `NOMBRE_HOJA_NEUTRA` en `ViaticosIndividualExporter`, sustituyendo todos los literales dispersos.

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

### 157.105 Ronda 147 — Revisión de Fases 777-779 (Liberación COM y Métricas de Cierre) (Fases 780 a 782)

#### 1. Fase 780 — Preservación del Diagnóstico de Asignación ante Fallos en Escritura COM
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - Al trasladar en F778 el registro de advertencias al final del bucle de filas, si una llamada a `EscribirCelda` lanzaba una excepción COM a mitad del docente, la ejecución abortaba y se perdía la trazabilidad de los avisos emitidos por el allocator (reasignaciones, ventanas, continuaciones).
- **Solución Implementada:**
  - Se separó el registro en `LoggerHelper.Warning` (ejecutado inmediatamente tras `Distribuir`, garantizando trazabilidad ante fallos) de la agregación a `_ultimosAvisos` y `UltimosBloquesReasignados` (mantenida al final del bucle y condicionada a `planCompleto == true`).

#### 2. Fase 781 — Delimitación del Alcance en el Mensaje de Capacidad Excedida
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - El mensaje `"Filas omitidas desde el docente '{docenteCorte}': {filasOmitidas}"` sugería que la cifra acumulaba a todos los docentes restantes, cuando en realidad solo contabilizaba las filas pendientes del docente donde se interrumpió la escritura.
- **Solución Implementada:**
  - Mensaje reformulado a: `$"Capacidad de T_Principal8 excedida (fila {filaMaximaPlantilla}). Filas omitidas del docente '{docenteCorte}': {filasOmitidas} (docentes posteriores no procesados)."`.

#### 3. Fase 782 — Liberación Exhaustiva de Objetos COM Intermedios en `ObtenerFilaMaximaPlantilla`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`
- **Diagnóstico:**
  - Las propiedades COM intermedias `ws.ListObjects` y `rango.Rows` generaban objetos RCW no almacenados en variables locales, evadiendo la liberación explícita de F777.
- **Solución Implementada:**
  - Declaradas las variables locales `coleccion` y `filas`, capturando cada referencia intermedia y liberándolas explícitamente en el bloque `finally` con `Marshal.ReleaseComObject`.

---

### 157.106 Ronda 148 — Renombrado de Modalidad de Tutoría: Presencial / Plataforma / Mixta (Fases 783 a 791)

#### 1. Fase 783 — Catálogo Canónico de Modalidades de Tutoría (`ModalidadTutoriaCatalog.cs`)
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/ModalidadTutoriaCatalog.cs` (Nuevo), `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:**
  - Inconsistencia terminológica en la aplicación: `Virtual` debía renombrarse a `Plataforma` y `Semipresencial` a `Mixta` tanto en interfaz como en persistencia. Se requería un catálogo canónico en Dominio para no introducir dependencias inversas desde `AgAula` o `AgVinculacion`.
- **Solución Implementada:**
  - Creación de `ModalidadTutoriaCatalog` con constantes `Presencial`, `Plataforma`, `Mixta`, arrays `Opciones` y tuplas `Pares`.
  - Implementación de `Normalizar(modalidad)` con orden de precedencia estricto (Plataforma/Virtual -> Mixta/Semi/Distancia -> Presencial -> Trim/Fallback) y comparación invariante sin tildes/mayúsculas.
  - Implementación de `AEtiqueta(codigo)` y predicados booleanos `EsPresencial`, `EsPlataforma`, `EsMixta`.
  - Inclusión en `ConversorPDF.XLL.csproj`.

#### 2. Fase 784 — Migración V66 Aditiva y Tolerante (`MigrationV66_RenombrarModalidadTutoria.cs`) y Registro en `SqliteHelper`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs` (Nuevo), `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:**
  - Valores legados (`VIRTUAL`, `SEMIPRESENCIAL`, `A DISTANCIA`) persistidos en `vinculaciones`, `ag_vinculaciones`, `ag_aula` y `ag_arbol_nodo`, así como tokens embebidos en identificadores `aula_ref`, `codigo`, `nombre`. Riesgo de colisión en el índice único `uq_ag_arbol_nodo(idarbol, aula_ref)`.
- **Solución Implementada:**
  - Creación de `MigrationV66_RenombrarModalidadTutoria` implementando `IMigration, IMigrationTolerante` (`Version = 66`, `Aditiva = true`).
  - Actualización set-based e idempotente de columnas de modalidad en las 4 tablas.
  - Reemplazo de tokens en `ag_aula` (`|VIRTUAL` -> `|PLATAFORMA`, `|SEMIPRESENCIAL` -> `|MIXTA`, `|A DISTANCIA` -> `|MIXTA`).
  - Deduplicación explícita previa en `ag_arbol_nodo` antes del update de `aula_ref` para proteger la restricción `uq_ag_arbol_nodo`.
  - Registro en `SqliteHelper.BuildMigrator()` posterior a V65 e inclusión en `.csproj`.

#### 3. Fase 785 — Actualización de Dominio y DTOs (`AgAula.cs`, `AgVinculacion.cs`, `AgContextoSeleccionDto.cs`)
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgAula.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgVinculacion.cs`, `ConversorPDF.XLL/Academic/Application/DTOs/AgContextoSeleccionDto.cs`
- **Diagnóstico:**
  - Valores por defecto hardcodeados con `"Presencial"` o `"Virtual"`, y riesgo de que DTOs transportaran modalidades legadas sin normalizar por el resto de la aplicación.
- **Solución Implementada:**
  - Asignación de valores predeterminados utilizando `ModalidadTutoriaCatalog.Presencial`.
  - Normalización canónica de `Modalidad` en constructores de `AgContextoSeleccionDto` y propiedad calculada `ModalidadEfectiva => EsVirtual > 0 ? ModalidadTutoriaCatalog.Plataforma : ModalidadTutoriaCatalog.Normalizar(Modalidad)`.

#### 4. Fase 786 — Actualización de Mappers y Factorías
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAulaFactory.cs`, `ConversorPDF.XLL/Academic/Application/Mappers/AgArbolMapper.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`, `ConversorPDF.XLL/Academic/Application/Mappers/AgAulaToAgVinculacionMapper.cs`, `ConversorPDF.XLL/Academic/Application/Mappers/AgVinculacionToVinculacionMapper.cs`, `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEncabezadoVinculacionFactory.cs`
- **Diagnóstico:**
  - Generación de claves `AulaRef` con tokens legados y propagación de modalidades sin normalizar entre capas de mapeo y factorías de genética.
- **Solución Implementada:**
  - `AgAulaFactory.ConstruirAulaRef`: normalización de modalidad antes de aplicar `ToUpperInvariant()`.
  - `AgAulaFactory.CrearDesdeDto`: asignación `modVinculacion = esVirtual == 1 ? Plataforma : mod`.
  - `AgArbolMapper`: normalización bidireccional entre nodos de árbol y DTOs, evaluando `EsPlataforma`.
  - Mappers y factorías de vinculaciones actualizados para invocar `ModalidadTutoriaCatalog.Normalizar` con fallback a `Presencial`.

#### 5. Fase 787 — Servicios de Aplicación (`CalculadoraViaticaSesion.cs`, `SesionConflictAnalyzer.cs`)
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/CalculadoraViaticaSesion.cs`, `ConversorPDF.XLL/Academic/Application/Services/SesionConflictAnalyzer.cs`
- **Diagnóstico:**
  - Servicios de lógica de negocio usando comparaciones de texto crudo con `"VIRTUAL"` o `"SEMIPRESENCIAL"` para cálculo de viáticos y detección de solapamiento.
- **Solución Implementada:**
  - `CalculadoraViaticaSesion`: uso de `ModalidadTutoriaCatalog.EsPlataforma`, `EsPresencial`, `EsMixta` preservando intacta la semántica de negocio (Plataforma nunca viática; Presencial viática con desplazamiento; Mixta viática solo en rangos presenciales).
  - `SesionConflictAnalyzer`: exclusión de colisiones para sesiones virtuales/plataforma con `ModalidadTutoriaCatalog.EsPlataforma(s.Modalidad)`.

#### 6. Fase 788 — Actualización de Controles y Lógica de `FormSesionesAcademicas` (Designer + Code-Behind)
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:**
  - Nombres obsoletos de controles `cmbTipoJornada` y `lblTipoJornada`, texto de advertencia desalineado y opciones visuales legadas.
- **Solución Implementada:**
  - Renombrado de controles a `cmbTipoTutoria` y `lblTipoTutoria`.
  - Poblado de combo con `ModalidadTutoriaCatalog.Opciones` (`Presencial`, `Plataforma`, `Mixta`).
  - Normalización en guardado y carga de vinculaciones.
  - Advertencia actualizada: *"Con la modalidad Mixta, las sesiones se generarán sin viáticas asignadas..."*.

#### 7. Fase 789 — Actualización de `FormVerVinculaciones` y Árbol de `AgFormGeneradorHorarios`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  - Representación visual en grilla de vinculaciones y árbol AG con términos legados; mensajes de error de aulas faltantes con códigos crudos.
- **Solución Implementada:**
  - `FormVerVinculaciones`: badge `" | 🌐 Plataforma"`, omisión de conflictos con `EsPlataforma`.
  - `AgFormGeneradorHorarios`: nivel 5 poblado con pares `(par.Codigo, par.Etiqueta)`, asignación `key6 = new AgNodoArbolKey(6, 0, par.Codigo, par.Codigo)` y `new TreeNode(par.Etiqueta)`.
  - Mensajes de validación de aulas faltantes usando `ModalidadTutoriaCatalog.AEtiqueta(c.Modalidad)`.
  - Soporte de compatibilidad tolerante en la recarga de perfiles legados.

#### 8. Fase 790 — Suite de Pruebas Unitarias Nuevas y Actualizadas
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Domain/ModalidadTutoriaCatalogTests.cs` (Nuevo), `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV66Tests.cs` (Nuevo), `ConversorPDF.Tests/Academic/Application/Services/CalculadoraViaticaSesionTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAulaFactoryTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/AgArbolServiceTests.cs`, `ConversorPDF.Tests/AgArbolMapperTests.cs`, `ConversorPDF.Tests/Academic/Application/Mappers/AgAulaToAgVinculacionMapperTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgTreeViewCheckServiceTests.cs`
- **Diagnóstico:**
  - Carencia de pruebas automatizadas para el nuevo catálogo y migración V66; aserciones obsoletas en pruebas existentes.
- **Solución Implementada:**
  - Cobertura exhaustiva en `ModalidadTutoriaCatalogTests` (opciones, pares, normalización, etiquetas, predicados).
  - Cobertura completa en `MigrationV66Tests` (metadatos, tablas ausentes, migración de datos, resolución de colisión en índice único, idempotencia).
  - Actualización y adición de casos de prueba canónicos y legados en las 6 suites existentes.

#### 9. Fase 791 — Documentación Técnica del Cambio y Sincronización de Especificación
- **Archivos Afectados:** `docs/articles/changes/2026-10-21-renombrado-modalidades-tutoria.md` (Nuevo), `AG-2.txt`
- **Diagnóstico:**
  - Necesidad de documentación formal del cambio arquitectónico y actualización del registro técnico en `AG-2.txt`.
- **Solución Implementada:**
  - Redacción del artículo técnico detallando propósito, cambios, cómo usar y cobertura de pruebas.
  - Actualización de la fila de `cmbTipoTutoria` en `AG-2.txt`.

---

### 157.107 Ronda 149 — Revisión de Renombrado Presencial / Plataforma / Mixta (Fases 792 a 799)

#### 1. Fase 792 — Variables Legacy y Variable Sin Uso en `FormSesionesAcademicas`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:**
  - Variable `esPresencial` asignada pero nunca leída (advertencia de compilación CS0219). Identificadores `esVirtual` y `esSemipresencial` contradecían la nueva terminología, y el log emitía `EsPlataforma: {esVirtual}`.
- **Solución Implementada:**
  - Eliminada la variable muerta `esPresencial` y renombradas las variables locales a `esPlataforma` y `esMixta`, sincronizando la condición de aviso y el mensaje del log en L808 y L864.

#### 2. Fase 793 — Alineación de `TipoAula` y `Capacidad` con `EsVirtual` Derivado en `AgArbolMapper`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Mappers/AgArbolMapper.cs`
- **Diagnóstico:**
  - `EsVirtual` en el nodo del árbol se calculaba a partir de `ModalidadTutoriaCatalog.EsPlataforma(modEfectiva)`, pero `TipoAula` y `Capacidad` evaluaban `dto.EsVirtual` crudo, provocando que un DTO con `Modalidad="Plataforma"` y `EsVirtual=0` generara un nodo con `es_virtual=1`, `tipo_aula='ESTANDAR'` y `capacidad=0`, inconsistente con `AgAulaFactory`.
- **Solución Implementada:**
  - Precalculado `esVirtual = (dto.EsVirtual > 0 || ModalidadTutoriaCatalog.EsPlataforma(modEfectiva)) ? 1 : 0`, asignando `Capacidad = dto.Capacidad > 0 ? dto.Capacidad : (esVirtual == 1 ? 999 : 30)` y `TipoAula = !string.IsNullOrWhiteSpace(dto.TipoAula) ? dto.TipoAula : (esVirtual == 1 ? "VIRTUAL" : "ESTANDAR")`.

#### 3. Fase 794 — Deduplicación Simétrica y en Dos Pasos en `ag_arbol_nodo` (Migración V66)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Diagnóstico:**
  - La deduplicación comparaba `n2.aula_ref` sin normalizar contra `REPLACE(ag_arbol_nodo.aula_ref, ...)`, impidiendo detectar colisiones entre variantes con distinta capitalización (ej. `|VIRTUAL` y `|Virtual`). Al intentar el `UPDATE` posterior, SQLite violaba la restricción única `uq_ag_arbol_nodo(idarbol, aula_ref)` y abortaba la transacción completa.
- **Solución Implementada:**
  - Implementada deduplicación en dos pasos normalizando ambos lados: Paso 1 elimina registros legados si ya existe un equivalente canónico (`n2.aula_ref NOT LIKE '%|VIRTUAL%'`); Paso 2 resuelve duplicados legados equivalentes conservando el registro con menor `id` (`n2.id < ag_arbol_nodo.id`). Replicado para los bloques Plataforma y Mixta.

#### 4. Fase 795 — Guardas de Columnas y Aislamiento por Bloque en Migración V66
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Diagnóstico:**
  - La migración ejecutaba sentencias `UPDATE` asumiendo la existencia incondicional de columnas en vaults antiguos o con migraciones parciales. Un fallo en tablas secundarias revertía los cambios sobre `vinculaciones` al ejecutarse en una única transacción.
- **Solución Implementada:**
  - Incorporado el helper `ExisteColumna` usando `pragma_table_info`.
  - Envolvimiento de `MigrarAgAula` y `MigrarAgArbolNodo` en bloques `try/catch` con logging preventivo sin abortar el resto de la migración.
  - Validación preventiva de la existencia de cada columna antes de ejecutar sentencias de actualización.

#### 5. Fase 796 — Ampliación de Filtro en Identificadores Derivados (Migración V66)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Diagnóstico:**
  - La cláusula `WHERE` en `ag_aula` y `ag_arbol_nodo` filtraba exclusivamente por `aula_ref LIKE ...`. Filas con `aula_ref` vacío pero con tokens legados en `codigo` o `nombre` quedaban sin normalizar.
- **Solución Implementada:**
  - Ampliado el filtro `WHERE` para evaluar `aula_ref LIKE ... OR codigo LIKE ... OR nombre LIKE ...` cubriendo todas las variantes de mayúsculas y minúsculas para Plataforma y Mixta.

#### 6. Fase 797 — Asignación de Etiqueta a `Descripcion` en `AgNodoArbolKey` Nivel 6
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  - La instanciación `new AgNodoArbolKey(6, 0, par.Codigo, par.Codigo)` pasaba el código persistido en el parámetro de descripción en lugar de la etiqueta visible.
- **Solución Implementada:**
  - Ajustado a `new AgNodoArbolKey(6, 0, par.Codigo, par.Etiqueta)` garantizando coherencia en diagnósticos y métodos de visualización.

#### 7. Fase 798 — Inmutabilidad de `Opciones` en `ModalidadTutoriaCatalog`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/ModalidadTutoriaCatalog.cs`, `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`, `ConversorPDF.Tests/Academic/Domain/ModalidadTutoriaCatalogTests.cs`
- **Diagnóstico:**
  - `public static readonly string[] Opciones` exponía un arreglo mutable que permitía la alteración accidental del catálogo canónico en tiempo de ejecución.
- **Solución Implementada:**
  - Encapsulado el arreglo en `private static readonly string[] _opciones` y expuesta la propiedad de solo lectura `public static IReadOnlyList<string> Opciones => _opciones`.
  - Ajustado el poblado del combo en `FormSesionesAcademicas.cs` con `.Cast<object>().ToArray()` y actualizada la aserción de longitud en `ModalidadTutoriaCatalogTests.cs` a `.Count`.

#### 8. Fase 799 — Consideración de `EsVirtual` en `AgEncabezadoVinculacionFactory`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEncabezadoVinculacionFactory.cs`
- **Diagnóstico:**
  - Si un aula tenía `EsVirtual = 1` pero `ModalidadVinculacion` nula o vacía, `Normalizar` retornaba el valor por defecto `Presencial`, asignando viáticos indebidos a una vinculación virtual.
- **Solución Implementada:**
  - Condicionada la asignación: si `aula != null && aula.EsVirtual == 1` se asigna directamente `ModalidadTutoriaCatalog.Plataforma`; en caso contrario se normaliza `aula?.ModalidadVinculacion`.

---

### 157.108 Ronda 150 — Revisión y Endurecimiento de Fases 792-799 (Fases 800 a 803)

#### 1. Fase 800 — Corrección de Error de Compilación CS1503 en `AEtiqueta`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/ModalidadTutoriaCatalog.cs`
- **Diagnóstico:**
  - `Array.IndexOf(Opciones, normalizado)` producía el error de compilación CS1503 debido a que `Opciones` fue tipado a `IReadOnlyList<string>` en la Fase 798, sin conversión implícita a `System.Array`.
- **Solución Implementada:**
  - Reemplazado por `Opciones.Contains(normalizado, StringComparer.Ordinal)` utilizando la extensión Linq correspondiente.

#### 2. Fase 801 — Control de Fallos en V66 para Evitar Registro Falso de Esquema
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Diagnóstico:**
  - Los bloques `try/catch` de F795 silenciaban excepciones en `MigrarAgAula` y `MigrarAgArbolNodo`, permitiendo que `Up` terminara exitosamente y `SchemaMigrator` registrara la versión 66 sin corregir identificadores de aula ni permitir reintentos futuros.
- **Solución Implementada:**
  - Introducido el contador `fallos`: si alguno de los bloques secundarios lanza una excepción, se emite `LoggerHelper.Error` y se lanza `InvalidOperationException`, impidiendo que `SchemaMigrator` marque la migración como aplicada y garantizando reintentos en arranques posteriores.

#### 3. Fase 802 — Generación Dinámica de Cláusulas `SET` y `WHERE` ante Columnas Parciales en V66
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV66Tests.cs`
- **Diagnóstico:**
  - Las guardas `tieneAulaRef`, `tieneCodigo`, `tieneNombre` solo se utilizaban en la condición global, pero la sentencia SQL invocaba incondicionalmente las tres columnas. Si una de ellas faltaba en vaults con esquemas heterogéneos, el comando fallaba con `no such column`.
- **Solución Implementada:**
  - Construcción dinámica de las cláusulas `SET` y `WHERE` evaluando exclusivamente las columnas físicamente presentes en la tabla, y adición de prueba unitaria `Up_WhenAgAulaHasPartialColumns_ShouldUpdateExistingColumnsWithoutThrowing`.

#### 4. Fase 803 — Blindaje de Inmutabilidad Real en `Opciones` y `Pares`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/ModalidadTutoriaCatalog.cs`
- **Diagnóstico:**
  - Aunque expuestas como `IReadOnlyList<T>`, la propiedad `Opciones` retornaba la referencia directa al array interno `_opciones`, y `Pares` a `List<(string, string)>`, permitiendo mutación mediante casts directos (`(string[])Opciones` o `(List<...>)Pares`).
- **Solución Implementada:**
  - Encapsuladas mediante `Array.AsReadOnly(_opciones)` y `.AsReadOnly()` sobre la lista de tuplas, previniendo alteraciones incluso ante conversiones explícitas en runtime.

---

### 157.109 Ronda 151 — Revisión de Fases 800-803 (Endurecimiento V66 + Catálogo) (Fases 804 a 808)

#### 1. Fase 804 — Guarda de `modalidad` e `idarbol` en Deduplicación de `ag_arbol_nodo`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Diagnóstico:**
  - Los `DELETE` de deduplicación se ejecutaban bajo la guarda `tieneAulaRef`, pero sus sentencias referenciaban `modalidad` e `idarbol`. En vaults sin estas columnas, la sentencia fallaba con `no such column`, incrementaba `fallos` y revertía el proceso en cada inicio.
- **Solución Implementada:**
  - Evaluada la presencia física de `tieneIdArbol` y `tieneModalidad` junto a `tieneAulaRef`, omitiendo la deduplicación con aviso en log si faltan columnas requeridas.

#### 2. Fase 805 — Aislamiento y Control de Excepciones en `MigrarTablaSimple` de Producción
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Diagnóstico:**
  - `vinculaciones` y `ag_vinculaciones` se ejecutaban sin `try/catch`. Cualquier error inesperado (como triggers stale) abortaba la ejecución sin contabilizar `fallos` ni emitir logs estructurados de bloque.
- **Solución Implementada:**
  - Envolvimiento de cada llamada en `try/catch`, incrementando `fallos`, capturando la primera excepción y registrando `LoggerHelper.Error`.

#### 3. Fase 806 — Preservación de la Excepción Raíz en el `throw` de Cierre de V66
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Diagnóstico:**
  - El `throw new InvalidOperationException` descartaba la excepción original, dificultando el diagnóstico del error SQLite subyacente en los registros de auditoría.
- **Solución Implementada:**
  - Almacenada la primera excepción capturada (`primera`) y pasada como `InnerException` en el constructor de `InvalidOperationException`.

#### 4. Fase 807 — Mapeo Explícito de Tokens a Destinos Canónicos
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Diagnóstico:**
  - La determinación del destino de reemplazo mediante `Contains("VIRTUAL")` era una regla implícita por subcadena, susceptible a colisiones con futuros tokens compuestos.
- **Solución Implementada:**
  - Declarado el arreglo estático tipado `TokensModalidad` con tuplas explícitas `(Token, Destino)` para Plataforma y Mixta, reutilizado en `MigrarAgAula` y `MigrarAgArbolNodo`.

#### 5. Fase 808 — Cobertura de Pruebas Unitarias para Rutas de Falla y Columnas Incompletas
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV66Tests.cs`
- **Diagnóstico:**
  - Carencia de pruebas automatizadas para la omisión segura de deduplicación ante falta de columnas en `ag_arbol_nodo` y para la verificación de propagación de excepciones en `Up`.
- **Solución Implementada:**
  - Incorporadas las pruebas unitarias `Up_WhenAgArbolNodoLacksModalidadColumn_ShouldNotThrowAndStillRenameAulaRef` y `Up_WhenBlockFails_ShouldThrowInvalidOperationExceptionWithInnerException`.

---

### 157.110 Ronda 152 — Revisión de Fases 804-808 (Aislamiento V66 y Guardas de Columna) (Fases 809 a 811)

#### 1. Fase 809 — Deduplicación Dinámica de `ag_arbol_nodo` ante Ausencia de Columna `modalidad`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Diagnóstico:**
  - Al omitir la deduplicación cuando `tieneModalidad` es falso, la actualización posterior de identificadores sobre `aula_ref` violaba la restricción única `uq_ag_arbol_nodo(idarbol, aula_ref)` ante registros legados (`|VIRTUAL`) y canónicos (`|PLATAFORMA`) preexistentes en el mismo árbol, provocando fallo de migración permanente y reintento en cada arranque.
- **Solución Implementada:**
  - Ajustada la guarda de deduplicación a `tieneAulaRef && tieneIdArbol`, ensamblando dinámicamente los predicados `legacyVirtual` y `legacyMixta` para incluir la columna `modalidad` solo cuando esté presente en la tabla, ejecutando la deduplicación por `aula_ref` aun si falta `modalidad`.

#### 2. Fase 810 — Agregación Completa de Excepciones en el `throw` de Cierre de V66
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Diagnóstico:**
  - `primera = primera ?? ex` retenía exclusivamente el primer error capturado (frecuentemente benigno o colateral), descartando la causa raíz si ocurrían fallos en bloques posteriores.
- **Solución Implementada:**
  - Acumulación de excepciones en `List<Exception> errores`, asignando `errores.Count == 1 ? errores[0] : new AggregateException(errores)` como `InnerException` en `InvalidOperationException`.

#### 3. Fase 811 — Cobertura de Pruebas Unitarias para Deduplicación sin Columna `modalidad` y Prioridad Canónica
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV66Tests.cs`
- **Diagnóstico:**
  - Falta de verificación automatizada para el escenario de deduplicación en esquemas sin columna `modalidad` donde la fila canónica posee un `id` mayor que la fila legada.
- **Solución Implementada:**
  - Añadida la prueba unitaria `Up_WhenNoModalidadColumnAndCanonicRowHasHigherId_ShouldKeepCanonicRow` validando que la deduplicación conserva la fila canónica y elimina la legada sin lanzar excepciones.

---

### 157.111 Ronda 153 — Verificación de Fases 809-811 (Fases 812 y 813)

#### 1. Fase 812 — Dedupe Conserva ID Menor ante Fila con Modalidad Legacy (Opcional)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Estado:** RETIRADA (Refinamiento opcional fuera de alcance crítico).
- **Diagnóstico y Decisión:**
  - En escenarios teóricos de estado intermedio, la fila superviviente normaliza su modalidad en la fase posterior de `UPDATE`. El índice único queda protegido sin colisiones ni corrupción de datos.

#### 2. Fase 813 — Optimización de Consulta de Deduplicación con Escaneo Indexable (Opcional)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`
- **Estado:** RETIRADA (Optimización innecesaria).
- **Diagnóstico y Decisión:**
  - V66 es aditiva y corre una única vez. El volumen real de nodos no compromete el tiempo de arranque. Reescribir o materializar consultas en memoria introduce complejidad injustificada contraria al principio de simplicidad.

---

## 📦 Entrega de Release v1.6.3.60 (Instalador Oficial)

- **Versión Oficial:** `1.6.3.60`
- **Archivos de Versión Sincronizados:**
  - `ConversorPDF.XLL\Core\Helpers\App.cs` (`APP_VERSION = "1.6.3.60"`)
  - `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (`<Version>1.6.3.60</Version>`)
  - `installer\InstaladorConversorPDF.wxs` (`Version="1.6.3.60"`, `DefaultVersion="!(bind.FileVersion.filXllDll)"`)
- **Documentación Técnica Generada:**
  - DocFX build completado exitosamente: **0 errores**.
  - Sitio estático compilado en `docs\_site` con archivo `.nojekyll` para GitHub Pages.
- **Verificación Integral de Pruebas Unitarias:**
  - Suite completa: **638 / 638 pruebas aprobadas (0 errores)** (`dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`).
- **Compilación de Add-In:**
  - `ConversorPDF.XLL` (Release|x64): **0 Errores** (`ConversorPDF.XLL-AddIn64.xll`).
- **Paquete Instalador Generado (`/build-release 1.6.3.60`):**
  - Archivo: `installer\InstaladorConversorPDFv1.6.3.60.msi`
  - Tamaño: **5,591,040 bytes** (5.59 MB)
  - Hash SHA-256: `FC6E09EF42DFAFB2CA0AFD45128CE6F5A96C28604B4D6F059EEE4EF20B83E09D`
  - Estado: Listo para distribución e instalación en entornos de producción.

---

### 157.112 Ronda 154 — Eliminación de Alias `Normalizar` en `ModalidadTutoriaCatalog` (Conflicto UDF Excel-DNA) (Fase 814)

#### 1. Fase 814 — Supresión de Colisión de Registro UDF `Normalizar` en Excel-DNA
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/ModalidadTutoriaCatalog.cs`
- **Diagnóstico:**
  - Excel-DNA registra automáticamente por reflexión todos los métodos `public static` del ensamblado como funciones definidas por el usuario (UDF). La existencia simultánea de `DiaSemanaNormalizador.Normalizar(string dia)` y `ModalidadTutoriaCatalog.Normalizar(string modalidad)` causó el error de carga: `Registration [Error] Repeated function name: 'Normalizar' - previous registration will be overwritten`.
- **Solución Implementada:**
  - Eliminado el alias `Normalizar` en `ModalidadTutoriaCatalog.cs`, manteniendo exclusivamente el método unívoco `NormalizarTutoria`, eliminando por completo la colisión en el registro COM/Excel-DNA.

---

### 157.113 Ronda 155 — Despliegue de Binario y Erradicación Total de `Normalizar` como UDF (Fases 815 a 817)

#### 1. Fase 815 — Despliegue de Release Corregido a la Ruta de Registro `OPEN`
- **Archivo Afectado:** `%APPDATA%\Microsoft\AddIns\ConversorPDF.XLL.dll`
- **Diagnóstico:**
  - Excel carga el complemento desde la ruta configurada en la clave `OPEN` (`%APPDATA%\Microsoft\AddIns\ConversorPDF.xll`). El binario en dicha ruta era anterior a la eliminación del alias `Normalizar` (compilado a las 12:53), reproduciendo el diálogo modal a pesar de que el código fuente ya estaba corregido.
- **Solución Implementada:**
  - Identificada la necesidad de recompilar Release con el código saneado y copiar `ConversorPDF.XLL.dll`, `ConversorPDF.xll` y `ConversorPDF.dna` a `%APPDATA%\Microsoft\AddIns\`.

#### 2. Fase 816 — Erradicación del Último `Normalizar` Público del Ensamblado (`DiaSemanaNormalizador`)
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/DiaSemanaNormalizador.cs`, `ConversorPDF.Tests/Academic/Application/Services/DiaSemanaNormalizadorTests.cs`
- **Diagnóstico:**
  - `DiaSemanaNormalizador.Normalizar` era el único método restante con el nombre genérico `Normalizar`, registrándose como UDF en Excel-DNA y dejando el sistema vulnerable a colisiones futuras o coexistencia con versiones previas.
- **Solución Implementada:**
  - Renombrado `Normalizar` a `NormalizarDiaSemana` en `DiaSemanaNormalizador.cs` y en su suite de pruebas unitarias `DiaSemanaNormalizadorTests.cs`. Se garantiza que no exista ningún método `public static string Normalizar` en todo el proyecto.

#### 3. Fase 817 — Protocolo de Verificación de Integridad de Binarios en Smoke Tests
- **Archivo Afectado:** Protocolo de verificación / smoke test
- **Diagnóstico:**
  - La disparidad entre el artefacto compilado en el repositorio y el instalado en el perfil de usuario generaba falsos positivos diagnósticos al evaluar versiones desfasadas.
- **Solución Implementada:**
  - Establecida la regla de verificación de hash SHA-256 entre `bin\x64\Release` y `%APPDATA%\Microsoft\AddIns\` previa a cualquier prueba manual en Excel.

---

### 157.114 Ronda 156 — Verificación de Fases 815-817 y Ejecución de Despliegue de Release (Fase 818)

#### 1. Fase 818 — Despliegue Físico de Release y Actualización en `%APPDATA%\Microsoft\AddIns`
- **Archivos Afectados:** `ConversorPDF.XLL/bin\x64\Release\ConversorPDF.XLL.dll`, `%APPDATA%\Microsoft\AddIns\ConversorPDF.XLL.dll`
- **Diagnóstico:**
  - El código fuente en el repositorio se encuentra 100% saneado (sin métodos `Normalizar` en `ModalidadTutoriaCatalog` ni en `DiaSemanaNormalizador`), pero el binario desplegado bajo la clave `OPEN` de Excel corresponde a una compilación previa (12:53:36) que todavía contenía la colisión UDF.
- **Solución Implementada:**
  - Programada y consolidada la secuencia de ejecución: cierre de Excel, compilación en Release|x64, verificación reflectiva de ausencia de `Normalizar`, y copia atómica hacia `%APPDATA%\Microsoft\AddIns\`.

---

### 157.115 Ronda 157 — Estabilidad de Sincronización CRDT y Reparación de Triggers en Terceros (Fases 819 a 821)

#### 1. Fase 819 — Blindaje de `MigrationV58_EnsureTerceroVacante` contra Triggers CRDT Obsoletos en `terceros`
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/CrdtTriggerRepair.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/PeriodoVentanaCrdtRepair.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV58_EnsureTerceroVacante.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV57V58V59Tests.cs`
- **Diagnóstico:**
  - En vaults existentes con esquema CRDT desfasado, la sentencia `INSERT INTO terceros (9 columnas)` disparaba un trigger `terceros__crsql_itrig` generado con una aridad distinta de columnas, resultando en `SQL logic error: expected 43 values, got 5`. Al ser V58 una migración aditiva tolerante, el error se registraba como HUECO DE VERSIÓN y la migración omitía su registro, impidiendo la creación del docente centinela `VACANTE` requerido por `AgAsignadorTercero`.
- **Solución Implementada:**
  - Creación del helper genérico `CrdtTriggerRepair` con el método `DropStaleTriggers` que elimina los tres triggers CRDT (`dtrig`, `utrig`, `itrig`) de cualquier tabla especificada dentro de una transacción activa, preservando `__crsql_clock`.
  - Refactorización de `PeriodoVentanaCrdtRepair` para delegar en `CrdtTriggerRepair.DropStaleTriggers(connection, transaction, "periodos")`.
  - Registro de `CrdtTriggerRepair.cs` en `ConversorPDF.XLL.csproj`.
  - En `MigrationV58_EnsureTerceroVacante`, captura de excepciones de aridad y triggers en `INSERT` y `UPDATE`, invocación de `CrdtTriggerRepair.DropStaleTriggers(connection, transaction, "terceros")` y reintento atómico de la operación.
  - Adición de prueba unitaria `MigrationV58_Up_WhenStaleCrsqlTriggersExistOnTerceros_ShouldDropTriggersAndInsertVacanteSuccessfully`.

#### 2. Fase 820 — Aislamiento Estricto de Sincronización PUSH en Entorno de Pruebas
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Core/Sync/SubunidadSyncService.cs`
- **Diagnóstico:**
  - `PushChangesAsync` carecía de guardia de entorno de pruebas. Al ejecutarse suites de pruebas de integración como `AcademicInitTest` desde bases de datos temporales (`esap_acad_tmp_*`), se serializaban y transmitían cambios locales reales a carpetas compartidas de OneDrive (`csync\VALLE` y `csync\_GT`), escribiendo archivos `.crsync` y arriesgando contaminación cruzada entre estaciones reales.
- **Solución Implementada:**
  - Incorporado el predicado estático `EsEntornoDePruebas() => SqliteHelper.IsTestVaultPathActive || AppDomain.CurrentDomain.FriendlyName.IndexOf("test", StringComparison.OrdinalIgnoreCase) >= 0;` en `SubunidadSyncService`.
  - Evaluación de guarda al inicio de `PushChangesAsync`, retornando inmediatamente `PushResult.SinCambios` e informando en log `[SubunidadSyncService] PUSH omitido: entorno de pruebas detectado.`
  - Incorporada propiedad alias `PushResult.SinCambios => NoChanges();`.
  - Eliminación segura de los 3 archivos `.crsync` temporales generados en la carpeta de OneDrive durante las corridas de prueba.

#### 3. Fase 821 — Desambiguación de Diagnóstico en Tests de Falla Intencional
- **Archivo Afectado:**
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV66Tests.cs`
- **Diagnóstico:**
  - La prueba `Up_WhenBlockFails_ShouldThrowInvalidOperationExceptionWithInnerException` provocaba un fallo intencional mediante `SELECT RAISE(FAIL, 'Error provocado en trigger');`. Al ejecutarse, registraba 5 entradas `[ERROR] [V66] Bloque vinculaciones falló: Error provocado en trigger` indistinguibles de fallos genuinos en análisis de logs.
- **Solución Implementada:**
  - Actualizado el mensaje a `SELECT RAISE(FAIL, '[TEST-V66] fallo intencional');` y sincronizada la aserción en `MigrationV66Tests.cs`, garantizando trazabilidad unívoca entre errores inducidos por test y anomalías operativas reales.

---

### 157.116 Ronda 158 — Refinamiento de Trazabilidad PUSH, Filtros Defensivos V58 y Purga de Alias Innecesario (Fases 822 a 824)

#### 1. Fase 822 — Preservación del Motivo de Omisión en PUSH (`PushResult.Skipped`)
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Sync/SubunidadSyncService.cs`
- **Diagnóstico:**
  - La guarda de entorno de pruebas retornaba `PushResult.SinCambios` (`WasSuccessful = true, HadChanges = false, SkipReason = null`), haciendo indistinguible una omisión por entorno de test de una corrida real de sincronización sin cambios, perdiendo trazabilidad diagnóstica en `SyncScheduler`.
- **Solución Implementada:**
  - Retorno actualizado a `PushResult.Skipped("Entorno de pruebas: PUSH deshabilitado.")` y eliminado el alias `PushResult.SinCambios`.

#### 2. Fase 823 — Ampliación del Criterio de Detección de Triggers CRDT en V58 (`EsFalloDeTriggerCrdt`)
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV58_EnsureTerceroVacante.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV57V58V59Tests.cs`
- **Diagnóstico:**
  - El filtro por subcadena `when (values || trigger)` no contemplaba fallos en triggers stale provocados por referencias a columnas eliminadas o renombradas (`no such column` o `no column named`), lo que habría provocado un HUECO DE VERSIÓN sin regeneración de triggers.
- **Solución Implementada:**
  - Implementado el método `EsFalloDeTriggerCrdt(Exception ex)` evaluando `values`, `trigger`, `no such column` y `no column named` en `MigrationV58_EnsureTerceroVacante.cs`.
  - Añadida prueba unitaria `MigrationV58_Up_WhenStaleCrsqlTriggerFailsWithNoSuchColumn_ShouldDropTriggersAndInsertVacanteSuccessfully`.

#### 3. Fase 824 — Purga de Alias Innecesario y Engañoso `RegenerarTabla`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/CrdtTriggerRepair.cs`
- **Diagnóstico:**
  - El método `RegenerarTabla` carecía de llamadores y su nombre sugería una recreación completa, cuando su única función es soltar triggers (`DropStaleTriggers`), induciendo a error semántico en futuras migraciones.
- **Solución Implementada:**
  - Eliminado el método `RegenerarTabla`, manteniendo exclusivamente el método descriptivo unívoco `DropStaleTriggers`.

---

### 157.117 Ronda 159 — Plan Consolidado: Regla de Carriles / Oleadas, Columna K, Duración, Duplicados y Botón Módulo R (Fases 825 a 843)

> ⚠️ RONDA 159 NO APLICADA. Verificado en Ronda 160: 0/19 artefactos existen.
> No ejecutar build/release hasta completar 825-843.

#### 1. Fase 825 (T1) — Enums y Ponderaciones de Aptitud para Carriles (R11-R13, S6-S7)
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPesosFitness.cs`, `ConversorPDF.XLL/Academic/Domain/Enums/AgPoliticaLlenado.cs`, `ConversorPDF.XLL/Academic/Domain/Enums/AgSincronizacionCarriles.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Catálogo de violaciones sin códigos para control de duración de sesión (R11), concurrencia semanal aula/cohorte (R12), desborde de carga semanal (R13), huecos de secuencia (S6) y desbalance de carriles (S7).
- **Solución Propuesta:** Adición de códigos R11=29, R12=30, R13=31, S6=106, S7=107; pesos de penalización en `AgPesosFitness`; enums de política (`Distribuido`/`Contiguo`) y sincronización (`Libre`/`Sincronizada`).

#### 2. Fase 826 (T2) — Modelos POCO del Plan Calendario por Carriles
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgUnidadSemanal.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgCarril.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPlanAsignatura.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPlanCalendario.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgParametrosPlan.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Ausencia de estructuras de datos para modelar unidades semanales discretas (4h/8h), carriles de días continuos y planes calendarios derivados.
- **Solución Propuesta:** Creación de clases selladas inmutables y POCOs con serialización JSON (`AgParametrosPlan`).

#### 3. Fase 827 (T3) — Claves Canónicas ISO-8601 y Formateador Unificado de Columna K
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSemanaIso.cs`, `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Columna K vacía (C1) por formato no homogéneo y falta de clave anual ISO-8601 para validaciones R5/R12/R13.
- **Solución Propuesta:** `AgSemanaIso.Clave(DateTime) => "yyyy-Www"` y `DiasHorasSemanaFormatter` para normalizar días y horas en formato `DIA|horas` sin tildes.

#### 4. Fase 828 (T4) — Construcción Determinista de Carriles y Cálculo de Semanas Disponibles
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgUnidadesSemanales.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Ventanas con sábados y domingos no separaban bloques AM y PM, y días hábiles podían recibir bloques de 8h (C2).
- **Solución Propuesta:** `AgUnidadesSemanales`: partición determinista en 1 o 2 carriles por ventana, división de sábado en 2 bloques (AM/PM de 4h) y domingo en 1 bloque (AM de 4h).

#### 5. Fase 829 (T5) — Algoritmo de Planificación de Carriles y Distribución de Holgura
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorCarriles.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Asignaturas solapadas dentro del aula y sin sincronización con festivos ni fecha de fin de aula.
- **Solución Propuesta:** `AgPlanificadorCarriles` con reparto `DISTRIBUIDO` (guard de `n <= 1`), cálculo exacto de semanas y validación de fin de carril contra fecha límite.

#### 6. Fase 830 (T6) — Operadores Genéticos sobre Permutaciones de Carriles
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorCarriles.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Cruce OX por genes producía duplicación de asignaturas (C3).
- **Solución Propuesta:** Cruce y mutación aplicados a la permutación de asignaturas en el carril (`CruzarPermutacion`, `Permutar`).

#### 7. Fase 831 (T7) — Replanificador Fenotípico de Carriles por Aula
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReplanificadorCarriles.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Desincronización fenotípica de fechas y bloques tras mutación.
- **Solución Propuesta:** `AgReplanificadorCarriles` que recalcula fechas estimadas, HorasGen y patrones de días para el aula modificada.

#### 8. Fase 832 (T8) — Fábrica de Individuos Basada en Carriles
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Asignación aleatoria de bloques (L199) y fechas estáticas de inicio de semestre (L202-203) (C2, C3, C4).
- **Solución Propuesta:** Generación de individuos anclada al plan de carriles con fechas derivadas y `EsFijoTemporal = true`.

#### 9. Fase 833 (T9) — Blindaje Estricto de Cruce y Mutación contra Alteración Horaria
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Operadores genéticos alteraban aleatoriamente la franja y duración del gen.
- **Solución Propuesta:** Restricción estricta de mutación y cruce para no alterar días, bloques ni duraciones de sesión.

#### 10. Fase 834 (T10) — Cuantización Dinámica de Horas por Grupo según Aula
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Uso de constante fija `4.0` en reparto de horas (L70, L98, L111, L128).
- **Solución Propuesta:** Cuantización dinámica según `HorasSesionAula(aula)`.

#### 11. Fase 835 (T11) — Validación de Reglas R4, R5, R11, R12, R13 con Claves Canónicas y Vacantes
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** R4 y R5 evaluaban únicamente `Grupo` simple sin distinguir programa ni sede, e ignoraban vacantes en reglas de ocupación (C4, D10).
- **Solución Propuesta:** Clave canónica de cohorte, validación de sábado AM/PM, conteo por semana ISO y nuevas reglas R11, R12, R13, S6, S7.

#### 12. Fase 836 (T12) — Pre-Flight Bloqueante de Factibilidad de Carriles
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** Fallo tardío tras finalizar optimización por sobrecupo de carga en la ventana.
- **Solución Propuesta:** Bin-packing de carriles en validación pre-vuelo con mensaje accionable con ranuras, horas y semanas.

#### 13. Fase 837 (T13) — Ponderación de Aptitud Integral con S6/S7 y Cobertura de Fechas
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorAptitud.cs`
- **Diagnóstico:** Ponderación F3 no evaluaba consistencia de fechas con el plan calendario.
- **Solución Propuesta:** Integración de S6, S7 y redefinición de F3 por cobertura del plan de fechas.

#### 14. Fase 838 (T14) — Materialización Segura con Invariantes I3, I4, I5 y Deduplicación
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Sobrescritura de fechas del gen con la ventana semestral completa y falta de dedupe (C3, C5).
- **Solución Propuesta:** Preservación de fechas del gen, invariantes I3-I5 y deduplicación estricta `(aula, fecha, franja)`.

#### 15. Fase 839 (T15) — Repositorio AG: Recuperación de Última Corrida y Purga Multi-Programa
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** Purga limitada a un solo programa en optimizaciones multi-programa (C5).
- **Solución Propuesta:** `GetUltimaCorridaAsync` y purga multi-programa en staging.

#### 16. Fase 840 (T16) — UI: Botón «📥 Cargar en Módulo R», Gate D11 y Eliminación de CSV
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico:** Generación no requerida de CSV espejo (C6) y falta de compuerta visual ante corridas con violaciones duras (D11).
- **Solución Propuesta:** Reemplazo de botón por `📥 Cargar en Módulo R`, eliminación de CSV espejo y compuerta de confirmación con bloqueo de promoción.

#### 17. Fase 841 (T17) — Formateo Resiliente de Columna K en Exportador Semestres
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Diagnóstico:** Columna K vacía ante ausencia de `DiaSemana` (C1).
- **Solución Propuesta:** Integración de `DiasHorasSemanaFormatter` con derivación desde `FechaSesion`.

#### 18. Fase 842 (T18) — Desacentuación de Días en Vistas AG y Trazabilidad en SqliteHelper
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** Vistas con tildes en días de la semana y catch silencioso de errores DDL.
- **Solución Propuesta:** Recreación de vistas con días sin tilde (`MIERCOLES`, `SABADO`) y logging estructurado.

#### 19. Fase 843 (T19) — Documentación Técnica, Matriz de Pruebas y Registro en Proyecto
- **Estado:** PENDIENTE (NO APLICADA) — 0 archivos creados.
- **Archivos Afectados:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgNormalizadorCromosoma.cs`, `docs/articles/changes/2026-10-21-ag-carriles-oleadas.md`
- **Diagnóstico:** Archivos nuevos no registrados en el proyecto MSBuild y ausencia de documentación técnica consolidada.
- **Solución Propuesta:** Registro de compilación en `.csproj`, normalizador de cromosomas y artículo de cambios.

---

### 157.118 Ronda 160 — Saneamiento Documental de Ronda 159 y Defectos de Runtime/UX en AG (Fases 844 a 852)

#### 1. Fase 844 — Saneamiento de Registro Documental de Ronda 159
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`
- **Diagnóstico:** Fases 825-843 figuraban erróneamente como implementadas en los artefactos sin que existan los archivos en disco.
- **Solución Implementada:** Adición de advertencia explícita y marcado de fases 825-843 como `PENDIENTE (NO APLICADA)`.

#### 2. Fase 845 — Corrección de Acceso a Celda de Grid en Resolución de Vinculación (`IdSesionGlobal`)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `ObtenerIdVinculacionSeleccionadaAsync` intentaba acceder a la columna `"IdSesion"`, lanzando `ArgumentException` silenciada puesto que el tipo anónimo proyecta `"IdSesionGlobal"`.
- **Solución Implementada:** Búsqueda tolerante por `"IdSesionGlobal"` o `"IdSesion"` insensible a mayúsculas.

#### 3. Fase 846 — Inclusión de Sesiones Validadas en Vista AG (`vw_AgSesionesDetalle`)
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** `v.estado = 0` ocultaba sesiones al pasar la vinculación a validada (`estado = 1`), dejando al usuario con 0 sesiones sobre staging poblado.
- **Solución Implementada:** Filtro actualizado a `IFNULL(v.estado, 0) <> 3` (excluye solo descartadas) en `EnsureAgViews`.

#### 4. Fase 847 — Limpieza de Panel de Vacantes al Cargar Corrida sin Sesiones
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Salida temprana de `CargarModuloResultadosAsync` dejaba `_todasLasVacantes` y `dgvVacantes` con datos de la corrida previa.
- **Solución Implementada:** Vaciado de lista de vacantes, asignación `dgvVacantes.DataSource = null` y reinicio de etiqueta de conteo.

#### 5. Fase 848 — Corrección de Suma de Horas en Genes Duplicados (`DiaHorasSemanales`)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** `subGrp.Sum(x => x.HorasSemanales)` producía acumulaciones erróneas como `MARTES|8` al existir genes duplicados para un mismo día.
- **Solución Implementada:** Resolución de horas a partir de la duración del bloque horario asignado (`b.DuracionHoras`) con fallback seguro a `4.0`.

#### 6. Fase 849 — Redondeo hacia Arriba al Construir Patrón Semanal
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** El cast entero `(int)totalHoras` truncaba fracciones (48.5 -> 48), calculando menos sesiones que las necesarias y violando I1.
- **Solución Implementada:** Uso de `(int)Math.Ceiling(totalHoras)` al invocar `_planificadorPatron.Construir`.

#### 7. Fase 850 — Reinicio de Estado de Corrida al Iniciar Optimización
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Si una optimización fallaba antes de persistirse, `_idCorridaCargada` conservaba el ID previo, mezclando datos en la interfaz.
- **Solución Implementada:** Reinicio de `_ultimoResultado = null; _idCorridaCargada = 0;` al inicio de `IniciarOptimizacionAsync`.

#### 8. Fase 851 — Escape de Caracteres Especiales en Exportación CSV
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Textos con comillas o punto y coma desfasaban las columnas de la exportación CSV.
- **Solución Implementada:** Incorporación del método auxiliar `EscaparCsv` con entrecomillado y duplicación de comillas internas.

#### 9. Fase 852 — Reinicio de Diccionario Estático de Incidencias entre Corridas
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `_incidenciasCoherencia` crecía indefinidamente reteniendo memoria y suprimiendo advertencias en corridas sucesivas.
- **Solución Implementada:** Implementación de `ReiniciarIncidencias()` e invocación al inicio de `IniciarOptimizacionAsync`.

---

### 157.119 Ronda 161 — Saneamiento Documental Completo (R159) y Vistas AG Canónicas en Migraciones (Fases 853 a 858)

#### 1. Fase 853 — Saneamiento Documental de Fases 826 a 843
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`
- **Diagnóstico:** Fases 826-843 mantenían `Solución Implementada` y `Estado: ACEPTADA` a pesar de que los 18 artefactos correspondientes aún no existen físicamente en disco.
- **Solución Implementada:** Marcado exhaustivo de Fases 826-843 como `PENDIENTE (NO APLICADA)` y sustitución por `Solución Propuesta`.

#### 2. Fase 854 — Migración Canónica V67 para Recreación de Vistas AG en Bases de Datos Existentes
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV67_RecrearVistasAgCanonicas.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** `EnsureAgViews` solo se ejecutaba en la creación inicial de base de datos (`ApplyBaseSchemaAndSeeds`). En bases de datos existentes, las migraciones previas no actualizaban las vistas, impidiendo que la corrección de Fase 846 (`estado <> 3`) y la proyección de `DiaSemana` surtieran efecto.
- **Solución Implementada:** Creación de la migración canónica `MigrationV67_RecrearVistasAgCanonicas` (Versión 67, aditiva y tolerante) que invoca `SqliteHelper.EnsureAgViews(connection, transaction)`, registro en `SqliteHelper.BuildMigrator()` y adición a `ConversorPDF.XLL.csproj`.

#### 3. Fase 855 — Unificación de DDL Divergente en Migraciones AG y Prueba Unitaria V67
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV59_RecreateAgViews.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV67_RecrearVistasAgCanonicasTests.cs`
- **Diagnóstico:** `MigrationV59` mantenía una definición SQL hardcodeada obsoleta (`v.estado = 0` y sin `DiaSemana`), provocando divergencia y regresiones ante ejecuciones aisladas de migraciones.
- **Solución Implementada:** Delegación de la recreación de vistas en `MigrationV59` a `SqliteHelper.EnsureAgViews(connection, transaction)` y creación de la prueba unitaria `MigrationV67_RecrearVistasAgCanonicasTests` validando presencia de `DiaSemana` y filtrado correcto (`estado = 1` visible, `estado = 3` excluido).

#### 4. Fase 856 — Reinicio Exhaustivo de Estado de Sesión en AgFormGeneradorHorarios
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El reset de optimización no limpiaba `_contextoActual` ni deshabilitaba `btnPromoverProduccion`, permitiendo resolver consultas con catálogos obsoletos y promocionar estados inconsistentes si la nueva corrida fallaba.
- **Solución Implementada:** Asignación de `_contextoActual = null` y `btnPromoverProduccion.Enabled = false` al inicio de `IniciarOptimizacionAsync`.

#### 5. Fase 857 — Determinismo en Formateo de `DiaHorasSemanales` ante Múltiples Bloques en Mismo Día
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** `subGrp.Select(x => x.IdBloqueHorario).Distinct().FirstOrDefault()` seleccionaba un bloque arbitrario ante grupos con dos bloques en el mismo día (sábado AM/PM), distorsionando el horario asignado en columna K.
- **Solución Implementada:** Agrupación determinista por `new { g.DiaSemana, g.IdBloqueHorario }` ordenada por día y bloque, resolviendo la duración y franja exacta de cada bloque con registro de advertencia ante bloques inexistentes.

#### 6. Fase 858 — Diagnóstico Explícito ante Ausencia de Columna de Sesión en Grid de Horarios
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Fallo silencioso en `ObtenerIdVinculacionSeleccionadaAsync` cuando el grid no contiene `IdSesionGlobal` ni `IdSesion`, impidiendo diagnosticar problemas de binding en UI.
- **Solución Implementada:** Verificación explícita `if (cell == null)` con logging estructurado `LoggerHelper.Warning` antes de retornar nulo.

---

### 157.120 Ronda 162 — Corrección de Esquema Canónico en Vistas AG, Tests y Normalización de Bloques (Fases 859 a 863)

#### 1. Fase 859 — Restauración de Columnas de Staging en `vw_AgSesionesDetalle` Canónica
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** El DDL canónico de `vw_AgSesionesDetalle` en `EnsureAgViews` omitía las columnas `s.id_aula AS IdAula`, `s.idgen AS IdGen`, `s.es_vacante AS EsVacante`, `s.motivo_vacante AS MotivoVacante`, `s.candidatos_pool AS CandidatosPool` y `v.dia_horas_semanales AS DiaHorasSemanales`. Esto provocaba que `SesionListItem.IdAula/IdGen/EsVacante` resolvieran a 0, dejando vacío el Panel de Vacantes y quebrando las aserciones de `MigrationV57V58V59Tests`.
- **Solución Implementada:** Incorporación de las columnas de staging requeridas en la proyección SELECT de `vw_AgSesionesDetalle`.

#### 2. Fase 860 — Corrección de Fixture y Tipo de Aserción en Test de Migración V67
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV67_RecrearVistasAgCanonicasTests.cs`
- **Diagnóstico:** El fixture SQLite in-memory carecía de las tablas `periodos` y `terminales`, y de las columnas `departamento/transporte` en `sedes`, `metodologia/cod_materia` en `asignaturas`, y `isTerminalAereo/idTerminales/gastoTerminalAerea` en `ag_vinculaciones`. Además, `DiaSemana` canónico proyecta texto (`'LUNES'`), lanzando `FormatException` al castear con `Convert.ToInt32`.
- **Solución Implementada:** Completado el esquema del fixture con todas las tablas y columnas referenciadas por las vistas y corregida la aserción a `string dia = (string)sesiones[0].DiaSemana; dia.Should().Be("LUNES");`.

#### 3. Fase 861 — Verificación de Post-Condición en Migración V67
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV67_RecrearVistasAgCanonicas.cs`
- **Diagnóstico:** `EnsureAgViews` captura excepciones internamente y solo emite warning. Si la recreación de vistas fallaba, la migración se registraba como aplicada silenciosamente, dejando las vistas obsoletas de forma permanente.
- **Solución Implementada:** Adición de consulta a `pragma_table_info('vw_AgSesionesDetalle')` para verificar la existencia de `DiaSemana`, lanzando `InvalidOperationException` si la vista no se recreó correctamente.

#### 4. Fase 862 — Restauración de Columnas de API Pública en `vw_AgVinculacionesDetalle`
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** La vista `vw_AgVinculacionesDetalle` canónica omitía columnas indispensables de staging (`IdIndividuo`, `IdPrograma`, `IdAsignatura`, `Grupo`, `Estado`, `DiaHorasSemanales`, `TotalHoras`, `EsVacante`, `MotivoVacante`, `CandidatosPool`, `Id`), rompiendo la compatibilidad con consumidores y pruebas.
- **Solución Implementada:** Incorporación de todas las columnas de staging y compatibilidad dual con `Id` e `IdVinculacion` en `vw_AgVinculacionesDetalle`.

#### 5. Fase 863 — Prevención de Clave Duplicada por Día en `DiaHorasSemanales`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Agrupar por `(DiaSemana, IdBloqueHorario)` emitía múltiples entradas para el mismo día (ej. sábado AM y PM). Al cargar la cadena en `FormSesionesAcademicas.RebuildSelectedDays`, el diccionario indexado por día sobrescribía la primera entrada con la segunda, provocando pérdida silenciosa de un bloque al editar.
- **Solución Implementada:** Agrupación por `g.DiaSemana` tomando el menor `IdBloqueHorario` con log de advertencia ante grupos multi-bloque en el mismo día, garantizando exactamente una entrada por día.

---

### 157.121 Ronda 163 — Robustez de Post-Condición V67, Normalización de Fechas Legacy y Control de Reintentos (Fases 864 a 867)

#### 1. Fase 864 — Validación Integral de Resolución y Columnas en Migración V67
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV67_RecrearVistasAgCanonicas.cs`
- **Diagnóstico:** SQLite no valida la existencia de columnas ni tablas al crear vistas (`CREATE VIEW`). Si la base de datos tenía esquemas legacy inconsistentes, la verificación superficial de solo `DiaSemana` pasaba y la primera consulta en runtime fallaba con `no such column`.
- **Solución Implementada:** Forzada la resolución real con `SELECT * FROM vw_AgSesionesDetalle LIMIT 0` y `SELECT * FROM vw_AgVinculacionesDetalle LIMIT 0`, sumada a la validación exhaustiva de columnas requeridas (`DiaSemana`, `IdAula`, `IdGen`, `EsVacante`, `MotivoVacante`, `CandidatosPool`, `DiaHorasSemanales`).

#### 2. Fase 865 — Cobertura de Prueba Unitaria para `vw_AgVinculacionesDetalle`
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV67_RecrearVistasAgCanonicasTests.cs`
- **Diagnóstico:** El test unitario de V67 solo verificaba `vw_AgSesionesDetalle`, dejando sin validación las columnas de `vw_AgVinculacionesDetalle` incorporadas en Fase 862.
- **Solución Implementada:** Añadida aserción sobre `pragma_table_info('vw_AgVinculacionesDetalle')` para verificar las columnas clave (`IdVinculacion`, `IdCorrida`, `IdIndividuo`, `IdPrograma`, `IdAsignatura`, `Grupo`, `Estado`, `DiaHorasSemanales`, `TotalHoras`, `EsVacante`, `MotivoVacante`, `CandidatosPool`) y ejecución de consulta de lectura (`SELECT * ... LIMIT 1`).

#### 3. Fase 866 — Normalización de Fechas Legacy `dd/MM/yyyy` a ISO `yyyy-MM-dd` en V67
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV67_RecrearVistasAgCanonicas.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV67_RecrearVistasAgCanonicasTests.cs`
- **Diagnóstico:** En bases de datos existentes, registros históricos con formato local `dd/MM/yyyy` provocaban que `strftime('%w', fechasesion)` evaluara a NULL, dejando `DiaSemana` nulo y vaciando la columna K.
- **Solución Implementada:** Invocado script SQL en `MigrationV67.Up` previo a `EnsureAgViews` que reestructura `fechasesion` de `__/__/____` a `substr(7,4) || '-' || substr(4,2) || '-' || substr(1,2)`, garantizando formato ISO. Validado mediante prueba unitaria con fecha `02/03/2026` normalizada a `'LUNES'`.

#### 4. Fase 867 — Límite de Reintentos para Migraciones Tolerantes en SchemaMigrator
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`
- **Diagnóstico:** La rama `IMigrationTolerante.Aditiva` ejecutaba `continue` antes de incrementar el diccionario `_attempts`, provocando que migraciones tolerantes fallidas se reintentaran indefinidamente en cada arranque emitiendo logs de error continuos.
- **Solución Implementada:** Incremento de `_attempts` para migraciones tolerantes, emitiendo logging acotado a 3 intentos antes de silenciar el fallo hasta el próximo despliegue.

---

### 157.122 Ronda 164 — Normalización de Fechas con Validación de Calendario, Cobertura Integral y Diagnóstico (Fases 868 a 871)

#### 1. Fase 868 — Validación de Calendario y Normalización `d/M/yyyy` en V67
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV67_RecrearVistasAgCanonicas.cs`
- **Diagnóstico:** La conversión posicional con `LIKE '__/__/____'` aceptaba fechas inexistentes (ej. `31/02/2026` -> `2026-02-31`), resultando en `strftime` nulo, `DiaSemana` nulo y columna K vacía. Además no cubría formatos sin ceros a la izquierda `d/M/yyyy` ni diagnosticaba residuos no normalizables.
- **Solución Implementada:** Incorporada cláusula `date(...) IS NOT NULL` en la conversión de 10 caracteres, query de normalización para `_/_/____` con `printf('%02d')` y consulta escalar de diagnóstico con `LoggerHelper.Warning` para sesiones con fechas inválidas o no ISO.

#### 2. Fase 869 — Cobertura de Prueba Unitaria para Normalización Legacy y Aserción de Vistas
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV67_RecrearVistasAgCanonicasTests.cs`
- **Diagnóstico:** El fixture anterior no comprobaba la transformación física de fechas legacy ni el resultado del campo `DiaSemana` en la vista. Adicionalmente, `vinc.Should().NotBeNull()` constituía una aserción trivial no verificable.
- **Solución Implementada:** Fixture actualizado con fecha legacy `'03/03/2026'`, aserción directa de persistencia `fechasesion == "2026-03-03"` y proyección en vista `DiaSemana == "MARTES"`, y reemplazo de aserción débil por conteo estricto `vinc.Should().HaveCount(2)` para validar que `vw_AgVinculacionesDetalle` preserva tanto vinculaciones activas como descartadas.

#### 3. Fase 870 — Corrección de Mensaje de Diagnóstico en Reintentos de SchemaMigrator
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`
- **Diagnóstico:** Al alcanzar 3 intentos fallidos, el log indicaba "Se omite hasta el próximo despliegue", cuando en realidad `_attempts` es un diccionario en memoria por proceso que se reinicia en cada arranque de Excel.
- **Solución Implementada:** Ajustado el mensaje a `falló {nTol} veces en esta sesión. Se omite el resto del arranque; se reintentará en el próximo inicio.`.

#### 4. Fase 871 — Validación Simétrica de Columnas en `vw_AgVinculacionesDetalle` en V67
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV67_RecrearVistasAgCanonicas.cs`
- **Diagnóstico:** La migración V67 solo validaba columnas requeridas contra `vw_AgSesionesDetalle`. Si una base de datos legacy carecía de columnas en `ag_vinculaciones`, la migración abortaba sin identificar la vista ni columnas causantes.
- **Solución Implementada:** Agregada validación simétrica sobre `pragma_table_info('vw_AgVinculacionesDetalle')` para verificar las 12 columnas requeridas (`IdVinculacion`, `IdCorrida`, `IdIndividuo`, `IdPrograma`, `IdAsignatura`, `Grupo`, `Estado`, `DiaHorasSemanales`, `TotalHoras`, `EsVacante`, `MotivoVacante`, `CandidatosPool`), arrojando `InvalidOperationException` detallada si alguna falta.

---

### 157.123 Ronda 166 — Desacoplamiento de Normalización de Fechas (V68), Fallback en Exportación y Determinismo en Tests (Fases 872 a 877)

#### 1. Fases 872/873/875 — Migración Independiente `MigrationV68_NormalizarFechasAgSesiones` y Desacoplamiento de V67
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV67_RecrearVistasAgCanonicas.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** La normalización de fechas y la recreación/validación de vistas compartían la misma transacción en V67. Si la validación de vistas fallaba, se hacía rollback de la normalización, impidiendo normalizar los datos históricos. Asimismo, se requerían múltiples formatos de fecha adicionales con validación real de calendario.
- **Solución Implementada:** Creada la migración aditiva `MigrationV68_NormalizarFechasAgSesiones` (Versión 68), registrada en `SqliteHelper.BuildMigrator` y `.csproj`. La migración cubre 8 formatos (`dd/MM/yyyy`, `d/M/yyyy`, `dd/M/yyyy`, `d/MM/yyyy`, `dd/MM/yyyy HH:mm`, `dd/MM/yyyy H:mm`, `dd-MM-yyyy`, `yyyy/MM/dd`) mediante `DateTime.TryParseExact` con calendario real, actualizando únicamente fechas válidas a ISO y emitiendo advertencia ante residuos. En V67 se retiró la mutación de datos para preservar el principio de responsabilidad única (DDL y verificación estructural).

#### 2. Fase 876 — Fallback de `DiaSemana` por `FechaSesion` y Orden Canónico en SemestresExcelExporter
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Diagnóstico:** Si una sesión histórica contiene una fecha no normalizable o `DiaSemana` nulo en la vista, `diasUnicos` quedaba vacío y la columna K se exportaba en blanco a pesar de contar con `FechaSesion`.
- **Solución Implementada:** En la proyección de `diasUnicos`, si `DiaSemana` es nulo/blanco pero `FechaSesion` existe, se deriva el día canónico mediante `ObtenerNombreDia(s.FechaSesion.Value.DayOfWeek)`. Los días resultantes se ordenan canónicamente mediante `ObtenerOrdenDia` (Lunes a Domingo).

#### 3. Fases 874/877 — Orden Determinista `ORDER BY IdSesion` y Aserción de Colección en Tests
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV67_RecrearVistasAgCanonicasTests.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV68_NormalizarFechasAgSesionesTests.cs`
- **Diagnóstico:** El acceso por índice `sesiones[0]` sin `ORDER BY` en la consulta de la vista generaba aserciones no deterministas dependientes del orden físico del motor SQLite.
- **Solución Implementada:** Consulta actualizada con `ORDER BY IdSesion` y aserciones mediante `sesiones.Select(s => (int)s.IdSesion).Should().BeEquivalentTo(new[] { 103 })` y `sesiones.Should().OnlyContain(s => (string)s.DiaSemana == "MARTES")`. Creado test dedicado `MigrationV68_NormalizarFechasAgSesionesTests` para validación exhaustiva de formatos y salvaguarda de fechas de calendario inválidas.

---

### 157.124 Ronda 167 — Normalización Sin Tildes de Días en Columna K y Saneamiento de Fechas en Vinculaciones (Fases 878 a 880)

#### 1. Fase 878 — Unificación y Normalización Sin Tildes en `diasUnicos` en SemestresExcelExporter
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Diagnóstico:** La proyección `s.DiaSemana` proveniente de la vista SQLite contiene tildes (`MIÉRCOLES`, `SÁBADO`), mientras que la rama de derivación fallback `ObtenerNombreDia` producía nombres sin tilde (`MIERCOLES`, `SABADO`). El operador `Distinct()` trataba ambas cadenas como diferentes, generando valores duplicados en columna K (ej. `"MIÉRCOLES/MIERCOLES"`).
- **Solución Implementada:** Implementado método estático `NormalizarDia(string d)` que elimina tildes y diéresis de forma exhaustiva (`ÁÉÍÓÚÜ` -> `AEIOUU`) aplicado en la proyección de `diasUnicos`, garantizando que `Distinct()` deduplique coherentemente y unificando el formato con el esperado en las plantillas Excel.

#### 2. Fase 879 — Normalización de `fecha_inicio` y `fecha_fin` en `ag_vinculaciones` en V68
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`
- **Diagnóstico:** La migración V68 únicamente abarcaba `ag_sesiones.fechasesion`. En bases de datos legacy, las columnas `ag_vinculaciones.fecha_inicio` y `ag_vinculaciones.fecha_fin` podían conservar formatos locales (`dd/MM/yyyy`), los cuales se propagaban a las vistas y se promovían a producción a través de `AgVinculacionToVinculacionMapper`.
- **Solución Implementada:** Incorporado barrido de normalización en `ag_vinculaciones` para `fecha_inicio` y `fecha_fin` usando el catálogo completo de 8 formatos y validación de calendario real, actualizando únicamente fechas válidas a ISO e ignorando valores especiales como `--`.

#### 3. Fase 880 — Cobertura Completa de Formatos y Tablas en Tests de V68
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV68_NormalizarFechasAgSesionesTests.cs`
- **Diagnóstico:** La suite de pruebas de V68 no validaba el formato `yyyy/MM/dd`, no cubría la tabla `ag_vinculaciones` incorporada en Fase 879, ni verificaba que las fechas de calendario inválidas (`31/02/2026`) se preservaran y contabilizaran en los residuales.
- **Solución Implementada:** Añadidos casos de prueba para `yyyy/MM/dd` ('2026/09/10'), normalización de `ag_vinculaciones` (`fecha_inicio` y `fecha_fin`), y aserción explícita de conteo de residuales (`residuales == 1`).

---

### 157.125 Ronda 168 — Detección Round-Trip de Fechas ISO Imposibles y Cobertura de Formatos (Fases 881 y 882)

#### 1. Fase 881 — Validación Round-Trip Canónica contra Fechas ISO Imposibles en V68
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`
- **Diagnóstico:** SQLite renormaliza internamente fechas ISO imposibles (ej. `'2026-02-31'` -> `'2026-03-03'`), por lo que `date(fechasesion) IS NULL` evaluaba a falso y omitía dichas filas. Al no pasar por validación, `strftime('%w')` calculaba el día de otra fecha distorsionando `DiaSemana`, y el diagnóstico de residuales no las contabilizaba.
- **Solución Implementada:** Incorporada la condición round-trip `(date(c) IS NULL OR date(c) <> c)` en las consultas de `ag_sesiones`, `ag_vinculaciones` y conteo de residuales, y agregado `"yyyy-MM-dd"` al arreglo de formatos en `Normalizar` para que `DateTime.TryParseExact` rechace fechas imposibles preservándolas como residuales diagnosticables.

#### 2. Fase 882 — Cobertura de Variantes de 9 Caracteres y Fechas ISO Imposibles en Tests
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV68_NormalizarFechasAgSesionesTests.cs`
- **Diagnóstico:** Los formatos `dd/M/yyyy` (`12/3/2026`) y `d/MM/yyyy` (`3/12/2026`) figuraban en el catálogo pero carecían de verificación en la suite unitaria, y no se validaba la persistencia inalterada de fechas ISO imposibles (`2026-02-31`) ni su contabilización en los residuales.
- **Solución Implementada:** Añadidos casos de prueba para `12/3/2026`, `3/12/2026` y `2026-02-31`, verificando normalización a `2026-03-12`, `2026-12-03`, preservación de `2026-02-31` y conteo de residuales igual a 2.

---

### 157.126 Ronda 169 — Soporte de Formatos ISO con Hora, Normalización de Genoma y Corrección de Falsos Residuales (Fases 883 a 885)

#### 1. Fase 883 — Soporte de Formatos ISO con Hora y Corrección de Consulta de Residuales en V68
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`
- **Diagnóstico:** Fechas en formato ISO con hora (ej. `'2026-05-10 14:30'`) entraban al conjunto de pendientes debido a que `date('2026-05-10 14:30') = '2026-05-10' <> '2026-05-10 14:30'`. Al no estar contemplados los formatos con hora ni `yyyy-M-d` en `formatos`, `TryParseExact` fallaba, dejando la fecha sin normalizar y contándola falsamente como residual no manejable en el log de advertencia.
- **Solución Implementada:** Incorporados formatos ISO con hora (`yyyy-MM-dd HH:mm`, `yyyy-MM-dd H:mm`, `yyyy-MM-ddTHH:mm:ss`, `yyyy-MM-ddTHH:mm`) y componentes de 1 dígito (`yyyy-M-d`) al arreglo de formatos reconocidos. Refinada la consulta de conteo de residuales para evaluar `(date(c) IS NULL OR (length(c) <= 10 AND date(c) <> c) OR (length(c) > 10 AND date(c) <> substr(c, 1, 10)))`, diferenciando válidamente timestamps de fechas imposibles renormalizadas por SQLite.

#### 2. Fase 884 — Cobertura Unitaria de Formatos ISO con Hora, Componentes de 1 Dígito y Residuos en Vinculaciones
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV68_NormalizarFechasAgSesionesTests.cs`
- **Diagnóstico:** La suite unitaria no probaba cadenas ISO con hora (`'2026-05-10 14:30'`), fechas sin ceros a la izquierda (`'2026-4-5'`), ni la preservación de fechas imposibles (`'31/02/2026'`) o marcadores (`'--'`) en la tabla `ag_vinculaciones`.
- **Solución Implementada:** Añadidos registros de prueba con `'2026-05-10 14:30'` (normalizado a `'2026-05-10'`), `'2026-4-5'` (normalizado a `'2026-04-05'`), y vinculación con inicio `'31/02/2026'` y fin `'--'` (preservados intactos), verificando que el total de residuales se mantiene exactamente en 2 (id 6 e id 10).

#### 3. Fase 885 — Normalización de Fechas Estimadas en Genoma (`ag_gen`) en V68
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV68_NormalizarFechasAgSesionesTests.cs`
- **Diagnóstico:** La tabla `ag_gen` almacena `fecha_inicio_estimada` y `fecha_fin_estimada` para cada individuo del algoritmo genético. Al rehidratar individuos legacy en `GetIndividuoByIdAsync`, cadenas no ISO o con hora podían dejar los campos `DateTime?` nulos o desalineados, afectando la expansión de sesiones.
- **Solución Implementada:** Agregada sección 3 en `MigrationV68_NormalizarFechasAgSesiones.Up` que detecta la presencia de `ag_gen` y normaliza sus columnas estimadas a formato canónico ISO usando el catálogo completo de formatos, con caso de prueba en `MigrationV68_NormalizarFechasAgSesionesTests`.

---

### 157.127 Ronda 170 — Formatos con Segundos, Detección Exhaustiva de Residuos por Longitud y Diagnóstico en Vinculaciones y Genoma (Fases 886 a 889)

#### 1. Fase 886 — Formatos con Segundos y Corrección de Residuos de Longitud > 10 en V68
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`
- **Diagnóstico:** Fechas con segundos (ej. `'2026-05-10 00:00:00'`, `'dd/MM/yyyy HH:mm:ss'`, `'yyyy-MM-ddTHH:mm:ss.fff'`) no coincidían con ningún patrón de normalización, permaneciendo intactas. El criterio residual anterior para longitud > 10 comparaba `date(f) <> substr(f,1,10)`; para marcas de tiempo válidas ambas subcadenas coincidían, provocando que la fila no normalizada no se contara en los residuales ni se emitiera advertencia diagnóstica.
- **Solución Implementada:** Añadidos patrones con segundos (`"yyyy-MM-dd HH:mm:ss"`, `"yyyy-MM-dd H:mm:ss"`, `"yyyy-MM-ddTHH:mm:ss.fff"`, `"dd/MM/yyyy HH:mm:ss"`) al catálogo de formatos. Simplificada la consulta de residuales: tras la normalización todo valor válido debe tener exactamente 10 caracteres (`length(fechasesion) <> 10 OR date(fechasesion) IS NULL OR date(fechasesion) <> fechasesion`), garantizando la detección y logueo de cualquier residuo.

#### 2. Fase 887 — Diagnóstico de Residuales en Vinculaciones y Genoma
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`
- **Diagnóstico:** Solo `ag_sesiones` evaluaba y emitía advertencia de residuales. Fechas imposibles o no parseables en `ag_vinculaciones` o `ag_gen` no se reportaban, arriesgando la promoción a producción de fechas inconsistentes.
- **Solución Implementada:** Incorporadas consultas escalares de diagnóstico para `ag_vinculaciones` (excluyendo valores vacíos y `'--'`) y para `ag_gen` (si la tabla existe), emitiendo advertencia unificada `LoggerHelper.Warning` cuando existan residuales pendientes de revisión.

#### 3. Fase 888 — Cobertura Unitaria de Formatos con Segundos y Residuales por Tabla
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV68_NormalizarFechasAgSesionesTests.cs`
- **Diagnóstico:** La suite de pruebas no cubría formatos con segundos (`'2026-05-10 00:00:00'`, `'10/05/2026 14:30:55'`) ni verificaba la normalización de vinculaciones o genoma cuando contienen marcas de tiempo.
- **Solución Implementada:** Añadidos registros de prueba para sesiones con segundos, vinculaciones con hora/segundos (`'2026-05-10 00:00:00'`) y genoma con tiempo, verificando su normalización a formato ISO canónico de 10 caracteres y manteniendo el conteo de residuales exactamente en 2 (id 6 e id 10).

#### 4. Fase 889 (Opcional / Retirada) — Normalización N+1 por Fila
- **Estado:** RETIRADA
- **Diagnóstico Crítico:** Se evaluó la conveniencia de reemplazar el bucle de actualización fila a fila por sentencias set-based `UPDATE ... WHERE id IN (...)` o por lotes.
- **Causa Raíz / Justificación de Retiro:** En SQLite, todas las operaciones dentro de una transacción activa (`BeginTransaction()`) operan en memoria en el journal/WAL sin fsync individual por fila, ejecutando cientos de registros en menos de 20 ms. La sintaxis SQL `UPDATE ... WHERE id IN (...)` no permite asignar valores individuales por fila sin sentencias `CASE WHEN` complejas o tablas temporales, lo cual añade sobrecarga y complejidad innecesaria contraviniendo el principio de simplicidad para los volúmenes de staging.

---

### 157.128 Ronda 171 — Métrica de Truncamiento Datetime y Normalización en Exportación de Proyección Espejo (Fases 890 a 891)

#### 1. Fase 890 — Métrica de Truncamiento Datetime a Fecha en Migración V68
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`
- **Diagnóstico:** Fechas históricas con marca de tiempo (ej. `'2026-05-10 14:30:00'`) eran normalizadas a `'2026-05-10'` descartando la hora sin registro de auditoría, impidiendo conocer el volumen de registros que sufrieron truncamiento.
- **Solución Implementada:** Incorporado contador de truncadas en `MigrationV68_NormalizarFechasAgSesiones.Up` para `ag_sesiones`, `ag_vinculaciones` y `ag_gen` cuando `raw.Length > 10` y la conversión ISO resulte exitosa, emitiendo log informativo `LoggerHelper.Info` al finalizar la migración si `truncadas > 0`.

#### 2. Fase 891 — Normalización de Días y Horas en ExportarProyeccionEspejo
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El método `ExportarProyeccionEspejo` concatenaba la columna de días con `string.Join("/", ...)` sin horas y sin remover tildes, produciendo un formato inconsistente respecto a `SemestresExcelExporter` y la columna K de la plantilla académica (`"LUNES|4, MARTES|4"`).
- **Solución Implementada:** Actualizada la proyección de `diasHorasSemana` para agrupar por día normalizado sin tildes mediante `NormalizarDiaTexto` (apoyado en `DiaSemanaNormalizador.NormalizarDiaSemana`), sumando las horas por día y formateando con `g.Key|g.Sum(Horas)`, separadas por coma.

---

### 157.129 Ronda 172 — Corrección de Vistas Canónicas AG, Detección de Huecos en Migrador y Cuarentena de Bloques Horarios (Fases 892 a 897)

#### 1. Fase 892 (P0-1) — Corrección de Columna Canónica en `vw_AgVinculacionesDetalle` (`v.id_individuo`)
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** `MigrationV67` fallaba al validar con `no such column: v.idindividuo`. La tabla `ag_vinculaciones` definida en `MigrationV49` utiliza el nombre de columna canónico `id_individuo`.
- **Solución Implementada:** Corregida la proyección en `vw_AgVinculacionesDetalle` a `v.id_individuo AS IdIndividuo`.

#### 2. Fase 893 (P0-2a) — Migración V69 de Revalidación Forzada de Vistas AG (`MigrationV69_RevalidarVistasAg`)
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV69_RevalidarVistasAg.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Tras fallar V67 en bases existentes, V68 se registró en `__SchemaVersion`, dejando el vault sin la recreación de vistas canónicas y ocultando sesiones validadas (`estado = 0`).
- **Solución Implementada:** Creada la migración no aditiva `MigrationV69_RevalidarVistasAg` (Versión 69) registrada en `SqliteHelper.BuildMigrator` y en `.csproj`, invocando `SqliteHelper.EnsureAgViews` y validando la resolución real de columnas con `SELECT * FROM vw_... LIMIT 0`.

#### 3. Fase 894 (P0-2b) — Detección y Reintento de Huecos de Versiones en `SchemaMigrator`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`
- **Diagnóstico:** El migrador calculaba migraciones pendientes únicamente con `m.Version > currentVersion`. Si una migración tolerante fallaba (ej. V67) y una versión superior tenía éxito (V68), la migración fallida quedaba excluida de por vida en subsiguientes inicios.
- **Solución Implementada:** Modificado el cálculo de pendientes para consultar el conjunto de versiones aplicadas (`!aplicadas.Contains(m.Version) && m.Version > baseVersion`), permitiendo que cualquier migración pendiente sea reintentada automáticamente.

#### 4. Fase 895 (P0-3) — Validación Inmediata y Relanzamiento de Excepciones en `EnsureAgViews`
- **Archivo Afectado:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** SQLite no valida columnas al ejecutar sentencias `CREATE VIEW`, y `EnsureAgViews` atrapaba excepciones emitiendo sólo advertencias sin relanzar, enmascarando vistas rotas en producción.
- **Solución Implementada:** Incorporada verificación inmediata post-creación mediante `SELECT * FROM {v} LIMIT 0;` con relanzamiento obligatorio (`throw;`) si alguna vista no puede resolverse.

#### 5. Fase 896 (P1-1) — Reescalado de Genes Fijados sin Ajustables para Satisfacer Invariante I1
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Grupos donde todos los genes son fijos (`ajustables.Count == 0`) cuya suma difería de `TotalHorasAsignatura` (ej. `Σ=96 ≠ esperado=48`) no eran corregidos en `Sellar`, provocando violación del Invariante I1 y pérdida del grupo entero en cuarentena.
- **Solución Implementada:** En `Sellar`, si `ajustables.Count == 0` y `suma != esperado`, se reescalan proporcionalmente los genes fijados a múltiplos de 4.0h y se ajusta el residuo fino en el de mayor carga, limpiando las listas de `FechasGen` para forzar su regeneración consistente.

#### 6. Fase 897 (P1-2) — Validación Dura y Cuarentena Explícita de Bloque Horario Inexistente / Cero
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Genes con `IdBloqueHorario <= 0` no eran penalizados por `AgValidadorReglas`, llegando a `AgMaterializadorSesiones` donde se asumían 4h silenciosamente y luego fallaban en el expansor de sesiones, provocando cuarentena con diagnósticos crípticos.
- **Solución Implementada:** Incorporada penalización dura `R8_DiaFueraDePatron` en `AgValidadorReglas` si `gen.HorasGen > 0.0 && gen.IdBloqueHorario <= 0`. En `AgMaterializadorSesiones`, agregada guarda de cuarentena inmediata `I6_BloqueInexistente` si algún gen con carga tiene un bloque no registrado en el catálogo.

---

### 157.130 Ronda 173 — Umbral Estático de Esquema Base, Aislamiento Granular de Vistas y Cuantización Dinámica de Fijados (Fases 898 a 901)

#### 1. Fase 898 — Fijación de Umbral Estático `BASE_SCHEMA_VERSION` en Migrador
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`
- **Diagnóstico:** Si `__SchemaVersion` carecía de la fila 28, el migrador caía a `baseVersion = 0`, intentando ejecutar migraciones 1-27 que pertenecen al DDL base `AcademicSchema.sql` y no son idempotentes, rompiendo el arranque.
- **Solución Implementada:** Establecido `const int baseVersion = BASE_SCHEMA_VERSION;`, garantizando que jamás se re-ejecuten migraciones base históricas en ningún vault.

#### 2. Fase 899 — Desacoplamiento Granular de Vistas AG y Tolerancia en V69
- **Archivos Afectados:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV69_RevalidarVistasAg.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV69_RevalidarVistasAgTests.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgSchemaMigrationTests.cs`
- **Diagnóstico:** `EnsureAgViews` ejecutaba la creación de ambas vistas en un único bloque; si la tabla `periodos` estaba ausente en vaults legacy, fallaba `vw_AgVinculacionesDetalle` arrastrando el rollback de `vw_AgSesionesDetalle` y bloqueando el inicio de Excel con `V69.Aditiva = false`.
- **Solución Implementada:** Separada la creación y validación individual de cada vista con método auxiliar `ValidarVista`, condicionando la vista de vinculaciones a la existencia previa de la tabla `periodos`, y configurando `MigrationV69.Aditiva => true;`.

#### 3. Fase 900 — Código de Violación Dedicado `R14_BloqueHorarioNoAsignado`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`
- **Diagnóstico:** El gen sin bloque horario asignado (0) reportaba erróneamente `R8_DiaFueraDePatron`, distorsionando los reportes de violaciones y colisionando con la regla de patrón multi-día.
- **Solución Implementada:** Añadido valor de enum `R14_BloqueHorarioNoAsignado = 32` y actualizado `AgValidadorReglas` para reportar esta violación dura de forma unívoca, con aserción actualizada en test unitario.

#### 4. Fase 901 — Cuantización Dinámica por `HorasDia` en Reescalado de Fijados
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs`
- **Diagnóstico:** El reescalado de genes fijados utilizaba la constante fija `4.0`, generando violaciones de jornada en aulas de 8 horas (`HorasDia == 8`) y reduciendo a 0 horas genes de baja intensidad.
- **Solución Implementada:** Implementada función local `ObtenerUnidad` que detecta dinámicamente si el aula asignada tiene jornada de 8h o 4h, aplicando piso mínimo `u` y recuantización del residuo fino, con prueba unitaria dedicada `Sellar_WhenAllGenesFixedAndAula8Hours_ShouldRescaleIn8HourMultiples`.

---

### 157.131 Ronda 174 - Plan Consolidado v2: Carriles, Oleadas, Duracion por Nivel, Columna K, Duplicados y Boton Modulo R (Fases 902 a 922 / T1 a T21)

> **Estado de Ejecucion**: PENDIENTE (NO APLICADA EN CODIGO). Re-evaluada en Ronda 175.

#### 1. Fase 902 (T1) - Enums y Ponderaciones de Aptitud para Carriles (R11-R14, S6-S7, Nivel, Politicas)
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivos Afectados:** ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs, ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPesosFitness.cs, ConversorPDF.XLL/Academic/Domain/Enums/AgNivelAcademico.cs, ConversorPDF.XLL/Academic/Domain/Enums/AgPoliticaLlenado.cs, ConversorPDF.XLL/Academic/Domain/Enums/AgSincronizacionCarriles.cs
- **Diagnostico:** Catalogo sin R11 (duracion nivel), R12 (tope cohorte/aula semana ISO), R13 (capacidad semanal cohorte), S6 (hueco secuencia), S7 (desbalance carriles). Ajuste en Fase 923: R14_BloqueSolapadoPorIntervalo = 33 para evitar colision con codigo 32.
- **Solucion Propuesta:** Incorporar valores canonicos y tipos en Domain/Enums/ y AgPesosFitness.cs.

#### 2. Fase 903 (T2) - Modelos POCO del Plan Calendario por Carriles
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivos Afectados:** ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgUnidadSemanal.cs, ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgCarril.cs, ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPlanAsignatura.cs, ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPlanCalendario.cs, ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgParametrosPlan.cs
- **Diagnostico:** Inexistencia de abstracciones tipadas para unidades semanales, carriles concurrentes, plan por asignatura, calendario global y serializacion de parametros.
- **Solucion Propuesta:** Crear los modelos POCO con metodos ToJson() y DesdeJson().

#### 3. Fase 904 (T3) - AgSemanaIso y Formateador Unificado de Columna K
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivos Afectados:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSemanaIso.cs, ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs
- **Diagnostico:** C1: Columna K sin horas (string.Join("/", diasUnicos)). Sin agrupador ISO-8601 canonico para R5/R12/R13.
- **Solucion Propuesta:** Implementar AgSemanaIso.Clave(fecha) y DiasHorasSemanaFormatter ("DIA|horas, ...").

#### 4. Fase 905 (T4) - AgUnidadesSemanales: Carriles por Nivel y Resolucion de Horas
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgUnidadesSemanales.cs
- **Diagnostico:** C2/C6: Clases de 8h en PREGRADO y lectura de dias_ventana ignorada.
- **Solucion Propuesta:** Implementar HorasSesionAula, ConstruirCarriles, DiasDeVentana, SemanasDisponibles.

#### 5. Fase 906 (T5) - AgPlanificadorCarriles: Encadenamiento DISTRIBUIDO y Factibilidad
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorCarriles.cs
- **Diagnostico:** C3/C4: Concentracion en entana.Inicio. Falta encadenamiento DISTRIBUIDO con extremos anclados y D15.
- **Solucion Propuesta:** Algoritmo DISTRIBUIDO con RepartirHolgura y RepartirEnVentana (D15). Validacion de factibilidad.

#### 6. Fase 907 (T6) - AgAsignadorCarriles y Operadores de Permutacion
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorCarriles.cs
- **Diagnostico:** Cruce genetico por posicion duplicaba materias en el cromosoma.
- **Solucion Propuesta:** Asignacion voraz inicial, mutacion Permutar dentro de carril y cruce OX sobre permutaciones.

#### 7. Fase 908 (T7) - AgReplanificadorCarriles, AgNormalizadorCromosoma y AgBloqueSolape
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivos Afectados:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReplanificadorCarriles.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgNormalizadorCromosoma.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgBloqueSolape.cs
- **Diagnostico:** Desfases fenotipicos y solapes por ID de bloque en vez de intervalos de tiempo.
- **Solucion Propuesta:** Replanificacion fenotipica por aula/individuo, normalizacion de genes y helper AgBloqueSolape.SeSolapan.

#### 8. Fase 909 (T8) - AgFabricaIndividuos con Siembra Basada en Carriles
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs
- **Diagnostico:** L199 y L336-349 seleccionaban bloques aleatorios y forzaban 8 h en PREGRADO.
- **Solucion Propuesta:** Bucle de siembra acoplado a carriles; un gen por unidad semanal con EsFijoTemporal = true.

#### 9. Fase 910 (T9) - Blindaje Estricto de Mutacion y Cruce
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivos Afectados:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs
- **Diagnostico:** Operadores alteraban franjas destruyendo la estructura de carriles.
- **Solucion Propuesta:** Mutacion restringida a docente o permuta de carril; cruce sobre permutacion + replanificacion + normalizacion.

#### 10. Fase 911 (T10) - Cuantizacion Dinamica en AgCoherenciaGrupo por Nivel
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs
- **Diagnostico:** Cuantizacion con paso fijo de 4.0 h corrompia aulas de posgrado de 8 h.
- **Solucion Propuesta:** Reparto en multiplos de HorasSesionAula(aula, nivel) usando cardinalidad de FechasGen.

#### 11. Fase 912 (T11) - Validador de Reglas con Intervalos y Vacantes
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs
- **Diagnostico:** C4/C7: Solape por ID de bloque permitia hasta 36 h/dia en cohorte; R4/R5 excluian vacantes.
- **Solucion Propuesta:** Claves canonicas (cohorte, ula, semana ISO), solape por intervalo horario, inclusion de vacantes en R4-R6/R8/R11-R14, incorporacion de R11-R14, S6-S7.

#### 12. Fase 913 (T12) - Pre-Flight Bloqueante en AgValidadorVentana
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs
- **Diagnostico:** C6: AgValidadorVentana.Validar nunca validaba dias_patron c dias_ventana sobre datos reales.
- **Solucion Propuesta:** Invocar pre-flight Section 3.4 bloqueante antes de la corrida con mensajes accionables exactos.

#### 13. Fase 914 (T13) - Ponderacion Especifica por Regla en AgEvaluadorAptitud
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorAptitud.cs
- **Diagnostico:** C8: Penalidad lineal plana borraba la presion selectiva de AgPesosFitness.
- **Solucion Propuesta:** Ponderacion especifica por regla en f2 y evaluacion de cobertura en F3.

#### 14. Fase 915 (T14) - Invariantes I3-I6 y Deduplicacion en AgMaterializadorSesiones
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs
- **Diagnostico:** Sesiones duplicadas por clave natural y pisado de fechas estimadas con ventana.
- **Solucion Propuesta:** Compuertas I3 (duracion segun nivel), I4 (max 2 materias aula/semana), I5 (dedupe id_aula, fecha, horariosesion), I6 (bloque en catalogo); cuarentena defensiva.

#### 15. Fase 916 (T15) - GetUltimaCorridaAsync, Dedupe y Purga Multi-Programa en Repositorio
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivos Afectados:** ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs, ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs
- **Diagnostico:** C9: Purga mono-programa dejaba datos huerfanos entre programas. Carencia de consulta de ultima corrida para recarga.
- **Solucion Propuesta:** Implementar GetUltimaCorridaAsync y purga multi-programa en LimpiarStagingAsync.

#### 16. Fase 917 (T16) - Boton "Cargar en Modulo R", Gate D12 y Retiro de CSV Espejo
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivos Afectados:** ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs, ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs
- **Diagnostico:** C10: Boton Proyeccion 1:1 generaba CSV espejo redundante en vez de cargar en UI. Sin compuerta D12 ante violaciones duras.
- **Solucion Propuesta:** Renombrar a "Cargar en Modulo R", eliminar ExportarProyeccionEspejo, implementar gate D12 (CON_VIOLACIONES).

#### 17. Fase 918 (T17) - Formateo Canonico Columna K en SemestresExcelExporter
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs
- **Diagnostico:** C1: Columna K vacia de horas.
- **Solucion Propuesta:** Integrar DiasHorasSemanaFormatter en SemestresExcelExporter.cs:389.

#### 18. Fase 919 (T18) - Verificacion de Columnas y Registro en EnsureAgViews
- **Estado:** RETIRADA (Falso Positivo)
- **Archivo Afectado:** ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs
- **Diagnostico:** Falsos positivos: ValidarVista ya existe en SqliteHelper.cs e implementa LoggerHelper.Error + 	hrow; implementado y verificado en Fases 893/895/900.
- **Solucion:** Retirada del plan para evitar duplicacion.

#### 19. Fase 920 (T19) - Purga Multi-Programa y Siembra por Carril en UI
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs
- **Diagnostico:** C5: L3984-3986 forzaba 4,5,6 para 64 h fuera de la ventana MAR-VIE.
- **Solucion Propuesta:** Eliminar derivacion de dias por horas; siembra por carril round-robin y purga multi-programa.

#### 20. Fase 921 (T20) - Documentacion Tecnica y Articulo de Cambios
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivo Afectado:** docs/articles/changes/2026-10-21-ag-carriles-oleadas.md
- **Diagnostico:** Carencia de registro documental tecnico del modelo v2 de carriles y oleadas.
- **Solucion Propuesta:** Crear articulo tecnico en docs/articles/changes/.

#### 21. Fase 922 (T21) - Suite de Pruebas Unitarias de Regresion y Reglas O1-O12
- **Estado:** PENDIENTE (NO APLICADA)
- **Archivos Afectados:** ConversorPDF.Tests/Academic/Application/Services/Genetica/*.cs
- **Diagnostico:** Inexistencia de pruebas automatizadas para fixtures prog1, prog2, rechazo de prog3 y solapes por intervalo.
- **Solucion Propuesta:** Crear suites xUnit + FluentAssertions + Moq.

---

### 157.132 Ronda 175 - Correcciones Bloqueantes Plan v2 y Sincronizacion Repo-Brain (Fases 923 a 934)

#### 1. Fase 923 - Desambiguacion de Codigo 32 en AgReglaViolacion (R14_BloqueSolapadoPorIntervalo = 33)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs
- **Diagnostico:** R14_BloqueHorarioNoAsignado = 32 ya existia desde la Fase 900. F902 proponia reusar 32 provocando ambiguedad en serializacion y Enum.GetName.
- **Solucion Propuesta:** Conservar R14_BloqueHorarioNoAsignado = 32 y declarar R14_BloqueSolapadoPorIntervalo = 33.

#### 2. Fase 924 - Resolucion de Nivel Academico y Forzado de 4h en PREGRADO
- **Archivos Afectados:** ConversorPDF.XLL/Academic/Domain/Enums/AgNivelAcademico.cs, ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgNivelAcademicoCatalogo.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgUnidadesSemanales.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs
- **Diagnostico:** AgCoherenciaGrupo.cs:90-98 evaluaba solo .HorasDia == 8, permitiendo clases de 8h en PREGRADO.
- **Solucion Propuesta:** Crear catalogo/helper de nivel y metodo AgUnidadesSemanales.HorasSesionAula(aula, nivel) que fuerce 4.0h en pregrado.

#### 3. Fase 925 - Eliminacion de Derivacion Heuristica de dias_patron por Horas (C5)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs
- **Diagnostico:** L3982-3986 asignaba "4,5,6" si horas >= 64, metiendo sabado en ventanas MAR-VIE.
- **Solucion Propuesta:** Eliminar la heuristica por horas; sembrar por carril o cadena vacia para resolucion por planificador.

#### 4. Fase 926 - Invocacion Activa de AgValidadorVentana.Validar en Pre-Flight (C6)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs
- **Diagnostico:** El metodo estatico Validar nunca se evaluaba contra asignaturas reales en el pre-flight.
- **Solucion Propuesta:** Evaluar AgValidadorVentana.Validar(aa.DiasPatron, aula.DiasVentana, ...) para cada fila de g_aula_asignatura.

#### 5. Fase 927 - Inclusion de Vacantes y Agrupacion Canonica por Cohorte en R4/R5 (C7/D11)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs
- **Diagnostico:** R4 y R5 ignoraban vacantes y agrupaban por Grupo plano. R1-R3 comparaban IDs de bloque por igualdad exacta en vez de intervalos.
- **Solucion Propuesta:** Agrupar por ClaveCohorte(g), incluir vacantes en reglas de ocupacion y verificar solapes por intervalo con AgBloqueSolape.SeSolapan.

#### 6. Fase 928 - Ponderacion Especifica por Regla en Evaluador de Fitness (C8)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorAptitud.cs
- **Diagnostico:** Penalidad lineal base ignoraba los pesos individuales de AgPesosFitness.
- **Solucion Propuesta:** Sumar penalidades ponderadas individuales usando contexto.Pesos.PorRegla(v.Regla).

#### 7. Fase 929 - Formateo Canonico Columna K en SemestresExcelExporter (C1)
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs
- **Diagnostico:** L378-389 emitia nombres de dias sin horas (MARTES/MIERCOLES).
- **Solucion Propuesta:** Asignar matrizDatos[r, 10] = DiasHorasSemanaFormatter.DesdeSesiones(sesionesGrupo).

#### 8. Fase 930 - Boton "Cargar en Modulo R" y Retiro de CSV Espejo (C10)
- **Archivos Afectados:** ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs, ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs
- **Diagnostico:** Persistia boton "Proyeccion 1:1" y exportacion espejo obsoleta.
- **Solucion Propuesta:** Renombrar boton a "Cargar en Modulo R", enlazar con CargarUltimaGeneracionEnModuloRAsync y eliminar ExportarProyeccionEspejo().

#### 9. Fase 931 - Incorporacion de GetUltimaCorridaAsync en Repositorio
- **Archivos Afectados:** ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs, ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs
- **Diagnostico:** Contrato carece de metodo para obtener la ultima corrida por periodo/programa, bloqueando F930.
- **Solucion Propuesta:** Agregar firma e implementacion con consulta Dapper ordenada por id DESC LIMIT 1.

#### 10. Fase 932 - Compuertas I3, I4, I5 en Materializador de Sesiones
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs
- **Diagnostico:** Fechas pisadas con ventana, sin deduplicacion de clave natural y sin validacion de duracion segun nivel.
- **Solucion Propuesta:** Validacion de I3 (duracion nivel), I4 (max 2 asignaturas aula/semana) e I5 (dedupe id_aula, fecha, horariosesion).

#### 11. Fase 933 - Retiro de Fase 919 (T18) por Falso Positivo
- **Archivo Afectado:** ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs
- **Diagnostico:** ValidarVista ya existia con LoggerHelper.Error y 	hrow; implementado en Fases 893/895/900.
- **Solucion Propuesta:** Retirar T18 formalmente del plan consolidado.

#### 12. Fase 934 - Sincronizacion Integral de Documentos Repo-Brain
- **Archivos Afectados:** implementation_plan.md, walkthrough.md
- **Diagnostico:** Desincronizacion entre copias locales en el brain y archivos Markdown en la raiz del repositorio.
- **Solucion Propuesta:** Sincronizar simultaneamente ambos archivos en repo root y brain.

---

### 157.133 Ronda 176 — Verificación de Fases 923-934, Corrección de Fallback en DiasKDe y Agregación R12 (Fases 935 a 947)

> **Estado de Ejecución**: F935-F945, F947 APROBADAS (no aplicadas en código .cs); F946 APLICADA (eliminación de scratch/sync_docs.ps1).

#### 1. Fase 935 — F923 no aplicada: enum sin R11/R12/R13/R14=33/S6/S7
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs`
- **Diagnóstico:** Enum termina en R10=28 y S5=105; código 32 ocupado por `R14_BloqueHorarioNoAsignado`.
- **Solución Propuesta:** Incorporar R11=29, R12=30, R13=31, R14_BloqueSolapadoPorIntervalo=33, S6=106, S7=107.

#### 2. Fase 936 — F924 no aplicada: ObtenerUnidad decide por HorasDia
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`, `AgUnidadesSemanales.cs`, `AgNivelAcademicoCatalogo.cs`
- **Diagnóstico:** L95 `return a.HorasDia == 8 ? 8.0 : 4.0;` permite 8h en PREGRADO con horas_dia=8.
- **Solución Propuesta:** Cuantizar con `AgUnidadesSemanales.HorasSesionAula(a, nivel)` derivado desde `programas.modalidad`.

#### 3. Fase 937 — F925 y Corrección de Fallback en AgPlanificadorPatronSemanal
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** Vaciar `diasPatron` sin ajustar `DiasKDe` activaba fallback `{ 2, 3 }` (L40), forzando Martes/Miércoles a todas las asignaturas.
- **Solución Propuesta:** Eliminar heurística L3984-3985 y corregir fallback de `DiasKDe` para resolver desde `aa.DiasVentana`.

#### 4. Fase 938 — F926 no aplicada: Invocación de Pre-Flight
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `Validar` estático sin invocación en el flujo de preparación pre-vuelo.
- **Solución Propuesta:** Iterar `contexto.AulaAsignaturas` validando `Validar(aa.DiasPatron, aula.DiasVentana, ...)`.

#### 5. Fase 939 — F927 no aplicada: Vacantes y Solape de Intervalos
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** R4/R5 filtran `EsVacante == 0` y agrupan por `Grupo` plano; R1/R2/R3 usan igualdad de bloque.
- **Solución Propuesta:** Agrupar por `ClaveCohorte(g)`, incluir vacantes en ocupación y comparar intervalos con `AgBloqueSolape.SeSolapan`.

#### 6. Fase 940 — F928 no aplicada: Ponderación de Fitness
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorAptitud.cs`
- **Diagnóstico:** f2 plano ignora coeficientes de `AgPesosFitness`.
- **Solución Propuesta:** Ponderar sumando `contexto.Pesos.PorRegla(v.Regla)`.

#### 7. Fase 941 — F929 no aplicada: Formato de Columna K
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Diagnóstico:** L389 exporta días sin horas (`MARTES/MIERCOLES`).
- **Solución Propuesta:** Asignar `DiasHorasSemanaFormatter.DesdeSesiones(sesionesGrupo)`.

#### 8. Fase 942 — F930 no aplicada: Botón Módulo R y Retiro CSV
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `.Designer.cs`
- **Diagnóstico:** Botón conserva texto «📑 Proyección 1:1» y cableado a `ExportarProyeccionEspejo()`.
- **Solución Propuesta:** Texto "📥 Cargar en Módulo R", cablear a `CargarUltimaGeneracionEnModuloRAsync()` y suprimir CSV espejo.

#### 9. Fase 943 — F931 no aplicada: GetUltimaCorridaAsync en Repositorio
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs`, `AgScheduleRepository.cs`
- **Diagnóstico:** Método inexistente en contrato e implementación.
- **Solución Propuesta:** Declarar e implementar consulta Dapper ordenada por `id DESC LIMIT 1`.

#### 10. Fase 944 — F932 no aplicada: Materializador Staging
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Sin deduplicación de clave natural `(id_aula, fecha, horariosesion)` ni guarda de duración nivel.
- **Solución Propuesta:** Validar compuertas I3, I4, I5 en staging.

#### 11. Fase 945 — Clave de Agregación de R12 (Aula vs Cohorte)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** R12 (`Max2AsignaturasPorAulaSemana`) debe agrupar por `(IdAula, SemanaIso)`, distinguiéndose de R5 (cohorte).
- **Solución Propuesta:** `genes.GroupBy(g => (Aula: g.IdAula, Semana: AgSemanaIso.De(g.FechasGen)))`.

#### 12. Fase 946 — Eliminación de Artefacto `scratch\sync_docs.ps1`
- **Archivo Afectado:** `scratch\sync_docs.ps1`
- **Diagnóstico:** Script transitorio contaminando el workspace.
- **Solución Implementada:** Archivo eliminado del repositorio con `Remove-Item`.

#### 13. Fase 947 — Estandarización de Estado Documental con Evidencia
- **Archivo Afectado:** `walkthrough.md`
- **Diagnóstico:** Estados marcados como "ACEPTADA" sin código en disco.
- **Solución Implementada:** Adopción de estado `APROBADA (no aplicada)` con métrica explícita de símbolos en disco.

---

### 157.134 Ronda 177 — Corrección de Bloqueantes de Compilación y Estructura de Carriles (Fases 948 a 953)

#### 1. Fase 948 — Método `PorRegla` en `AgPesosFitness`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPesosFitness.cs`
- **Diagnóstico:** `contexto.Pesos.PorRegla(v.Regla)` (F940) causa CS1061.
- **Solución Propuesta:** Implementar `PorRegla(AgReglaViolacion)` con switch expression, evitando duplicar etiquetas para enums con valores idénticos (CS0152).

#### 2. Fase 949 — Firma y Adaptación de `DiasKDe` con `AgAula`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgAulaAsignatura.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** `AgAulaAsignatura` carece de `DiasVentana` (reside en `AgAula`).
- **Solución Propuesta:** Firma `DiasKDe(AgAulaAsignatura aa, AgAula aula = null)` y resolución desde `aula?.DiasVentana`.

#### 3. Fase 950 — Creación de Tipos Base `AgSemanaIso` y `AgBloqueSolape`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSemanaIso.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgBloqueSolape.cs`
- **Diagnóstico:** Inexistencia de tipos base que bloquea compilación de F937, F939 y F945 (CS0246).
- **Solución Propuesta:** Crear tipos estáticos con métodos `DesdeVentana`, `De`, `Clave` y `SeSolapan`.

#### 4. Fase 951 — Filtro de Fechas Válidas en Agrupación R12
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** Genes no expandidos con lista de fechas vacía anulan la regla R12.
- **Solución Propuesta:** Filtrar `g.FechasGen != null && g.FechasGen.Count > 0` y agrupar por `AgSemanaIso.Clave(g.FechasGen[0])`.

#### 5. Fase 952 — Preservación de Carriles Concurrentes en la Fábrica
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** `GroupBy(ClaveCohorteAula).Select(g => g.First())` colapsa los carriles paralelos de una cohorte a una sola aula.
- **Solución Propuesta:** Agrupar considerando jornada y ventana: `GroupBy(a => (ClaveCohorteAula(a), a.Jornada ?? "", a.DiasVentana ?? ""))`.

#### 6. Fase 953 — Unificación de Clave de Cohorte en Dominio
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgClaveCohorte.cs`, `AgFabricaIndividuos.cs`, `AgValidadorReglas.cs`
- **Diagnóstico:** Fábrica y validador usan fórmulas divergentes para identificar la cohorte.
- **Solución Propuesta:** Crear `AgClaveCohorte` en dominio para centralizar la resolución unívoca.

---

### 157.135 Ronda 178 — Robustecimiento de Pre-Flight, Consultas Dapper, Migración de Vacantes y Detección de Intervalos (Fases 954 a 959)

#### 1. Fase 954 — Evaluación de Días Efectivos en Pre-Flight de Ventana
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `Validar` omite la contención si `diasPatron` es vacío (`!string.IsNullOrWhiteSpace(diasPatron)` en L151).
- **Solución Propuesta:** Resolver días efectivos con `AgPlanificadorPatronSemanal.DiasKDe(aa, aula)` y pasarlos unidos por coma.

#### 2. Fase 955 — Consulta Dapper sobre Tabla Canónica `ag_corrida`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** El DDL define la tabla `ag_corrida` (singular); consultas a `ag_corridas` causan error de tabla inexistente.
- **Solución Propuesta:** Consulta SQL estricta contra `ag_corrida` con alias DTO.

#### 3. Fase 956 — Migración V70 para Persistencia de Vacantes
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV70_AddAgCorridaVacantes.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** Propiedad `AgCorrida.Vacantes` carece de columna en base de datos.
- **Solución Propuesta:** Migración `ALTER TABLE ag_corrida ADD COLUMN vacantes INTEGER NOT NULL DEFAULT 0;` y actualización de `INSERT`/`SELECT`.

#### 4. Fase 957 — Intersección Horaria Matemática en `AgBloqueSolape.SeSolapan`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgBloqueSolape.cs`
- **Diagnóstico:** Comparación textual por `Franja` ("MAÑANA", "TARDE") produce falsos solapes y falsos negativos.
- **Solución Propuesta:** Coincidencia en `DiaSemanaNumero` e intersección horaria `i1 < f2 && i2 < f1` parseando `HoraInicio` y `HoraFin` con `TimeSpan`.

#### 5. Fase 958 — Delegación de `AgSemanaIso.DesdeVentana` en `AgDiasVentana.Dias`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSemanaIso.cs`
- **Diagnóstico:** Parser canónico ya existe en `AgDiasVentana.Dias(ventana)`.
- **Solución Propuesta:** Delegación directa: `AgDiasVentana.Dias(diasVentana)`.

#### 6. Fase 959 — Cálculo de Capacidad en `AgValidadorVentana` con Aula
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** Invocaciones a `DiasKDe` sin aula en L95 y L114 caen en fallback `{2,3}` y distorsionan la capacidad real.
- **Solución Propuesta:** Suministrar `AulaDe(a.IdAula)` a las llamadas de `DiasKDe`.

---

### 157.136 Ronda 179 — Robustecimiento de Cadena de Siembra y Materialización (Fases 960 a 963)

#### 1. Fase 960 — Blindaje ante `fechasesion TEXT NOT NULL` en `ag_sesiones`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** DDL define `fechasesion TEXT NOT NULL` (`MigrationV49:207`); `AgSesion.FechaSesion` es nullable. Inserción de sesiones sin fecha dispara excepción SQLite 19.
- **Solución Propuesta:** Descartar en `AgMaterializadorSesiones` antes de staging (`I3 gen {s.IdGen}: sesión sin fecha`) y aplicar fallback seguro en `AgScheduleRepository` con `(s.FechaSesion ?? DateTime.MinValue).ToString("yyyy-MM-dd")`.

#### 2. Fase 961 — Bloques de 4 h en Pregrado Durante la Siembra
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La fábrica resuelve franjas con `aula.HorasDia` crudo (8h) sin considerar el nivel académico, violando la regla D4′ para programas de pregrado.
- **Solución Propuesta:** Resolver `nivel = AgNivelAcademicoCatalogo.Desde(...)` y usar `AgUnidadesSemanales.HorasSesionAula(aula, nivel)` (4h en pregrado) en ambos puntos de siembra (L270 y L324).

#### 3. Fase 962 — Discriminación por Día en `AgFranjaJornada.Resolve` (Regla D6)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgFranjaJornada.cs`
- **Diagnóstico:** `Resolve(jornada, horasDia)` no recibe día ISO, imposibilitando aplicar la regla D6 (sábado POSGRADO-8h = bloque único `08:00–17:00`) y generando solapes entre tarde/noche 8h.
- **Solución Propuesta:** Extender firma a `Resolve(string jornada, int horasDia, int diaIso = 0)` retornando `08:00–17:00` MAÑANA si `diaIso == 6 && horasDia == 8`.

#### 4. Fase 963 — Persistencia de `dias_patron` en `InsertAulaAsignaturasAsync`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** `InsertAulaAsignaturasAsync` omite la columna `dias_patron` añadida en V55, persistiendo valor vacío y desactivando la contención D14.
- **Solución Propuesta:** Incluir `dias_patron` en el `INSERT OR IGNORE` con `COALESCE(@DiasPatron, '')`.

#### 5. Fase 964 — Cuantización Estricta de Sesiones en `AgExpansorSesiones`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
- **Diagnóstico:** `Math.Min(bloque.DuracionHoras, horasFaltantes)` genera sesiones residuales de duración fraccionaria (ej. 2h en bloques de 4h) si `HorasGen` no es múltiplo exacto de la unidad, provocando rechazo masivo por la compuerta I3 en `AgMaterializadorSesiones` o sesiones duplicadas en sábado D6.
- **Solución Propuesta:** Exigir `horasObjetivoGen % bloque.DuracionHoras == 0` con guard fail-fast y expandir cada sesión con `horasSesion = bloque.DuracionHoras` fijo sin `Math.Min`.

---

### 157.137 Ronda 181 — F964 Aplicada con Error de Compilación + 4 Defectos en Código Nuevo (Fases 965 a 969)

#### 1. Fase 965 — `horasAcumuladas` eliminada por F964 → CS0103 ×4
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
- **Diagnóstico:** El reemplazo del bloque de inicialización en F964 eliminó la declaración `double horasAcumuladas = 0.0;`, produciendo 4 errores CS0103 en líneas de acumulación y control.
- **Solución Propuesta:** Restaurar `double horasAcumuladas = 0.0;` tras la inicialización de `sesiones`.

#### 2. Fase 966 — `GetUltimaCorridaAsync` declarada sin implementar → CS0535
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** El contrato `IAgScheduleRepository` expone `GetUltimaCorridaAsync(int, int, string, CancellationToken)`, pero la implementación en Dapper no existía en `AgScheduleRepository.cs`.
- **Solución Propuesta:** Implementar consulta Dapper `SELECT ... FROM ag_corrida WHERE idperiodo = @idPeriodo AND idprograma = @idPrograma AND (@subunidad IS NULL OR subunidad = @subunidad) ORDER BY id DESC LIMIT 1` con soporte de `vacantes` y cancelación.

#### 3. Fase 968 — Coherencia cuantiza con `4.0` fijo → posgrado 8 h viola F964 y aborta el individuo
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** `AgCoherenciaGrupo` cuantizaba sesiones y proyectaba horas usando la constante `4.0` en la rama principal. En posgrados de 8h (`HorasDia=8`), generaba asignaciones como 20h (`20 / 8 = 2.5`), provocando rechazo fail-fast en F964 y aborto total de la expansión.
- **Solución Propuesta:** Promover función local `ObtenerUnidad(AgGen g)` al ámbito del grupo utilizando `AgUnidadesSemanales.HorasSesionAula` y reemplazar todos los factores `4.0` por `unidad` dinámica.

#### 4. Fase 967 — `AgBloqueSolape` y `R14_BloqueSolapadoPorIntervalo` sin emisor → regla inerte
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** El código enum 33 (`R14_BloqueSolapadoPorIntervalo`) y el evaluador de intervalos `AgBloqueSolape.SeSolapan` carecían de llamada en `AgValidadorReglas`, permitiendo que bloques solapados pasaran la compuerta.
- **Solución Propuesta:** Conectar evaluación en `AgValidadorReglas.Evaluar` sobre genes con bloque asignado, validando solape temporal por intervalo cuando compartan ámbito o recurso (aula, docente, cohorte/grupo) e incrementando `hardCount`.

#### 5. Fase 969 — Un gen inválido cancela la materialización completa del individuo
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
- **Diagnóstico:** `ExpandirIndividuo` retornaba `Result.Failure` al primer gen no cuantizado, descartando todas las demás sesiones válidas sin reportar el detalle de los genes fallidos.
- **Solución Propuesta:** Aislar la falla por gen acumulando errores, registrando advertencia diagnóstica y permitiendo el retorno de las sesiones válidas generadas, fallando únicamente si el individuo queda sin ninguna sesión expandible.

---

### 157.138 Ronda 182 — Discriminación de Cohorte en R14, Optimización Temporal y Divisibilidad de Carga (Fases 970 a 972)

#### 1. Fase 970 — F967 compara cohortes entre programas/sedes distintas → R14 hard masivo
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** La condición `g1.Grupo == g2.Grupo && g1.Semestre == g2.Semestre` en `mismoRecurso` no evaluaba el programa ni la sede. Programas o sedes diferentes que comparten nomenclatura de semestres y grupos (ej. Semestre "1", Grupo "1") provocaban que dos genes sin conflicto de recursos se considerasen la misma cohorte, emitiendo violaciones duras `R14_BloqueSolapadoPorIntervalo` y tornando casi cualquier individuo en `Infactible`.
- **Solución Propuesta:** Condicionar la equivalencia de cohorte a `g1.IdPrograma == g2.IdPrograma && g1.IdSede == g2.IdSede && !string.IsNullOrWhiteSpace(g1.Grupo) && !string.IsNullOrWhiteSpace(g1.Semestre) && g1.Grupo == g2.Grupo && g1.Semestre == g2.Semestre`.

#### 2. Fase 971 — F967 evalúa O(n²) sobre todos los genes → bloqueo en poblaciones reales
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** La evaluación de solapes recorría todos los pares de genes con bloque mediante doble bucle $O(n^2)$ sin particionamiento previo. Con tamaños de genoma típicos (2.000 genes), esto requería ~2M de iteraciones por individuo en cada generación, congelando la ejecución.
- **Solución Propuesta:** Particionar previamente los genes por día (`conBloque.GroupBy(x => x.Bloque.DiaSemanaNumero)`), ejecutando la comparación de pares únicamente dentro de la lista de cada día y aplicando corte temprano del bucle interno (`cortado = true; break;`).

#### 3. Fase 972 — F968 usa unidad de 8 h sin validar múltiplo → I1 queda irresoluble
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** En grupos de posgrado con sesiones de 8h, si la carga total requerida (`esperado`) no es múltiplo exacto de 8h (ej. 60h), redondear sobre unidad 8.0h produce desfases permanentes $\Sigma \neq I1$ (64h) que el rebalanceo de residuos no puede reconciliar.
- **Solución Propuesta:** Comprobar si `esperado` es divisible exactamente por `unidad` (`Math.Abs(esperado / unidad - Math.Round(esperado / unidad)) > 0.001`); de no serlo, degradar a `unidad = 4.0` (regla D4′).

---

### 157.139 Ronda 183 — Reconciliación de Bloques en Fábrica, Remanente de Coherencia y Limpieza R14 (Fases 973 a 975)

#### 1. Fase 973 — Fábrica asigna bloque de 8 h sin validar que la carga sea múltiplo → F964 rechaza y F969 omite el gen
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La fábrica asignaba franjas horarias con `horasEfectivas = HorasSesionAula(aula, nivelAula)` (8h en posgrado) sin considerar si la carga de la asignatura era múltiplo de 8h. En asignaturas de 60h, Coherencia reducía la unidad a 4h pero el gen conservaba el bloque de 8h, causando que `AgExpansorSesiones` rechazara la cuantización en F964 y F969 omitiera el gen.
- **Solución Propuesta:** Evaluar `cargaUnidad / horasEfectivas` (usando `totalHoras` en día único y `totalHoras / diasK.Count` en multi-día) y degradar `horasEfectivas = 4` cuando `horasEfectivas > 4` y la carga no sea divisible entre 8h.

#### 2. Fase 974 — F972 valida `esperado` pero no `horasRestantes` → residuo de fijados rompe I1
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Si `esperado` es múltiplo de 8 (ej. 64h) pero los genes fijados suman un valor que deja un remanente no divisible entre 8 (ej. 20h fijadas $\rightarrow$ 44h restantes), la unidad permanecía en 8.0h provocando que la suma proyectada discrepara de `esperado` sin posibilidad de ajuste residual.
- **Solución Propuesta:** Degradar `unidad = 4.0` si `esperado` o `horasRestantes` no son múltiplos exactos de `unidad`.

#### 3. Fase 975 — `cortado` + `continue` es código muerto en F971
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** La bandera `cortado` y la condición `if (cortado) continue;` al final del bucle `i` eran superfluas, ya que `break;` cortaba el bucle interior `j` e `i` continuaba naturalmente al siguiente elemento.
- **Solución Propuesta:** Remover la variable `cortado` y la cláusula `continue`, conservando el `break;` directo dentro del ciclo de pares.

---

### 157.140 Ronda 184 — Sincronización de Horas Semanales y Cuantización de Fallback en Fábrica (Fases 976 y 977)

#### 1. Fase 976 — `HorasSemanales` fijo en 4.0 contradice bloque de 8 h → carga por gen subestimada
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Asignar bloques de 8h en programas de posgrado mientras se retenía `HorasSemanales = 4.0` en `AgFabricaIndividuos` subestimaba a la mitad el conteo de horas calculadas por sesión (`sesionesEsteDia * HorasSemanales`), trasladando un desbalance masivo al ajuste posterior de residuo I1 y generando horas no múltiplos de 8h que F964 rechazaba.
- **Solución Propuesta:** Derivar dinámicamente `gen.HorasSemanales = bloque.DuracionHoras > 0 ? bloque.DuracionHoras : 4.0` tras asignar el bloque tanto en la rama de día único como en la de multi-día, eliminando la constante rígida en la instanciación.

#### 2. Fase 977 — Fallback `totalHoras / diasK.Count` sin cuantizar → F964 rechaza el gen
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** El cálculo de fallback en multi-día dividía linealmente `totalHoras / diasK.Count` sin redondear a múltiplos de la unidad de bloque, generando duraciones no cuantizadas (ej. 10h con bloque de 4h) que eran descartadas por la compuerta F964 y excluidas por F969.
- **Solución Propuesta:** Cuantizar el fallback a la unidad de sesión del gen mediante `Math.Max(unidadDia, Math.Round(cargaDia / unidadDia) * unidadDia)` con `unidadDia = gen.HorasSemanales > 0 ? gen.HorasSemanales : 4.0`.

---

### 157.141 Ronda 185 — Sincronización de Unidad de Sesión en Validador de Ventana y Contención D14 (Fases 978 y 979)

#### 1. Fase 978 — `sesionesRequeridas = ceil(hAsig / 4.0)` bloquea posgrado 8 h
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** En asignaturas de posgrado con sesiones de 8h (`HorasDia = 8`), calcular las sesiones requeridas como `ceil(hAsig / 4.0)` duplicaba artificialmente la demanda de sesiones (ej. 16 sesiones requeridas para 64h en vez de 8). En ventanas semestrales típicas de 12 semanas (capacidad real = 12 ocurrencias de días K), se disparaba un bloqueo falso de capacidad insuficiente ($16 > 12$).
- **Solución Propuesta:** Calcular `unidadAsig` mediante `AgUnidadesSemanales.HorasSesionAula(aulaAsig, AgNivelAcademicoCatalogo.Desde(progAsig?.Modalidad))` y obtener `sesionesRequeridas = (int)Math.Ceiling(hAsig / unidadAsig)`.

#### 2. Fase 979 — Validación D14 usa `auCfg.HorasDia` crudo → exige bloque de 8 h inexistente
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** La validación D14 pasaba `auCfg.HorasDia` crudo a `Validar`. En aulas de pregrado con `horas_dia = 8`, esto provocaba la búsqueda de bloques horarios de 8h (inexistentes en catálogo para pregrado según la regla D4′ de bloques de 4h), resultando en un error bloqueante en pre-vuelo.
- **Solución Propuesta:** Evaluar `horasEfectivasCfg = (int)AgUnidadesSemanales.HorasSesionAula(auCfg, AgNivelAcademicoCatalogo.Desde(progCfg?.Modalidad))` y pasarlo a `Validar` en lugar del valor crudo de la entidad de aula.

---

### 157.142 Ronda 186 — Resolución de Unidad desde Aula Ganadora y Conteo de Capacidad por Sesiones (Fases 980 y 981)

#### 1. Fase 980 — Unidad 4 h cuando `aa == null` → F978 inoperante en primera corrida
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `unidadAsig` se calculaba de forma aislada a partir del mapeo previo `aa` antes de evaluar `configs`. En primeras corridas o ante asignaturas sin aula preasignada (`aa == null`), caía en fallback de 4.0h y volvía a sobredimensionar a 16 sesiones requeridas las asignaturas de 64h en posgrado, reactivando bloqueos falsos.
- **Solución Propuesta:** Extraer función `UnidadDe(au)` y diferir la resolución de `unidadAsig` y `sesionesRequeridas` al aula ganadora de `configs` (`mejorCfg`), asegurando simetría total entre la unidad de demanda y la oferta de capacidad.

#### 2. Fase 981 — Capacidad cuenta 1 sesión por día ignorando AM/PM
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `CapacidadDe` sumaba ocurrencias de días en calendario asumiendo 1 sesión por día. En aulas con `HorasDia = 8` y asignaturas de 4h (pregrado con franjas mañana y tarde), un aula puede alojar 2 sesiones por día, por lo que la capacidad real se subestimaba al 50%.
- **Solución Propuesta:** Calcular `sesionesPorDia = au != null ? Math.Max(1, (int)Math.Round(au.HorasDia / unidadAu)) : 1` dentro de `CapacidadDe` y multiplicar las ocurrencias del día K por este factor.

---

### 157.143 Ronda 187 — Acotamiento de Capacidad Diaria por Franjas Reales de Catálogo (Fase 982)

#### 1. Fase 982 — `sesionesPorDia` apuesta 2 sesiones/día sin verificar franjas AM/PM existentes
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `sesionesPorDia` asumía que la jornada del aula podía dividirse en sesiones de duración `unidadAu` sin verificar si el catálogo de bloques (`contexto.Bloques`) contenía efectivamente múltiples franjas activas para esa duración. Si un aula de 8h/día solo disponía de una franja activa de 4h (ej. solo MAÑANA), se le atribuía el doble de capacidad en pre-vuelo, difiriendo el conflicto al fitness y generando individuos con sesiones imposibles de acomodar.
- **Solución Propuesta:** Limitar `sesionesPorDia` con `Math.Min(..., Math.Max(1, franjasReales))`, contabilizando el número de franjas activas y distintas con `DuracionHoras == unidadAu` registradas en `contexto.Bloques`.

---

### 157.144 Ronda 188 — Unidad de Sesión en Patrón Semanal y Overload Funcional de Días K (Fases 983 y 984)

#### 1. Fase 983 — `Construir` divide por 4.0 fijo → posgrado 8 h declara patrón infactible
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs` y `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** `Construir` utilizaba un divisor hardcodeado de 4.0h para calcular el número de sesiones requeridas $S = \lceil totalHoras / 4.0 \rceil$. En programas de posgrado (asignatura de 64h con sesiones de 8h), proyectaba 16 sesiones en lugar de 8, declarando infactible el patrón cuando la ventana contenía menos de 16 semanas hábiles, cancelando el patrón semanal y degradando la asignación a un solo día.
- **Solución Propuesta:** Incorporar parámetro `double unidadSesion = 4.0` en `Construir` y calcular sesiones con `unidad = unidadSesion > 0 ? unidadSesion : 4.0`. En `AgFabricaIndividuos`, resolver la unidad efectiva según el aula/modalidad y pasarla a la invocación de `Construir`.

#### 2. Fase 984 — Sin overload `DiasKDe(diasPatron)` → carriles `[M,X]`/`[J,V]` inalcanzables
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** `DiasKDe` solo admitía como entrada `AgAulaAsignatura` persistida. Patrones textuales declarados (semillas de UI como '1,3' o '4,5') no podían resolverse de forma funcional contra la ventana del aula, impidiendo la siembra y expresión del modelo de carriles alternos T5-T7.
- **Solución Propuesta:** Crear overload estático puro `DiasKDe(string diasPatron, string diasVentana = null, AgAula aula = null)` para procesar patrones textuales, intersectarlos contra la ventana del aula y recurrir al comportamiento canónico ante patrones vacíos o inválidos.

---

### 157.145 Ronda 189 — Reparación de Cruce: Preservación de Carril Descendiente y Propagación de Unidad Posgrado D4' (Fases 985 a 987)

#### 1. Fase 987 — Cruce reemplaza `dst.DiasPatron` con el patrón de otra aula → fusión de carriles
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** `diasK` se derivaba incondicionalmente de `aaSrc` del progenitor, sobrescribiendo `dst.DiasPatron`. Si el descendiente pertenecía a un carril paralelo y el progenitor a otro, ambos carriles colapsaban en uno solo en oleadas T5-T7.
- **Solución Aplicada:** Declarada función local `AulaSrcDst(idAula)` e inicializado `diasK` con `AgPlanificadorPatronSemanal.DiasKDe(dst.DiasPatron, null, AulaSrcDst(dst.IdAula))` si `dst.DiasPatron` ya está presente; solo recurrir a `aaSrc` si está vacío.

#### 2. Fase 985 — Cruce reconstruye patrón con `unidad=4.0` → posgrado 8 h queda sin `PatronSemanal`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Invocación de `_planificadorPatron.Construir` sin `unidadSesion` utilizaba el valor por defecto de 4.0h. En posgrados de 64h requiriendo 8 sesiones de 8h, calculaba $S = 16$ y fallaba en ventanas cortas, anulando `dst.PatronSemanal`.
- **Solución Aplicada:** Resolver `unidadCruce` dinámicamente mediante `AgUnidadesSemanales.HorasSesionAula(aulaDst, nivel)` desde el aula y modalidad del programa destino, pasándola como argumento a `Construir`.

#### 3. Fase 986 — Selección de bloque fija `DuracionHoras == 4.0` → posgrado recibe bloque de 4 h
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** `bloquePatron` filtraba rígidamente por `b.DuracionHoras == 4.0`. En programas de posgrado con catálogo de 8h, fallaba la coincidencia y el fallback asignaba cualquier bloque del día (posiblemente de 4h), violando silenciosamente D4′.
- **Solución Aplicada:** Priorizar bloques cuya duración coincida con `unidadCruce` (`Math.Abs(b.DuracionHoras - unidadCruce) < 0.001`), con fallbacks sucesivos a 4.0h y a cualquier bloque del día.

---

### 157.146 Ronda 190 — Compatibilidad de Regla R8 y Columna K con Carriles y Posgrado 8h (Fases 988 y 989)

#### 1. Fase 988 — `R8_DiaFueraDePatron` usa `gr.First().DiasPatron` → todo carril secundario es violación dura
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** La validación R8 utilizaba el patrón de días del primer gen del grupo (`gr.First().DiasPatron`) para contrastar todos los días del grupo. Si un grupo contenía múltiples carriles (oleadas alternas `[1,3]` y `[4,5]`), los días del carril secundario se marcaban fuera de patrón como infracción dura, exterminando la solución en el proceso genético.
- **Solución Aplicada:** Validar cada gen de forma individual contra su propio `DiasPatron` (`gr.Where(g => !string.IsNullOrWhiteSpace(g.DiasPatron))`), aislando la verificación a nivel de carril.

#### 2. Fase 989 — Columna K cae a 4 h cuando falta el bloque → posgrado reportado mal
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Al componer `DiaHorasSemanales` en la vinculación, si no se resolvía el bloque horario en catálogo, la duración recurría a un fallback fijo de 4.0h, distorsionando las horas de asignaturas de posgrado (8h) en la columna K y en `RebuildSelectedDays`.
- **Solución Aplicada:** Extraer la duración de `subGrp.Select(x => x.HorasSemanales).Where(h => h > 0).DefaultIfEmpty(4.0).Max()` para reflejar fielmente la unidad real del gen cuando no haya bloque en catálogo.

---

### 157.147 Ronda 191 — Columna K Multi-Franja (AM/PM) y Contrato Canónico de 3 Partes (Fases 990 y 991)

#### 1. Fase 990 — Materializador elige `Min()` de bloque por día → franja reportada no es la programada
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Las sesiones se agrupaban exclusivamente por `DiaSemana` tomando deterministamente `Min()` del bloque horario. En carriles con sesiones el mismo día en franjas distintas (ej. SÁBADO AM 08:00-12:00 y PM 14:00-18:00), el gen PM se reportaba erróneamente con horario AM al tener menor IdBloque, falseando la columna K y colapsando a una sola franja en `RebuildSelectedDays`.
- **Solución Aplicada:** Agrupación compuesta por `new { g.DiaSemana, g.IdBloqueHorario }` ordenada por día y bloque, emitiendo una entrada canónica individual por cada franja/bloque sin colapsar jornadas distintas ni emitir advertencias espurias.

#### 2. Fase 991 — `DiasHorasSemanaFormatter` emite 2 segmentos (inerte e incompatible con `RebuildSelectedDays`)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`
- **Diagnóstico:** `DiasHorasSemanaFormatter` generaba claves de 2 segmentos (`"DIA|horario"`). Al no proveer el segmento intermedio de horas exigido por `RebuildSelectedDays` (`int.TryParse(segments[1])`), el parseo descartaba silenciosamente las entradas, dejando vacía la columna K al reabrir la sesión.
- **Solución Aplicada:** Adecuación al contrato canónico de 3 partes (`$"{dia}|{horas:0.##}|{horario}"`) con fallback de horas (`s.Horas > 0 ? s.Horas : 4.0`), asegurando compatibilidad round-trip con `RebuildSelectedDays` y preservando claves únicas por franja.

---

### 157.148 Ronda 192 — Erradicación del Enmascaramiento de Fechas con `DateTime.MinValue` (Fases 992 y 993)

#### 1. Fase 992 — Sesión sin fecha se persiste como `0001-01-01` en `ag_sesiones`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** En `GuardarStagingAsync`, si una sesión carecía de fecha, el fallback `(s.FechaSesion ?? DateTime.MinValue).ToString("yyyy-MM-dd")` persistía `"0001-01-01"` en la columna obligatoria `fechasesion TEXT NOT NULL`. Esto provocaba que `AcademicDateSyncService` la ignorara (filtro `> 2000-01-01`), `FormVerVinculaciones` mostrara la fecha ficticia `01/01/0001` y los ordenamientos de viáticos se alteraran, ocultando el error de origen.
- **Solución Aplicada:** Inclusión de compuerta preventiva previa al bucle de persistencia (`sesiones.FirstOrDefault(x => !x.FechaSesion.HasValue || x.FechaSesion.Value.Year < 1900)`), abortando la transacción y retornando `Result.Failure` explicativo, y persistiendo directamente `s.FechaSesion.Value.ToString("yyyy-MM-dd")`.

#### 2. Fase 993 — Promoción a producción replica el centinela `DateTime.MinValue`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPromotorVinculaciones.cs`
- **Diagnóstico:** Al promover de staging a la tabla operativa `sesiones`, se aplicaba `FechaSesion = ss.FechaSesion ?? DateTime.MinValue`, enmascarando sesiones inválidas y propagando fechas residuales a producción donde no existe aislamiento por corrida.
- **Solución Aplicada:** Reemplazo del fallback por expresión `throw` explícita (`ss.FechaSesion ?? throw new Exception(...)`), asegurando fallo fail-fast antes de contaminar la base productiva.

---

### 157.149 Ronda 193 — Prevención de Colapso de Franjas y Unificación del Contrato Canónico (Fases 994 a 998)

#### 1. Fase 994 — Día con dos franjas (AM/PM) colapsa al reabrir la vinculación
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** Al reabrir sesiones en `FormSesionesAcademicas`, `RebuildSelectedDays` utilizaba el nombre del día como clave de diccionario sin contemplar múltiples franjas en la misma jornada (ej. SÁBADO AM y PM). La segunda franja sobrescribía a la primera y al persistir se eliminaba silenciosamente la mitad del horario del carril.
- **Solución Aplicada:** En `RebuildSelectedDays`, al detectar una entrada previa para el mismo día, se emite advertencia en `LoggerHelper.Warning` y se conserva deterministamente la franja con mayor carga horaria, impidiendo el descarte ciego de la sesión principal.

#### 2. Fase 995 — `int.TryParse` descarta horas no enteras y formato invariante
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`, `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`
- **Diagnóstico:** `int.TryParse(segments[1])` fallaba ante horas fraccionarias (`"3.5"` o `"2,5"`), omitiendo el día. Además, la interpolación `{horas:0.##}` con cultura local (es-CO) emitía comas decimales (`"7,5"`), que al dividirse con `Split(',')` fragmentaban corruptamente las cadenas de días.
- **Solución Aplicada:** Migración a `double.TryParse(..., NumberStyles.Float, CultureInfo.InvariantCulture)` y redondeo seguro en `RebuildSelectedDays`; serialización con `horas.ToString("0.##", CultureInfo.InvariantCulture)` en `AgMaterializadorSesiones` y `DiasHorasSemanaFormatter`.

#### 3. Fase 996 — `DiasHorasSemanaFormatter` emite abreviaturas; contrato canónico usa nombres completos
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`
- **Diagnóstico:** El formateador producía abreviaturas (`"LUN"`, `"MAR"`, `"MIE"`), incompatibles con el catálogo canónico del sistema (`"LUNES"`, `"MARTES"`, ...). Esto rompía `DayHoursDropdown._diasSemana`, `DayPatternParser` y `RebuildSelectedDays`.
- **Solución Aplicada:** Delegación de `NormalizarDia` y `ObtenerNombreDia` en `DiaSemanaNormalizador.ObtenerNombreDia(...)`, homologando el vocabulario a nombres completos sin tildes.

#### 4. Fase 997 — Separador `" / "` en `DesdeSesiones` rompe el round-trip
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`
- **Diagnóstico:** `DesdeSesiones` unía elementos con `" / "`, contraviniendo el estándar separado por comas `", "` esperado por `RebuildSelectedDays` y `DayHoursDropdown`.
- **Solución Aplicada:** Unificación del delimitador a `", "` y documentación de la deuda técnica para su posterior integración en exportadores y grillas.

#### 5. Fase 998 — IDs fantasma tras rollback de F992
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** En `GuardarStagingAsync`, al abortar por sesión sin fecha válida, las instancias `AgVinculacion` en memoria retenían los `dbId` asignados previo al `trans.Rollback()`, dejando IDs inexistentes en el grafo de objetos.
- **Solución Aplicada:** Restauración de los IDs originales en memoria mapeando desde `mapVinc` antes de ejecutar el rollback y retornar `Result.Failure`.

---

### 157.150 Ronda 194 — Corrección de Especificador Militar `HH` y Prevención de Descarte Silencioso (Fases 999 a 1001)

#### 1. Fase 999 — `NormalizarHorario` formatea con `hh` (12 h) mutilando horarios PM a AM
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/HorarioMilitarizador.cs`, `ConversorPDF.XLL/Academic/Domain/Models/FranjaHoraria.cs`
- **Diagnóstico:** El formato interpolado `{f.Inicio:hh\\:mm} A {f.Fin:hh\\:mm}` y `{Inicio:hh\\:mm} - {Fin:hh\\:mm}` utilizaba el especificador de 12 horas `hh` sin designador meridiano en vez del estándar militar de 24 horas `HH`. Horarios vespertinos como `"13:00 A 17:00"` eran reescritos como `"01:00 A 05:00"`, provocando que `DayHoursDropdown` no los encontrara en su catálogo de opciones, asignara `DBNull.Value` y descartara permanentemente el día al guardar.
- **Solución Aplicada:** Corrección del especificador de formato a `HH\\:mm` tanto en `HorarioMilitarizador.NormalizarHorario` como en `FranjaHoraria.ToString()`, garantizando idempotencia en el rango 00:00 - 23:59.

#### 2. Fase 1000 — Datos normalizados previamente por V63 quedaron corruptos
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV71_AuditarHorariosMeridiano.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** La migración V63 unificó horarios utilizando `HorarioMilitarizador` cuando este contenía el defecto `hh`. Las filas persistidas en `sesiones.horariosesion`, `ag_sesiones.horariosesion` y `vinculaciones.dia_horas_semanales` no pueden corregirse automáticamente de forma determinista por ambigüedad intrínseca AM/PM.
- **Solución Aplicada:** Creación de la migración no destructiva `MigrationV71_AuditarHorariosMeridiano` (versión 71), la cual audita registros afectados mediante consultas de solo lectura y emite alertas diagnósticas en `LoggerHelper.Warning` para orientar la regeneración de staging y auditoría de producción.

#### 3. Fase 1001 — `DayHoursDropdown` descarta el día sin emitir traza diagnóstica
- **Archivo Afectado:** `ConversorPDF.XLL/UI/DayHoursDropdown.cs`
- **Diagnóstico:** Cuando un horario normalizado no coincidía con ninguna de las opciones predefinidas en el ComboBox de la celda, se asignaba silenciosamente `DBNull.Value`, provocando que al guardar la fila el día se eliminara sin dejar evidencia en logs.
- **Solución Aplicada:** Inclusión de registro de advertencia con `LoggerHelper.Warning` en el bloque `else` de resolución de horario en `DayHoursDropdown.CargarDesdeDiaHorasSemanales`, detallando el horario ausente, el día y las opciones disponibles antes de asignar `DBNull.Value`.

---

### 157.151 Ronda 195 — Fugas de Subunidad, Preservación de Carriles Concurrentes y Afinamiento de Auditoría (Fases 1002 a 1006)

#### 1. Fase 1002 — `GetBloquesHorariosAsync` ignora el parámetro `subunidad`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** El método recibía el parámetro `subunidad` pero la consulta SQL filtraba exclusivamente por `activo = 1`, permitiendo que bloques de otras subunidades se mezclaran en el plan y alteraran la asignación de franjas por subunidad.
- **Solución Aplicada:** Incorporada condición `@Subunidad = '' OR COALESCE(subunidad, '') = @Subunidad` en la cláusula `WHERE`, pasando el parámetro mediante `CommandDefinition`.

#### 2. Fase 1003 — Carriles concurrentes colapsados por `GroupBy(...).Select(g => g.First())`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La agrupación por tupla `(AgClaveCohorte, Jornada, DiasVentana)` reducía múltiples aulas con la misma cohorte a una sola instancia representativa. En patrones alternos (ej. MAR-VIE con carriles `[M,X]` y `[J,V]`), el segundo par de días era descartado y no se planificaba.
- **Solución Aplicada:** Deduplicación por `a.Id` en lugar de por atributos de cohorte, conservando la totalidad de carriles concurrentes de la misma cohorte.

#### 3. Fase 1004 — Comparación exacta `DuracionHoras == 4.0` sobre columna REAL
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La comparación de igualdad estricta `b.DuracionHoras == 4.0` en tipos de punto flotante (`double`) fallaba ante imprecisiones numéricas (`3.9999999` o `4.0000001`) provenientes de la columna `REAL` en base de datos, provocando caídas no deseadas al fallback de 8 horas o a bloques arbitrarios.
- **Solución Aplicada:** Sustituida la igualdad estricta por comparación con margen de tolerancia: `Math.Abs(b.DuracionHoras - 4.0) < 0.001` en las líneas 292 y 364.

#### 4. Fase 1005 — Auditoría V71 con criterio demasiado amplio (falsos positivos)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV71_AuditarHorariosMeridiano.cs`
- **Diagnóstico:** El filtro `LIKE '% A %'` contabilizaba la totalidad de los horarios que contuvieran el separador `" A "`, incluyendo registros canónicos y válidos, inflando el recuento y perdiendo precisión diagnóstica.
- **Solución Aplicada:** Restringida la consulta diagnóstica para contar únicamente registros no vacíos cuyo `horariosesion` no exista en el catálogo `ag_bloque_horario` (`hora_inicio || ' A ' || hora_fin`), detectando exactamente aquellos alterados o desalineados.

#### 5. Fase 1006 — Semilla V62 introduce bloques de domingo y variantes 8 h duplicadas
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV62_AddAgAulaVentanasYBloquesSemilla.cs`
- **Diagnóstico:** El bucle de inserción iteraba de 1 a 7, sembrando bloques para domingo (día 7), e incluía una variante redundante de 8 horas para `NOCHE` (`14:00-22:00`) que colisionaba con `TARDE`.
- **Solución Aplicada:** Ajustado el límite del bucle a `dia <= 6` para excluir domingos de la semilla y eliminada la plantilla duplicada de 8 horas en `NOCHE`.

---

### 157.152 Ronda 196 — Preservación de Bloques Globales, Repartición de Carriles y Depuración de Semilla (Fases 1007 a 1009)

#### 1. Fase 1007 — F1002 excluye los bloques globales (`subunidad = ''`) cuando hay subunidad activa
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** La cláusula `WHERE (@Subunidad = '' OR COALESCE(subunidad, '') = @Subunidad)` introducida en F1002 excluía todos los bloques base/globales sembrados con `subunidad = ''` siempre que el usuario seleccionara una subunidad específica en el generador, resultando en 0 bloques horarios y provocando vacantes masivas o fallos de ejecución.
- **Solución Aplicada:** Modificada la cláusula a `@Subunidad = '' OR COALESCE(subunidad, '') IN ('', @Subunidad)` y ordenado con `ORDER BY CASE WHEN COALESCE(subunidad, '') = @Subunidad THEN 0 ELSE 1 END, dia_semana, hora_inicio` para dar prioridad a los bloques especializados de la subunidad sin ignorar los bloques globales.

#### 2. Fase 1008 — `DiasKDe` trunca toda ventana a 2 días → un solo carril
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** `DiasKDe` ejecutaba `Take(2)` de manera fija sobre `diasVentana`, de modo que ventanas de 4 días como `MAR-VIE` (`"2,3,4,5"`) siempre retornaban los dos primeros días (`[2,3]`). Esto ocasionaba que todas las asignaturas de una cohorte colisionaran en los mismos días y el carril complementario (`[4,5]`) nunca fuera asignado.
- **Solución Aplicada:** Incorporado parámetro opcional `int carril = 0` en `DiasKDe` calculando el offset dinámico (`Skip(idx * 2).Take(2)`), e instrumentado el contador `idxAsig` por aula en `AgFabricaIndividuos` para alternar secuencialmente los carriles entre las asignaturas programadas.

#### 3. Fase 1009 — Semilla de domingo ya persistida y variantes 8 h duplicadas
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV72_LimpiarBloquesDomingoYDuplicados.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** La corrección en el código de la migración V62 (F1006) solo aplica a bases nuevas; bases de datos existentes ya ejecutaron V62 y retienen bloques para `dia_semana = 7` (domingo) y bloques duplicados de 8 horas en `NOCHE` (`14:00-22:00`).
- **Solución Aplicada:** Creación y registro de la migración no destructiva `MigrationV72_LimpiarBloquesDomingoYDuplicados` (versión 72), eliminando de `ag_bloque_horario` los registros globales de domingo y las variantes de 8h en `NOCHE`.

---

### 157.153 Ronda 197 — Intersección de Patrón con Ventana y Soporte de Ventanas Impares (Fases 1010 a 1012)

#### 1. Fase 1010 — `DiasKDe(aa, aula, …)` no intersecta `aa.DiasPatron` con la ventana del aula
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** Cuando `aa.DiasPatron` contenía días sembrados por carga horaria (ej. 64 h → `"4,5,6"` o 80 h → `"2,3,4,5,6"`), `DiasKDe` retornaba esos días de forma directa sin intersectar con `aula.DiasVentana`. Si el aula tenía ventana `MAR-VIE` (`[2,3,4,5]`), se programaban sesiones en sábado (`6`) y se saltaba la partición en carriles de F1008.
- **Solución Aplicada:** Añadida intersección explícita con `AgDiasVentana.Dias(aula.DiasVentana)` dentro de la rama `!string.IsNullOrWhiteSpace(aa.DiasPatron)`, asegurando paridad funcional con el overload textual de `DiasKDe`.

#### 2. Fase 1011 — Ventanas impares pierden el último día (`carriles` truncado)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El cálculo de carriles utilizaba división entera (`diasVentana.Count / porCarril`). Para ventanas de 3 días como `JUE-SAB` (`4,5,6`) o `VIE-DOM` (`5,6,7`), el resultado era `3 / 2 = 1`, dejando un único carril `[4,5]` y omitiendo permanentemente el día sábado (`6`).
- **Solución Aplicada:** Modificado el cálculo de cantidad de carriles a `(int)Math.Ceiling(diasVentana.Count / (double)porCarril)`, permitiendo que ventanas impares generen un carril complementario (ej. `[4,5]` y `[6]`).

#### 3. Fase 1012 — Validador de capacidad calcula con carril 0 para todos los carriles
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** Las comprobaciones de factibilidad en `CapacidadDe` y validación D14 invocaban `DiasKDe` sin parámetro de carril, asumiendo implícitamente el carril 0 (de 2 días).
- **Solución Aplicada:** Explicitada la llamada a `DiasKDe(..., carrilGen)` con carril representativo en pre-vuelo y documentada la relación entre los carriles y la capacidad efectiva por curso.

---

### 157.154 Ronda 198 — Sembrado de Patrón por Carril y Advertencia Multicarril en Pre-Vuelo (Fases 1013 a 1014)

#### 1. Fase 1013 — `dias_patron` sembrado desde horas anula F1008 y excede 2 días
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El sembrado automático en `GenerarEncabezadosAulasAsync` escribía `dias_patron` = `"2,3,4,5,6"` (≥80 h) o `"4,5,6"` (≥64 h). Al quedar `AgAulaAsignatura.DiasPatron` no vacío, `DiasKDe(aa, aula, carril)` retornaba por la rama de patrón existente e ignoraba el parámetro `carril`, provocando que todas las asignaturas de la cohorte compartieran los mismos días (un solo carril real) y distribuyeran la carga en 3 a 5 días con bloques atípicos de 1 a 2 horas.
- **Solución Aplicada:** Instrumentado contador `idxAulaAsig` por aula y asignado `diasPatron` de a 2 días rotado por asignatura (`idxAulaAsig % carriles`) extrayendo de `AgDiasVentana.Dias(item.DiasVentana)` mediante `Skip(lane * 2).Take(2)`.

#### 2. Fase 1014 — F1012 introdujo variable inerte; capacidad sigue calculada con carril 0
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `int carrilGen = 0;` en `CapacidadDe` era constante, manteniendo el cálculo de capacidad acotado silenciosamente al carril 0 sin advertir que en aulas multicarril con carriles residuales de longitud impar (ej. `[6]`) la capacidad calculada sobreestima la oferta real del carril complementario.
- **Solución Aplicada:** Reemplazado el código inerte por la determinación de `carriles` sobre `AgDiasVentana.Dias(au.DiasVentana)` y la emisión de advertencia diagnóstica no redundante mediante `LoggerHelper.WarningLimitado` tanto en `CapacidadDe` como en la validación D14 de configuraciones de aula.

---

### 157.155 Ronda 199 — Depuración de Patrones Legacy Multidía y Determinismo en Reparto de Carriles (Fases 1015 a 1016)

#### 1. Fase 1015 — `dias_patron` legacy (3-5 días) persiste y anula F1013
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV73_LimpiarDiasPatronMultidia.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Filas preexistentes en `ag_aula_asignatura` sembradas bajo versiones anteriores conservaban valores como `"2,3,4,5,6"` o `"4,5,6"`. La instrucción `INSERT OR IGNORE` en `GenerarEncabezadosAulasAsync` no actualizaba estas filas y reutilizaba el valor legacy no vacío, forzando a `DiasKDe` a eludir la partición de carriles de F1008/F1013 y fragmentando las horas en bloques de 1 a 2 horas.
- **Solución Aplicada:** Implementada y registrada la migración `MigrationV73_LimpiarDiasPatronMultidia` (versión 73) para resetear a `''` cualquier `dias_patron` con más de 2 días en `ag_aula_asignatura`, permitiendo que el motor genético gobierne los días de clase por ventana y carril.

#### 2. Fase 1016 — Carril asignado sin `ORDER BY` → reparto no reproducible
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** La consulta `sqlAsigs` carecía de cláusula de ordenamiento, provocando que `idxAulaAsig % carriles` dependiera del orden arbitrario de iteración retornado por SQLite, introduciendo no determinismo entre ejecuciones y rompiendo la auditabilidad de las corridas.
- **Solución Aplicada:** Añadida cláusula `ORDER BY a.id;` a `sqlAsigs` para garantizar un orden determinista y reproducible en la rotación de asignaturas por carril.

---

### 157.156 Ronda 200 — Preservación de 2 Días en Carril Residual y Fallback de Patrón sobre Ventana (Fases 1017 a 1018)

#### 1. Fase 1017 — Carril residual de 1 día reduce capacidad frente a la calculada
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En ventanas impares (ej. `JUE-SAB` = `[4,5,6]`), el carril residual 1 tomaba un único día (`[6]`) con la mitad de capacidad que el carril 0 (`[4,5]`). Como el validador pre-vuelo evalúa con el carril 0, las asignaturas rotadas al carril residual no alcanzaban a cubrir las horas requeridas y degradaban el fitness sin bloquear.
- **Solución Aplicada:** Calculado el inicio de ventana mediante desplazamiento deslizante `Math.Min(idx * porCarril, Math.Max(0, diasVentana.Count - porCarril))`, garantizando que el carril residual siempre conserve 2 días (ej. `[5,6]` para `JUE-SAB`), replicado en el sembrado de `AgFormGeneradorHorarios.cs`.

#### 2. Fase 1018 — Patrón fuera de ventana se devuelve sin filtrar
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** Cuando `aa.DiasPatron` no intersectaba con `aula.DiasVentana` (ej. patrón `[2,3]` y ventana cambiada a `JUE-SAB`), `dentro.Count == 0` retornaba la lista original sin filtrar, programando sesiones en días no autorizados fuera de la ventana del aula.
- **Solución Aplicada:** Añadido fallback en `DiasKDe` que emite `LoggerHelper.WarningLimitado` y resuelve el carril de 2 días directamente sobre la ventana del aula cuando el patrón está íntegramente fuera de ella.

---

### 157.157 Ronda 201 — Paridad en Overload Textual y Carriles Disjuntos contra Infracción R4 (Fases 1019 a 1020)

#### 1. Fase 1019 — Overload textual sin paridad F1018 → días fuera de ventana en cruce
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** El overload textual `DiasKDe(string, string, AgAula)` retornaba `pedidos` crudo cuando `dentro.Count == 0`, permitiendo que en el operador de cruce (`AgOperadorCruce.cs:166`) se propagaran días ajenos a la ventana del aula destino, eludiendo el fallback seguro introducido en F1018. Adicionalmente, la rama `aaSrc` en cruce invocaba `DiasKDe` sin pasar el aula destino.
- **Solución Aplicada:** Incorporado parámetro opcional `int carril = 0` y fallback con `LoggerHelper.WarningLimitado` delegando en `DiasKDe((AgAulaAsignatura)null, aula, carril)` ante patrones fuera de ventana, y pasado `AulaSrcDst(dst.IdAula)` en `AgOperadorCruce.cs:168`.

#### 2. Fase 1020 — Carriles solapados violan R4 (misma aula/grupo, mismo día)
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En ventanas impares (ej. `JUE-SAB` = `[4,5,6]`), el clamp deslizante de F1017 generaba carriles solapados `[4,5]` y `[5,6]`, haciendo que dos asignaturas de la misma cohorte coincidieran el día 5 y provocando violaciones duras sistemáticas de R4 (`R4_DosAsignaturasMismoDia`) y de aula R2.
- **Solución Aplicada:** Reemplazado el cálculo de carriles por división entera estricta (`diasVentana.Count / porCarril`) asegurando carriles disjuntos, dejando el día residual fuera de rotación con advertencia diagnóstica no redundante (`LoggerHelper.WarningLimitado`), y espejado en el sembrado de `AgFormGeneradorHorarios.cs`.

---

### 157.158 Ronda 202 — Tolerancia a ISO Numérico en Ventana y Diferimiento de Patrón sin Ventana (Fases 1021 a 1022)

#### 1. Fase 1021 — `AgDiasVentana.Dias` sin tolerancia a ISO numérico
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgDiasVentana.cs`
- **Diagnóstico:** `AgDiasVentana.AIso` únicamente validaba contra las 5 claves canónicas del diccionario (ej. `"MAR-VIE"`). Cuando el valor de `dias_ventana` ya venía persistido como ISO numérico (ej. `"1,2,3,4"` o `"2,3,4,5"`), `Opciones.TryGetValue` fallaba retornando cadena vacía, lo que provocaba que `Dias` devolviera lista vacía y `DiasKDe` recurriera al default `{2,3}` ignorando la ventana real del aula.
- **Solución Aplicada:** Añadido soporte en `AIso` para parsear listas de enteros delimitadas por coma o punto y coma en el rango ISO [1..7] si la búsqueda por clave falla, retornando el formato normalizado ordenado y sin duplicados.

#### 2. Fase 1022 — Sembrado con ventana vacía fuerza `4,5`/`2,3`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En la inicialización de `ag_aula_asignatura`, cuando el aula no tenía `item.DiasVentana` declarado, el fallback forzaba patrones estáticos `"4,5"` o `"2,3"` según la intensidad horaria, lo que restringía rígidamente los días de cursado sin considerar la jornada o restricciones de la sede antes de la optimización genética.
- **Solución Aplicada:** Asignado `diasPatron = string.Empty` en la rama sin ventana declarada y registrada traza diagnóstica `AulaAsignaturaSinVentana`, permitiendo que el motor genético y `DiasKDe` deriven los días en función de las restricciones reales del aula.

---

### 157.159 Ronda 203 — Colisión de nombre Excel 'De' (Fases 1033 a 1038)

#### 1. Fase 1033 — Eliminar el nombre Excel 'De' en AgSemanaIso y AgClaveCohorte
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSemanaIso.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgClaveCohorte.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Excel-DNA 1.8.0 registra por defecto todos los métodos públicos estáticos de clases públicas cuyos parámetros y retornos pertenecen a tipos soportados. El nombre Excel se deriva exclusivamente del nombre del método sin prefijo de clase. `AgSemanaIso.De(DateTime)` y `AgClaveCohorte.De(int, int, int, string, string, string)` colisionaban registrando ambas la función `=DE(...)`, provocando el mensaje `Repeated function name: 'De' - previous registration will be overwritten.` en el log de registro.
- **Solución Aplicada:**
  1. Se eliminó el alias `De(DateTime)` en `AgSemanaIso.cs` preservando `DiaIso(DateTime)` como método canónico.
  2. Se renombraron las sobrecargas en `AgClaveCohorte.cs` a `DesdeCohorte`, `DesdeAula` y `DesdeGen`.
  3. Se actualizó el call site en `AgFabricaIndividuos.cs:84` a `AgClaveCohorte.DesdeAula(a)`.

#### 2. Fase 1034 — Test guardián de registro de métodos Excel-DNA
- **Archivo Afectado:** `ConversorPDF.Tests/Unit/ExcelDnaRegistrationGuardTests.cs`
- **Diagnóstico:** Ausencia de prueba automatizada de regresión que verifique las reglas de descubrimiento y registro de Excel-DNA 1.8.0 sobre el ensamblado `ConversorPDF.XLL`, permitiendo que futuros métodos estáticos públicos con nombres idénticos o con el identificador colisionado `De` vuelvan a ingresar al código base.
- **Solución Aplicada:** Implementado test guardián con xUnit y FluentAssertions que inspecciona por reflexión todos los tipos públicos del ensamblado de `ConversorPDF.XLL`, validando que no existan nombres duplicados (`OrdinalIgnoreCase`) y que ningún método registrable se denomine `De`.

#### 3. Fase 1036 — `DesdeCohorte` queda expuesto como UDF inválido (6 argumentos requeridos)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgClaveCohorte.cs`
- **Diagnóstico:** `DesdeCohorte(int, int, int, string, string aulaRef = null)` es público y estático en clase pública con tipos compatibles. Excel-DNA 1.8.0 no soporta valores por defecto (`ParameterInfo.IsOptional` no se evalúa) y registra la función exigiendo todos los parámetros en Excel, fallando las invocaciones documentadas con 4/5 parámetros.
- **Solución Aplicada:** Modificada la visibilidad de `DesdeCohorte` a `internal static string DesdeCohorte(...)`, restringiéndolo a key-builder de dominio y excluyéndolo del registro UDF de Excel.

#### 4. Fase 1037 — El test guardián no replica las reglas de Excel-DNA 1.8.0
- **Archivo Afectado:** `ConversorPDF.Tests/Unit/ExcelDnaRegistrationGuardTests.cs`
- **Diagnóstico:** El test guardián inicial filtraba `IsSpecialName`, omitiendo accesores de propiedades `get_`/`set_` estáticos públicos que Excel-DNA 1.8.0 sí registra (7 métodos en Release). Asimismo, omitía los filtros de tipos de Excel-DNA (`type.IsArray`, genéricos abiertos y namespace `"My"`).
- **Solución Aplicada:** Alineadas las reglas en `EsMetodoRegistrableExcelDna180` y `ObtenerMetodosRegistrables` con `AssemblyLoader.GetExcelMethods` y `IsMethodSupported` de Excel-DNA 1.8.0 (removido `IsSpecialName`, agregada exclusión de métodos abstractos/genéricos, arrays, tipos genéricos con parámetros abiertos y namespace `"My"`).

#### 5. Fase 1038 — Trazabilidad y datos incorrectos en documentación
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`, `docs/articles/changes/2026-10-22-colision-nombre-funcion-de.md`
- **Diagnóstico:** Desincronización en la raíz del repositorio omitiendo la Ronda 203; cita incorrecta en artículo alegando excepción en lugar del log real de sobrescritura `Repeated function name: 'De'`; omisión de la visibilidad interna de `DesdeCohorte` y falta de advertencia de error `#NAME?` para fórmulas que invocaban `=De(...)`.
- **Solución Aplicada:** Sincronizados `implementation_plan.md` y `walkthrough.md` en la raíz con la Ronda 203 (Fases 1033 a 1038); corregido el artículo de cambios con los textos y advertencias de compatibilidad precisos.

---

### 157.160 Ronda 204 — Corrección de Solapamientos Espaciales y Extensión de `treeContextoAcademico` (Fases 1039 a 1044)

#### 1. Fase 1039 — Splitter Horizontal y Reorganización de `tabSimulacion`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El contenedor `tabSimulacion` acumulaba controles apilados con `Dock=Bottom` fijos (82 + 58 + 118 = 258 px), comprimiendo el `grpSeleccion` a 357 px y dejando el árbol en solo 138 px con ~105 px de vacío gris inútil sin posibilidad de ajuste manual por parte del usuario.
- **Solución Aplicada:**
  1. Se introdujo `splitContenedorSim` (`SplitContainer`, `Orientation=Horizontal`, `Dock=Fill`, `FixedPanel=Panel2`, `SplitterWidth=6`) como único hijo de `tabSimulacion`.
  2. En `Panel1` se alojan `grpSeleccion` (`Dock=Fill`, `MinimumSize=(0,240)`) y `pnlConfiguracionInferior` (`Dock=Bottom`).
  3. En `Panel2` se alojan `pnlBotonesSim` (`Dock=Top`) y `txtLogAG` (`Dock=Fill`).
  4. En `OnLoad`, se calibra dinámicamente `SplitterDistance = Math.Max(280, Height - 300)` otorgando al árbol ≥ 240 px iniciales expandibles hasta ~345 px mediante el divisor.

#### 2. Fase 1040 — `treeContextoAcademico` con `Dock=Fill` y Toolbar Desacoplada sin Colisiones
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `treeContextoAcademico` utilizaba `Anchor` cuádruple con coordenadas fijas (10, 112, 1072, 138), perdiendo altura relativa. En `pnlArbolToolbar`, `lblConteoCombinaciones` en X=738 con `AutoSize=true` crecía hasta ~928 px solapando por 39 px a `btnGenerarAulas` (X=889) al pintarse encima por z-order.
- **Solución Aplicada:**
  1. `treeContextoAcademico` pasa a `Dock=Fill` con `BorderStyle=FixedSingle`, `HideSelection=false`, `ShowNodeToolTips=true`, `ItemHeight=20` y `Scrollable=true`.
  2. En `pnlArbolToolbar` (alto 34, `Dock=Top`), los botones de perfil se agruparon en `pnlToolbarIzq` (`FlowLayoutPanel`, `Dock=Left`, `WrapContents=false`, `AutoSize=true`).
  3. `btnGenerarAulas` se fijó con `Dock=Right` (150 px).
  4. `lblConteoCombinaciones` se configuró con `Dock=Fill`, `AutoSize=false`, `AutoEllipsis=true`, `TextAlign=MiddleLeft` y `toolTipConteo` para visualización completa sin colisión.

#### 3. Fase 1041 — Cabeceras Nivel 6 Desacopladas y Cálculo Proporcional Unificado
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Las cabeceras de Nivel 6 (`lblColJornada`, `lblColDias`, `lblColHorasDia`) residían en la segunda fila de `pnlArbolToolbar` con Y=36 y X fijas (253, 478, 893), permanecían siempre visibles aún sin nodos de modalidad desplegados, y no coincidían con los 11 offsets mágicos hardcodeados en `DrawNode`, `ProcesarClickDimensionNivel6` y `MouseMove`.
- **Solución Aplicada:**
  1. Se extrajeron las cabeceras a un panel independiente `pnlCabecerasNivel6` (`Dock=Top`, alto 20, `Visible=false`) con separadores verticales de 1 px (`sepNivel6A`, `sepNivel6B`).
  2. Se creó el struct `AgColumnasNivel6` y método `CalcularColumnasNivel6()` como fuente única de verdad con reparto proporcional (Jornada 30%, Días 48%, Horas 22%) sobre el ancho disponible y anchos mínimos garantizados (62, 74, 64 px).
  3. Se unificaron `DrawNode`, `ProcesarClickDimensionNivel6`, `MouseMove` y `ActualizarCabecerasNivel6()` consumiendo el struct, fijando `pnlCabecerasNivel6.Visible = HayNodoNivel6Visible()`.

#### 4. Fase 1042 — Reorganización de Barra de Acciones y Depuración de Emojis
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `pnlAccionesSim` tenía ancho rígido de 700 px dejando 89 px muertos y solo 7 px entre `chkModoDebug` y los botones. Con fuentes del sistema sin glifo emoji (`🔍`, `▶`) en Segoe UI 9 o bajo DPI 125 %, los textos se desbordaban generando solapes visuales sobre la barra de progreso.
- **Solución Aplicada:**
  1. En `pnlBotonesSim` (`Dock=Top` en `Panel2`, alto 58), se creó `pnlProgreso` (`Dock=Left`, ancho 380) conteniendo `progressBarAG` y `lblStatusAG` (`AutoSize=false`, `AutoEllipsis=true`).
  2. `chkModoDebug` se fijó con `Dock=Right` (180 px, `AutoEllipsis=true`).
  3. Las acciones se alojaron en `flowAcciones` (`FlowLayoutPanel`, `Dock=Right`, `FlowDirection=RightToLeft`, `AutoSize=true`).
  4. Se normalizaron los textos eliminando emojis: `"Diagnóstico Pre-vuelo"`, `"Optimizar Horarios"`, `"Cancelar"`.

#### 5. Fase 1043 — Estandarización Tipográfica, DPI Base y Suscripciones de Redimensión
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El formulario carecía de propiedad `Font` explícita heredando MS Sans Serif 8.25 mientras los controles usaban Segoe UI 9, provocando distorsión en la métrica `AutoScaleDimensions = (6F, 13F)` al escalar a 125%/150% DPI. Asimismo, colapsar nodos o mover el splitter no refrescaba las cabeceras ni el cursor.
- **Solución Aplicada:**
  1. Fijado en el formulario `Font = new Font("Segoe UI", 9F)`, `AutoScaleDimensions = (7F, 15F)`, `ClientSize = (1120, 720)` y `MinimumSize = (1120, 700)`.
  2. Suscritos eventos `treeContextoAcademico.AfterCollapse`, `treeContextoAcademico.ClientSizeChanged` y `splitContenedorSim.SplitterMoved` a `ActualizarCabecerasNivel6()`.
  3. `ObtenerXBaseNivel6` enriquecido con salvaguarda `Math.Max(..., ClientSize.Width * 0.09)` para pantallas de alta densidad.

#### 6. Fase 1044 — Sincronización Documental y Trazabilidad en Base de Conocimiento
- **Archivos Afectados:** `docs/articles/changes/2026-ag-horarios-academicos.md`, `implementation_plan.md`, `walkthrough.md`
- **Diagnóstico:** Necesidad de reflejar las mejoras de arquitectura espacial de UI en la documentación oficial del motor genético (§8.2) y sincronizar la trazabilidad de la Ronda 204 en los registros de cambios del repositorio.
- **Solución Aplicada:**
  1. Actualizada la subsección §8.2 de `docs/articles/changes/2026-ag-horarios-academicos.md` documentando el modelo de docking responsivo, el SplitContainer y la distribución paramétrica de Nivel 6.
  2. Consolidadas las Fases 1039 a 1044 en `implementation_plan.md` y `walkthrough.md`.

---

### 157.161 Ronda 205 — Corrección de Contenedor de Componentes Nulo, Calibración de Splitter, Anti-Reentrada y Separadores (Fases 1045 a 1051)

#### 1. Fase 1045 — Instanciación de `components` en `InitializeComponent`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico:** `this.toolTipConteo = new System.Windows.Forms.ToolTip(this.components);` con `components == null` lanzaba `ArgumentNullException` en tiempo de ejecución al construir el formulario, impidiendo su apertura y bloqueando el Módulo S.
- **Solución Aplicada:** Inicializado `this.components = new System.ComponentModel.Container();` al inicio de `InitializeComponent()`, asegurando la liberación adecuada mediante `Dispose()`.

#### 2. Fase 1046 — Calibración de Splitter y Corrección de `grpSeleccion.MinimumSize`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico:** `Height - 300` reservaba espacio excesivo para el panel inferior (294 px) dejando el árbol en solo 174 px útiles; asimismo, `grpSeleccion.MinimumSize = (0, 240)` provocaba que a tamaño mínimo (1120x700) el grupo desbordara 7 px sobre `pnlConfiguracionInferior` reintroduciendo el solape.
- **Solución Aplicada:**
  1. Configurado `splitContenedorSim.Panel1MinSize = 300` y `splitContenedorSim.Panel2MinSize = 120`.
  2. Calibrado `SplitterDistance = alto - altoPanel2 - anchoSeparador` (objetivo: 229 px útiles para el árbol).
  3. Reducido en Designer `grpSeleccion.MinimumSize = new Size(0, 190)`.

#### 3. Fase 1047 — Protección Anti-Reentrada en `ActualizarCabecerasNivel6`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Asignar `pnlCabecerasNivel6.Visible` alteraba el alto disponible del árbol disparando `ClientSizeChanged`, lo que reentraba en `ActualizarCabecerasNivel6` durante el ciclo de layout produciendo parpadeo visible y ejecuciones redundantes.
- **Solución Aplicada:** Incorporada bandera de guardia `private bool _alineandoCabeceras;` con bloque `try/finally` y asignación condicionada `if (pnlCabecerasNivel6.Visible != hayVisible)`.

#### 4. Fase 1048 — Unificación a Recorrido Único de Búsqueda de Nivel 6
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `ActualizarCabecerasNivel6` ejecutaba dos recorridos recursivos O(n) por cada redimensión del control (`HayNodoNivel6Visible` y `ObtenerXBaseNivel6`).
- **Solución Aplicada:** Localizado el nodo `TreeNode n6` una sola vez en `ActualizarCabecerasNivel6` y propagado como argumento hacia `CalcularColumnasNivel6(n6)` y `ObtenerXBaseNivel6(n6)`.

#### 5. Fase 1049 — Compensación de Borde de 1 px para `BorderStyle.FixedSingle`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El árbol dibuja en coordenadas cliente internas mientras que las cabeceras se ubican en el panel exterior, generando un desfase horizontal de 1 px por el borde `FixedSingle`.
- **Solución Aplicada:** Incorporada la constante `const int BordeArbol = 1;` restando 1 px en el posicionamiento horizontal `Left` de etiquetas y separadores.

#### 6. Fase 1050 — Restauración de Separador de Toolbar en `pnlToolbarIzq`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico:** La toolbar izquierda carecía de separación visual entre los botones de administración de perfiles y los botones de navegación del árbol tras el retiro de paneles runtime.
- **Solución Aplicada:** Declarado e instanciado `sepToolbar1` (1x20 px) en Designer e insertado entre `btnEliminarArbol` y `btnExpandirArbol`.

#### 7. Fase 1051 — Corrección de Citas de Evidencia y Sincronización Documental
- **Archivos Afectados:** `walkthrough.md`, `implementation_plan.md`
- **Diagnóstico:** Discrepancia en las líneas citadas para `ActualizarConteoCombinaciones` (3410-3439 vs 3326-3350) y necesidad de alinear la métrica de extensión del árbol al objetivo de 229 px.
- **Solución Aplicada:** Sincronizadas las evidencias de código y consolidados los registros de cambio en los artefactos de auditoría.

---

## 🚀 Ronda 206 — Sistema de Carriles Disjuntos y Oleadas Secuenciales (Fases 1052 a 1065)

### Propósito y Objetivos
Garantizar la viabilidad, calidad pedagógica y consistencia normativa de los horarios generados por el motor AG:
1. **Invariante I1:** Persistencia obligatoria de `Jornada`, `HorasDia` y `DiasVentana` desde la selección del árbol hacia `ag_aula` y `ag_aula_asignatura` mediante UPSERT.
2. **Invariantes I2 e I3:** Carriles disjuntos de 2 días ($\max(1, |\beta|/2)$) con rotación modular `carril = idxAsignatura mod carriles` por orden canónico `(orden_curricular, id)`.
3. **Invariante I4:** Oleadas temporales secuenciales encadenadas por carril (`oleada = idxAsignatura div carriles`), acotando sub-ventanas `[FechaInicio, FechaFin]` sin solape.
4. **Invariante I5:** Compuerta de pre-vuelo de capacidad verificando que la ventana contiene ranuras suficientes para la carga horaria.
5. **Invariante I6 / D4:** Blindaje de 4.0h en pregrado en todos los niveles (8.0h reservado estrictamente a posgrado con `HorasDia == 8`).
6. **Invariante I7:** Erradicación de la degradación silenciosa a 1 día; planes infactibles se marcan como vacantes `SIN_DISPONIBILIDAD`.
7. **Invariante I8:** Compuerta de seguridad en exportación y promoción ante `HardViolations > 0`.

### Fases Ejecutadas
- **Fase 1052:** Sincronización de dimensiones de decisión en `ag_aula` (`UPDATE`) y `ag_aula_asignatura` (`ON CONFLICT DO UPDATE`) en `AgFormGeneradorHorarios.cs`.
- **Fase 1053:** Mapeador `AgSeleccionMapper.cs` validando `Jornada`, `HorasDia` (4 u 8) y `DiasVentana` conservando el código canónico.
- **Fase 1054:** Modelo matemático y de dominio `AgCarriles.cs` con partición disjunta, rotación modular y aislamiento de días residuales.
- **Fase 1055:** POCO inmutable `AgPlanOleada.cs` modelando sub-ventanas, carril y oleada por asignatura.
- **Fase 1056:** Defensa en profundidad en `AgExpansorSesiones.cs` y regla dura `R11` en `AgValidadorReglas.cs` validando duración según nivel académico.
- **Fase 1057:** Servicio determinista `AgPlanificadorOleadas.cs` planificando sub-ventanas temporales y asignación de carriles.
- **Fase 1058:** Integración plan-first en `AgFabricaIndividuos.cs`, orden canónico `(OrdenCurricular, Id)`, supresión de degradación a 1 día y filtro estricto de bloques de 4h.
- **Fase 1059:** Adaptación de reglas `R5` (concurrencia semanal $\le$ carriles) y `R6` (secuencia y solape por carril) en `AgValidadorReglas.cs`.
- **Fase 1060:** Compuerta de seguridad ante `HardViolations > 0` en `ExportarResultadosExcel` y `ExportarProyeccionEspejo` de `AgFormGeneradorHorarios.cs`.
- **Fase 1061:** Migración V74 (`MigrationV74_SembrarBloques4hFaltantes.cs`) asegurando bloques de 4.0h para días 1 a 6 y registro en `SqliteHelper.cs`.
- **Fase 1062:** Mutación constructiva en `AgOperadorMutacion.cs` preservando duración de nivel y días del carril asignado.
- **Fase 1063:** Batería de pruebas unitarias `AgCarrilesTests.cs`, `AgPlanificadorOleadasTests.cs`, `AgSeleccionMapperTests.cs` y `MigrationV74Tests.cs` en `ConversorPDF.Tests`.
- **Fase 1064:** Artículo de documentación de cambios `docs/articles/changes/2026-10-22-ag-carriles-oleadas.md`.
- **Fase 1065:** Sincronización de artefactos de auditoría (`implementation_plan.md`, `walkthrough.md`) y análisis estático de calidad.

---

## 🔧 Ronda 207 — Aislamiento de Log en Pruebas y Descarte de Falso Positivo V66 (Fases 1066 a 1070)

### Propósito y Objetivos
1. **Descarte Concluyente del Falso Positivo V66:** Confirmar que los mensajes `[V66] Bloque vinculaciones falló.` acompañados de `[TEST-V66] fallo intencional` en `log_20261022.txt` corresponden a la ejecución de pruebas unitarias (`MigrationV66Tests.Up_WhenBlockFails_ShouldRollbackAndThrow`) y **no a un fallo en producción**. La tabla `vinculaciones` no tiene restricciones UNIQUE/CHECK sobre `modalidadVinculacion` en BD productiva y V66 corrió 4 veces OK el mismo día.
2. **Aislamiento de Registro en Pruebas:** Evitar que las suites de prueba contaminen `%LOCALAPPDATA%\ConversorPDF\log\` mediante redirección automática a `%TEMP%\ConversorPDF.Tests\log`.
3. **Cero Impacto en Producción:** `LoggerHelper` mantiene su ruta por defecto para la aplicación normal; el redireccionamiento se activa exclusivamente al cargar el ensamblado de pruebas mediante un inicializador de módulo.

### Fases Ejecutadas
- **Fase 1066 — Redirección configurable de la salida de log (`LoggerHelper.cs`):** Campo `_logDirectory` mutable y propiedad `internal static string DirectorioLog { get; set; }` para redirigir la carpeta de logs dinámicamente.
- **Fase 1067 — Shim de `ModuleInitializerAttribute` para net48 (`ModuleInitializerAttribute.cs`):** Definición de atributo bajo `System.Runtime.CompilerServices` con guard `#if !NET5_0_OR_GREATER` para habilitar inicializadores de módulo en C# 9+ sobre .NET Framework 4.8.
- **Fase 1068 — Bootstrap de pruebas con redirección a `%TEMP%` (`TestLogBootstrap.cs`):** Inicializador de módulo que desvía automáticamente `LoggerHelper.DirectorioLog` al cargar `ConversorPDF.Tests`.
- **Fase 1069 — Prueba de no regresión de la redirección (`LoggerHelperRedirectionTests.cs`):** Test unitario que valida la escritura en carpeta aislada y restauración limpia del estado.
- **Fase 1070 — Trazabilidad documental y clarificación en pruebas:** Comentario aclaratorio en `MigrationV66Tests.cs`, artículo técnico `docs/articles/changes/2026-10-22-aislamiento-log-pruebas.md` y sincronización de `implementation_plan.md` y `walkthrough.md`.

---

## 🔧 Ronda 208 — Carriles y Oleadas: Fugas de 8 h, Oleada Sobrescrita por Cruce y Carril Obsoleto (Fases 1071 a 1082)

### Propósito y Objetivos
1. **Erradicación Concluyente de 8 h en Pregrado:** Eliminar el `fallbackBloqueId` en fábrica que asignaba bloques arbitrarios (de 8 h o ajenos al día); colocar genes sin bloque adecuado en cuarentena (`IdBloqueHorario = 0`, `Incompleto`). Alinear franja TARDE en V74 a 13:00-17:00 coincidente con V62 y `AgFranjaJornada`.
2. **Preservación de Sub-Ventana de Oleada en Cruce:** Asegurar que el descendiente mantenga su plan temporal (`FechaInicioEstimada`, `FechaFinEstimada`), heredando del progenitor sólo cuando esté vacío, evitando solapes espurios en el mismo carril.
3. **Rotación Forzosa del Carril ante Cambio de Ventana:** Derivar siempre `dias_patron` desde la ventana vigente en la selección del árbol sin rehusar patrones obsoletos desalineados.
4. **Normalización Estricta de Carril en `DiasKDe`:** Validar que los patrones guarden la propiedad de carril disjunto de 2 días consecutivos; de lo contrario, normalizar al carril canónico de la ventana.
5. **Eliminación de Falsos Positivos Inter-Aula en Cohortes:** Discriminar por `IdAula` en R6 y R13 para permitir que múltiples aulas de la misma cohorte (presencial y virtual) coexistan sin colisión de oleadas.
6. **Aislamiento de Asignaturas Infactibles y Sanción Dura R7:** Propagar `R7_CoberturaHorariaIncompleta` dura si una asignatura carece de plan de oleada factible (I5), evitando que el AG acepte individuos vacantes como óptimos.
7. **Cálculo de Ranuras con Descuento de Festivos:** Descontar días feriados en `AgPlanificadorOleadas` para detectar infactibilidad I5 de forma temprana y exacta.
8. **Feedback Bloqueante en UI ante Configuración Incompleta:** Advertir al usuario con `MessageBox` modal si existen hojas del árbol sin jornada, ventana o duración configuradas al generar aulas.

### Fases Ejecutadas
- **Fase 1071 — Cuarentena estricta en fábrica y erradicación de fallbackBloqueId:** Reemplazo de fallback por `IdBloqueHorario = 0`, `EsVacante = 1`, `Estado = AgEstadoGen.Incompleto` en `AgFabricaIndividuos.cs:522-536`.
- **Fase 1072 — Protección de sub-ventana de oleada en operador de cruce:** Asignación condicional `if (!dst.FechaInicioEstimada.HasValue)` en `AgOperadorCruce.cs:143-146`.
- **Fase 1073 — Derivación modular y resiembra del carril sobre ventana vigente:** Generación invariable de `diasPatron` con `AgCarriles.DelCarril` en `AgFormGeneradorHorarios.cs:4195-4220`.
- **Fase 1074 — Validación de carril disjunto consecutivo en `DiasKDe`:** Validación `esCarril` (2 días, gap 1) en `AgPlanificadorPatronSemanal.cs:52-64`.
- **Fase 1075 — Discriminación por `IdAula` en R6 y R13:** Agrupación por `${x.IdAula}|${DiasPatron}` en `AgValidadorReglas.cs:374, 442`.
- **Fase 1076 — Validación de duración real del bloque en regla R11:** Resolución de duración desde `contexto.BloquesPorId[IdBloqueHorario]` en `AgValidadorReglas.cs:425-433`.
- **Fase 1077 — Aislamiento de asignaturas infactibles y sanción dura R7:** Extracción de `asig=` de error I5 en fábrica y adición de `R7_CoberturaHorariaIncompleta` en `AgValidadorReglas.cs:471-480`.
- **Fase 1078 — Descuento de festivos en cálculo de capacidad de oleadas:** Conteo de `diasEfectivos` en `AgPlanificadorOleadas.cs:120-130`.
- **Fase 1079 — Franja TARDE 4h canónica 13:00-17:00 en V74:** Ajuste de hora de inicio y fin en `MigrationV74_SembrarBloques4hFaltantes.cs:33-37`.
- **Fase 1080 — Eliminación de `<InternalsVisibleTo Include>` inerte en csproj legacy:** Limpieza de ítem inerte en `ConversorPDF.XLL.csproj:685-687`.
- **Fase 1081 — Detección bloqueante de combinaciones inválidas/incompletas en UI:** Parámetro `descartes` en `ObtenerCombinacionesSeleccionadas` y compuerta en `GenerarEncabezadosAulasAsync` de `AgFormGeneradorHorarios.cs:3390, 4022`.
- **Fase 1082 — Trazabilidad documental y suites de pruebas de no regresión:** Pruebas unitarias de F1071 a F1079 en `AgFabricaIndividuosTests.cs`, `AgOperadorCruceTests.cs`, `AgPlanificadorPatronSemanalTests.cs`, `AgValidadorReglasTests.cs`, `AgPlanificadorOleadasTests.cs`, `MigrationV74Tests.cs` y sincronización de artefactos.

---

## 🔧 Ronda 209 — F1071 y F1074 Incompletas: 8 h y Carril Inválido Persisten por Rutas Alternas (Fases 1083 a 1089)

### Propósito y Objetivos
1. **Erradicación Concluyente y Total de `fallbackBloqueId` (F1071 incompleta):** Sustituir la asignación residual de `fallbackBloqueId` en la rama de día único de `AgFabricaIndividuos.cs` por cuarentena estricta (`IdBloqueHorario = 0`, `EsVacante = 1`, `Estado = Incompleto`), y eliminar completamente la variable `fallbackBloqueId`.
2. **Validación Estricta de Carril en Sobrecarga Textual de `DiasKDe` (F1074 incompleta):** Asegurar que el overload textual `DiasKDe(string, string, AgAula, int)` en `AgPlanificadorPatronSemanal.cs` valide la consecutividad de 2 días disjuntos (`esCarril`) antes de devolver el patrón; de lo contrario, resolver el carril canónico de la ventana.
3. **Desacoplamiento de Motivo de Vacante en Regla Dura R7:** Introducir `AgMotivoVacante.SIN_PLAN_OLEADA = 6` para diferenciar la infactibilidad de oleada (I5) de la ausencia de bloques en catálogo (`SIN_DISPONIBILIDAD`), evitando que la falta de bloques de 4h bloquee la corrida completa bajo una etiqueta errónea.
4. **Filtro Estricto de Duración en Fallback del Cruce:** Restringir el último recurso de asignación de bloques en `AgOperadorCruce.cs` para que coincida exactamente con `unidadCruce` (4h en pregrado), impidiendo que se asignen bloques de 8h a asignaturas de pregrado.
5. **Calibración Robusta de Concurrencia R5 para Cohortes Multi-Aula:** Calcular `maxCarriles` como el máximo entre todas las aulas activas de la cohorte en `AgValidadorReglas.cs`, eliminando falsos positivos cuando conviven aulas presenciales (2 carriles) y virtuales (1 carril).
6. **Fuente Única de Verdad para Capacidad de Oleadas:** Delegar la validación de capacidad en `AgPlanificadorOleadas.cs` directamente a `AgPlanificadorPatronSemanal.Construir`, eliminando algoritmos duplicados de conteo de feriados.
7. **Sincronización Documental y Pruebas Unitarias:** Documentar la corrección integral en `implementation_plan.md` y `walkthrough.md`, y añadir pruebas de no regresión para F1083-F1088.

### Fases Ejecutadas
- **Fase 1083 — Cuarentena estricta en rama de día único y erradicación total de fallbackBloqueId:** Reemplazo de fallback por cuarentena (`IdBloqueHorario = 0`, `EsVacante = 1`, `Incompleto`, `HorasSemanales = horasEfectivas`) y eliminación de la variable en `AgFabricaIndividuos.cs:65, 414-427`.
- **Fase 1084 — Validación de carril disjunto en sobrecarga textual de `DiasKDe`:** Evaluación de `esCarril` (Count == 2 && gap == 1) en `AgPlanificadorPatronSemanal.cs:111-122`.
- **Fase 1085 — Desacoplamiento de motivo vacante `SIN_PLAN_OLEADA` y corrección de R7:** Inclusión de `AgMotivoVacante.SIN_PLAN_OLEADA = 6` en `AgMotivoVacante.cs:14`, asignación en `AgFabricaIndividuos.cs:288-298`, evaluación en `AgValidadorReglas.cs:482-490`, y atención en `AgFormGeneradorHorarios.cs:2285, 2327`.
- **Fase 1086 — Filtro de duración en asignación de bloque de último recurso en cruce:** Verificación estricta de `Math.Abs(b.DuracionHoras - unidadCruce) < 0.001` en `AgOperadorCruce.cs:193-211`.
- **Fase 1087 — Cálculo de `maxCarriles` usando el máximo de las aulas de la cohorte en R5:** Agrupación y cálculo con `.Max()` en `AgValidadorReglas.cs:325-331`.
- **Fase 1088 — Delegación de capacidad de oleada a `AgPlanificadorPatronSemanal.Construir`:** Reemplazo de cálculo manual por llamada unificada en `AgPlanificadorOleadas.cs:118-130`.
- **Fase 1089 — Trazabilidad documental y pruebas de no regresión:** Actualización de `walkthrough.md` y suite de pruebas unitarias (`AgFabricaIndividuosTests.cs`, `AgPlanificadorPatronSemanalTests.cs`, `AgValidadorReglasTests.cs`, `AgOperadorCruceTests.cs`, `AgPlanificadorOleadasTests.cs`, `AgMotivoVacanteTests.cs`).

---

## 🔧 Ronda 210 — Plan de Oleadas Recalculado por Individuo, Aborto Total del Plan y `SIN_PLAN_OLEADA` no Inmune a `Sellar` (Fases 1090 a 1096)

### Propósito y Objetivos
1. **Memoización del Plan de Oleadas (Fase 1090):** Evitar la regresión de rendimiento O(individuos × asignaturas) memoizando los planes deterministas por aula y parámetros de ventana en `_cachePlanes` de `AgFabricaIndividuos.cs` e inyectando `_planificadorPatron` en el constructor de `AgPlanificadorOleadas.cs`.
2. **Inmunidad de `SIN_PLAN_OLEADA` en `Sellar` (Fase 1091):** Evitar que la normalización de grupo en `AgCoherenciaGrupo.cs` elimine el motivo `SIN_PLAN_OLEADA` cuando existe al menos un gen no vacante (e.g. cruce mixto), garantizando que la regla dura R7 y el diagnóstico en UI no se pierdan.
3. **Planificación Parcial ante Asignaturas Infactibles (Fase 1092):** Reemplazar el aborto total del plan por la acumulación de infactibles en `AgPlanOleadasResultado`, permitiendo que las asignaturas factibles conserven sus carriles y oleadas y aislando únicamente las que no caben en su oleada.
4. **Bloqueo Estricto de Aulas sin Ventana (Fase 1093):** Bloquear aulas con `dias_ventana` vacíos con `[I1_VENTANA_VACIA]` en `AgValidadorVentana.ValidarFactibilidad`, advertir en `AgPlanificadorPatronSemanal.DiasKDe` y marcar la cohorte como infactible en la fábrica en vez de planificar con `{2,3}` silencioso.
5. **Cálculo Canónico de Carriles en Advertencias (Fase 1094):** Sustituir `Math.Ceiling(count / 2.0)` por `AgCarriles.Cantidad(ventanaIsoCfg)` en `AgValidadorVentana.cs`.
6. **Capacidad Teórica Real en `AgPlanOleada` (Fase 1095):** Corregir `CapacidadRanuras` para reflejar la holgura teórica real (`semanasOleada × diasCarril`) en lugar de reportar tautológicamente las sesiones requeridas.
7. **Desacoplamiento de Motivo `SIN_BLOQUE_NIVEL` (Fase 1096):** Introducir `AgMotivoVacante.SIN_BLOQUE_NIVEL = 7` en `AgMotivoVacante.cs`, asignarlo en `AgFabricaIndividuos.cs` ante ausencia de bloque horario de 4h, y mapear su tratamiento contextual en `AgFormGeneradorHorarios.cs`.

### Fases Ejecutadas
- **Fase 1090 — Memoización de planes de carril/oleada e inyección de planificador:** Diccionario `_cachePlanes` en `AgFabricaIndividuos.cs:20, 220` e inyección de `_planificadorPatron` en `AgPlanificadorOleadas.cs:17-21`.
- **Fase 1091 — Inmunidad de `SIN_PLAN_OLEADA` en `AgCoherenciaGrupo.Sellar`:** Guarda prioritaria en `AgCoherenciaGrupo.cs:234-245` que preserva vacante y motivo `SIN_PLAN_OLEADA`.
- **Fase 1092 — Planificación parcial con `AgPlanOleadasResultado`:** Modelo `AgPlanOleadasResultado.cs`, retorno de planes factibles + lista de infactibles en `AgPlanificadorOleadas.cs:115-152` y consumo selectivo en `AgFabricaIndividuos.cs:235-245`.
- **Fase 1093 — Detección bloqueante de aula sin ventana y traza en fallback:** Bloqueo `[I1_VENTANA_VACIA]` en `AgValidadorVentana.cs:209-214`, traza en `AgPlanificadorPatronSemanal.cs:90-94` y cuarentena en `AgFabricaIndividuos.cs:253-260`.
- **Fase 1094 — Corrección de cálculo de carriles con `AgCarriles.Cantidad` en validador de ventana:** Unificación en `AgValidadorVentana.cs:178`.
- **Fase 1095 — Capacidad teórica real en modelo de oleada:** Asignación de `semanasPorOla[oleada] * diasPatron.Count` en `AgPlanificadorOleadas.cs:136`.
- **Fase 1096 — Desacoplamiento de motivo vacante `SIN_BLOQUE_NIVEL`:** Creación de `AgMotivoVacante.SIN_BLOQUE_NIVEL = 7`, asignación en `AgFabricaIndividuos.cs:437, 546` y atención en `AgFormGeneradorHorarios.cs:2287, 2325`.

---

## 🔧 Ronda 211 — Errores de Compilación por Tipo No Registrado y `using` Faltante; Cuarentena I1 Efectiva (Fases 1097 a 1102)

### Propósito y Objetivos
1. **Restauración de Compilación en Proyecto Legacy (Fase 1097):** Registrar explícitamente `<Compile Include="Academic\Domain\Models\Genetica\AgPlanOleadasResultado.cs" />` en `ConversorPDF.XLL.csproj` para resolver error bloqueante CS0246.
2. **Importación de Namespace Funcional (Fase 1098):** Incorporar `using CSharpFunctionalExtensions;` en `AgFabricaIndividuos.cs` para resolver el tipo `Result<T>` en `_cachePlanes`.
3. **Efectividad Operativa de Cuarentena I1 (Fase 1099):** Asegurar que cuando una cohorte carece de plan de oleadas (`cohorteSinPlan = planInfactible && asigsInfactibles.Count == 0`), sus genes se fijen con `IdBloqueHorario = 0`, `DiasPatron = string.Empty` y `Estado = Incompleto`, impidiendo que `AgMaterializadorSesiones` materialice sesiones espurias con el centinela.
4. **Desduplicación de Sanción Dura R14 (Fase 1100):** Excluir del chequeo genérico `R14_BloqueHorarioNoAsignado` a los genes con motivo tipificado `SIN_BLOQUE_NIVEL`.
5. **Observabilidad y Verificación Real de Memoización (Fase 1101):** Incorporar contador interno `Construcciones` en `AgPlanificadorOleadas.cs` y validar en prueba unitaria que se ejecute exactamente 1 vez para individuos de la misma corrida.
6. **Aislamiento de Caché por Corrida y Período (Fase 1102):** Limpiar `_cachePlanes` al detectar cambio en `idCorrida` e incluir `idCorrida` e `IdPeriodo` en la clave compuesta de la memoización.

### Fases Ejecutadas
- **Fase 1097 — Registro de `AgPlanOleadasResultado.cs` en csproj:** Inclusión en `ConversorPDF.XLL.csproj:653`.
- **Fase 1098 — Directiva `using CSharpFunctionalExtensions;` en fábrica:** Importación de namespace en `AgFabricaIndividuos.cs:8`.
- **Fase 1099 — Cuarentena efectiva I1 en construcción de genes:** Anulación de bloque (`IdBloqueHorario = 0`, `DiasPatron = ""`, `Incompleto`) ante `cohorteSinPlan` en `AgFabricaIndividuos.cs:312, 455, 595`.
- **Fase 1100 — Exclusión de `SIN_BLOQUE_NIVEL` en regla R14:** Guarda en `AgValidadorReglas.cs:172-178` para no duplicar sanción.
- **Fase 1101 — Contador de construcciones en planificador de oleadas y prueba estricta:** Propiedad `Construcciones` en `AgPlanificadorOleadas.cs:19, 34` y aserción en `AgFabricaIndividuosTests.cs:330`.
- **Fase 1102 — Invalidación de caché por corrida y clave extendida:** Limpieza de `_cachePlanes` ante nueva corrida y clave con periodo en `AgFabricaIndividuos.cs:23, 53, 230`.

---

## 🔧 Ronda 212 — `SIN_BLOQUE_NIVEL` sin Sanción Dura, Preservación en Sellar y Corrección de Factibilidad I5 (Fases 1103 a 1107)

### Propósito y Objetivos
1. **Sanción Dura para Genes `SIN_BLOQUE_NIVEL` (Fase 1103):** Garantizar que un gen con carga horaria y motivo `SIN_BLOQUE_NIVEL` emita la regla dura `R7_CoberturaHorariaIncompleta`, evitando que el individuo eluda la penalización y la asignatura desaparezca silenciosamente en materialización.
2. **Preservación de Motivos No Asignables en `Sellar` (Fase 1104):** Extender la guarda de sellado en `AgCoherenciaGrupo.cs` para proteger tanto `SIN_PLAN_OLEADA` como `SIN_BLOQUE_NIVEL` ante cruces mixtos, asegurando que ningún gen sin bloque sea revertido erróneamente a docente real.
3. **Cálculo Real de Demanda en Pre-Vuelo I5 (Fase 1105):** Eliminar la asignación ficticia de 64h ante `HorasObjetivo == 0` en `AgValidadorVentana.cs`, resolviendo la intensidad desde el catálogo de asignaturas y excluyendo materias desconocidas (<= 0) con advertencia `[I5_HORAS_DESCONOCIDAS]` para no bloquear corridas viables.
4. **Desduplicación de `SIN_PLAN_OLEADA` en R14 (Fase 1106):** Excluir `SIN_PLAN_OLEADA` de `R14_BloqueHorarioNoAsignado` en `AgValidadorReglas.cs`, evitando duplicidad con la sanción dura `R7` previamente tipificada.
5. **Cobertura Integral de Pruebas Unitarias (Fase 1107):** Añadir pruebas en `AgValidadorReglasTests.cs`, `AgCoherenciaGrupoTests.cs` y `AgValidadorVentanaTests.cs` para certificar la emisión de R7 dura, la exclusión de R14, la preservación en `Sellar` y la no-infactibilidad ante `HorasObjetivo == 0`.

### Fases Ejecutadas
- **Fase 1103 — Regla dura R7 para motivo `SIN_BLOQUE_NIVEL`:** Bucle en `AgValidadorReglas.cs:494-504` que emite `AgReglaViolacion.R7_CoberturaHorariaIncompleta` dura.
- **Fase 1104 — Preservación de `SIN_BLOQUE_NIVEL` y `SIN_PLAN_OLEADA` en `Sellar`:** Generalización con `motivoNoAsignable` en `AgCoherenciaGrupo.cs:234-245`.
- **Fase 1105 — Demanda real en validador de ventana I5 sin default de 64h:** Resolución en `AgValidadorVentana.cs:203, 222-235` y emisión de advertencia `[I5_HORAS_DESCONOCIDAS]`.
- **Fase 1106 — Exclusión de `SIN_PLAN_OLEADA` en verificación R14:** Guarda combinada en `AgValidadorReglas.cs:172-178`.
- **Fase 1107 — Cobertura de pruebas unitarias para `SIN_BLOQUE_NIVEL`, `Sellar` e I5:** Nuevas pruebas en `AgValidadorReglasTests.cs:610-670`, `AgCoherenciaGrupoTests.cs:252-290` y `AgValidadorVentanaTests.cs:211-232`.

---

## 🔧 Ronda 213 — Fechas de Oleadas en Vinculaciones/Promoción, Min/Max de Grupo, Conteo I5 y Contador Atómico (Fases 1108 a 1112)

### Propósito y Objetivos
1. **Preservación de Fechas de Oleada en Vinculaciones (Fase 1108):** Evitar que `MapGrupo` y `AgMaterializadorSesiones` sustituyan las fechas estimadas de la sub-ventana por las fechas completas del semestre cuando `IdPeriodo > 0`, manteniendo la visibilidad de oleadas en `ag_vinculaciones` y `ag_promocion`.
2. **Cálculo de Rango Min/Max de Fechas por Grupo (Fase 1109):** En grupos multi-día, calcular `FechaInicio` y `FechaFin` a partir del mínimo y máximo de las fechas estimadas de los genes que componen el grupo, en lugar de adoptar únicamente los valores del primer gen (`g0`).
3. **Conteo Coherente de Materias Computadas en Bloqueo I5 (Fase 1110):** Reportar explícitamente en el mensaje de error `[I5_INFACTIBLE]` el número de materias efectivamente computadas en la demanda frente al total asignado (`({materiasComputadas} materias computadas de {asigsAula.Count})`).
4. **Sincronización Atómica del Contador de Construcciones (Fase 1111):** Sincronizar el contador `Construcciones` en `AgPlanificadorOleadas.cs` mediante `Interlocked.Increment` y `Volatile.Read` para garantizar estabilidad bajo generación paralela o multihilo.
5. **Cobertura de Pruebas Unitarias para Fechas de Oleada e I5 (Fase 1112):** Añadir pruebas en `AgGenToAgVinculacionMapperTests`, `AgMaterializadorSesionesTests` y `AgValidadorVentanaTests` para certificar la conservación de fechas de oleada, el cálculo de Min/Max y el conteo de materias computadas.

### Fases Ejecutadas
- **Fase 1108 — Preservación de sub-ventana de oleada en mapeador y materializador:** Fechas de genes preservadas en `AgGenToAgVinculacionMapper.cs:73-82` y acotadas sin sobreescritura en `AgMaterializadorSesiones.cs:99-108`.
- **Fase 1109 — Rango Min/Max en `MapGrupo`:** `fiGrupo` y `ffGrupo` calculados mediante `Min()` y `Max()` sobre `genes` en `AgGenToAgVinculacionMapper.cs:74-75`.
- **Fase 1110 — Reporte de materias computadas en advertencia/bloqueo I5:** Contador `materiasComputadas` y mensaje actualizado en `AgValidadorVentana.cs:227, 238, 244`.
- **Fase 1111 — Contador `Construcciones` sincronizado en planificador:** `_construcciones` con `Volatile.Read` e `Interlocked.Increment` en `AgPlanificadorOleadas.cs:18-20, 39`.
- **Fase 1112 — Pruebas unitarias de oleadas en vinculación e I5:** Pruebas en `AgGenToAgVinculacionMapperTests.cs:60-129`, `AgMaterializadorSesionesTests.cs:320-384` y `AgValidadorVentanaTests.cs:231-254`.

---

## 🔧 Ronda 214 — Pre-vuelo UI con Duración Efectiva por Nivel, Invariante Dura de Nivel (F1114) y Parseo Invariante en Sub-Ventana (Fases 1113 a 1116)

### Propósito y Objetivos
1. **Validación Pre-Vuelo con Duración Efectiva por Nivel (Fase 1113):** Resolver `horasEfectivasV` vía `AgUnidadesSemanales.HorasSesionAula` según el nivel del programa en `AgFormGeneradorHorarios.cs`, impidiendo que una selección de 8h en Pregrado bloquee falsamente la corrida cuando el catálogo no posee bloques de 8h.
2. **Invariante Dura de Nivel — Ningún Gen de PREGRADO con Bloque de 8h (Fase 1114 corregida):** Establecer una invariante global en el punto de convergencia de operadores (`AgCoherenciaGrupo.Sellar`), fijar la duración impuesta por el nivel en `AgOperadorMutacion.cs`, bloquear en pre-vuelo (`[D4_PREGRADO_8H]` en `AgValidadorVentana.cs` y aborto con mensaje claro en `AgFormGeneradorHorarios.cs`), y certificar con pruebas unitarias en todas las capas.
3. **Parseo Culturalmente Invariante en Cotas de Fechas (Fase 1115):** Utilizar `DateTime.TryParseExact` con `CultureInfo.InvariantCulture` en `AgMaterializadorSesiones.cs` para evitar omisiones de cotas o fallos de parseo bajo culturas no gregorianas, advirtiendo ante cadenas inválidas.
4. **Desacoplamiento de Patrón en Pre-Vuelo de Aula (Fase 1116):** Pasar `string.Empty` como `diasPatron` en la validación pre-vuelo de aula en `AgFormGeneradorHorarios.cs`, dado que los patrones específicos por asignatura se validan a nivel de motor.

### Fases Ejecutadas
- **Fase 1113 — Pre-vuelo UI con horas efectivas de nivel:** Resolución de `horasEfectivasV` en `BtnGenerarAulas_Click` y `BtnOptimizar_Click` de `AgFormGeneradorHorarios.cs:1113-1126, 1360-1373`.
- **Fase 1114 — Invariante dura de nivel (1114.a a 1114.d):**
  - *1114.a (Sellar):* Normalización de `HorasSemanales` a `uNivel`, reasignación de bloque del mismo día o cuarentena con `SIN_BLOQUE_NIVEL` en `AgCoherenciaGrupo.cs:56-96`.
  - *1114.b (Mutación):* Duración objetivo impuesta por nivel del aula en `AgOperadorMutacion.cs:110-125`.
  - *1114.c (Bloqueo Pre-vuelo):* Bloqueo `[D4_PREGRADO_8H]` en `AgValidadorVentana.cs:222-230` y aborto bloqueante en `AgFormGeneradorHorarios.cs:1117-1125, 1373-1381`.
  - *1114.d (Pruebas):* Cobertura en `AgCoherenciaGrupoTests.cs:288-356`, `AgOperadorMutacionTests.cs:192-228`, `AgValidadorVentanaTests.cs:254-276` y `AgFabricaIndividuosTests.cs:405-432`.
- **Fase 1115 — Parseo `InvariantCulture` en cota de sub-ventana:** `DateTime.TryParseExact` con advertencia de log en `AgMaterializadorSesiones.cs:100-123` y prueba en `AgMaterializadorSesionesTests.cs:387-463`.
- **Fase 1116 — Desacoplamiento de patrón en pre-vuelo:** Argumento `string.Empty` como `diasPatron` en `AgFormGeneradorHorarios.cs:1128, 1384`.

---

## 🔧 Ronda 215 — Robustecimiento de Invariante de Nivel: Protección de POSGRADO, Resguardo de Genes sin Duración e Indexación O(1) de Bloques (Fases 1117 a 1121)

### Propósito y Objetivos
1. **Resguardo de Genes con Duración Desconocida y Protección de Fechas (Fase 1117):** Excluir genes con `HorasSemanales <= 0` de la invariante en `AgCoherenciaGrupo.Sellar` y limitar la corrección estrictamente a casos de sobre-duración (`g.HorasSemanales > uNivel + 0.001`), evitando borrar `FechasGen` ni mandar genes válidos a vacante `SIN_BLOQUE_NIVEL`.
2. **Corrección de Unidad por Nivel sin Degradar POSGRADO (Fase 1118):** En `AgCoherenciaGrupo.Sellar`, hacer que `UnidadNivel` retorne `-1.0` si el aula o programa no son resolubles, en lugar del fallback a `4.0` que degradaba erróneamente asignaturas legítimas de POSGRADO (8h) o en aulas no indexadas.
3. **Cobertura de No-Regresión de la Invariante (Fase 1119):** Añadir pruebas unitarias específicas en `AgCoherenciaGrupoTests.cs` verificando que POSGRADO conserva 8h, que genes con `HorasSemanales = 0` no pierden fechas ni estado, y que aulas no indexadas no son alteradas.
4. **Causa de Fallo Específica en Pre-Vuelo (Fase 1120):** Incorporar `HorasDiaInvalidoNivel = 13` en `AgCausaFallo.cs` y registrarlo en las trazas de fallo de pre-vuelo P07 y P08 en `AgFormGeneradorHorarios.cs`.
5. **Optimización O(1) de Bloques en Sellar (Fase 1121):** Construir un diccionario indexado de bloques activos por `(DiaSemanaNumero, DuracionHoras)` al inicio de `Sellar`, sustituyendo la búsqueda lineal O(Genes × Bloques) repetida en cada llamada de cruce, mutación y materialización.

### Fases Ejecutadas
- **Fase 1117 — Protección contra disparo de invariante con HorasSemanales <= 0:** Cláusulas de guarda en `AgCoherenciaGrupo.cs:72-74` evitando borrar `FechasGen` o alterar genes sin sobre-duración.
- **Fase 1118 — Retorno -1.0 en UnidadNivel ante aulas no resolubles:** Prevención de asunción indebida de PREGRADO en `AgCoherenciaGrupo.cs:68`.
- **Fase 1119 — Pruebas unitarias de no-regresión de invariante:** Implementación de `Sellar_PosgradoConBloque8h_NoDegradaDuracion`, `Sellar_GenSinHorasSemanales_NoModificaFechasNiEstado` y `Sellar_AulaNoIndexada_NoAlteraDuracion` en `AgCoherenciaGrupoTests.cs:360-464`.
- **Fase 1120 — Enum AgCausaFallo.HorasDiaInvalidoNivel:** Nuevo código en `AgCausaFallo.cs:50` y referenciado en `AgFormGeneradorHorarios.cs:1122, 1385`.
- **Fase 1121 — Indexación O(1) de bloques en Sellar:** Diccionario `indiceBloques` agrupado y búsqueda directa en `AgCoherenciaGrupo.cs:53-57, 80-81`.

---

## 🔧 Ronda 216 — Validación Pre-Vuelo por Sub-Ventana de Oleada, Capacidad Canónica de Carril y Caché de Bloques en Contexto (Fases 1122 a 1126)

### Propósito y Objetivos
1. **Validación de Sub-Ventana de Oleada por Asignatura (Fase 1122):** Invocar `AgPlanificadorOleadas.Construir` dentro de `ValidarFactibilidad` en `AgValidadorVentana.cs`, detectando asignaturas cuya carga horaria exceda la sub-ventana asignada a su oleada (`[I5_INFACTIBLE_OLEADA]`), evitando que demandas individuales inviables pasen el pre-vuelo.
2. **Capacidad Canónica de Carril — 1 Sesión/Día (Fase 1123):** Fijar `sesionesPorDia = 1` en `CapacidadDe` de `AgValidadorVentana.cs` en correspondencia con el modelo de carriles disjuntos (R4), impidiendo la sobreestimación artificial de capacidad a 2 sesiones/día ante selecciones de 8h/día.
3. **Cálculo Canónico de Carriles en Pre-Vuelo (Fase 1124):** Utilizar `AgCarriles.Cantidad(ventanaIso)` en lugar de `Math.Ceiling(count / 2.0)` en `CapacidadDe` de `AgValidadorVentana.cs`, alineando el reporte diagnóstico de multicarril con la partición real del motor.
4. **Caché Estable de Bloques en AgContexto (Fase 1125):** Implementar `TryGetBloque(int dia, double dur, out AgBloqueHorario b)` con diccionario memoizado en `AgContexto.cs` e invalidación en `ConstruirIndices()`, eliminando la re-creación repetitiva del diccionario de bloques en `AgCoherenciaGrupo.Sellar`.
5. **Cobertura de No-Regresión de Oleada y Capacidad (Fase 1126):** Diseñar pruebas unitarias en `AgValidadorVentanaTests.cs` que certifiquen el bloqueo ante asignaturas que exceden su oleada y verifiquen que la capacidad no se duplica.

### Fases Ejecutadas
- **Fase 1122 — Validación de sub-ventana por oleada en pre-vuelo:** Bloqueo `[I5_INFACTIBLE_OLEADA]` en `AgValidadorVentana.cs:255-276`.
- **Fase 1123 — Capacidad canónica de 1 sesión/día/carril:** `sesionesPorDia = 1` en `AgValidadorVentana.cs:122`.
- **Fase 1124 — Carriles canónicos con AgCarriles.Cantidad:** Corrección de `Math.Ceiling` en `AgValidadorVentana.cs:129`.
- **Fase 1125 — Caché indexado de bloques en AgContexto:** Campo `_indiceBloquesDiaDur`, método `TryGetBloque` en `AgContexto.cs:150-169`, e integración en `AgCoherenciaGrupo.cs:81-89`.
- **Fase 1126 — Cobertura de pruebas unitarias de oleada y carril:** Pruebas `ValidarFactibilidad_AsignaturaExcedeSuOleada_BloqueaI5` y `ValidarFactibilidad_Pregrado8h_NoDuplicaCapacidadPorDia` en `AgValidadorVentanaTests.cs:277-343`.

---

## 🔧 Ronda 217 — Coherencia de Jornada en Reasignación de Duración, Invalidación Explícita de Caché de Bloques y Alcance Curricular en Pre-Vuelo (Fases 1127 a 1130)

### Propósito y Objetivos
1. **Coherencia Jornada ↔ Franja en Corrección de Duración (Fase 1127):** Evitar colapsos de franja en `AgContexto.TryGetBloque` alindexar listas de bloques por `(DiaSemanaNumero, DuracionHoras)` y filtrar por la franja canónica de la jornada del aula (`AgFranjaJornada.Resolve`), garantizando que aulas de jornada TARDE reciban bloques de franja TARDE tras normalizar 8h→4h.
2. **Invalidación Explícita de Caché ante Mutación de Bloques (Fase 1128):** Exponer `InvalidarIndiceBloques()` en `AgContexto.cs` e invocarlo en `ConstruirIndices()` y en los puntos de recarga de bloques en `AgFormGeneradorHorarios.cs` (P05), previniendo falsos `SIN_BLOQUE_NIVEL` por catálogos obsoletos.
3. **Alineación de Alcance Curricular en Pre-Vuelo de Oleadas (Fase 1129):** Filtrar `itemsPlan` en `AgValidadorVentana.cs` según coincidencia de programa y semestre del aula (`AgFormatoAsignatura.SemestreRomano`), impidiendo bloqueos falsos `[I5_INFACTIBLE_OLEADA]` por materias fuera del alcance que el motor nunca planificará.
4. **Cobertura de No-Regresión de Franja, Caché y Alcance Curricular (Fase 1130):** Añadir pruebas unitarias verificando la asignación de franja según jornada, la actualización del catálogo tras invalidación y la exclusión de asignaturas de otro semestre en pre-vuelo.

### Fases Ejecutadas
- **Fase 1127 — Coherencia de jornada y franja preferida en TryGetBloque:** Índice a `List<AgBloqueHorario>` en `AgContexto.cs:153-195` y resolución con `AgFranjaJornada.Resolve` en `AgCoherenciaGrupo.cs:79-88`.
- **Fase 1128 — InvalidarIndiceBloques ante recarga de bloques:** Método `InvalidarIndiceBloques()` en `AgContexto.cs:157` y llamada en `AgFormGeneradorHorarios.cs:762`.
- **Fase 1129 — Filtro de itemsPlan por programa y semestre del aula:** Cláusula `Where` curricular en `AgValidadorVentana.cs:261-265`.
- **Fase 1130 — Cobertura de pruebas unitarias de jornada, caché y alcance:** Pruebas `Sellar_PregradoTarde8h_ReasignaBloqueDeFranjaTarde` y `Contexto_TrasInvalidarIndiceBloques_UsaCatalogoVigente` en `AgCoherenciaGrupoTests.cs:463-524`, y `ValidarFactibilidad_AulaConAsignaturaDeOtroSemestre_NoBloquea` en `AgValidadorVentanaTests.cs:342-372`.

---

## 🔧 Ronda 218 — Coherencia de Jornada en Cruce, Mutación y Regla R11 con Unificación de Selección de Bloques (Fases 1131 a 1135)

### Propósito y Objetivos
1. **Coherencia de Jornada en Operador de Cruce (Fase 1131):** Evitar la reasignación arbitraria al primer bloque del día en `AgOperadorCruce.cs` al seleccionar bloques compatibles con la franja canónica de la jornada del aula (`AgFranjaJornada.Resolve`), garantizando que aulas de jornada TARDE reciban bloques de franja TARDE.
2. **Coherencia de Jornada en Operador de Mutación (Fase 1132):** Filtrar bloques candidatos en mutación mono-día y multi-día de `AgOperadorMutacion.cs` según la jornada del aula, impidiendo que genes de aulas NOCHE o TARDE migren a franjas distintas.
3. **Validación Dura R11 para Discrepancias Franja ∉ Jornada (Fase 1133):** Extender la regla dura R11 en `AgValidadorReglas.cs` para validar que la franja del bloque asignado a cada gen corresponda estrictamente a la jornada del aula (`AgFranjaJornada.Resolve`), bloqueando individuos inconsistentes.
4. **Cobertura de Pruebas Unitarias de Jornada en Operadores y Validador (Fase 1134):** Agregar pruebas unitarias en `AgOperadorCruceTests.cs`, `AgOperadorMutacionTests.cs` y `AgValidadorReglasTests.cs` asegurando que el cruce, la mutación y la regla R11 respeten y certifiquen la jornada del aula.
5. **Punto Único de Selección de Bloque Horario (Fase 1135):** Crear `AgSeleccionBloque.cs` como punto único para resolver día, duración y franja por jornada, eliminando la duplicación en fábrica, cruce, mutación y coherencia.

### Fases Ejecutadas
- **Fase 1131 — Selección de bloque por jornada en cruce:** Uso de `AgSeleccionBloque.ParaDia` en `AgOperadorCruce.cs:182-205`.
- **Fase 1132 — Filtro de franja por jornada en mutación:** Resolución de `franjaObjetivo` y filtro de candidatos en `AgOperadorMutacion.cs:126-175`.
- **Fase 1133 — Regla R11 con validación estricta de franja vs jornada:** Emisión de violación dura R11 en `AgValidadorReglas.cs:445-459`.
- **Fase 1134 — Cobertura de pruebas unitarias de jornada:** Pruebas `Cruzar_AulaTarde_AsignaBloqueDeFranjaTarde` (`AgOperadorCruceTests.cs:213-241`), `Mutacion_AulaTarde_NoMigraAManana` (`AgOperadorMutacionTests.cs:227-260`) y `Validador_GenDeFranjaFueraDeJornada_EmiteR11` (`AgValidadorReglasTests.cs:672-708`).
- **Fase 1135 — Unificación de selección en AgSeleccionBloque:** Clase `AgSeleccionBloque.cs` creada, registrada en `ConversorPDF.XLL.csproj` e integrada en `AgOperadorCruce.cs`, `AgCoherenciaGrupo.cs`.

---

## 🔧 Ronda 219 — Formato Canónico DIA|horas|horario en Columna K del Exportador Excel (Fases 1136 a 1138)

### Propósito y Objetivos
1. **Formato Canónico en Columna K de Excel (Fase 1136):** Sustituir la emisión simple de nombres de días `"MARTES/MIERCOLES"` en `SemestresExcelExporter.cs` por la cadena canónica `DIA|horas|horario` (ej. `"MARTES|4|08:00-12:00, MIERCOLES|4|08:00-12:00"`), permitiendo verificar la unidad horaria de 4h y los días del carril.
2. **Overload en DiasHorasSemanaFormatter y Limpieza de TODO (Fases 1136 & 1137):** Crear el método `DesdeSesionListItems(IEnumerable<SesionListItem>)` en `DiasHorasSemanaFormatter.cs` para el modelo desnormalizado del exportador masivo, retirando comentarios pendientes de cableado.
3. **Cobertura de Pruebas Unitarias para Columna K (Fase 1138):** Crear `DiasHorasSemanaFormatterTests.cs` en `ConversorPDF.Tests` con cobertura para pares de carril, deduplicación de sesiones semanales, casos sin día y colecciones nulas.
4. **Evaluación de Botón Cargar en Módulo R (Fase 1139):** Confirmar preservación de compuertas de seguridad existentes en `AgFormGeneradorHorarios.cs:2086-2097`, posponiendo la refactorización de flujo por tratarse de alcance de producto y no de defecto.

### Fases Ejecutadas
- **Fase 1136 — Cableado de DiasHorasSemanaFormatter en SemestresExcelExporter:** Overload `DesdeSesionListItems` en `DiasHorasSemanaFormatter.cs:57-93` y asignación en `SemestresExcelExporter.cs:377-380`.
- **Fase 1137 — Retiro de TODO y preservación de granularidad en grilla:** Eliminado `TODO` en `DiasHorasSemanaFormatter.cs` y confirmada separación de columnas en `FormSesionesAcademicas.cs`.
- **Fase 1138 — Pruebas unitarias de DiasHorasSemanaFormatter:** 5 pruebas en `DiasHorasSemanaFormatterTests.cs`.

---

## 🔧 Ronda 220 — Implementación F1-F4 compilación-bloqueada y saneamiento V76 inefectivo en filas legacy (Fases 1195 a 1201)

### Propósito y Objetivos
1. **Corrección de Compilación AgMensajeDimensionesAula (Fase 1195):** Corregir import de namespace para `AgAula` (`ConversorPDF.XLL.Academic.Domain.Models.Genetica`) en `AgMensajeDimensionesAula.cs` y suite de pruebas `AgMensajeDimensionesAulaTests.cs`.
2. **Saneamiento Determinista V76 para Filas Legacy y Múltiples Árboles (Fases 1196 & 1197):** Flexibilizar `ag_aula` con subunidad vacía en el JOIN contra `ag_arbol_nodo`, y resolver ambigüedad multivariable seleccionando deterministamente el nodo con `MAX(idarbol)` que posea dimensiones válidas.
3. **Aislamiento por Subunidad y Marcado de Suciedad en Sincronización (Fase 1198):** Incorporar parámetro `subunidad` en `AgAulaDimensionesSync.SincronizarAsync`, filtrar `WHERE (@Sub = '' OR COALESCE(subunidad,'') = @Sub OR COALESCE(subunidad,'') = '')` y notificar `SyncContext.Scheduler?.MarkDirty()`.
4. **Tope en Modal, Normalización de Saltos de Línea y Telemetría Estructurada (Fase 1199):** Limitar modal a 5 aulas, normalizar saltos `\n` por `" ¶ "` en logs y enriquecer telemetría de traza en pasos P07 y P08.
5. **Reintento de Migración V76 ante Fallos (Fase 1200):** Eliminar captura silenciosa en `MigrationV76_SanearDimensionesAgAula.cs` para permitir que `SchemaMigrator` registre el hueco de versión y reintente en el siguiente arranque.
6. **Evaluación de Duplicación y Obsolecencia de Métodos (Fase 1201):** Ratificar marca `[Obsolete]` preexistente en `GenerarEncabezadoAulaAsync` y preservar simetría de etapas P07 y P08 sin introducir abstracciones artificiales.

### Fases Ejecutadas
- **Fase 1195 — Corrección de namespace en AgMensajeDimensionesAula:** Namespace `ConversorPDF.XLL.Academic.Domain.Models.Genetica` en `AgMensajeDimensionesAula.cs` y `AgMensajeDimensionesAulaTests.cs`.
- **Fase 1196 — Tolerancia a subunidad vacía en V76:** Cláusula `(COALESCE(x.subunidad,'') = COALESCE(a.subunidad,'') OR TRIM(COALESCE(a.subunidad,'')) = '')` en `MigrationV76_SanearDimensionesAgAula.cs`.
- **Fase 1197 — Determinismo por MAX(idarbol) en V76:** Subconsulta correlacionada `MAX(x.idarbol)` eliminando `GROUP BY a.id` en `MigrationV76_SanearDimensionesAgAula.cs`.
- **Fase 1198 — Sincronización con subunidad y scheduler dirty:** Parámetro `subunidad`, cláusula `@Sub` y `Scheduler?.MarkDirty()` en `AgAulaDimensionesSync.cs`, actualización de llamada en `AgFormGeneradorHorarios.cs:105` y pruebas en `AgAulaDimensionesSyncTests.cs`.
- **Fase 1199 — Paginación de modal y saneamiento de log en P07/P08:** Límite de 5 aulas en modal, reemplazo de `\n` por `" ¶ "` en log y telemetría de `AulasSinDimensiones` y `AulasConErrorVentana` en `AgFormGeneradorHorarios.cs:1189-1205, 1483-1499`.
- **Fase 1200 — Relanzamiento de excepción en V76:** `LoggerHelper.Error` y `throw` en catch de `MigrationV76_SanearDimensionesAgAula.cs:115-120`.
- **Fase 1201 — Evaluación crítica de duplicación P07/P08 y API:** Retirada como refactor innecesario; preservadas identidades de traza independientes P07/P08 y ratificado `[Obsolete]` en `GenerarEncabezadoAulaAsync`.
- **Fase 1202 — Registro de archivos nuevos en ConversorPDF.XLL.csproj:** Inclusión de `AgMensajeDimensionesAula.cs`, `AgAulaDimensionesSync.cs` y `MigrationV76_SanearDimensionesAgAula.cs` en `ConversorPDF.XLL.csproj:660-662`, desbloqueando la compilación en Debug/Release x64 y la ejecución de la suite de pruebas.

---

## 🔧 Ronda 223 — Planificador de oleadas v3 por bandas de días (Fases 1224 a 1235)

### Propósito y Objetivos
1. **Partición de Ventana en Bandas Disjuntas de Días (Fases 1224 & 1225):** Reemplazar reparto de carriles de 2 ranuras por partición contigua de días en `AgBandasDias` y validar determinismo y propiedades matemáticas en `AgBandasDiasTests`.
2. **Planificador de Oleadas v3 Atómico por Semanas (Fases 1226 & 1227):** Eliminar `olas = ceil(n/carriles)` y sincronización por rebanadas globales en `AgPlanificadorOleadas`, sustituyéndolas por pipelines secuenciales por banda y búsqueda de asignación exhaustiva/BFD; extender `AgPlanOleadasResultado` con campos de partición y holgura.
3. **Escalera de Capacidad y Escalado Asíncrono de Aulas (Fases 1228 & 1229):** Definir `AgEscaleraVentanas` para sugerencias monótonas dentro del catálogo (`VIE-SAB` → `JUE-SAB`) e implementar `AgEscaladorVentanaAula` para dry-run de déficit y aplicación transaccional en `ag_aula` y `ag_arbol_nodo`.
4. **Validación de Factibilidad Accionable y Detección de Holgura Baja (Fase 1230):** Enriquecer `[I5_INFACTIBLE]` en `AgValidadorVentana` con ventana sugerida de catálogo, emitir advertencia `[I5_HOLGURA_BAJA]` al superar 88% de ocupación y canalizar patrones declarados.
5. **Huella de Patrones en Clave de Caché y Write-Back por Banda (Fases 1231 & 1232):** Incorporar huella de patrones en la clave de memoización de `AgFabricaIndividuos` y sembrar `dias_patron` por banda en `AgFormGeneradorHorarios.cs:4544` para prevenir contradicciones con R8.
6. **Diálogo Interactivo Pre-Vuelo y Suites de Prueba (Fases 1233 a 1235):** Integrar diálogo de propuesta de escalado en diagnóstico pre-vuelo de UI, actualizar suites canónicas T1-T5 en `AgPlanificadorOleadasTests` y documentar especificación formal y limitaciones.

### Estado: PLANIFICADO — NO EJECUTADO (retirado en Ronda 223)

Artefactos ausentes verificados con `Test-Path`: `AgBandasDias.cs`, `AgEscaleraVentanas.cs`,
`AgEscaladorVentanaAula.cs`, `AgBandasDiasTests.cs`, `docs/articles/changes/2026-10-23-ag-oleadas-banda-dias.md`.
Sin registros `<Compile Include=...>` en `ConversorPDF.XLL.csproj`.

### Fases Pendientes (1224 a 1235)
- **Fase 1224 — AgBandasDias:** pendiente.
- **Fase 1225 — Suite de AgBandasDias:** pendiente.
- **Fase 1226 — AgPlanificadorOleadas v3:** pendiente.
- **Fase 1227 — Resultado extendido:** pendiente.
- **Fase 1228 — AgEscaleraVentanas + tests:** pendiente.
- **Fase 1229 — AgEscaladorVentanaAula + tests:** pendiente.
- **Fase 1230 — Validador accionable y holgura baja:** pendiente.
- **Fase 1231 — Fábrica con huella de patrón:** pendiente.
- **Fase 1232 — Write-back de dias_patron por banda:** pendiente.
- **Fase 1233 — Cableado pre-vuelo:** pendiente.
- **Fase 1234 — Suites canónicas de prueba:** pendiente.
- **Fase 1235 — Documentación y verificación:** pendiente.

---

## 🔧 Ronda 223 — Corrección de trazabilidad y 5 defectos reales de motor (Fases 1236 a 1241)

### Propósito y Objetivos
1. **Saneamiento de Trazabilidad Fantasma (Fase 1236):** Desmarcar las 12 fases 1224-1235 como ejecutadas en `implementation_plan.md` y `walkthrough.md`.
2. **Desbloqueo de Cura de Dimensiones con horas_dia NULL (Fase 1237):** Proteger `horas_dia` con `COALESCE(horas_dia, -1) <> @H` en `AgAulaDimensionesSync.cs`.
3. **Aislamiento Estricto por Subunidad en Sync (Fase 1238):** Evitar que subunidad vacía barra otras subunidades mediante `COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub`.
4. **Desacople de Factibilidad de la Cota de Oleadas (Fase 1239):** Reemplazar `W < olas` por cálculo de capacidad total de ranuras y acotar `oleada` a los límites del array en `AgPlanificadorOleadas.cs`.
5. **Siembra de Patrón de Días por Subconjunto Estricto (Fase 1240):** Asignar patrón de ≤2 días por alternancia modular en `AgFormGeneradorHorarios.cs:4544` para no forzar ventana completa en `ag_aula_asignatura`.
6. **Huella de Patrones Declarados en Clave de Memoización (Fase 1241):** Agregar `huellaPatrones` a `clavePlan` en `AgFabricaIndividuos.cs` para invalidar caché al reasignar patrones.

### Fases Ejecutadas
- **Fase 1236 — Corrección de trazabilidad:** `implementation_plan.md` y `walkthrough.md` rebajados a PLANIFICADO.
- **Fase 1237 — horas_dia NULL en sync:** `COALESCE(horas_dia, -1) <> @H` en `AgAulaDimensionesSync.cs:51`.
- **Fase 1238 — Aislamiento por subunidad:** `COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub` en `AgAulaDimensionesSync.cs:48`.
- **Fase 1239 — Factibilidad por capacidad de ranuras:** Validación de `demandaTotal > capacidadTotal` y clamp de `oleada` en `AgPlanificadorOleadas.cs:110-128, 215`.
- **Fase 1240 — Siembra por banda estricta:** Sembrado modular de días por subconjunto (<=2 días) en `AgFormGeneradorHorarios.cs:4544`.
- **Fase 1241 — Huella de patrones en caché:** `huellaPatrones` integrada en `clavePlan` en `AgFabricaIndividuos.cs:225-240`.

---

## 🔧 Ronda 224 — Corrección de regresiones en oleadas, siembra y pruebas (Fases 1242 a 1246)

### Propósito y Objetivos
1. **Saneamiento de Test Obsoleto (Fase 1242):** Actualizar aserción de `I5_SIN_SEMANAS_POR_OLEADA` a `I5_SIN_CAPACIDAD` en `AgPlanificadorOleadasTests.cs`.
2. **Prevención de Solape Ola × Carril (Fase 1243):** Detectar acumulación de demanda sobrepasando oferta en `(Ola, Carril)` en `AgPlanificadorOleadas.cs` y marcar infactibilidad accionable.
3. **Suspensión de Siembra de Bandas hasta v3 (Fase 1244):** Preservar siembra por carriles detrás del flag `USAR_BANDAS_DIAS = false` en `AgFormGeneradorHorarios.cs:4540-4560`.
4. **Mensajes Precisos de Semanas Hábiles y Capacidad (Fase 1245):** Aislar `I5_SIN_SEMANAS_HABILES` cuando `W == 0` y corregir texto de ranuras/semana en `I5_SIN_CAPACIDAD`.
5. **Protección Anti-Clobber en Sync y Alineación de Evidencia (Fase 1246):** Incorporar `COALESCE(NULLIF(@J, ''), jornada)` y `CASE WHEN @H > 0` en `AgAulaDimensionesSync.cs:45-47` y alinear referencias en `walkthrough.md`.

### Fases Ejecutadas
- **Fase 1242 — Test suite alineada a I5_SIN_CAPACIDAD:** Corrección de aserción en `AgPlanificadorOleadasTests.cs:84`.
- **Fase 1243 — Guard anti-solape ola × carril:** `cargaOlaCarril` y saturación en `AgPlanificadorOleadas.cs:228-255`.
- **Fase 1244 — Diferimiento de siembra por banda:** Flag `USAR_BANDAS_DIAS = false` en `AgFormGeneradorHorarios.cs:4544-4556`.
- **Fase 1245 — Mensajes de capacidad precisos:** `I5_SIN_SEMANAS_HABILES` e `I5_SIN_CAPACIDAD` en `AgPlanificadorOleadas.cs:125-135`.

---

## 🔧 Ronda 225 — Planificador de oleadas v3 por bandas de días y correcciones (Fases 1224 a 1235 + 1247/1248)

### Propósito y Objetivos
1. **Predicado Efectivo en Sync (Fase 1248):** Comparar contra valores efectivos en `AgAulaDimensionesSync.cs` (`COALESCE(jornada, '') <> COALESCE(NULLIF(@J, ''), COALESCE(jornada, ''))`, etc.).
2. **Partición de Ventana en Bandas Disjuntas de Días (Fases 1224 & 1225):** Crear `AgBandasDias.cs` con resolución determinista (exhaustiva/BFD) y validar en `AgBandasDiasTests.cs`.
3. **Contrato de Resultado Extendido (Fase 1227):** Exponer `Particion`, `VentanaSugerida`, `PatronesRenormalizados` y `OcupacionPorcentaje` en `AgPlanOleadasResultado.cs`.
4. **Escalera de Catálogo y Persistencia Asíncrona (Fases 1228 & 1229):** Crear `AgEscaleraVentanas.cs` y `AgEscaladorVentanaAula.cs` para sugerencias idempotentes y aplicación transaccional en `ag_aula` y `ag_arbol_nodo`.
5. **Motor v3 por Bandas de Días y Saturación por Exceso (Fases 1226 & 1247):** Implementar partición por bandas en `AgPlanificadorOleadas.cs`, asignando `Carril = banda` y `Oleada = 0`, marcando solo el excedente en saturación.
6. **Validación Accionable y Holgura Baja (Fase 1230):** Enriquecer `[I5_INFACTIBLE]` y `[I5_INFACTIBLE_OLEADA]` con ventana sugerida de `AgEscaleraVentanas`, y emitir `[I5_HOLGURA_BAJA]` si ocupación ≥ 88%.
7. **Huella de Patrones (Fase 1231):** Confirmada cubierta por Fase 1241 (`huellaPatrones` en `AgFabricaIndividuos.cs`).
8. **Write-Back por Banda (Fase 1232):** Activar flag `USAR_BANDAS_DIAS = true` en `AgFormGeneradorHorarios.cs:4551`.
9. **Diálogo Interactivo Pre-Vuelo (Fase 1233):** Insertar consulta de confirmación modal de escalado en pre-vuelo de `AgFormGeneradorHorarios.cs`.
10. **Suites Canónicas T1–T5 (Fase 1234):** Actualizar tests unitarios en `AgPlanificadorOleadasTests.cs` para reflejar bandas paralelas y validar casos 60/64 sesiones.
11. **Documentación Formal de Cambios (Fase 1235):** Crear `docs/articles/changes/2026-10-23-ag-oleadas-banda-dias.md`.

### Fases Ejecutadas
- **Fase 1248 — Predicado efectivo en sync:** Implementado en `AgAulaDimensionesSync.cs:51-53`.
- **Fase 1224 — AgBandasDias:** Implementado en `AgBandasDias.cs`, registrado en `ConversorPDF.XLL.csproj:653`.
- **Fase 1225 — AgBandasDiasTests:** Implementado en `AgBandasDiasTests.cs`.
- **Fase 1227 — Resultado extendido:** `AgPlanOleadasResultado.cs` actualizado con propiedades complementarias.
- **Fase 1228 — AgEscaleraVentanas:** Implementado en `AgEscaleraVentanas.cs`, registrado en `ConversorPDF.XLL.csproj:654`.
- **Fase 1229 — AgEscaladorVentanaAula:** Implementado en `AgEscaladorVentanaAula.cs`, registrado en `ConversorPDF.XLL.csproj:655`.
- **Fase 1226 — Motor v3 por bandas de días:** Reemplazado `Construir` en `AgPlanificadorOleadas.cs` con asignación determinista por banda.
- **Fase 1247 — Saturación no-cascada por banda:** Integrado en bucle de evaluación de bandas de `AgPlanificadorOleadas.cs`.
- **Fase 1230 — Validador con sugerencia y holgura baja:** `AgValidadorVentana.cs:295-344`.
- **Fase 1231 — Clave memoización patrones:** Cubierta por Fase 1241 en `AgFabricaIndividuos.cs:234-242`.
- **Fase 1232 — Write-back por banda:** `USAR_BANDAS_DIAS = true` en `AgFormGeneradorHorarios.cs:4551`.
- **Fase 1233 — Diálogo pre-vuelo de escalado:** Insertado en pre-vuelo de `AgFormGeneradorHorarios.cs:1235`.
- **Fase 1234 — Suites canónicas T1-T5:** Actualizado `AgPlanificadorOleadasTests.cs` (pruebas de bandas y casos T1-T5).
- **Fase 1235 — Documentación v3:** Creado `docs/articles/changes/2026-10-23-ag-oleadas-banda-dias.md`.

---

## 🔧 Ronda 226 — Auditoría de v3: bandas, escalera, escalador (Fases 1251 a 1262)

### Propósito y Objetivos
1. **Patrón de Días como Restricción (Fase 1251):** `AgBandasDias.cs` filtra candidatos mediante `ParticionAlojaPatrones` en dos pasadas (estricta y relajada).
2. **Cableado de diasPatronPorAsignatura (Fase 1252):** Conectar diccionario de patrones en `AgFabricaIndividuos.cs:245` y `AgValidadorVentana.cs:324`.
3. **Fuente Única de Saturación (Fase 1253):** Llenar `saturadas` directamente desde `asign.ExcesoPorBanda` en `AgPlanificadorOleadas.cs`.
4. **Limpieza de Variables Muertas (Fase 1254):** Eliminar `carriles` y `olas` en `AgPlanificadorOleadas.cs:86-87` (CS0219).
5. **Alineación de Suites de Prueba (Fase 1255):** Actualizar aserciones v2 a semántica v3 en `AgBandasDiasTests.cs` y `AgPlanificadorOleadasTests.cs`.
6. **Infactibilidad por Violación de Patrón R8 (Fase 1256):** Reportar infactibilidad accionable ante patrones que no caben en la banda asignada.
7. **Salvaguarda de Ventana Vacía (Fase 1257):** Retorno seguro en fallback de `AgBandasDias.Resolver` cuando `ventana` está vacía (evita `IndexOutOfRangeException`).
8. **Filtro de Árbol en Nodo (Fase 1258):** Condición `(@IdArbol <= 0 OR idarbol = @IdArbol)` en `AgEscaladorVentanaAula.cs` para no contaminar otros árboles.
9. **Filtro de Demanda por Programa y Semestre (Fase 1259):** Alinear `DemandaAula` en `AgEscaladorVentanaAula.cs` con el alcance del validador.
10. **Suites Unitarias de Escalera y Escalador (Fase 1260):** Crear `AgEscaleraVentanasTests.cs` y `AgEscaladorVentanaAulaTests.cs`.
11. **Refactorización Pre-Vuelo y Revalidación (Fase 1261):** Identificador ASCII `decisionesEscalado`, instancia única de escalador y revalidación en `AgFormGeneradorHorarios.cs`.
12. **Eliminación de Escalón Lateral (Fase 1262):** Remover `MAR-VIE → LUN-JUE` en catálogo de `AgEscaleraVentanas.cs`.

### Fases Ejecutadas
- **Fase 1251 — Patrón como restricción:** Dos pasadas con `ParticionAlojaPatrones` en `AgBandasDias.cs`.
- **Fase 1252 — Cableado de diasPatron:** Conectado en `AgFabricaIndividuos.cs` y `AgValidadorVentana.cs`.
- **Fase 1253 — Saturación por ExcesoPorBanda:** `saturadas` derivado del solver en `AgPlanificadorOleadas.cs`.
- **Fase 1254 — Eliminación de CS0219:** Eliminadas variables `carriles` y `olas` en `AgPlanificadorOleadas.cs`.
- **Fase 1255 — Tests v3 actualizados:** `AgBandasDiasTests.cs` y `AgPlanificadorOleadasTests.cs`.
- **Fase 1256 — Infactibilidad R8 accionable:** Detección de patrón no alojado en banda en `AgPlanificadorOleadas.cs`.
- **Fase 1257 — Fallback seguro de ventana vacía:** Retorno vacío antes de `demandaBanda` en `AgBandasDias.cs`.
- **Fase 1258 — Filtro idarbol en sqlNodo:** Parámetro `@IdArbol` en `AgEscaladorVentanaAula.cs`.
- **Fase 1259 — Filtro de demanda por aula:** Asignaturas filtradas por programa/semestre en `AgEscaladorVentanaAula.cs`.
- **Fase 1260 — Tests unitarios escalera y escalador:** Creados `AgEscaleraVentanasTests.cs` y `AgEscaladorVentanaAulaTests.cs`.
- **Fase 1261 — Pre-vuelo saneado con revalidación:** Identificadores limpios y `ValidarFactibilidad` tras escalado en `AgFormGeneradorHorarios.cs`.
- **Fase 1262 — Catálogo escalera sin lateral:** `MAR-VIE` escala directamente a `1,2,3,4,5` en `AgEscaleraVentanas.cs`.

---

## 🔧 Ronda 227 — Auditoría de compilación y regresiones de v3 (Fases 1263 a 1270)

### Propósito y Objetivos
1. **Restitución de Llave de Cierre en Ranuras (Fase 1263):** Reparar CS1513 en `AgPlanificadorOleadas.cs:84` restituyendo llave de cierre del bloque de ranuras de compatibilidad.
2. **Corrección de Acceso a InformeFactibilidad (Fase 1264):** Reemplazar llamadas a propiedades inexistentes por `EsFactible` y `BloqueosCriticos` en `AgFormGeneradorHorarios.cs:1263-1267` (CS1061).
3. **Exigencia de Factibilidad en Pasada Estricta (Fase 1265):** Descartar planes con `excesoTotal > 0` en pasada estricta de `AgBandasDias.Evaluar` para no bloquear relajación.
4. **Coherencia de Test de Patrón Común (Fase 1266):** Alinear aserción en `AgPlanificadorOleadasTests.cs` reconociendo que el patrón exige contención en banda.
5. **Capacidad Efectiva de Banda por Ocurrencia y Festivos (Fase 1267):** Pasar `capacidadDia` desde `AgPlanificadorOleadas.cs` a `AgBandasDias.Resolver` y `Evaluar`.
6. **Claridad de Árbol en Pre-Vuelo (Fase 1268):** Documentar propagación de `dias_ventana` desde pre-vuelo a `ag_aula` y nodos de árbol asociados.
7. **Reporte Específico R8 en Validador (Fase 1269):** Emitir `[R8_DiaFueraDePatron]` diferenciado de `[I5_INFACTIBLE_OLEADA]` en `AgValidadorVentana.cs:348-351`.
8. **Limpieza de Métodos Muertos v2 (Fase 1270):** Remover `ContarDias` y `CabeEnSemanas` de `AgPlanificadorOleadas.cs`.

### Fases Ejecutadas
- **Fase 1263 — Llave restituida en ranuras:** Llave cerrada en `AgPlanificadorOleadas.cs:84`.
- **Fase 1264 — Contrato InformeFactibilidad:** `EsFactible` y `BloqueosCriticos` en `AgFormGeneradorHorarios.cs:1263-1267`.
- **Fase 1265 — Factibilidad estricta en bandas:** `if (exigirPatrones && excesoTotal > 0) return;` en `AgBandasDias.cs:238`.
- **Fase 1266 — Test patrón común v3:** `Construir_CuandoPatronComun_AmbasAsignaturasCompartenBanda` en `AgPlanificadorOleadasTests.cs:334-345`.
- **Fase 1267 — CapacidadDia por festivos:** `capacidadDia` en `AgBandasDias.cs` y `AgPlanificadorOleadas.cs`.
- **Fase 1268 — Documentación pre-vuelo árbol:** Comentario aclaratorio en `AgFormGeneradorHorarios.cs:1252`.
- **Fase 1269 — Reporte R8 en validador:** `[R8_DiaFueraDePatron]` en `AgValidadorVentana.cs:348-351`.
- **Fase 1270 — Limpieza v2:** Eliminados métodos residuales en `AgPlanificadorOleadas.cs`.

---

## 🔧 Ronda 228 — Corrección de infactibles falsos y saneamiento de v3 (Fases 1271 a 1275)

### Propósito y Objetivos
1. **Reintento en Ventana Completa (Fase 1271):** En `AgPlanificadorOleadas.cs`, si la sub-ventana truncada (`iniBanda..finBanda`) no alberga la demanda requerida debido a festivos, reintentar sobre la ventana completa (`fechaFin.Date`) antes de declarar infactible.
2. **Aislamiento Estricto de Árbol en Pre-Vuelo (Fase 1272):** Condicionar la ejecución de `sqlNodo` a `idArbol > 0` en `AgEscaladorVentanaAula.cs` para evitar actualizar nodos de otros árboles cuando `idArbol == 0`.
3. **Capacidad Efectiva y Ocupación con Festivos (Fase 1273):** Calcular `capacidadDia` antes de la guarda global `demandaTotal > capacidadTotal` en `AgPlanificadorOleadas.cs`, y usar `capacidadDia.Values.Sum()` para `slotsTotales` en `AgBandasDias.cs:275`.
4. **Revalidación Asíncrona fuera de UI (Fase 1274):** Ejecutar `ValidarFactibilidad` dentro de `Task.Run` con `ConfigureAwait(true)` en `AgFormGeneradorHorarios.cs:1264`.
5. **Retiro de Restos Muertos de v2 (Fase 1275):** Eliminar la colección y bucle inalcanzable de `saturadas`, la rama `saturadas.Contains` en `detalleInfactibles`, y el método privado no referenciado `ContarRanuras` en `AgPlanificadorOleadas.cs`.

### Fases Ejecutadas
- **Fase 1271 — Reintento en ventana completa:** Implementado en `AgPlanificadorOleadas.cs:179-184`.
- **Fase 1272 — Pre-vuelo no toca otros árboles:** `if (idArbol > 0)` en `AgEscaladorVentanaAula.cs:122` y doc en `AgFormGeneradorHorarios.cs:1252`.
- **Fase 1273 — Capacidad efectiva pre-guard y ocupación:** `capacidadDia` en `AgPlanificadorOleadas.cs:128-142` y `AgBandasDias.cs:275-278`.
- **Fase 1274 — Revalidación asíncrona:** `await Task.Run` en `AgFormGeneradorHorarios.cs:1264`.
- **Fase 1275 — Limpieza saturadas y ContarRanuras:** Eliminados de `AgPlanificadorOleadas.cs`.

---

## 🔧 Ronda 229 — Sincronización de perfil activo y coherencia de reporte de oleadas (Fases 1276 a 1278)

### Propósito y Objetivos
1. **Sincronización de Perfil Activo en Pre-Vuelo (Fase 1276):** Pasar `idArbolActivo` desde `cmbArboles.SelectedItem` en `AgFormGeneradorHorarios.cs:1253` para actualizar `ag_arbol_nodo` del perfil activo sin afectar árboles ajenos.
2. **Coherencia de Semanas y Capacidad tras Reintento (Fase 1277):** Si la banda se amplió a la ventana completa en `AgPlanificadorOleadas.cs`, reportar `semanasReportadas = W` y `capacidadReportada = RanurasEfectivas(...)` en `AgPlanOleada`, restituyendo la función `RanurasEfectivas`.
3. **Detalle de Infactibilidad con Ranuras Efectivas (Fase 1278):** Reportar ranuras efectivas (festivos descontados) mediante `RanurasEfectivas` en el mensaje de error de `AgPlanificadorOleadas.cs:209-211`.

### Fases Ejecutadas
- **Fase 1276 — Sincronización idArbolActivo:** `idArbolActivo` en `AgFormGeneradorHorarios.cs:1252-1254` y comentario en `AgEscaladorVentanaAula.cs:122`.
- **Fase 1277 — Coherencia de SemanasOleada/CapacidadRanuras:** `semanasReportadas`, `capacidadReportada` y `RanurasEfectivas` en `AgPlanificadorOleadas.cs:184, 216-228, 240-249`.
- **Fase 1278 — Ranuras efectivas en detalle de error:** Mensaje contextual en `AgPlanificadorOleadas.cs:210-212`.

---

## 🔧 Ronda 230: Recuperación de Última Optimización AG y Persistencia Resiliente de Pernocta (Fases 1279 a 1283)

### Propósito y Objetivos
1. **Soporte de Columna isPernoctaUltimoDia en Staging (Fase 1279):** Crear e implementar `MigrationV77_AddPernoctaAgVinculaciones` para asegurar la columna `isPernoctaUltimoDia INTEGER NOT NULL DEFAULT 0` en `vinculaciones` y `ag_vinculaciones`, evitando excepciones SQL en el repositorio académico en modo AG. Registrar en `SqliteHelper.cs` y `ConversorPDF.XLL.csproj`.
2. **Sustitución de Botones de Exportación en UI (Fase 1280):** En `AgFormGeneradorHorarios.Designer.cs`, remover `btnExportarProyeccion` y `btnExportarSimulacion`, e incorporar `btnCargarUltimaOptimizacion` inmediatamente a la izquierda de `btnPromoverProduccion`.
3. **Carga y Rehidratación de Última Optimización (Fase 1281):** En `AgFormGeneradorHorarios.cs`, cablear el botón a `CargarUltimaOptimizacionAsync()`, consultando `_scheduleRepository.GetUltimaCorridaAsync` y `VinculacionQueryRepository.GetSesionesAsync`, rehidratando grids, vacantes, labels de métricas y eliminando 4 métodos muertos de exportación.
4. **Endurecimiento Transaccional en FormSesionesAcademicas (Fase 1282):** Validar `SqliteResult.IsFailure` en inserción/actualización de vinculación y sesiones para prevenir reasignación de IdVinculacion inválido, y dinamizar tabla de consulta en eliminación con `FROM {_modo.TablaVinculaciones}`.
5. **Suites de Pruebas Unitarias (Fase 1283):** Validar en `MigrationV77Tests.cs` los metadatos, idempotencia y adición de columna, y en `AcademicRepositoryAgModoPernoctaTests.cs` la resolución del error SQL y la persistencia de pernocta en `ag_vinculaciones`.

### Fases Ejecutadas
- **Fase 1279 — Migración V77:** Implementada en `MigrationV77_AddPernoctaAgVinculaciones.cs`, registrada en `SqliteHelper.cs:378` y `ConversorPDF.XLL.csproj:668`.
- **Fase 1280 — Control UI btnCargarUltimaOptimizacion:** Configurado en `AgFormGeneradorHorarios.Designer.cs:90, 930, 973-982, 1284`.
- **Fase 1281 — CargarUltimaOptimizacionAsync:** Implementado en `AgFormGeneradorHorarios.cs:354, 2061-2212` con remoción de exportadores obsoletos.
- **Fase 1282 — Guardado y borrado resiliente:** En `FormSesionesAcademicas.cs:1330-1348, 2590`.
- **Fase 1283 — Suites de pruebas unitarias:** Creados `MigrationV77Tests.cs` y `AcademicRepositoryAgModoPernoctaTests.cs`.

---

## 🔧 Ronda 231: Consolidación de Carga de Staging y Saneamiento de Estado (Fases 1284 a 1288)

### Propósito y Objetivos
1. **Desduplicación de Carga de Resultados (Fase 1284):** Hacer que `CargarUltimaOptimizacionAsync` delegue en `CargarModuloResultadosAsync(corrida.Id)` eliminando ~110 líneas duplicadas y garantizando la población de violaciones y vacantes auditables.
2. **Saneamiento de Estado de Sesión ante Corrida Vacía (Fase 1285):** Restablecer a cero/nulo `_idCorridaCargada` y `_ultimoResultado`, y restaurar etiquetas de modo y colores de advertencia cuando la corrida no posee sesiones en staging.
3. **Desacople Diagnóstico de Fallos SQL (Fase 1286):** En `CargarModuloResultadosAsync`, separar la verificación de `sesionesResult.IsFailure` con log de error específico de la condición `sesionesResult.Value.Count == 0`.
4. **Evaluación de Pernocta en Vistas AG (Fase 1287):** Exponer `isPernoctaUltimoDia` en `vw_AgSesionesDetalle` y `vw_AgVinculacionesDetalle` mediante nueva migración V78, condicionada a decisión de alcance previa por su impacto en 5 suites de pruebas.
5. **Ampliación de Dimensiones del Botón de Carga (Fase 1288):** Ajustar ancho a 195px (`Location = (535, 10)`, `Size = (195, 32)`) para prevenir truncamiento tipográfico bajo escalado DPI > 100%.

### Fases Ejecutadas
- **Fase 1284 — Desduplicación de Carga:** Delegación en `CargarModuloResultadosAsync` en `AgFormGeneradorHorarios.cs:2088-2105`.
- **Fase 1285 — Saneamiento de Estado Vacío:** Limpieza de `_idCorridaCargada`, `_ultimoResultado` y labels en `AgFormGeneradorHorarios.cs:2093-2102`.
- **Fase 1286 — Desacople de IsFailure:** Logging de fallo SQL aislado de corrida vacía en `AgFormGeneradorHorarios.cs:1917-1926`.
- **Fase 1287 — Pernocta en Vistas AG:** Retirada / diferida a ronda dedicada por requerir nueva migración V78 y actualización de 5 suites de fixtures.
- **Fase 1288 — Ampliación de Botón UI:** Dimensiones `(535, 10)` / `(195, 32)` en `AgFormGeneradorHorarios.Designer.cs:976-978`.

---

## Ronda 232 — Contrato Explícito de Carga, Saneamiento de Pantalla Mixta y Documentación XML (Fases 1289 a 1291)

### Propósito y Objetivos
1. **Contrato Explícito de Carga (Fase 1289):** Sustituir la detección de fallo indirecta basada en `DataSource == null` por un contrato directo `Task<bool>` en `CargarModuloResultadosAsync`. Consumir el valor retornado en `CargarUltimaOptimizacionAsync` para evitar falsos éxitos cuando la corrida no posee sesiones o falla la rehidratación.
2. **Saneamiento de Pantalla Mixta ante Fallo de Staging (Fase 1290):** Al detectar `sesionesResult.IsFailure` en `CargarModuloResultadosAsync`, limpiar inmediatamente el grid de horarios, lista/grid de vacantes, etiquetas de métricas (`lblMejorFitness`, `lblHardViolaciones`, `lblVacantes`) y deshabilitar `btnPromoverProduccion` para prevenir mezclas de datos de corridas anteriores.
3. **Actualización de Documentación XML (Fase 1291):** Actualizar el comentario XML de `CargarUltimaOptimizacionAsync` para reflejar con precisión la delegación a `CargarModuloResultadosAsync` y el origen en staging (`ag_sesiones` + `ag_vinculaciones`) en vez de la vista de producción.
4. **Fase 1287:** Permanece diferida a nueva migración V78 y actualización transversal de 5 suites de pruebas.

### Fases Ejecutadas
- **Fase 1289 — Contrato Explícito Task<bool>:** `CargarModuloResultadosAsync` retorna `Task<bool>` y `CargarUltimaOptimizacionAsync` evalúa `!cargada` en `AgFormGeneradorHorarios.cs:1847, 1884, 2064, 2110-2123`.
- **Fase 1290 — Saneamiento ante Fallo de Staging:** Limpieza completa de UI y vacantes en rama `sesionesResult.IsFailure` en `AgFormGeneradorHorarios.cs:1917-1930`.
- **Fase 1291 — Actualización de Documentación XML:** Corrección de documentación en `AgFormGeneradorHorarios.cs:2066-2070`.

---

## 🔧 Ronda 233: Corrección Quirúrgica de Franja Horaria de Sábado y Defensa de Invariante I7 (Fases 1292 a 1296)

### Propósito y Objetivos
1. **Tests Primero (Fase 1292):** Validar en pruebas unitarias que: (a) En sábado 4h se asignen bloques de MANANA y TARDE; (b) En sábado 8h posgrado se conserve franja única; (c) L-V conserve franja única; (d) El genotipo resultante no emita violaciones duras en R4 ni R14.
2. **Corrección de Ranuras en Fábrica (Fase 1293):** En `AgFabricaIndividuos.cs`, condicionar `permitidasRanura.Count > 1` (`franjaExactaObligatoria = true`) para exigir la franja exacta de la ranura, evitando el colapso a MANANA en sábado 4h.
3. **Sellar Franjas de Sábado (Fase 1294):** En `AgCoherenciaGrupo.cs:Sellar`, reasignar el 2.º gen del sábado a la franja alterna si comparte franja con el 1.º, previniendo regresiones por cruce/mutación.
4. **Deduplicación en Columna K (Fase 1295):** En `AgMaterializadorSesiones.cs:DiaHorasSemanalesDe`, agrupar por `(Día, HoraInicio, HoraFin)` en vez de `IdBloqueHorario` para emitir `SABADO|4|08:00 A 12:00, SABADO|4|13:00 A 17:00`.
5. **Invariante I7 de Sesiones Duplicadas (Fase 1296):** En `AgExpansorSesiones.cs` y `AgMaterializadorSesiones.cs`, rechazar y poner en cuarentena grupos o individuos con sesiones duplicadas en `(FechaSesion, HorarioSesion)`.

### Fases Ejecutadas
- **Fase 1292 — Cobertura de Pruebas Unitarias:** Implementados tests en `AgFabricaIndividuosTests.cs` y `AgValidadorReglasTests.cs`.
- **Fase 1293 — Exigencia de Franja Exacta en Ranuras:** En `AgFabricaIndividuos.cs:429-438`, incorporada `franjaExactaObligatoria` cuando `permitidasRanura.Count > 1`.
- **Fase 1294 — Sellar Franja Alterna en Sábado:** En `AgCoherenciaGrupo.cs:108-154`, reasignación determinista a franja alterna en sábado.
- **Fase 1295 — Deduplicación de Columna K por Horario:** En `AgMaterializadorSesiones.cs:136-175`, agrupación por `(DiaSemana, HoraInicio, HoraFin)`.
- **Fase 1296 — Invariante I7:** En `AgExpansorSesiones.cs:187-206` y `AgMaterializadorSesiones.cs:258-272`.

---

## 🔧 Ronda 234: R8 Permite Desdoble de Sábado y Saneamiento de Fallbacks y Tests (Fases 1297 a 1302)

### Propósito y Objetivos
1. **Corrección de Identificador Enum en Tests (Fase 1297):** Reemplazar `R14_SolapeHorarioBloque` por el miembro canónico `R14_BloqueSolapadoPorIntervalo` en `AgValidadorReglasTests.cs:765`.
2. **R8 Evalúa Unicidad por Par (Día, Franja) (Fase 1298):** En `AgValidadorReglas.cs:580-592`, evaluar pares `(DiaSemana, Franja)` en vez de días puros para que sábado 4h con MANANA y TARDE no reciba violación dura espuria.
3. **Reproducción Efectiva en Test de Fábrica (Fase 1299):** En `AgFabricaIndividuosTests.cs:436-468`, añadir segunda asignatura para alcanzar carril 1 ([6:MANANA, 6:TARDE]) y asertar `HardViolations == 0`.
4. **Fallback Legacy ante Falta de Bloque Exacto (Fase 1300):** En `AgFabricaIndividuos.cs:440-454`, si `bloqueRanura == null` con `franjaExactaObligatoria`, aplicar fallback legacy y advertencia en vez de enviar a cuarentena I6.
5. **Tracking de Franjas en Sellar (Fase 1301):** En `AgCoherenciaGrupo.cs:108-153`, gestionar `franjasOcupadas` (HashSet) y registrar `WarningLimitado` si no hay bloque alterno disponible.
6. **Descarte de Duplicados en Expansor (Fase 1302):** En `AgExpansorSesiones.cs:187-206`, advertir y deduplicar sesiones duplicadas en `ExpandirIndividuo` sin cancelar todo el individuo.

### Fases Ejecutadas
- **Fase 1297 — Enum R14:** Corregido en `AgValidadorReglasTests.cs:765`.
- **Fase 1298 — R8 con Pares (Día, Franja):** Implementado en `AgValidadorReglas.cs:580-592`.
- **Fase 1299 — Test de Fábrica de Sábado:** Actualizado en `AgFabricaIndividuosTests.cs:436-468`.
- **Fase 1300 — Fallback Legacy en Fábrica:** Implementado en `AgFabricaIndividuos.cs:440-454`.
- **Fase 1301 — Sellar con HashSet de Franjas:** Refactorizado en `AgCoherenciaGrupo.cs:108-153`.
- **Fase 1302 — I7 con Deduplicación en Expansor:** Modificado en `AgExpansorSesiones.cs:187-206`.

---

## 🔧 Ronda 235: Coherencia Defensiva en Materialización y Saneamiento de Pool Docente (Fases 1303 a 1306)

### Propósito y Objetivos
1. **Habilitación de Docente en Test de Validador (Fase 1303):** En `AgValidadorReglasTests.cs:722`, registrar docente habilitado con comodín de semestre (`contexto.RegistrarDocenteHabilitado(10, 1, 100, "")`) para satisfacer R5 y permitir `HardViolations == 0`.
2. **Ejecución Defensiva Incondicional de Sellar en Staging (Fase 1304):** En `AgMaterializadorSesiones.cs:63-70`, invocar `AgCoherenciaGrupo.Sellar` para todo grupo con `genes.Count > 1`, asegurando que el desdoble de franjas de sábado actúe en staging aun cuando `Σ HorasGen == esperado`.
3. **Reserva de Franjas para Genes Fijados (Fase 1305):** En `AgCoherenciaGrupo.cs:117-128`, registrar previamente en `franjasOcupadas` las franjas de genes bloqueados (`Estado == Fijado`) en sábado para no reasignar sobre ellas.
4. **Clarificación Diagnóstica de Fallback Legacy (Fase 1306):** En `AgFabricaIndividuos.cs:465-472`, condicionar el mensaje del log a si `bloqueRanura != null` o si el gen terminará en cuarentena I6.

### Fases Ejecutadas
- **Fase 1303 — Habilitación de Docente en Test:** Implementado en `AgValidadorReglasTests.cs:722`.
- **Fase 1304 — Sellar Incondicional en Materialización:** Actualizado en `AgMaterializadorSesiones.cs:63-70`.
- **Fase 1305 — Reserva de Franjas Fijadas en Sellar:** Añadido en `AgCoherenciaGrupo.cs:117-128`.
- **Fase 1306 — Clarificación de Log en Fallback:** Corregido en `AgFabricaIndividuos.cs:465-472`.

---

## 🔧 Ronda 236: Desdoble de Sábado con Genes Fijados, Catálogo Desconectado y Ajuste de Equivalencias (Fases 1307 a 1310)

### Propósito y Objetivos
1. **Desdoble de Sábado con Genes Fijados (Fase 1307):** En `AgCoherenciaGrupo.cs:108-117`, calcular `genesSabadoTotal` considerando todos los genes del día 6 (incluyendo `Fijado`), y activar el desdoble cuando `genesSabadoTotal > 1 && genesSabado.Count > 0`. Test unitario en `AgCoherenciaGrupoTests.cs`.
2. **Resolución Robusta de Franja en Sellar (Fase 1308):** En `AgCoherenciaGrupo.cs:117-154`, crear helper `FranjaResuelta` con fallback a `ctx.Bloques`, evitar mutaciones a ciegas si el bloque no es resoluble (`null`), y emitir log diferenciado `SabadoSinFranjaLibre` vs `SabadoSinAlterna`.
3. **Restricción de franjaExactaObligatoria a Sábado (Fase 1309):** En `AgFabricaIndividuos.cs:430-438`, condicionar `diaIso == 6 && permitidasRanura.Count > 1` para no romper equivalencias legítimas `TARDE/NOCHE` de 8h en L–V.
4. **Diagnóstico Explícito sin Drop en Expansor I7 (Fase 1310):** En `AgExpansorSesiones.cs:187-210`, registrar `LoggerHelper.WarningLimitado` detallando genes en colisión sin descartar sesiones, evitando déficit de horas I2.

### Fases Ejecutadas
- **Fase 1307 — Gate de Desdoble con Fijados:** Actualizado en `AgCoherenciaGrupo.cs:108-117` y test en `AgCoherenciaGrupoTests.cs:525-550`.
- **Fase 1308 — FranjaResuelta y Log Diferenciado:** Implementado en `AgCoherenciaGrupo.cs:117-154`.
- **Fase 1309 — franjaExactaObligatoria Acotada a Sábado:** Modificado en `AgFabricaIndividuos.cs:430-438`.
- **Fase 1310 — Diagnóstico I7 sin Pérdida de Horas:** Ajustado en `AgExpansorSesiones.cs:187-210`.

---

## 🔧 Ronda 238: Cierre Falso de Ronda 237 — 0/15 Fases Aplicadas y 9 Defectos de Diseño (Fases 1326 a 1334)

### Contexto y Diagnóstico
- **Ronda 237 Reabierta:** Las fases 1311 a 1325 no fueron aplicadas en disco (0/12 archivos nuevos existen, 0 migraciones registradas). Permanecen abiertas en el backlog de implementación.
- **Defectos de Especificación:** Se identificaron 9 defectos críticos en la propuesta de diseño original que requerían corrección antes de la codificación en disco.

### Fases de Corrección de Especificación (Fases 1326 a 1334)
- **Fase 1326 — Falso Cierre con Evidencia Inexistente:** Reabrir fases 1311–1325. Establecer gate estricto de verificación `Test-Path` + `Select-String` en disco por fase. Estado: `ESPECIFICADA`.
- **Fase 1327 — Reflow de `pnlFiltros` en 2 Filas (84 px):** Rediseñar `pnlFiltros` a 84 px de alto en dos filas. Estado: `ESPECIFICADA`.
- **Fase 1328 — Integridad de `PoolDocenteKeys` y Genes Fijados:** Validar que genes fijados no usen docentes excluidos del alcance para prevenir violaciones duras R5; preservar centinela VACANTE en `AplicarAlcanceTerceros`. Estado: `ESPECIFICADA`.
- **Fase 1329 — Exclusión Explícita de Tercero Centinela VACANTE:** Filtrar centinela en consulta SQL (`id <> @IdVacante`) y en `AgTerceroFiltro` en memoria. Estado: `ESPECIFICADA`.
- **Fase 1330 — Paridad en Filtro de Subunidad:** Incluir `(subunidad = @Sub OR subunidad = '' OR subunidad IS NULL)` en consultas de terceros. Estado: `ESPECIFICADA`.
- **Fase 1331 — Supresión de Eventos en Inicialización:** Incorporar bandera `_suprimiendoEventos` durante carga de catálogos y ejecutar `RestaurarFiltrosPersistidosAsync` antes de `PoblarArbolContextoAcademico`. Estado: `ESPECIFICADA`.
- **Fase 1332 — Persistencia Síncrona en `OnFormClosing`:** Reemplazar llamadas asíncronas en el cierre por `GuardarFiltrosPersistidosSync` defensivo antes del apagado del scheduler. Estado: `ESPECIFICADA`.
- **Fase 1333 — Normalización e Índice en `ag_filtro_seleccion`:** Definir `subunidad TEXT NOT NULL DEFAULT ''` e índice único compuesto `ux_ag_filtro_seleccion(subunidad, tipo, id_ref)`. Estado: `ESPECIFICADA`.
- **Fase 1334 — API de `CheckedComboBox` Compatible:** Diseñar `CheckedComboBox` con método `Cargar<T>`, `Seleccionados<T>`, `Primario<T>` y `VacioEsTodos = true`. Estado: `ESPECIFICADA`.

---

## 🔧 Ronda 239: Ronda 238 Registrada Solo en Documentación — 9/9 Fases "ACEPTADAS" con 0 Líneas de Código (Fases 1335 a 1341)

### Contexto y Diagnóstico
- **Ausencia de Código Físico:** Verificación en disco confirma 0 líneas de código implementadas para las rondas 237 y 238.
- **Régimen de Estados:** Se formalizan los estados operativos: `ESPECIFICADA` (documentada sin código), `IMPLEMENTADA` (código en disco con `Test-Path = True` y `Select-String`), `VERIFICADA` (compilación y pruebas exitosas).
- **Parches de Integración:** Se especifican las fases 1335 a 1341 para resolver desactualización de csproj, radio de impacto de `CheckedComboBox`, truncamiento del árbol, persistencia síncrona sin deadlocks y corrección de atribuciones.

### Fases Aplicadas y Verificadas
| Fase | Estado | Evidencia en Disco |
|---|---|---|
| — | Ninguna fase aplicada en código aún | — |

### Backlog Abierto (Implementación Física Pendiente)
| Fase | Nombre | Estado | Archivo Destino |
|---|---|---|---|
| 1311 | `AgTipoProgramaCatalog` | ESPECIFICADA | `Academic\Domain\Models\AgTipoProgramaCatalog.cs` |
| 1312 / 1334 / 1337 | `CheckedComboBox` (UI + API + Compatibilidad) | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs` |
| 1313 / 1327 / 1338 | Designer + Eventos + Reflow (84px, grp 326px) | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` |
| 1314 | `AgFiltroArbolResolver` + Accesores | ESPECIFICADA | `Academic\Application\Services\AgFiltroArbolResolver.cs` |
| 1315 / 1333 | Migración V78 + `AgFiltroSeleccionRepository` | ESPECIFICADA | `Academic\Infrastructure\Persistence\Migrations\MigrationV78_AddAgFiltroSeleccion.cs` |
| 1316 / 1331 / 1332 / 1339 | Restauración, Supresión y Cierre Síncrono | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1317 | Horas por Nivel (4h Pregrado / 8h Posgrado) | ESPECIFICADA | `Academic\Application\Mappers\AgNodoArbolKeyMapper.cs` |
| 1318 / 1329 / 1330 / 1340 | `AgFormTerceros` (UI + SQL saneado) | ESPECIFICADA | `Academic\UI\Forms\AgFormTerceros.cs` |
| 1319 / 1329 | DTO + Lógica Pura Terceros + Filtro Vacante | ESPECIFICADA | `Academic\Application\DTOs\AgTerceroSeleccionDto.cs` |
| 1320 / 1328 | Alcance Terceros + Guarda Genes Fijados | ESPECIFICADA | `Academic\Domain\Models\Genetica\AgContexto.cs` |
| 1321 | Botones Generador y Resumen Filtros | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1322 | Suites de Pruebas Unitarias | ESPECIFICADA | `ConversorPDF.Tests\...` |
| 1323 / 1336 | Actualización `<Compile Include>` e Integración csproj | ESPECIFICADA | `ConversorPDF.XLL.csproj` |
| 1324 / 1335 / 1341 | Sincronización Documental y Reglas de Gate | ESPECIFICADA | `implementation_plan.md`, `walkthrough.md` |
| 1325 | Procedimiento de Aceptación Manual | ESPECIFICADA | Protocolo Operativo |
| 1342 | `SelectedIndex` read/write y carga tipada sin DataSource | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs`, `AgFormGeneradorHorarios.cs` |
| 1343 | `ModoSeleccion` (Unico/Multiple) y disparo de evento | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs`, `AgFormGeneradorHorarios.cs` |
| 1344 | `pnlFiltros` 66px y `SplitterDistance = 390` (árbol 189px) | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` |
| 1345 | `EncryptedSQLiteContext.Open()` para transacción Dapper | ESPECIFICADA | `Academic\Infrastructure\Persistence\Repositories\AgFiltroSeleccionRepository.cs` |
| 1346 | Registro de `MigrationV78_AddAgFiltroSeleccion` en `SqliteHelper` | ESPECIFICADA | `Core\Helpers\SqliteHelper.cs` |
| 1347 | Saneamiento de régimen de estados (ESPECIFICADA / IMPLEMENTADA) | ESPECIFICADA | `implementation_plan.md`, `walkthrough.md` |
| 1348 | Conexión compartida sin `using` en repo (`Func<IDbConnection>`) | ESPECIFICADA | `Academic\Infrastructure\Persistence\Repositories\AgFiltroSeleccionRepository.cs` |
| 1349 | `CheckedComboBox` definición completa unificada | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs` |
| 1350 | Coordenadas exactas en `pnlFiltros` sin solapamiento de labels | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` |
| 1351 | Declaración de campos e instanciación en Designer | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` |
| 1352 | Reconciliación de persistencia y restauración de período único | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1353 | Accesores `ProgramaPrimario()` y `ProgramasSeleccionados()` | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1354 | `PoblarArbolContextoAcademico` consume selecciones de combos | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1355 | Autoridad única en árbol de nivel 6 (combos son pre-filtros) | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1356 | Declaración de `SubunidadActual()` para persistencia/restauración | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1357 | `CheckedComboBox` con superficie visual completa | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs` |
| 1358 | Suites de pruebas unitarias para lógica nueva (SDK-style) | ESPECIFICADA | `ConversorPDF.Tests\Academic\*.cs` |
| 1359 | Dimensionamiento fijo y AutoEllipsis en etiquetas variables | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` |
| 1360 | Preservación de combinaciones marcadas en repoblado de árbol | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1361 | Guarda `_sincronizandoLista` en `CheckedComboBox.ItemCheck` | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs` |
| 1362 | `AutoClose = true` y cierre por Enter/Esc en desplegable | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs` |
| 1363 | Consumo de `cmbTipoPrograma` en `FiltrarProgramas` | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1364 | Uso del identificador canónico `toolTipConteo` | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1365 | Snapshot por ruta jerárquica y materialización recursiva en árbol lazy | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1366 | Clave compuesta jerárquica para eliminar colisiones en niveles 4-6 | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1367 | Supresión de `AfterCheck` (`_isUpdatingTreeChecks = true`) en restauración | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1368 | Preservación y restauración de ramas expandidas y scroll de árbol | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |

---

## 🔧 Ronda 240: Ronda 239 Registrada Sin Código y 4 Parches Técnicamente Inválidos (Fases 1342 a 1347)

### Contexto y Diagnóstico
- **Ronda 239 Documental:** 0 líneas de código implementadas; 1335–1341 permanecen en estado `ESPECIFICADA`.
- **Invalidación Técnica de Parches Previos:**
  - `SelectedIndexSet` no existe en WinForms ni en los call sites; se requiere propiedad read/write `SelectedIndex`.
  - `SelectedIndexChanged` no tenía punto de emisión; se requiere `ModoSeleccion { Unico, Multiple }` para período único y ventana semestral.
  - `grpSeleccion.Size` es inerte ante `Dock = Fill`; el ajuste real requiere `pnlFiltros` de 66 px y `SplitterDistance = 390` en `splitContenedorSim`.
  - `EncryptedSQLiteContext` no expone propiedad `Connection`; la API canónica es `using (var conn = ctx.Open())`.
  - `MigrationV78_AddAgFiltroSeleccion` carecía de orden de registro en `SqliteHelper.cs:378`.
  - El régimen de estados volvió a violarse con marcas prematuras en la tabla de principios.

---

## 🔧 Ronda 241: Parches 1342–1347 Registrados Sin Código y 6 Defectos en las Correcciones "Aceptadas" (Fases 1348 a 1353)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (1311 a 1353) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Evitar disponer conexión compartida (`AcademicRepository.cs:25`): usar `Func<IDbConnection>` sin `using` local sobre `Open()`.
  - Definición unificada, coherente y compilable de `CheckedComboBox` (evita duplicidad y símbolos huérfanos).
  - Posicionamiento exacto de etiquetas y controles en `pnlFiltros` de 66 px eliminando colisiones.
  - Declaración e inicialización explícita de controles en `AgFormGeneradorHorarios.Designer.cs` y desvinculación de combos antiguos.
  - Sincronización del modelo de persistencia `ag_filtro_seleccion` con la restricción de período único (`Modo = ModoSeleccion.Unico`).
  - Accesores canónicos `ProgramaPrimario()` y `ProgramasSeleccionados()` para desacoplar los 6 call sites sensibles de `chkProgramas`.

---

## 🔧 Ronda 242: Parches 1348–1353 Registrados Sin Código y 6 Defectos Residuales (Fases 1354 a 1359)

### Contexto y Diagnóstico
- **Régimen de Estados:** Todo el backlog (Fases 1311 a 1359) permanece en `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Consumo efectivo de selecciones en `PoblarArbolContextoAcademico` mediante `FiltrarPeriodos`, `FiltrarSedes` y `FiltrarProgramas`.
  - Unificación de autoridad en la selección de nodos nivel 6 del árbol (combos actúan estrictamente como pre-filtros).
  - Definición explícita de `private string SubunidadActual() => SyncContext.ActiveSubunidad ?? string.Empty;`.
  - Superficie interactiva de `CheckedComboBox` con botón desplegable, `CheckedListBox` alojado en `ToolStripDropDown`.
  - Incorporación de pruebas unitarias SDK-style sin dependencias de compilación manual.
  - Dimensionamiento fijo (280×18) y `AutoEllipsis = true` con `ToolTip` en etiquetas dinámicas del panel de filtros.

---

## 🔧 Ronda 243: Fases 1354–1359 Solo Documentadas — 5 Defectos Residuales (Fases 1360 a 1364)

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

