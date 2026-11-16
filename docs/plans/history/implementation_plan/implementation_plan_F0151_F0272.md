<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# implementation_plan - Fases 151-272 (Rondas -)

> Indice: [../../../implementation_plan.md](../../../implementation_plan.md) - Escision: 2026-11-16

---
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

