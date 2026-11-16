<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# implementation_plan - Fases 777-964 (Rondas 44)

> Indice: [../../../implementation_plan.md](../../../implementation_plan.md) - Escision: 2026-11-16

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

