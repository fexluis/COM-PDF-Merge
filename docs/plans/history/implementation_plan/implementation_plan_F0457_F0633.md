<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# implementation_plan - Fases 457-633 (Rondas 44)

> Indice: [../../../implementation_plan.md](../../../implementation_plan.md) - Escision: 2026-11-16

---
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

