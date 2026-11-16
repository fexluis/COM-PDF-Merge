<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# walkthrough - Fases 620-710 (Rondas 102..117)

> Indice: [../../../walkthrough.md](../../../walkthrough.md) - Escision: 2026-11-16

---
# Walkthrough — Movimiento de Usuarios entre Vaults al Cambiar de Subunidad (MUV)

**Fecha:** 2026-10-29
**Plan:** `.agents/plans/2026-10-29-mover-usuario-entre-vaults-plan.md`
**Estado:** COMPLETED
**Verificación Full:** 964 de 964 pruebas superadas (0 fallos, 0 regresiones vs baseline de 952).

## Resumen de la Ejecución

Se completó la implementación y verificación de las 6 tareas del plan MUV:
1. **T1 (integrator, L0):** Definición de contratos `IUsuarioVaultMover`, ampliación de `IMoverUsuarioSubunidadService`, DTOs `MovimientoVaultResultado` y `MovimientoUsuarioResultado`, esqueletos e inclusión en `ConversorPDF.XLL.csproj`.
2. **T2 (tdd-red-writer):** Suite de 12 pruebas automatizadas en `ConversorPDF.Tests` (Unit, Integration SQL en `:memory:` y E2E con dos vaults reales cifrados). Estado RED certificado (fallos controlados por esqueletos con Failure).
3. **T3 (implementer):** Implementación de `UsuarioVaultMover` con `UpsertEnVault` idempotente (`ON CONFLICT(id) DO UPDATE`), persistencia atómica en el vault destino (`AcademicData_<DESTINO>.db.enc`) y preservación estricta de perfil, activo y firma. 5 tests en verde.
4. **T4 (implementer):** Implementación de `MoverUsuarioSubunidadService` con la política R-MUV-02: `config.subunidad` se actualiza exclusivamente cuando el usuario editado es el usuario actual de Windows. 7 unit tests en verde.
5. **T5 (implementer):** Cableado en `FormUsuarios.btnGuardar` con persistencia secuencial (primero destino, luego origen) y recarga en caliente de `FormConfig` mediante `RecargarConfiguracion()`.
6. **T6 (build-verifier + task-reviewer):** Verificación Full de 964 tests aprobados, auditoría de arquitectura (0 `System.Data` en `Application`, 0 `DELETE FROM usuarios`, 0 C# 12) y generación de documentación técnica en `docs/articles/changes/2026-10-29-mover-usuario-entre-vaults.md`.

---

# Walkthrough — Generación de Documentación Técnica y Empaquetado Release 1.6.3.52

## Resumen de Ejecución

### 1. Documentación Técnica — `/build-docs`
- **Comando:** `powershell -ExecutionPolicy Bypass -File ".\scripts\build_docs.ps1"`
- **Acciones Realizadas:**
  - Compilación y extracción de XML doc comments sin errores (0 errores).
  - Construcción del sitio web estático DocFX (`docs\_site\`).
  - Sincronización del artículo conceptual `docs\articles\academic-viaticos.md` con la regla de comisiones y pernocta en último día (`FV-FO-033` / `GF-FO-023`).

### 2. Empaquetado de Instalador Release — `/build-release 1.6.3.52`
- **Comando:** `powershell -ExecutionPolicy Bypass -File ".\scripts\build_installer.ps1" -NewVersion "1.6.3.52"`
- **Acciones Realizadas:**
  - Sincronización y verificación de versión canónica `1.6.3.52` en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
  - Limpieza de instaladores MSI previos en `installer\`.
  - Ejecución integral de la suite de pruebas unitarias:
    - **Total Tests:** 624
    - **Superados:** 624
    - **Errores:** 0
    - **Omitidos:** 0
  - Compilación del Add-In en modo `Release|x64` y empaquetado de dependencias Excel-DNA (0 errores).
  - Generación del instalador MSI mediante WiX 4 Toolset (`InstaladorConversorPDF.wixproj`).
- **Artefacto Generado:** `installer\InstaladorConversorPDFv1.6.3.52.msi` (5,464,064 bytes).

---

## Verificación de Artefactos

| Artefacto | Estado | Detalle |
| :--- | :---: | :--- |
| `installer\InstaladorConversorPDFv1.6.3.52.msi` | Generado | Instalador MSI WiX 4 con auto-activación en Excel para la versión 1.6.3.52 (5.46 MB). |
| `ConversorPDF.Tests` | 100% Pass | 624 de 624 pruebas unitarias aprobadas (0 fallos). |
| `ConversorPDF.XLL` | Compilado | Ensamblado x64 empaquetado sin errores (`ConversorPDF.XLL-AddIn64.xll`). |
| `docs\_site\` | Generado | Documentación técnica DocFX actualizada y sincronizada. |

---

## Ronda 102 — E-03: Residuo 'Docentes catedra' y Renombres en Cascada por Desincronización Posicional de Copia

### Resumen de Cambios Aplicados

#### 1. Fase 620 — Resolución Determinista de Copia por Diferencia de Nombres
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Sustitución de `wb.Worksheets[antesImp + 1]` por captura previa de nombres en `wb.Sheets` e identificación de la copia recién creada por diferencia de conjuntos.

#### 2. Fase 621 — Paracaídas Temporal Visible y Desprotección en `EliminarHojaTemporal`
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Desprotección defensiva (`Unprotect`) previa al borrado y creación de hoja paracaídas visible (`__tmp_guard_{token}`) si la hoja objetivo es la única visible en el libro.

#### 3. Fase 622 — Renombrado Defensivo de Residuo No Eliminable (`LimpiarPreviaPlantillas`)
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Renombrado automático a `__res_{token}_{name}` de hojas previas que no se puedan eliminar para liberar el nombre canónico y prevenir renombres forzados `(n)`.

#### 4. Fase 623 — Reconocimiento de Prefijos Auxiliares en `LimpiarHojasPlantilla`
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Inclusión de prefijos `__tmp_guard_` y `__res_` en el recolector de hojas pendientes de eliminación para garantizar su purga al final del ciclo.

### Archivos Afectados
- `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`

---

## Ronda 103 — Regresiones y Riesgos Residuales en Fases 620–623

### Resumen de Cambios Aplicados

#### 1. Fase 624 — Inserción al Final del Paracaídas Temporal
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Inserción explícita de `wb.Worksheets.Add(Type.Missing, wb.Worksheets[(int)wb.Worksheets.Count])` al final para evitar desplazar índices en bucle descendente.

#### 2. Fase 625 — Purga en Bloque Finally del Paracaídas y Token Extendido
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Token de 8 caracteres con warning si falla renombrado y bloque `finally` para purgar el paracaídas si no es la única hoja visible.

#### 3. Fase 626 — Identidad por Nombre Real de Copia y Aislamiento de Hojas del Usuario
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Creación y paso de `nombresCopias` a `LimpiarHojasPlantilla`. Eliminación exclusiva de nombres reales importados; desambiguación canónica restringida a pre-limpieza (`nombresImportados == null`).

#### 4. Fase 627 — Inversión de Log Falso Positivo y Desprotección en Purga
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** `Unprotect` antes de `sh.Delete()`, log `Info` únicamente tras éxito y warning detallado con nombre e índice ante fallo.

#### 5. Fase 628 — Validación Defensiva de Fallback Posicional
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Validación estricta de que la hoja en `antesImp + 1` no perteneciera a `nombresAntes` antes de aceptarla como copia de plantilla.

### Archivos Afectados
- `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`

---

## Ronda 104 — Reconciliación de Supervivientes, Robustez en Paracaídas y Sincronización XML

### Resumen de Cambios Aplicados

#### 1. Fase 629 — Reconciliación de Nombres en Hojas Auxiliares Supervivientes
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Bloque de reconciliación al final de `LimpiarHojasPlantilla` que restaura el nombre canónico en hojas `__res_` supervivientes y renombra hojas paracaídas `__tmp_guard_` a `"Hoja1"` cuando son la única hoja restante del libro.

#### 2. Fase 630 — Supresión de Lectura Problemática de `.Name` en Catch de Paracaídas
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Se eliminó la evaluación de `paracaidas.Name` dentro del catch de renombrado para asegurar la visibilidad y borrado de la hoja temporal sin interrupciones COM secundarias.

#### 3. Fase 631 — Sincronización de Documentación XML en `LimpiarHojasPlantilla`
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Actualización del comentario XML `<param name="nombresImportados">` reflejando la semántica de nombres reales de copias (Fase 626).

### Archivos Afectados
- `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`

---

## Ronda 105 — Logging Explícito en Reconciliación y Centralización de Constantes Técnicas

### Resumen de Cambios Aplicados

#### 1. Fase 632 — Trazabilidad de Fallo en Reconciliación y Normalización de Mensajes
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Se capturan excepciones de renombre con warnings explícitos reportando nombre técnico y causa; se normalizó el mensaje a `"(superviviente de limpieza)"`.

#### 2. Fase 633 — Centralización de Prefijos Auxiliares y Nombre Neutro en Constantes
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
- **Cambio:** Definición y uso de `PREFIJO_PARACAIDAS`, `PREFIJO_RESIDUO` y `NOMBRE_HOJA_NEUTRA` eliminando literales dispersos.

### Archivos Afectados
- `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`

---

## Ronda 106 — Aislamiento y Resiliencia en Exportación Concurrente Asignaturas / Semestres

### Resumen de Cambios Aplicados

#### 1. Fase 634 — Captura de Nombre de Libro Inicial y Restauración de Foco
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Cambio:** Se captura `_nombreLibroInicial` en `OnLoad` como cadena (sin retener referencias COM) y se implementa `RestaurarFocoLibroInicial()` para reactivar el libro de origen liberando todas las referencias COM intermedias.

#### 2. Fase 635 — Verificación Preventiva de Hoja Plantilla/Protegida en `btnExportar_Click`
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Cambio:** En la exportación `PorSemestres`, si la hoja activa pertenece a la plantilla (`PROGRAMACIÓN`, `NOTOCAR`) o está protegida, intenta restaurar el foco al libro inicial y, si permanece en la plantilla, crea una nueva hoja en blanco para no alterar ni destruir los datos presupuestales de la plantilla.

#### 3. Fase 636 — Fallback Defensivo con Inserción de Hoja Limpia ante Libros Protegidos o Vacíos en `SemestresExcelExporter`
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Cambio:** Manejo proactivo si `ws == null` creando un nuevo libro con `Workbooks.Add()`, y si `ws.ProtectContents == true` insertando automáticamente una hoja limpia al final del libro activo con `wb.Worksheets.Add(Type.Missing, ultimoSheet)` y transfiriendo el ownership COM de `ws` antes de ejecutar `allCells.Clear()`.

### Archivos Afectados
- `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`

---

## Ronda 107 — Inserción Determinista No Destructiva de Hoja Semestres y Cero Fugas RCW

### Resumen de Cambios Aplicados

#### 1. Fase 637 — Aislamiento Estricto de Plantilla FV-FO-033 en Exportación Semestral
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Cambio:** Si falla la restauración del libro inicial y el libro activo es la plantilla `FV-FO-033`, se crea un nuevo libro independiente mediante `_excelApp.Workbooks.Add()`, impidiendo la adición de hojas dentro de la plantilla presupuestal.

#### 2. Fase 638 — Liberación Determinista de RCW de `ActiveWorkbook` en `OnLoad`
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- **Cambio:** Se captura `wbInicial` en un bloque `finally` con llamada obligatoria a `Marshal.ReleaseComObject(wbInicial)`, eliminando la retención invisible de RCW al inicializar el formulario.

#### 3. Fase 639 — Creación Determinista de Hoja Nueva "Semestres", Validación Temprana y Liberación COM
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`
- **Cambio:** Validación temprana de fechas antes de tocar Excel; inserción determinista de una nueva hoja `"Semestres"` (con sufijo `(n)` si ya existe) al final del libro, protegiendo las hojas previas del usuario y liberando los objetos COM intermedios (`hojas`, `ultimoSheet`, `wb`) en bloques `finally`.

### Archivos Afectados
- `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
- `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`

---

## Ronda 108 — Validación del Diagnóstico y Plan Consolidado (Fases 640–644)

### Resumen de Cambios Aplicados

#### 1. Fase 640 — Filtro de Cohortes en Dos Pasadas
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambio:** Pasada 1 por coincidencia exacta de `AulaRef` (incluyendo `idprograma`); Pasada 2 tolerante a homónimos únicamente para cohortes que no tuvieron coincidencia exacta, evitando sobre-inclusión de aulas no solicitadas como la 263 y la 291.

#### 2. Fase 641 — Compuertas P07/P08 Acotadas a Cohortes Solicitadas
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambio:** Restricción de la validación de asignaturas vinculadas (`ag_aula_asignatura`) únicamente a las aulas del alcance solicitado (`refsSolicitadas`), emitiendo advertencia de diagnóstico para huérfanas fuera de la selección sin bloquear la ejecución.

#### 3. Fase 642 — Traza Temprana Real de Aulas Huérfanas en P05
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambio:** Cálculo y reporte detallado de `aulasHuerfanas` en `ConstruirContextoAsync`, registrando advertencia `AulasHuerfanasAlcance` con el listado explícito de aulas sin materias vinculadas.

#### 4. Fase 643 — Homónimos Distinguibles en el Árbol (P03)
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambio:** Adición del sufijo `(ID {prog.Id})` al nombre visible del nodo de programa cuando existan múltiples programas homónimos, facilitando la distinción visual sin afectar el enlace interno por `Tag`.

#### 5. Fase 644 — Desbloqueo Operativo Inmediato (Dato, no código)
- **Archivo:** Base de datos SQLite del operador (`ag_aula`)
- **Cambio:** Script SQL documentado para verificar dependencias en `ag_sesiones` e inactivar (`UPDATE ag_aula SET estado = 0`) o eliminar el aula huérfana 263.

### Archivos Afectados
- `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

---

## Ronda 109 — Tolerancia a Homónimos en Validación de Anclaje y Consistencia Documental (Fases 645–646)

### Resumen de Cambios Aplicados

#### 1. Fase 645 — Validación de Anclaje P07/P08 con Fallback de Cohorte
- **Archivos:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAulaFactory.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAulaFactoryTests.cs`
- **Cambio:** Implementación de `AgAulaFactory.ExtraerCohorte(aulaRef)` para deduplicar la lógica entre capas; eliminación de la función local en `ConstruirContextoAsync`; y actualización de las compuertas de anclaje multi-combinación en P07 y P08 incorporando fallback por cohorte para tolerar aulas persistidas con ID de programa homónimo. Cobertura completa de pruebas unitarias agregada.

#### 2. Fase 646 — Consistencia Documental de Artefactos de la Ronda 108
- **Archivos:** `implementation_plan.md`, `walkthrough.md`
- **Cambio:** Corrección de identificadores de método en la documentación de la Ronda 108: `EjecutarOptimizacionAsync` corregido a `IniciarOptimizacionAsync`, y `PrepararContextoAsync` corregido a `ConstruirContextoAsync`.

### Archivos Afectados
- `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAulaFactory.cs`
- `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAulaFactoryTests.cs`
- `implementation_plan.md`
- `walkthrough.md`

---

## Ronda 110 — Anclaje Mono-Combinación con Tolerancia a Homónimos (Fase 647)

### Resumen de Cambios Aplicados

#### 1. Fase 647 — Anclaje de Combinación Única (P05) con Fallback de Cohorte Homónima
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambio:** Incorporación de fallback por cohorte homónima con `AgAulaFactory.ExtraerCohorte(aulaRefEsperada)` cuando no existe coincidencia exacta por `AulaRef` en el camino mono-combinación de `ConstruirContextoAsync`, anclando `IdAulaCorrida` y evitando falsos bloqueos pre-vuelo en P07 y P08.

### Archivos Afectados
- `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- `implementation_plan.md`
- `walkthrough.md`

---

## Ronda 111 — Plan Unificado v4: Decisiones β/γ/Horas y Unificación Militar de Horarios (Fases 648–656)

### Evaluación Crítica y Resumen de Fases

#### 1. Fase 648 — Contrato Puro: `AgFranjaJornada`, `AgDiasVentana`, `HorarioMilitarizador` y Ampliación de DTOs/Modelos
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Asimetría entre formatos en UI (12 h) y motor AG (militar 24 h); carencia de mapeo determinista para franjas de 8 h continuas (TARDE/NOCHE 14:00-22:00) y de las 5 combinaciones de días.
  - *Causa Raíz:* Históricamente, las franjas se ensamblaron ad-hoc sin un contrato puro desacoplado de infraestructura.
  - *Impacto Confirmado:* La UI descarta horarios militares o los mapea a DBNull, perdiendo la persistencia en `diasHoras`.
- **Acción Planificada:**
  - Crear `AgFranjaJornada.Resolve(jornada, horasDia)`.
  - Crear `AgDiasVentana` con las 5 opciones: `LUN-JUE` (1,2,3,4), `MAR-VIE` (2,3,4,5), `JUE-SAB` (4,5,6), `VIE-SAB` (5,6), `VIE-DOM` (5,6,7).
  - Crear `HorarioMilitarizador` basado en `HorarioSesionParser.TryParse`.
  - Extender `AgNodoArbolKey`, `AgContextoSeleccionDto` y `AgAula` con `Jornada`, `HorasDia`, `DiasVentana`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgFranjaJornada.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgDiasVentana.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/HorarioMilitarizador.cs`
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgNodoArbolKey.cs`
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgContextoSeleccionDto.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgAula.cs`
  - `ConversorPDF.Tests/Academic/Domain/AgFranjaJornadaTests.cs`
  - `ConversorPDF.Tests/Academic/Domain/AgDiasVentanaTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/HorarioMilitarizadorTests.cs`

#### 2. Fase 649 — Migración Aditiva de Schema y Semilla de 42 Bloques (`MigrationV62`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La tabla `ag_bloque_horario` solo tiene 8 bloques registrados y faltan columnas de decisión en `ag_aula` y `ag_arbol_nodo`.
  - *Causa Raíz:* La semilla inicial de V49/V55 no contemplaba sesiones de 8 h ni cubría todos los días de la semana (1..7).
  - *Impacto Confirmado:* El AG no encuentra bloques para franjas nocturnas o de 8 h en días específicos.
- **Acción Planificada:**
  - Crear migración aditiva `MigrationV62_AddAgAulaVentanasYBloquesSemilla` añadiendo columnas `jornada`, `horas_dia`, `dias_ventana` a `ag_aula` y `ag_arbol_nodo`.
  - Sembrar de forma idempotente 6 bloques/día × 7 días = 42 bloques en `ag_bloque_horario` mediante `WHERE NOT EXISTS`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV62_AddAgAulaVentanasYBloquesSemilla.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/Migrations/MigrationV62Tests.cs`

#### 3. Fase 650 — Migración de Datos para Unificación a Formato Militar (`MigrationV63`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Datos históricos en `vinculaciones.dia_horas_semanales` y `sesiones.horariosesion` almacenan cadenas en 12 h AM/PM (`"1:00 PM A 5:00 PM"`).
  - *Causa Raíz:* Concurrencia histórica de formatos entre legacy e importaciones Excel.
  - *Impacto Confirmado:* Desalineación y fallos de match al recargar vinculaciones y sesiones en los controles.
- **Acción Planificada:**
  - Crear `MigrationV63_UnificarHorariosMilitar` para normalizar `sesiones.horariosesion` y la 3ra parte de `vinculaciones.dia_horas_semanales` a `"HH:mm A HH:mm"` usando `HorarioMilitarizador`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV63_UnificarHorariosMilitar.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/Migrations/MigrationV63Tests.cs`

#### 4. Fase 651 — Sincronización Militar en `DayHoursDropdown` y Serialización 3-Partes en `AgMaterializadorSesiones`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `DayHoursDropdown` descarta valores militares por validación estricta contra items 12h; `AgMaterializadorSesiones` genera 2 partes (`DIA|HORAS`) omitiendo el horario militar y asume sesiones fijas de 4h en I2 (`sesionesGrupo * 4.0`).
  - *Causa Raíz:* Omisión de la tercera parte en serialización y rigidez de 4.0 h en el validador I2.
  - *Impacto Confirmado:* Desfase visual en el dropdown y falsos positivos de cuarentena I2 en sesiones de 8 h.
- **Acción Planificada:**
  - Actualizar colecciones permitidas en `DayHoursDropdown.cs` a formato militar.
  - Serializar `dia_horas_semanales` como `DIA|HORAS|HORARIO` militar en `AgMaterializadorSesiones.cs`.
  - Actualizar Invariante I2 para evaluar `Sum(s => s.Horas) == esperado`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/UI/DayHoursDropdown.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs`

#### 5. Fase 652 — Respeto de Decisiones en Fábrica AG y Operadores Genéticos
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La mutación y el cruce alteran bloques horarios incluso cuando el usuario ha fijado una decisión explícita de jornada y horas en el aula.
  - *Causa Raíz:* Ausencia de flag de congelamiento temporal en `AgGen`.
  - *Impacto Confirmado:* Deriva del horario fijado durante la optimización genética.
- **Acción Planificada:**
  - Agregar `AgGen.EsFijoTemporal`.
  - En `AgFabricaIndividuos`: asignar bloque fijo resuelto por `AgFranjaJornada.Resolve` cuando la decisión esté presente.
  - Proteger genes con `EsFijoTemporal` en `AgOperadorMutacion` y `AgOperadorCruce`, permitiendo optimizar únicamente la asignación docente.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorMutacionTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorCruceTests.cs`

#### 6. Fase 653 — OwnerDraw en Árbol Académico (Nivel 6: 5 Grupos de Radios y Cabeceras)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Imposibilidad de que el usuario seleccione la decisión de ventana, jornada y horas por cohorte en el árbol visual.
  - *Causa Raíz:* El árbol solo permitía selección binaria por nodo sin pseudo-columnas.
  - *Impacto Confirmado:* No existía interfaz para expresar la decisión β/γ/Horas antes de la corrida.
- **Acción Planificada:**
  - Agregar cabeceras en `pnlArbolToolbar`.
  - Implementar OwnerDraw (`treeContextoAcademico_DrawNode`) para Nivel 6 con `RadioButtonRenderer`.
  - Implementar hit-test con exclusividad por dimensión en `treeContextoAcademico_NodeMouseClick`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`

#### 7. Fase 654 — Persistencia de Perfiles en `ag_arbol_nodo` y Sincronización en `AgArbolMapper`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al guardar perfiles del árbol en `ag_arbol`, las decisiones de jornada, horas_dia y dias_ventana no se persistían.
  - *Causa Raíz:* `AgArbolMapper` y el repositorio ignoraban los nuevos campos de `AgNodoArbolKey`.
  - *Impacto Confirmado:* Pérdida de configuración al recargar perfiles.
- **Acción Planificada:**
  - Mapear bidireccionalmente los nuevos campos en `AgArbolMapper` y `AgArbolRepository`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgArbolMapper.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgArbolRepository.cs`
  - `ConversorPDF.Tests/Academic/Application/Mappers/AgArbolMapperTests.cs`

#### 8. Fase 655 — Pre-Vuelo: Validación de Ventana y Existencia de Bloque Horario
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Corridas con configuraciones inconsistentes (`dias_patron ⊄ ventana` o bloque no sembrado) avanzaban al motor AG y fallaban tardíamente.
  - *Causa Raíz:* Ausencia de compuerta pre-vuelo específica para la consistencia ventana/patrón y disponibilidad de bloques.
  - *Impacto Confirmado:* Fallos tardíos de optimización o cromosomas vacíos.
- **Acción Planificada:**
  - Crear e integrar `AgValidadorVentana` en diagnósticos pre-vuelo P07 y P08 de `AgFormGeneradorHorarios.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs`

#### 9. Fase 656 — Documentación Técnica de Arquitectura de Horarios
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Desalineación de la documentación formal frente a la unificación de franjas y formato militar.
  - *Causa Raíz:* Evolución sucesiva de versiones del diseño.
  - *Impacto Confirmado:* Documentación desactualizada para futuros desarrollos.
- **Acción Planificada:**
  - Actualizar `docs/articles/changes/2026-ag-horarios-academicos.md` consolidando el diseño definitivo.
- **Archivos Afectados:**
  - `docs/articles/changes/2026-ag-horarios-academicos.md`
  - `implementation_plan.md`
  - `walkthrough.md`

### Verificaciones de Principios (Ronda 111)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F648 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F649 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F650 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F651 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F652 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F653 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F654 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F655 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F656 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 112 — Plan v4: Gaps de Lógica en Congelamiento Temporal (Fases 657–659)

### Evaluación Crítica y Resumen de Fases

#### 1. Fase 657 — Preservación de `EsFijoTemporal` en `AgGen.Clonar()`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La clonación de cromosomas en `AgOperadorCruce` provoca que los genes de los hijos pierdan el flag `EsFijoTemporal`.
  - *Causa Raíz:* Omisión de la propiedad `EsFijoTemporal` en la inicialización de objeto dentro de `AgGen.Clonar()`.
  - *Impacto Confirmado:* A partir de la primera generación post-cruce, los genes congelados temporalmente vuelven a mutar y a intercambiar franjas.
- **Acción Planificada:**
  - Agregar `EsFijoTemporal = this.EsFijoTemporal` en `AgGen.Clonar()`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`
  - `ConversorPDF.Tests/Academic/Domain/AgGenTests.cs`

#### 2. Fase 658 — Cobertura Integral de Mutación para Genes con `EsFijoTemporal`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El operador de mutación sigue alterando `IdAula` en genes con decisión explícita, desplazándolos a otra cohorte.
  - *Causa Raíz:* El guard en F652 solo se contempló para la mutación de bloque (L108-122), omitiendo el bloque de mutación de aula (L80-106).
  - *Impacto Confirmado:* Ruptura de coherencia entre la cohorte seleccionada y la ventana horaria decidida.
- **Acción Planificada:**
  - Insertar `if (gen.EsFijoTemporal) continue;` al inicio del bucle por gen en `AgOperadorMutacion` tras L77, protegiendo aula y bloque en conjunto sin afectar la mutación docente previa (L28-73).
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorMutacionTests.cs`

#### 3. Fase 659 — Protección de Genes Congelados en Cascada de Cruce `HeredarOX`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El operador de cruce sobrescribe bloque horario y patrón semanal en genes congelados durante la recombinación OX.
  - *Causa Raíz:* Los genes congelados tienen `Estado == Valido` (no `Fijado`), evadiendo el filtro de L107 y cayendo en la cascada L121-214.
  - *Impacto Confirmado:* La recombinación genética destruye la decisión temporal fijada por el usuario.
- **Acción Planificada:**
  - Insertar `if (dst.EsFijoTemporal) continue;` tras L107 en `AgOperadorCruce.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorCruceTests.cs`

---

 ## Ronda 114 — Fase 660: Colisión de UDF `Normalizar` por Auto-Registro Excel-DNA

 ### Resumen de Cambios Aplicados

 #### 1. Fase 660 — Renombrado de Métodos `Normalizar` a Nombres Específicos
 - **Estado:** ACEPTADA
 - **Hallazgo:**
   - *Diagnóstico Crítico:* Al cargar el complemento en Excel, aparecían múltiples advertencias modales `Repeated function name: 'Normalizar'`.
   - *Causa Raíz:* `ConversorPDF.XLL-AddIn64.dna` auto-registra todos los métodos `public static` compatibles como UDFs al no tener `ExplicitRegistration="true"`. Coexistían `DiaSemanaNormalizador.Normalizar`, `AgFranjaJornada.Normalizar` y `HorarioMilitarizador.Normalizar`.
   - *Impacto Confirmado:* Error de inicialización y ambigüedad de UDFs en Excel.
 - **Acción Aplicada:**
   - `AgFranjaJornada.cs`: renombrado `Normalizar` a `NormalizarJornada`.
   - `HorarioMilitarizador.cs`: renombrado `Normalizar` a `NormalizarHorario`.
   - `MigrationV63_UnificarHorariosMilitar.cs`: actualizadas llamadas en L36, L69, L105, L120.
   - `DayHoursDropdown.cs`: actualizada llamada en L176.
   - `HorarioMilitarizadorTests.cs`: actualizadas llamadas en pruebas unitarias.
 - **Archivos Afectados:**
   - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgFranjaJornada.cs`
   - `ConversorPDF.XLL/Academic/Application/Services/HorarioMilitarizador.cs`
   - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV63_UnificarHorariosMilitar.cs`
   - `ConversorPDF.XLL/UI/DayHoursDropdown.cs`
   - `ConversorPDF.Tests/Academic/Application/Services/HorarioMilitarizadorTests.cs`

 ### Verificaciones de Principios (Ronda 114)

 | Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
 |---|---|---|---|---|
 | F660 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 115 — Fase 661: Corrección de UX en Nivel 6 de Árbol Académico (Etiquetas, Separadores, Alineación y Repintado)

### Resumen de Cambios Aplicados

#### 1. Fase 661 — Alineación Dinámica, Etiquetas Completas y Repintado de Fila en Nivel 6
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En el Nivel 6 del árbol, los radio buttons se dibujaban con etiquetas abreviadas corridas (`"M"`, `"T"`, `"N"`, etc.) sin separadores entre dimensiones de decisión, las cabeceras en `pnlArbolToolbar` tenían textos estáticos abreviados en posiciones fijas desalineadas respecto a `bounds.X`, el ancho de ventana (984 px) no permitía desplegar las columnas completas (~1060 px), y al hacer clic solo se invalidaba `e.Node.Bounds` (el texto de la modalidad), impidiendo el repintado visual inmediato de los radios.
  - *Causa Raíz:* Renderizado compacto de `DrawNode` con hit-test impreciso, coordenadas absolutas estáticas de cabeceras en la toolbar e invalidación parcial de nodo en `NodeMouseClick`.
  - *Impacto Confirmado:* Deficiente experiencia de usuario, recorte visual de opciones, desalineación espacial entre cabeceras y columnas y falta de respuesta visual inmediata tras el clic.
- **Acción Aplicada:**
  - `AgFormGeneradorHorarios.Designer.cs`:
    - Actualizado `ClientSize` a `(1120, 661)` y contenedores hijos (`tabControlPrincipal`, `grpSeleccion`, `pnlArbolToolbar`, `treeContextoAcademico`) a 1072 px de ancho.
    - Asignado `treeContextoAcademico.Indent = 14`.
    - Cabeceras renombradas a textos limpios: `"Jornada:"`, `"Días:"`, `"Horas/día:"`.
  - `AgFormGeneradorHorarios.cs`:
    - En `TreeContextoAcademico_DrawNode`: renderizado de textos completos (`"Mañana"`, `"Tarde"`, `"Noche"`, `"Lun-Jue"` ... `"Vie-Dom"`, `"4h"`, `"8h"`), espaciado ampliado y nuevo helper `DibujarSeparadorVertical` en X=+430 y X=+945.
    - En `NodeMouseClick`: invalidación de fila completa `new Rectangle(0, r.Y, treeContextoAcademico.ClientSize.Width, r.Height)` para refresco inmediato de radios.
    - En `ProcesarClickDimensionNivel6`: rangos exactos de hit-test por grupo con clamp superior estricto y no-op defensivo fuera de zona.
    - En `ConfigurarEventos`: adición de paneles separadores de 1 px en la toolbar (`_pnlSepToolbar1`, `_pnlSepToolbar2`), evento `MouseMove` para cursor `Hand` en zonas interactivas, y suscripción a `AfterExpand` para sincronizar cabeceras.
    - Método dinámico `AlinearCabecerasNivel6()` para alinear perfectamente cabeceras y separadores de toolbar con el X de las columnas de Nivel 6.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`

### Verificaciones de Principios (Ronda 115)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F661 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F662 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F663 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

#### 2. Fase 662 — Compresión Compacta de Columnas de Nivel 6 (Prevención de Desborde a la Derecha)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La suma del X base (`~103`) y el offset de `8h` (`1015`) alcanzaba `1118 px`, sobrepasando el ancho cliente utilizable del árbol (`~1055 px` con scroll vertical) y recortando el grupo Horas/día y el extremo de Vie-Dom.
  - *Causa Raíz:* Espaciados horizontales excesivamente holgados entre radios (100 px en grupo Días).
  - *Impacto Confirmado:* Recorte de controles en el margen derecho e imposibilidad de visualizar/interactuar limpiamente con 4h/8h.
- **Acción Aplicada:**
  - `AgFormGeneradorHorarios.cs`:
    - Comprimidos offsets a: Jornada (+150, +215, +275), Sep (+360), Días (+375, +450, +525, +600, +675), Sep (+775), Horas (+790, +855). Extremo derecho en 983 px (< 1055 px).
    - Actualizados rangos de hit-test en `ProcesarClickDimensionNivel6`.
    - Actualizadas posiciones dinámicas en `AlinearCabecerasNivel6`.
    - Ajustado rango interactivo de cursor en `MouseMove`.
  - `AgFormGeneradorHorarios.Designer.cs`:
    - Actualizadas posiciones iniciales en diseñador de `lblColJornada` (253), `lblColDias` (478) y `lblColHorasDia` (893).
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`

#### 3. Fase 663 — Render Condicionado al `Checked` de Modalidad y Selección Automática de Fila en Nivel 6
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Todas las filas de Nivel 6 (`Presencial`, `Virtual`, `Semipresencial`) renderizaban los radios seleccionados (`Mañana`, `Mar-Vie`, `4h`), dando la impresión errónea de que las 3 modalidades estaban simultáneamente activas.
  - *Causa Raíz:* `DrawNode` leía las opciones almacenadas en `Tag` sin condicionar el estado marcado al flag `e.Node.Checked`.
  - *Impacto Confirmado:* Confusión de usuario al ver radios encendidos en filas no seleccionadas.
- **Acción Aplicada:**
  - `AgFormGeneradorHorarios.cs`:
    - En `TreeContextoAcademico_DrawNode`: radios condicionados a `filaActiva && ...`. Filas desmarcadas muestran los radios sin seleccionar.
    - En `ProcesarClickDimensionNivel6`: marcar automáticamente la modalidad si no estaba seleccionada y aplicar exclusividad de modalidad.
    - En `NodeMouseClick` y `TreeContextoAcademico_AfterCheck`: invalidar las filas de todas las modalidades hermanas para refrescar inmediatamente el estado visual de sus radios.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

---

## Ronda 116 — Fase 664: Aislamiento Estricto de Asignaturas por Programa y Cohorte en AG (Prevención de Cruces Multi-Programa)

### Resumen de Cambios Aplicados

#### 1. Fase 664 — Aislamiento Curricular por Programa en Fábrica AG, Operadores Genéticos, Validador y Staging
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al abrir `FormVinculaciones` en Modo AG y seleccionar registros para revisar en `FormSesionesAcademicas`, asignaturas de un programa aparecían asignadas a cohortes de otro programa (ej. en APT-53 Semestre VII B se asignó `FINANZAS PUBLICAS` de APT-52, y en APT-52 Semestre X A se asignó `Opción de Grado` de APT-53). En consecuencia, al abrir `FormSesionesAcademicas` el selector de asignaturas quedaba en blanco porque filtraba por el catálogo del programa del aula y no encontraba la asignatura foránea; además, la asignatura propia de APT-52 (`TRABAJO DE GRADO`) quedaba excluida.
  - *Causa Raíz:* En `AgFabricaIndividuos.CrearAleatorio`, al iterar `gruposAula` y `contexto.Asignaturas`, el filtro emparejaba asignaturas únicamente por el numeral romano del semestre (`semAulaRom == semAsigRom`), omitiendo validar que `aula.IdPrograma == asig.IdPrograma` y sin restringir a las asignaturas de `ag_aula_asignatura` si existían. De manera similar, `AgOperadorMutacion` no filtraba `a.IdPrograma` al mutar aulas, `AgOperadorCruce` (`HeredarOX`) no verificaba coincidencia de programa antes de heredar `IdAsignatura`, y `AgValidadorReglas` (R4) no penalizaba discrepancias de programa entre aula, gen y asignatura.
  - *Impacto Confirmado:* Generación de soluciones híbridas con materias cruzadas entre programas en la base de datos temporal, combos en blanco en `FormSesionesAcademicas` e inconsistencias en la grilla de vinculaciones.
- **Acción Aplicada:**
  - `AgFabricaIndividuos.cs`:
    - Incorporado helper `ProgramaDeAula(AgAula a)` que extrae `IdPrograma` de la entidad o como fallback de `AulaRef` (`partes[2]`).
    - Actualizado `ClaveCohorteAula` para incorporar `IdPrograma` en la clave natural (`$"{prog}|{sem}|{grp}"`), impidiendo colisiones entre cohortes homónimas de programas distintos.
    - En el bucle de aulas y asignaturas:
      1. Restricción estricta a `asigsValidasAula` obtenidas de `contexto.AulaAsignaturas` (`ag_aula_asignatura`) si la cohorte posee asignaturas configuradas.
      2. Validación incondicional de programa: `idProgAula > 0 && idProgAsig > 0 && idProgAula != idProgAsig => continue`.
      3. Conservación de filtro por semestre romano.
    - Al instanciar `AgGen` (tanto en gen elemental como en multi-día): fijar `IdPrograma = asig.IdPrograma > 0 ? asig.IdPrograma : ...` sincronizado a la asignatura.
  - `AgOperadorMutacion.cs`:
    - En la selección de aulas candidatas para mutación de aula (L90): añadido filtro `(a.IdPrograma <= 0 || gen.IdPrograma <= 0 || a.IdPrograma == gen.IdPrograma)`.
  - `AgOperadorCruce.cs`:
    - En `HeredarOX` (L113-120): añadido guard de programa `progAsig > 0 && progAula > 0 && progAsig != progAula => continue` antes de transferir `IdAsignatura`.
  - `AgValidadorReglas.cs`:
    - En regla R4 (`CompatibilidadAula`): añadido `bool progOk = aula.IdPrograma <= 0 || gen.IdPrograma <= 0 || aula.IdPrograma == gen.IdPrograma`.
    - Añadida validación de correspondencia entre `asigEval.IdPrograma` y `gen.IdPrograma`, registrando violación dura si difieren.
  - `AgGenToAgVinculacionMapper.cs`:
    - En `MapGrupo`: resolución defensiva de `IdPrograma` a partir de `contexto.AsignaturasPorId` de la asignatura.
  - `AgMaterializadorSesiones.cs`:
    - Añadida guarda defensiva fail-fast previa a la persistencia en staging: si `aCohorte.IdPrograma != agVinc.IdPrograma`, abortar materialización con `Result.Failure`.
  - `AgFabricaIndividuosTests.cs`:
    - Añadida prueba unitaria `CrearAleatorio_WhenMultiPrograma_ShouldNotAssignSubjectsFromOtherPrograms`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`

### Verificaciones de Principios (Ronda 116)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F664 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 117 — Fases 665 - 669: Soporte de Pernocta en Último Día (`isPernoctaUltimoDia`) y Ajuste de Duración y Rango en Bloques de Viáticos

### Resumen de Cambios Aplicados

#### 1. Fase 665 — Propagación de Columna y Vista (Plumbing de Persistencia y DTO)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La vista `vw_SesionesDetalle` no exponía `isPernoctaUltimoDia` y `SesionListItem` no la capturaba, impidiendo que el motor de viáticos conociera el flag al evaluar sesiones.
  - *Causa Raíz:* La columna no estaba mapeada en la vista SQL ni en el DTO `SesionListItem`.
  - *Impacto Confirmado:* Imposibilidad de propagar el flag desde la base de datos hacia los bloques de viáticos.
- **Acción Aplicada:**
  - En `AcademicSchema.sql`: agregada `v.isPernoctaUltimoDia AS IsPernoctaUltimoDia` a `vw_SesionesDetalle`.
  - Creada `MigrationV64_AddPernoctaUltimoDia.cs`: agrega columna idempotente en `vinculaciones` y recrea la vista `vw_SesionesDetalle`.
  - Registrada en `SqliteHelper.cs` dentro de `BuildMigrator()`.
  - En `SesionListItem.cs`: añadida propiedad `public int IsPernoctaUltimoDia { get; set; }`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Sql/AcademicSchema.sql`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV64_AddPernoctaUltimoDia.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/SesionListItem.cs`

#### 2. Fase 666 — Lógica de Negocio del Bloque de Viáticos (`ViaticoBloque.cs`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La duración y visualización de viáticos omitían el día extra cuando el docente pernoctaba el último día de clase.
  - *Causa Raíz:* `DuracionTotal` y `DisplayText` calculaban directamente sobre `FechaMaxima` sin considerar el flag de pernocta.
  - *Impacto Confirmado:* Cálculo subestimado de días de viático (1.5 en vez de 2.5 días en fines de semana con pernocta).
- **Acción Aplicada:**
  - `ViaticoBloque.cs`:
    - Incorporada propiedad `TienePernoctaUltimoDia => Sesiones != null && Sesiones.Any(s => s.IsPernoctaUltimoDia == 1)`.
    - Incorporada propiedad `FechaFinComision => FechaMaxima?.AddDays(TienePernoctaUltimoDia ? 1 : 0)`.
    - Preservada `FechaMaxima` intacta para el algoritmo de consecutividad `AgruparConsecutivos`.
    - Recalculada `DuracionTotal` usando `FechaFinComision`.
    - Actualizado `DisplayText` con `FechaFinComision`.
    - Limpieza de campos cacheados en `Agregar` e `InvalidateCache`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/ViaticoBloque.cs`

#### 3. Fase 667 — Consumidores del Rango y Duración de Viáticos
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La UI del árbol, el exportador Excel/PDF, el timeline y el analizador de conflictos mostraban la fecha final sin el día de pernocta y usaban redondeo bancario perdiendo el día extra (`Round(2.5) == 2`).
  - *Causa Raíz:* Consumo directo de `bloque.FechaMaxima` y uso de `Math.Round(b.DuracionTotal)` sin midpoint rounding.
  - *Impacto Confirmado:* Discrepancias visuales en el árbol y en los documentos oficiales generados.
- **Acción Aplicada:**
  - `FormVerVinculaciones.cs`: `fechaFinStr` proyectado desde `bloque.FechaFinComision ?? bloque.FechaMaxima`.
  - `ViaticosIndividualExporter.cs`: `fFin = b.FechaFinComision ?? b.FechaMaxima ?? DateTime.Now` y `Math.Round(b.DuracionTotal, MidpointRounding.AwayFromZero)`.
  - `ViaticoTimelineMapper.cs`: `fechaFin = bloque.FechaFinComision ?? bloque.FechaMaxima ?? fechaIni` y rango de tooltip actualizado.
  - `ViaticoConflictAnalyzer.cs`: rango en tooltip y método `SeSolapan` extendidos con `FechaFinComision ?? FechaMaxima`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`
  - `ConversorPDF.XLL/VisTimeline/Application/Services/ViaticoTimelineMapper.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/ViaticoConflictAnalyzer.cs`

#### 4. Fase 668 — Prerrequisito UI y Persistencia (`FormSesionesAcademicas`, `Vinculacion`, `AcademicRepository`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El usuario no podía marcar ni persistir la pernocta desde `FormSesionesAcademicas`.
  - *Causa Raíz:* Inexistencia del control `chkPernoctaUltimoDia`, omisión de la propiedad en `Vinculacion.cs` y en las consultas de `AcademicRepository.cs`.
  - *Impacto Confirmado:* Imposibilidad de activar el flag para docentes que pernoctan.
- **Acción Aplicada:**
  - `FormSesionesAcademicas.Designer.cs`: control `chkPernoctaUltimoDia` en `Point(356, 143)` y ajuste armónico de `btnGenerar` a `Point(495, 137)` con tamaño `(145, 30)` sin solapamiento.
  - `FormSesionesAcademicas.cs`: guardado en `PrepararVinculacionDesdeForm` y restauración en `CargarEstadoPrevio`.
  - `Vinculacion.cs`: agregada propiedad `public int IsPernoctaUltimoDia { get; set; }`.
  - `AcademicRepository.cs`: soporte en `GetVinculaciones()`, `InsertVinculacion()`, `GetVinculacionById()` y `UpdateVinculacion()`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.Designer.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Vinculacion.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`

#### 5. Fase 669 — Pruebas Unitarias de ViaticoBloque y Migración V64
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Necesidad de validar matemáticamente la Regla A y garantizar la resiliencia e idempotencia de la migración V64.
  - *Causa Raíz:* Nuevas reglas de cálculo y evolución de esquema relacional.
  - *Impacto Confirmado:* Blindaje contra regresiones y fallos en bases de datos legadas.
- **Acción Aplicada:**
  - `ViaticoBloquePernoctaTests.cs`: 6 pruebas unitarias completas (sin flag 1.5, con flag 2.5, 1 día 1.5, dos bloques recurrentes, bloque ajeno, preservación de `FechaMaxima`).
  - `MigrationV64Tests.cs`: 4 pruebas unitarias (metadatos, BD previa, idempotencia, exposición en vista).
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Domain/ViaticoBloquePernoctaTests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV64Tests.cs`

### Verificaciones de Principios (Ronda 117)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F665 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F666 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F667 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F668 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F669 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.76 Ronda 118 — Verificación de Fases 665-669 (Pernocta Último Día) y Bloqueante de Compilación

### Evaluación Crítica de las Fases

#### 1. Fase 670 — Migración V64 no registrada en el proyecto (BLOCKER de Compilación)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La clase `MigrationV64_AddPernoctaUltimoDia` es instanciada en `SqliteHelper.cs:365`, pero el archivo no formaba parte de `<Compile Include>` en `ConversorPDF.XLL.csproj`.
  - *Causa Raíz:* El proyecto `ConversorPDF.XLL` no es SDK-style y no incluye wildcards automáticos.
  - *Impacto Confirmado:* Error de compilación `CS0246` en cualquier build del proyecto.
- **Acción Aplicada:** Agregada la línea `<Compile Include="Academic\Infrastructure\Persistence\Migrations\MigrationV64_AddPernoctaUltimoDia.cs" />` en `ConversorPDF.XLL.csproj`.
- **Archivo Afectado:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj`

#### 2. Fase 671 — `chkPernoctaUltimoDia` sin reset en la ruta "sin vinculación previa"
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Si `repo.GetVinculacionById(idVinculacion)` falla (no existe vinculación previa), el early-return no restablecía los checkboxes a su estado por defecto.
  - *Causa Raíz:* Se omitió el reseteo explícito de controles de pernocta y transporte aéreo en la rama de fallo.
  - *Impacto Confirmado:* Fuga de estado en instancias recicladas o reabiertas de `FormSesionesAcademicas`.
- **Acción Aplicada:** Incorporado reseteo vía `this.Invoke` de `chkTerminalAereo.Checked = false; chkPernoctaUltimoDia.Checked = false; pnlTransporteAereo.Visible = false;`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

#### 3. Fase 672 — Caché de `ViaticoBloque` dependiente de `IsPernoctaUltimoDia`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La caché `_cachedDuracionTotal` no reaccionaba ante cambios in-place del flag `IsPernoctaUltimoDia` en sesiones ya presentes en la lista interna.
  - *Causa Raíz:* La invalidación dependía exclusivamente de `_dirty` establecido en `Agregar()` o `InvalidateCache()`.
  - *Impacto Confirmado:* Lectura de valor obsoleto de días tras activar pernocta en memoria.
- **Acción Aplicada:** Incorporado `_cachedPernocta` en `ViaticoBloque` y verificación en `DuracionTotal`. Se añadió prueba unitaria `DuracionTotal_WhenSesionPernoctaMutatesInPlace_ShouldRecalculateWithoutCallingAgregar`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/ViaticoBloque.cs`
  - `ConversorPDF.Tests/Academic/Domain/ViaticoBloquePernoctaTests.cs`

#### 4. Fase 673 — Clarificación semántica en `ViaticosIndividualExporter`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `Math.Round(b.DuracionTotal, MidpointRounding.AwayFromZero)` transforma `0.5 -> 1`, representando días calendario de comisión para la carátula/objeto del viaje.
  - *Causa Raíz:* Diferencia conceptual entre días calendario de comisión y días calculados de viático.
  - *Impacto Confirmado:* Sin impacto en runtime (`dur > 1` se comporta igual), pero requería documentación en código.
- **Acción Aplicada:** Comentario explicativo en línea 435 de `ViaticosIndividualExporter.cs`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/ViaticosIndividualExporter.cs`

#### 5. Fase 674 — Acotación de `SeSolapan` en `ViaticoConflictAnalyzer`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Evaluar solape con `FechaFinComision` (+1 día de pernocta) creaba solapes con sesiones de otra sede en días sin sesión real en el bloque A, donde `EsCruceSalvado` no podía operar porque busca pares por `IdSesion`.
  - *Causa Raíz:* Conflicto de semántica entre ventana de traslado vs. conflicto de sesiones de clase concurrentes.
  - *Impacto Confirmado:* Falsos positivos no salvables por el usuario en la UI.
- **Acción Aplicada:** Acotado el método `SeSolapan` para comparar `FechaMaxima` (días efectivos de clase), manteniendo `FechaFinComision` en la descripción del cruce.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/ViaticoConflictAnalyzer.cs`

#### 6. Fase 675 — Margen de seguridad UI entre `chkPernoctaUltimoDia` y `btnGenerar`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Posible roce o solape visual entre el texto del checkbox y `btnGenerar` en configuraciones DPI específicas.
  - *Causa Raíz:* Separación horizontal de solo 4-8px entre ambos controles.
  - *Impacto Confirmado:* Riesgo de corte tipográfico o solape de foco.
- **Acción Aplicada:** Reubicado `btnGenerar` en `Location = (500, 137)` y `Size = (140, 30)` en `FormSesionesAcademicas.Designer.cs`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.Designer.cs`

#### 7. Fase 676 — Documental: MODO AG sin soporte de pernocta
- **Estado:** ACEPTADA (Documental)
- **Hallazgo:**
  - *Diagnóstico:* `ag_vinculaciones` no contiene la columna `isPernoctaUltimoDia`. El modo AG funciona con valor 0 por defecto sin error de runtime.
  - *Acción:* Registrado como observación pendiente para paridad en MODO AG si el negocio lo requiere.

---

### Verificaciones de Principios (Ronda 118)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F670 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F671 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F672 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F673 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F674 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F675 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F676 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.77 Ronda 119 — Verificación de Fases 670-676 (Correcciones Ronda 118)

### Evaluación Crítica de las Fases

#### 1. Fase 677 — `TienePernoctaUltimoDia` se reevalúa en cada acceso a la duración
- **Estado:** RETIRADA (Diseño intencional)
- **Hallazgo:**
  - *Diagnóstico Crítico:* La propuesta de cachear `TienePernoctaUltimoDia` invalidaría la reactividad ante mutaciones in-place del flag de pernocta en las sesiones del bloque (test `DuracionTotal_WhenSesionPernoctaMutatesInPlace_ShouldRecalculateWithoutCallingAgregar`), provocando falsos valores cacheados a menos que se llame explícitamente `InvalidateCache()`.
  - *Causa Raíz:* La clave `_cachedPernocta` introducida en F672 fue concebida precisamente para detectar cuándo `TienePernoctaUltimoDia` cambia sin requerir invalidación manual.
  - *Impacto Confirmado:* `Sesiones.Any` sobre un máximo de ≤8 elementos toma tiempo despreciable (sub-microsegundo). Se descarta para proteger la reactividad in-place.

#### 2. Fase 678 — `this.Invoke` del reset sin guarda de handle y fuera del `try`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La invocación `this.Invoke(...)` antes del bloque `try` arriesgaba lanzar `InvalidOperationException` si el handle aún no estaba creado, abortando la asignación del `id_periodo` por defecto.
  - *Causa Raíz:* Se ubicó el reset fuera del `try` existente en F671.
  - *Impacto Confirmado:* Blindaje defensivo implementado.
- **Acción Aplicada:** El reseteo de `chkTerminalAereo`, `chkPernoctaUltimoDia` y `pnlTransporteAereo` fue trasladado dentro del `try` y condicionado a `if (IsHandleCreated)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

---

### Verificaciones de Principios (Ronda 119)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F677 | [X] N/A (Retirada) | [X] Sí | [X] Sí | [X] Sí |
| F678 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.78 Ronda 120 — Fases 679 a 684: "Pernocta Último Día" en Nivel 2 y Menú Contextual (FormVerVinculaciones)

### Evaluación Crítica de las Fases

#### 1. Fase 679 — Repositorio y Dominio Base
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La actualización de pernocta desde el menú contextual requería un método atómico en el repositorio sin necesidad de reconstruir la entidad completa `Vinculacion`, y `ViaticoBloque` requería exponer los `IdsVinculacion` involucrados.
  - *Causa Raíz:* Funcionalidad inexistente en la interfaz del repositorio y en el modelo de bloque.
  - *Impacto Confirmado:* Permite alternar pernocta atómicamente por lote de vinculaciones.
- **Acción Aplicada:** Agregada la firma `SetPernoctaUltimoDia` en `IAcademicRepository`, implementada en `AcademicRepository` y expuesta la propiedad `IdsVinculacion` en `ViaticoBloque`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IAcademicRepository.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/ViaticoBloque.cs`

#### 2. Fase 680 — Nivel 2: "Cierre" y Formato del Bloque
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El nodo de bloque requería reflejar el horario de cierre (último día y horario más tardío) y el sufijo condicional `| Pernocta Ultimo Dia`.
  - *Causa Raíz:* Ausencia de servicio de formateo dedicado para bloques de viático.
  - *Impacto Confirmado:* Formato homogéneo, modular y testeable sin dependencias de WinForms.
- **Acción Aplicada:** Creado `ViaticoBloqueFormatter` con `ObtenerTextoCierre` y `FormatearNodoNivel2`. Registrado en `ConversorPDF.XLL.csproj`. Integrado en `FormVerVinculaciones.ConstruirArbolViaticos`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/ViaticoBloqueFormatter.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`

#### 3. Fase 681 — Menú Contextual en `FormVerVinculaciones`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* No existía acción contextual en el TreeView para marcar o desmarcar la pernocta en bloques de viático.
  - *Causa Raíz:* Menú contextual solo soportaba imprimir viáticos o editar sesiones individuales.
  - *Impacto Confirmado:* Interacción rápida en 1 clic desde la vista de Viáticos.
- **Acción Aplicada:** Creado `itemPernocta` en `ConfigurarMenuContextual`, con visibilidad condicionada a `_vistaActual == VistaVinculaciones.Viaticos && !_modo.EsModoAg` y texto dinámico ("🛏️ Quitar..." / "🛏️ Marcar...").
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`

#### 4. Fase 682 — Handler `AlternarPernoctaUltimoDiaAsync`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La alternancia debía actualizar la base de datos de manera transaccional, realizar flush al vault, reflejarse en memoria en `_items` y reconstruir el árbol manteniendo la posición y foco del usuario.
  - *Causa Raíz:* Nuevo flujo interactivo de mutación en la vista de consulta.
  - *Impacto Confirmado:* Flujo robusto con confirmación interactiva, rollback ante fallos y actualización instantánea en UI.
- **Acción Aplicada:** Implementado `AlternarPernoctaUltimoDiaAsync` con confirmación dialogada, transacción SQL, persistencia en vault, mutación en memoria y recarga del árbol con restauración de nodo seleccionado.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`

#### 5. Fase 683 — MODO AG (Aislamiento y Documental)
- **Estado:** ACEPTADA (Documental)
- **Hallazgo:**
  - *Diagnóstico:* MODO AG no contiene la columna `isPernoctaUltimoDia` en `ag_vinculaciones`.
  - *Acción:* La opción se mantiene oculta en MODO AG (`!_modo.EsModoAg`). Registrado como deuda técnica documental.
- **Archivo Afectado:** `implementation_plan.md` / `walkthrough.md`

#### 6. Fase 684 — Pruebas Unitarias
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Necesidad de validar el algoritmo de cierre (ordenamiento por inicio más tardío), normalizaciones 12h/24h, formato Nivel 2 y ciclo on/off de pernocta.
  - *Causa Raíz:* Nuevos componentes y métodos de dominio.
  - *Impacto Confirmado:* Cobertura exhaustiva al 100% de la lógica nueva sin dependencias UI.
- **Acción Aplicada:** Creado `ViaticoBloqueFormatterTests.cs` (6 tests) y extendido `ViaticoBloquePernoctaTests.cs` (2 tests: `IdsVinculacion` y `AlternarPernocta`).
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Application/Services/ViaticoBloqueFormatterTests.cs`
  - `ConversorPDF.Tests/Academic/Domain/ViaticoBloquePernoctaTests.cs`

---

### Verificaciones de Principios (Ronda 120)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F679 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F680 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F681 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F682 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F683 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F684 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.79 Ronda 121 — Fases 685 a 689: Robustecimiento de Pernocta y Cobertura de Repositorio

### Evaluación Crítica de las Fases

#### 1. Fase 685 — Mensaje de Confirmación con Duración Dinámica
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La duración mostrada en el mensaje de confirmación (`cambioDias`) asumía bloques fijos de 2 días (`1,5 ➔ 2,5 días`), resultando equívoca para bloques de 1 día u otras extensiones.
  - *Causa Raíz:* Se usaron cadenas fijas en lugar de calcular a partir de `bloque.DuracionTotal`.
  - *Impacto Confirmado:* Corrección de presentación interactiva fiel a la duración real del bloque seleccionado.
- **Acción Aplicada:** Cálculo dinámico: `double durNueva = valorNuevo == 1 ? durActual + 1 : durActual - 1; string cambioDias = $"{durActual} ➔ {durNueva} días";`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`

#### 2. Fase 686 — Aislamiento Cultural en Pruebas de Formatter
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Pruebas de formateo con interpolación de comas decimales y separadores de miles fallaban si el ejecutor de pruebas usaba otra configuración regional (ej. `en-US`).
  - *Causa Raíz:* Dependencia de la cultura de hilo sin fijar explícitamente en el test.
  - *Impacto Confirmado:* Tests 100% reproducibles y deterministas en cualquier entorno CI.
- **Acción Aplicada:** Fijado `CultureInfo.DefaultThreadCurrentCulture` y `CultureInfo.DefaultThreadCurrentUICulture` a `es-CO` en el constructor de `ViaticoBloqueFormatterTests`.
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Application/Services/ViaticoBloqueFormatterTests.cs`

#### 3. Fase 687 — Notificación `MarkDirty()` Post-Commit
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Las actualizaciones atómicas ejecutaban `NotifyMarkDirty()` dentro de la transacción activa por cada UPDATE, violando el patrón transaccional del proyecto.
  - *Causa Raíz:* Falta del parámetro `suppressMarkDirty: true` en la inicialización de `AcademicRepository`.
  - *Impacto Confirmado:* El scheduler no captura cambios no confirmados.
- **Acción Aplicada:** Se pasa `suppressMarkDirty: true` al constructor y se invoca `SyncContext.Scheduler?.MarkDirty()` inmediatamente tras `tx.Commit()`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`

#### 4. Fase 688 — Restauración de Selección por Coincidencia Exacta
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La comprobación `nDoc.Text.Contains(docenteSeleccionado)` podía emparejar erróneamente un docente cuyo nombre fuera subcadena de otro.
  - *Causa Raíz:* Comparación por subcadena sobre el texto decorado del nodo padre.
  - *Impacto Confirmado:* Selección siempre exacta del bloque alterado.
- **Acción Aplicada:** Emparejamiento por igualdad ordinal exacta en `b.Sesiones.Any(s => string.Equals(s.DocenteNombre, docenteSeleccionado, StringComparison.Ordinal))` y retorno anticipado tras posicionar el nodo.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`

#### 5. Fase 689 — Cobertura Unitaria de `SetPernoctaUltimoDia`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* No existían pruebas directas sobre la persistencia y actualización SQL de la columna `isPernoctaUltimoDia` a nivel de repositorio.
  - *Causa Raíz:* Omitido en la suite de pruebas previa.
  - *Impacto Confirmado:* Verificación automatizada de la consulta `UPDATE` y comportamiento ante IDs inexistentes y valores atípicos.
- **Acción Aplicada:** Creado `AcademicRepositorySetPernoctaTests.cs` con 4 casos de prueba SQLite en memoria.
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AcademicRepositorySetPernoctaTests.cs`

---

### Verificaciones de Principios (Ronda 121)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F685 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F686 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F687 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F688 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F689 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.80 Ronda 122 — Fases 690 y 691 (Fase 666 y 666.b): Rango de Período Académico como Advertencia Confirmable y Ampliación de Ventana Maestra

### Evaluación Crítica de las Fases

#### 1. Fase 690 (Fase 666) — Validación de Rango de Fechas del Período (Advertencia Confirmable)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El diálogo informativo con `MessageBoxButtons.OK` y salida temprana obligatoria impedía guardar vinculaciones con sesiones reales que desbordaban la ventana originalmente seleccionada.
  - *Causa Raíz:* Bloqueo duro en validación previa de guardado.
  - *Impacto Confirmado:* La validación ahora advierte con `MessageBoxButtons.YesNo` y, tras confirmación positiva del usuario, autoriza continuar guardando.
- **Acción Aplicada:** Reemplazado el bloqueo por advertencia interactiva con confirmación `YesNo` y bandera `ampliarRangoConfirmado`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

#### 2. Fase 691 (Fase 666.b) — Ampliación Automática de la Ventana Maestra
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al autorizar guardar sesiones fuera de rango, si `dtpInicio.Value` y `dtpFinal.Value` permanecían estrechos, la entidad `Vinculacion` se persistía con fechas maestras discordantes respecto a sus sesiones hijas.
  - *Causa Raíz:* No se ajustaban los controles maestros antes del guardado persistente.
  - *Impacto Confirmado:* Sincronización íntegra: la ventana maestra en BD y UI abraza exactamente el rango `[minFecha, maxFecha]`.
- **Acción Aplicada:** Ampliación previa a la persistencia con desuscripción y resuscripción de `dtpInicio.ValueChanged` para evitar efectos colaterales de regeneración.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

---

### Verificaciones de Principios (Ronda 122)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F690 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F691 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.81 Ronda 123 — Verificación de Fases 690/691 y Robustecimiento de Edición de Fechas (Fases 692 a 695)

### Evaluación Crítica de las Fases

#### 1. Fase 692 — Instrucción de Remediación en Salida "No"
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al responder "No" en el diálogo de advertencia de fechas fuera de rango, el usuario no recibía orientación sobre cómo proceder.
  - *Causa Raíz:* Omisión de la alternativa de ajuste en el texto del mensaje interactivo.
  - *Impacto Confirmado:* Mayor claridad operativa y UX mejorada.
- **Acción Aplicada:** Incorporada la instrucción `"Si selecciona 'No', ajuste las fechas del grid o modifique [Fecha Inicio]/[Fecha Final] antes de guardar.\n\n"`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

#### 2. Fase 693 — Tolerancia de Formatos y Objetos DateTime en `GridCellValueParser`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La edición de fechas en el grid provocaba reversiones silenciosas si el valor contenía hora (`dd/MM/yyyy HH:mm:ss`), objeto `DateTime` nativo o año de 2 dígitos (`10/10/26`).
  - *Causa Raíz:* `TryParseFecha` solo contemplaba un conjunto cerrado de 4 formatos exactos sin recortar horas ni admitir `DateTime`.
  - *Impacto Confirmado:* Edición de fechas en el grid 100% tolerante y fluida.
- **Acción Aplicada:** Soportado `value is DateTime dt`, recorte de componente horario y agregados formatos `"d/M/yy", "dd/MM/yy", "yyyy/MM/dd", "yyyyMMdd"`. Actualizadas pruebas unitarias en `GridCellValueParserTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Mappers/GridCellValueParser.cs`
  - `ConversorPDF.Tests/Unit/GridCellValueParserTests.cs`

#### 3. Fase 694 — Evaluación de Semántica de Período vs Ventana de Generación
- **Estado:** RETIRADA (Diseño intencional)
- **Hallazgo:**
  - *Diagnóstico:* La comparación contra `[dtpInicio, dtpFinal]` es la regla de negocio prevista para acotar la ventana específica de la vinculación/asignatura, no el semestre global.
  - *Acción:* Se mantiene la comparación contra los controles maestros de la vinculación.

#### 4. Fase 695 — Alineación de Trazabilidad Código vs Rondas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Discrepancia entre los comentarios de código ("Fase 666 / Fase 666.b") y la numeración oficial del plan ("Fases 690 / 691").
  - *Causa Raíz:* Renumeración en la ronda previa sin reflejar en los comentarios del archivo fuente.
  - *Impacto Confirmado:* Coherencia documental y trazabilidad 1:1 entre código y artefactos.
- **Acción Aplicada:** Actualizados los comentarios en `FormSesionesAcademicas.cs` a `Fase 690 (plan 666)` y `Fase 691 (plan 666.b)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

---

### Verificaciones de Principios (Ronda 123)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F692 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F693 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F694 | [X] N/A (Retirada) | [X] Sí | [X] Sí | [X] Sí |
| F695 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.82 Ronda 124 — Robustecimiento de Fechas de 2 Dígitos, Separadores Whitespace y Terminología de Ventana (Fases 696 a 698)

### Evaluación Crítica de las Fases

#### 1. Fase 696 — Normalización de Años de 2 Dígitos
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La introducción de `d/M/yy` y `dd/MM/yy` con `CultureInfo.InvariantCulture` resolvía fechas mayores o iguales a 30 (ej. `"15/10/30"`) a 1930 debido al límite predeterminado `TwoDigitYearMax = 2029`.
  - *Causa Raíz:* Comportamiento estándar del calendario gregoriano en BCL para años cortos.
  - *Impacto Confirmado:* Combinado con la advertencia confirmable de F690, se evitaba el bloqueo y permitía persistir fechas anómalas de siglo pasado.
- **Acción Aplicada:** Agregada normalización post-parseo `if (fecha.Year < 2000) fecha = fecha.AddYears(100);` y prueba unitaria en `GridCellValueParserTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Mappers/GridCellValueParser.cs`
  - `ConversorPDF.Tests/Unit/GridCellValueParserTests.cs`

#### 2. Fase 697 — Terminología Precisa en Mensaje de Validación
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El diálogo y los logs aún mencionaban "fuera del período académico", manteniendo ambigüedad terminológica frente al usuario.
  - *Causa Raíz:* Residuo de redacción previo a la confirmación de alcance de la ventana de vinculación.
  - *Impacto Confirmado:* Mensaje claro y coherente con el modelo de vinculación.
- **Acción Aplicada:** Actualizado texto a "fuera de la ventana de la vinculación", título "Fechas Fuera de la Ventana" y log respectivo.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

#### 3. Fase 698 — Recorte de Separadores Horarios Whitespace Múltiples (Espacio, Tab, NBSP)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `IndexOf(' ')` fallaba al recortar componentes horarios cuando el separador era tabulador o espacio no separable (`\u00A0`), común al pegar desde navegadores o Excel.
  - *Causa Raíz:* Búsqueda rígida de un único carácter de espacio ASCII.
  - *Impacto Confirmado:* Tolerancia completa ante datos copiados/pegados de fuentes web y ofimáticas.
- **Acción Aplicada:** Implementado recorte mediante `txt.Split` con separadores `{' ', '\t', '\r', '\n', '\u00A0'}` y toma del primer token no vacío. Pruebas unitarias agregadas.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Mappers/GridCellValueParser.cs`
  - `ConversorPDF.Tests/Unit/GridCellValueParserTests.cs`

---

### Verificaciones de Principios (Ronda 124)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F696 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F697 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F698 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.83 Ronda 125 — Auto-ampliación Silenciosa de Ventana Maestra al Grid y Sincronización de Snapshot de Fechas (Fase 699)

### Evaluación Crítica de las Fases

#### 1. Fase 699 — Auto-ampliación de Ventana Maestra y Actualización FI/FF en Snapshot
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La edición de `Fecha Inicio`/`Fecha Final` disparaba falsamente "Falta Generar Sesiones" al momento de guardar debido al desacople con el hash de `_ultimoEstadoGenerado`. Además, reducir la ventana por debajo del rango de sesiones del grid provocaba advertencias redundantes.
  - *Causa Raíz:* `ParametrosEstanDesincronizados()` comparaba literalmente `FI` y `FF` congelados durante la última generación, sin refrescar las fechas tras ajustes manuales y sin auto-ajustar los selectores a las sesiones existentes.
  - *Impacto Confirmado:* La edición de fechas maestras refresca de manera aislada los campos `FI/FF` en el snapshot (preservando `T/H/D`), mientras que la ventana se auto-amplía silenciosamente para envolver todas las sesiones del grid sin disparar advertencias innecesarias.
- **Acción Aplicada:**
  - Añadida bandera `_ajustandoFechaMaestra`.
  - Implementados métodos auxiliares `TryObtenerRangoSesionesGrid`, `AmpliarVentanaMaestraAlGrid` y `SincronizarFechasEnEstadoGenerado`.
  - Actualizados los manejadores de eventos `DtpInicio_ValueChanged` y `DtpFinal_ValueChanged`.
  - Supresión simétrica de eventos para ambos selectores en `Generar`, `Fase 691`, `Post-commit` y `CargarEstadoPrevio`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

---

### Verificaciones de Principios (Ronda 125)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F699 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.84 Ronda 126 — Blindaje de Rango DateTimePicker, Filtro de Filas Marcadas y Desduplicación de Logs (Fases 700 a 703)

### Evaluación Crítica de las Fases

#### 1. Fase 700 — Acotamiento contra MinDate/MaxDate en Auto-ampliación
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La asignación directa de fechas atípicas del grid a `dtpInicio.Value` o `dtpFinal.Value` podía arrojar `ArgumentOutOfRangeException` dentro del ciclo de eventos de WinForms.
  - *Causa Raíz:* Límites `1753-01-01` (`MinDate`) y `9998-12-31` (`MaxDate`) de `DateTimePicker` no verificados antes de la asignación.
  - *Impacto Confirmado:* Blindaje total contra excepciones de runtime en handlers de UI.
- **Acción Aplicada:** Clamp previo a la asignación: `primera = Math.Max(primera, dtpInicio.MinDate)` y `ultima = Math.Min(ultima, dtpFinal.MaxDate)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

#### 2. Fase 701 — Criterio de Selección de Filas Marcadas en `TryObtenerRangoSesionesGrid`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `TryObtenerRangoSesionesGrid` recorría todas las filas, permitiendo que sesiones desmarcadas influyeran en la ampliación de la ventana de vinculación.
  - *Causa Raíz:* Falta de verificación de `ColChecked` en el bucle lector.
  - *Impacto Confirmado:* Coherencia absoluta con la persistencia en base de datos (`PrepararSesionesDesdeGrid`), que solo preserva filas marcadas.
- **Acción Aplicada:** Agregado filtro `if (!SafeCellBool(row.Cells["ColChecked"].Value)) continue;`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

#### 3. Fase 702 — Actualización Documental del Snapshot `_ultimoEstadoGenerado`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Comentarios desactualizados en el campo `_ultimoEstadoGenerado`.
  - *Causa Raíz:* La introducción del refresco de `FI/FF` en F699 no había quedado reflejada en el XML doc del campo.
  - *Impacto Confirmado:* Claridad de mantenimiento y trazabilidad de arquitectura.
- **Acción Aplicada:** Actualizado comentario XML explicando el rol de `SincronizarFechasEnEstadoGenerado`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

#### 4. Fase 703 — Supresión de Logs Redundantes por Idempotencia de Snapshot
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Disparos dobles de logs idénticos de sincronización de snapshot al ajustar `Fecha Inicio` debido a la re-evaluación del invariante `dtpFinal >= dtpInicio`.
  - *Causa Raíz:* Falta de verificación de cambio real antes de emitir log.
  - *Impacto Confirmado:* Logs limpios y concisos sin ruido redundante.
- **Acción Aplicada:** Agregado guard `if (nuevo == _ultimoEstadoGenerado) return;`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

---

### Verificaciones de Principios (Ronda 126)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F700 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F701 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F702 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F703 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.85 Ronda 127 — Sincronización Inmediata del Snapshot tras Ampliación Pre-Guardado (Fase 704)

### Evaluación Crítica de las Fases

#### 1. Fase 704 — Sincronización de Snapshot tras Mutación de Fechas en F691
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La ampliación pre-guardado de `dtpInicio`/`dtpFinal` en F691 se ejecutaba con eventos suprimidos. Si `GuardarAsync` fallaba por contención de base de datos o cancelación, el snapshot `_ultimoEstadoGenerado` no se actualizaba, bloqueando reintentos con el mensaje "Falta Generar Sesiones".
  - *Causa Raíz:* La única sincronización de snapshot en el flujo de guardado residía en la rama exitosa post-commit.
  - *Impacto Confirmado:* Coherencia atómica: tan pronto los controles maestros se adaptan a las fechas confirmadas por el usuario, el snapshot refleja `FI/FF` actualizados, permitiendo reintentar el guardado sin regenerar.
- **Acción Aplicada:** Invocado `SincronizarFechasEnEstadoGenerado()` inmediatamente tras el bloque `finally` de re-suscripción en F691.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

---

### Verificaciones de Principios (Ronda 127)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F704 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.86 Ronda 128 — Consolidación Integral del Snapshot Pre-Persistencia (Fase 705)

### Evaluación Crítica de las Fases

#### 1. Fase 705 — Actualización de Snapshot tras Ajuste de Horas y Ampliación de Fechas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La actualización de `txtTotalHoras.Text` en la confirmación de discrepancia de horas (F2.4) dejaba `_ultimoEstadoGenerado` con el valor `H=` anterior. Ante un fallo de base de datos o cancelación en `GuardarAsync`, el usuario recibía falsamente "Falta Generar Sesiones" al reintentar guardar.
  - *Causa Raíz:* Las mutaciones pre-guardado de horas no se reflejaban de inmediato en el snapshot de validación.
  - *Impacto Confirmado:* Consolidación completa e inmediata de todos los parámetros de control (`T/H/D/FI/FF`) antes de despachar la tarea asíncrona de guardado.
- **Acción Aplicada:** Invocado `_ultimoEstadoGenerado = ObtenerEstadoParametrosCriticos()` inmediatamente antes de inicializar `_guardadoCts`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

---

### Verificaciones de Principios (Ronda 128)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F705 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.87 Ronda 129 — Consistencia Terminológica en Aviso de Parámetros Desincronizados (Fase 706)

### Evaluación Crítica de las Fases

#### 1. Fase 706 — Exclusión de "Fechas" del Aviso "Falta Generar Sesiones"
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El diálogo `Falta Generar Sesiones` indicaba que la causa de bloqueo podía ser la edición de `Fechas`. Desde la implementación de F699, las fechas maestras no bloquean el guardado ni desincronizan el hash generador, induciendo confusión operativa en el usuario.
  - *Causa Raíz:* Texto descriptivo hardcodeado previo a la flexibilización de fechas maestras.
  - *Impacto Confirmado:* Mensaje alineado 100% con las reglas vigentes: solo cambios reales en `Docente`, `Horas` o `Distribución de días` demandan regenerar el grid.
- **Acción Aplicada:** Eliminada la referencia `"o Fechas"` del cuerpo del mensaje en `btnGuardar_Click`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`

---

### Verificaciones de Principios (Ronda 129)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F706 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.88 Ronda 130 — Rango de Comisión de Viáticos con Pernocta en Último Día (Fases 707 a 710)

