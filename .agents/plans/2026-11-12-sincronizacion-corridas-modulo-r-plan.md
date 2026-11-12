---
numbering:
  registry: .agents/plans/REGISTRY.md
  registry_row: R-041
  ronda_base: 319
  fase_base: 1606
  waves: 7
  tasks: 7
---

# Plan R-041: Sincronización del Selector de Corridas del Módulo R

**Generado**: 2026-11-12 · **Autor**: Antigravity/Swarm-Planner · **Estado**: reserved · **Depende de**: R-040 (Fases 1603–1605)

## Execution envelope
- branch: `refactor/academic-modularization`
- workspace: `c:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`
- preflight: `powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace "c:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan .agents/plans/2026-11-12-sincronizacion-corridas-modulo-r-plan.md -AllowOneDrive`
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: `tests/baseline/test_baseline.txt (TOTAL: 952)`
- verify: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- build_gate: serializado (un único build-verifier a la vez; Excel-DNA bloquea bin/)
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0:
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Core/Helpers/App.cs
  - ConversorPDF.XLL/**/Migrations/MigrationV*.cs
  - ConversorPDF.XLL/**/Persistence/SchemaMigrator.cs
  - ConversorPDF.XLL/Ribbon/RibbonController.cs
  - Composition Root / IoC Container
  - docs/toc.yml
- stop_conditions:
  - Modificar comportamiento de FormVinculaciones o eliminar comodín "-- Todas las corridas --"
  - Degradación forzada a SelectedIndex = 0 que sobreescriba corridas activas
  - Refrescos duplicados o en bucle infinito
  - DELETE en tablas de staging o producción
  - Intento de push, merge o publicación a remoto
- ledger_owner: swarm-parallel-task
- no_push: true
- numbering:
  registry: .agents/plans/REGISTRY.md
  registry_row: R-041
  ronda_base: 319
  fase_base: 1606
  waves: 7
  tasks: 7

## Parallel execution groups

| Wave | Ronda | Tasks | Can start when |
|---|---|---|---|
| 1 | 319 | T1 (integrator, L0) | preflight `0` (`-AllowOneDrive` en este árbol) |
| 2 | 320 | T2 (tdd-red-writer) | T1 completado |
| 3 | 321 | T3 (implementer) | T2 completado |
| 4 | 322 | T4 (implementer) | T3 completado |
| 5 | 323 | T5 (implementer) | T4 completado |
| 6 | 324 | T6 (build-verifier) | T5 completado |
| 7 | 325 | T7 (implementer) | T6 completado |

Leases disjuntos: T1 opera en L0 (`ConversorPDF.XLL.csproj`) y crea `AgSeleccionModuloR.cs` junto al stub en `AgCorridaSelector.cs`; T2 opera en `ConversorPDF.Tests`; T3 en `AgCorridaSelector.cs` y `AgFormGeneradorHorarios.cs`; T4 en `AgFormGeneradorHorarios.cs`; T5 en `AgFormGeneradorHorarios.cs`; T6 en verificación con `files_lock: []`; T7 en documentación técnica en `docs/articles/changes/`.

---

## Tareas

### T1: Esqueleto de tipo AgSeleccionModuloR e inclusión en csproj (integrator-first)
- **id**: T1
- **fase**: 1606
- **role**: integrator
- **depends_on**: []
- **location**: `ConversorPDF.XLL/Academic/Application/DTOs/AgSeleccionModuloR.cs`
- **files_lock**:
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgSeleccionModuloR.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs`
- **description**: Crear el archivo `AgSeleccionModuloR.cs` cumpliendo la regla Single Type Per File con las propiedades inmutables (`CorridaStagingItem ItemSeleccionado`, `int IndiceSeleccionado`, `bool RequiereCargaInicial`, `bool CorridaActivaNoEncontrada`), registrarlo como `<Compile Include="..." />` en `ConversorPDF.XLL.csproj` (Lease L0), y declarar en `AgCorridaSelector.cs` la firma del método estático stub `public static AgSeleccionModuloR ResolverSeleccionModuloR(IEnumerable<CorridaStagingItem> corridas, long? idCorridaActiva)` lanzando `throw new NotImplementedException();`.
- **validation**: Compilación exitosa en modo Build (`verify.ps1 -Mode Build`) con 0 errores y el nuevo tipo disponible para tests.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build`
- **evidence**: `_evidence/20261112-071607/summary.md`
- **status**: done
- **log**: Esqueleto AgSeleccionModuloR.cs creado, registrado en ConversorPDF.XLL.csproj y stub estático en AgCorridaSelector.cs. Build exit 0.
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgSeleccionModuloR.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs`

### T2: Caracterización y pruebas unitarias RED en AgCorridaSelectorTests
- **id**: T2
- **fase**: 1607
- **role**: tdd-red-writer
- **depends_on**: [T1]
- **location**: `ConversorPDF.Tests/Academic/Application/Services/AgCorridaSelectorTests.cs`
- **files_lock**:
  - `ConversorPDF.Tests/Academic/Application/Services/AgCorridaSelectorTests.cs`
- **description**: Escribir pruebas unitarias para `AgCorridaSelector.ResolverSeleccionModuloR` cubriendo los 6 casos obligatorios: 1) Activa presente (selecciona la activa, no recarga); 2) Activa ausente con corrida candidata con staging (sin selección, no autoriza carga); 3) Activa ausente con solo cabeceras (sin selección, no autoriza carga); 4) Sin activa y staging disponible (selecciona primera con staging, autoriza carga inicial); 5) Sin activa y solo cabeceras (sin selección, no autoriza carga); 6) Lista vacía o nula. Preservar todas las pruebas existentes de `ResolverSeleccionInicial` y del comodín para `FormVinculaciones`.
- **validation**: La solución compila correctamente y las nuevas pruebas fallan con `NotImplementedException` (RED auténtico por lógica no implementada).
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgCorridaSelector"`
- **evidence**: `_evidence/20261112-071854/summary.md`
- **status**: done
- **log**: 6 pruebas unitarias RED implementadas en AgCorridaSelectorTests.cs. Build exitoso, 18 pruebas previas superadas, 6 nuevas fallando exactamente con NotImplementedException.
- **files edited/created**:
  - `ConversorPDF.Tests/Academic/Application/Services/AgCorridaSelectorTests.cs`

### T3: Implementación de la política segura ResolverSeleccionModuloR y enlace inicial
- **id**: T3
- **fase**: 1608
- **role**: implementer
- **depends_on**: [T2]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **description**: Implementar la lógica completa de `AgCorridaSelector.ResolverSeleccionModuloR`. Integrarlo en `RefrescarComboCorridasAsync` eliminando la degradación forzada a `idx >= 0 ? idx : 0` y la llamada incondicional a `CargarCorridaDesdeComboAsync`. Si la activa está presente, seleccionarla sin recargar; si la activa no está en la lista, fijar `SelectedIndex = -1` y actualizar `lblAgModo` informando que la corrida cargada no aparece con los filtros actuales; si no hay activa pero hay staging, seleccionar la primera y autorizar la carga inicial una sola vez; si solo hay cabeceras, dejar `SelectedIndex = -1` sin carga automática.
- **validation**: Las pruebas de T2 pasan a GREEN; ninguna regresión en `AgCorridaSelectorTests` ni en `FormVinculaciones`.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgCorridaSelector"`
- **evidence**: `_evidence/20261112-072421/summary.md`
- **status**: done
- **log**: ResolverSeleccionModuloR implementado en AgCorridaSelector.cs y cableado en RefrescarComboCorridasAsync en AgFormGeneradorHorarios.cs. Fallback destructor a 0 eliminado. 24/24 pruebas en GREEN.
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

### T4: Refresco seguro del selector y protección de concurrencia
- **id**: T4
- **fase**: 1609
- **role**: implementer
- **depends_on**: [T3]
- **location**: `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **description**: Corregir el evento `DropDown` para solicitar siempre una lectura fresca (eliminando el guard `ClaveFiltroCorridas() == _claveFiltroCorridas`). Capturar los filtros y el ID de corrida activa al inicio de la consulta asíncrona; no dar por válida una respuesta si durante el await cambiaron los filtros o el formulario fue dispuesto (`IsDisposed`). Proteger contra colisiones de peticiones concurrentes: si llega una nueva solicitud mientras una está en curso, registrar la solicitud pendiente y procesar la más reciente al finalizar, en lugar de descartarla con un return ciego. Ante error de lectura, preservar el catálogo y la corrida activa previos sin vaciar la selección.
- **validation**: Apertura de DropDown reconsulta siempre; respuestas obsoletas no sobreescriben la selección reciente; sin excepciones por controles dispuestos.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgCorridaSelector"`
- **evidence**: `_evidence/20261112-073145/test.trx`
- **status**: done
- **log**: DropDown reconsulta siempre; bucle de coalescencia con _solicitudPendienteRefrescoCombo, descarte de respuestas obsoletas ante cambio de filtro durante await, chequeos de IsDisposed/Disposing y preservación de catálogo ante error. Tests en PASS (24/24).
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

### T5: Ciclo de vida de corrida, sincronización post-ejecución y retorno de diálogo
- **id**: T5
- **fase**: 1610
- **role**: implementer
- **depends_on**: [T4]
- **location**: `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **description**: 1) Centralizar el refresco post-ejecución del AG en un único bloque protegido en `finally` que no herede el CancellationToken cancelado del AG, refrescando el combo únicamente si se alcanzó a crear o modificar una cabecera de corrida. 2) Tras una carga exitosa en `CargarModuloResultadosAsync`, actualizar `lblAgModo` con `#N` también en la ruta de ejecución directa. 3) Comprobar el valor devuelto (`bool`) de `CargarModuloResultadosAsync`: si retorna `false`, limpiar `_idCorridaCargada` y las vistas para no dejar una corrida falsamente activa. 4) En `AbrirVinculacionesModoAgAsync`, verificar el resultado de recargar `_idCorridaCargada`: si la corrida fue eliminada o quedó vacía en el diálogo, limpiar el estado y solicitar refresco del selector para resolver la selección inicial limpia.
- **validation**: Ejecución directa de AG actualiza combo y `lblAgModo` con la misma corrida #N; cancelación o fallo post-cabecera no activan corridas ajenas; cierre de diálogo con corrida eliminada limpia la interfaz.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~Ag"`
- **evidence**: `_evidence/20261112-074004/test.trx`
- **status**: done
- **log**: Sincronización post-ejecución centralizada en finally con idCorridaCreada > 0 y token independiente. CargarModuloResultadosAsync resetea _idCorridaCargada = 0 ante fallo de staging y actualiza lblAgModo en ruta directa. AbrirVinculacionesModoAgAsync verifica recarga y refresca combo. Tests 617/617 PASS.
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

### T6: Verificación integral de la suite completa
- **id**: T6
- **fase**: 1611
- **role**: build-verifier
- **depends_on**: [T5]
- **location**: `scripts/verify.ps1`
- **files_lock**: []
- **description**: Ejecutar la suite completa de pruebas unitarias e integración en modo Full mediante `verify.ps1` para asegurar 0 fallos y 0 regresiones respecto al baseline oficial (TOTAL: 952). Generar reporte TRX y summary.md bajo `_evidence/`.
- **validation**: Exit code 0, 0 tests fallidos, 0 regresiones.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- **evidence**: `_evidence/20261112-074457/summary.md`
- **status**: done
- **log**: Verificación integral Full ejecutada con éxito. Total: 1172, Passed: 1172, Failed: 0, Skipped: 0. 0 regresiones vs baseline (952). Veredicto PASS.
- **files edited/created**: []

### T7: Documentación técnica del cambio
- **id**: T7
- **fase**: 1612
- **role**: implementer
- **depends_on**: [T6]
- **location**: `docs/articles/changes/`
- **files_lock**:
  - `docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md`
- **description**: Redactar el artículo de cambios documentando la causa raíz del desfasaje entre el selector y los resultados, la política pura `ResolverSeleccionModuloR`, la gestión de concurrencia y el ciclo de vida de la corrida.
- **validation**: Archivo de documentación creado y coherente con las evidencias.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build`
- **evidence**: `_evidence/20261112-071607/summary.md`
- **status**: done
- **log**: Esqueleto AgSeleccionModuloR.cs creado, registrado en ConversorPDF.XLL.csproj y stub estático en AgCorridaSelector.cs. Build exit 0.
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgSeleccionModuloR.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs`

---

## Criterios de aceptación
1. Tras ejecutar o cargar una corrida en Módulo R, el selector `cmbCargarCorridas`, la etiqueta `lblAgModo` y las grillas representan la misma corrida `#N`.
2. Al desplegar el ComboBox (`DropDown`), se realiza una consulta fresca sin ser bloqueada por la caché de filtros.
3. Si la corrida activa no está en la lista resultante por efecto de los filtros, `SelectedIndex = -1`, se conserva el resultado cargado y se muestra un mensaje informativo en `lblAgModo` sin sobrescribir con el índice 0 ni borrar resultados válidos.
4. Al abrir el formulario sin corrida activa, se selecciona y carga automáticamente la primera corrida con staging; si solo hay cabeceras sin staging, no se selecciona ninguna automáticamente.
5. El contrato y comportamiento de `AgCorridaSelector.ResolverSeleccionInicial` y de `FormVinculaciones` se mantiene intacto con su comodín `-- Todas las corridas --`.
6. La ejecución y cancelación del algoritmo genético gestionan el selector de manera protegida en `finally` sin heredar el token cancelado.
7. Suite completa de tests (`verify.ps1 -Mode Full`) en PASS con 0 regresiones.
