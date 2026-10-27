# Reparación de Esquema CRDT en Periodos y Resolución de Hueco de Migración V50/V51

**Fecha:** 2026-10-19  
**Módulo:** Académico (`ConversorPDF.XLL.Academic`)  
**Rondas:** 140 (Fases 744 a 751), 141 (Fases 752 a 756) y 142 (Fases 757 a 759)

## Purpose
Resolver el error de esquema incompleto en la tabla `periodos` (`"El esquema no soporta ventana temporal (periodo_inicio/periodo_fin). Aplique las migraciones."`) que impedía persistir o visualizar las fechas maestras de inicio y fin semestral. La causa raíz fue un fallo en las migraciones aditivas tolerantes V50 y V51 provocado por triggers obsoletos de CR-SQLite (`expected 13 values, got 9`), que revirtió la transacción y generó un hueco de versión donde `__SchemaVersion` avanzó hasta 64 sin que las columnas existieran físicamente en el vault. La Ronda 141 blindó la condición de drop de triggers y restauró triggers ausentes; la Ronda 142 desacopló la regla V56 para ejecutarse de forma incondicional en V65 (F757), eliminó triggers remanentes antes de `crsql_as_crr` (F758) y unificó la lógica `COALESCE` en V50 y V51 (F759).

## Changes
1. **Helper de Triggers CRDT (`PeriodoVentanaCrdtRepair.cs`, Fase 744):**
   - Creación del helper estático `PeriodoVentanaCrdtRepair.DropStaleCrsqlTriggers` para soltar de forma segura los tres triggers de sincronización (`periodos__crsql_dtrig`, `periodos__crsql_utrig`, `periodos__crsql_itrig`).
   - Preserva íntegra la tabla `periodos__crsql_clock` para no perder historial de cambios entre sedes.
   - Los triggers son regenerados automáticamente con el nuevo esquema por `SyncInitializer.EnableAllCrdts` tras marcarse el vault como migrado.

2. **Nueva Migración V65 (`MigrationV65_PeriodoVentanaCrdtRepair.cs`, Fases 745, 752, 753, 757, 760):**
   - Agrega físicamente `periodo_inicio TEXT` y `periodo_fin TEXT` a la tabla `periodos` si faltan.
   - Drop condicionado de triggers (Fase 752): solo se sueltan los triggers CRDT si efectivamente se requiere ejecutar sentencias DML de backfill (`requiereBackfill`). En vaults sanos con fechas completas no se tocan los triggers.
   - Backfill canónico parametrizado con `COALESCE(NULLIF(..., ''), @Val)` (Fase 753) para `idperiodo = 1, 2, 3`.
   - Regla V56 incondicional (Fase 757): la actualización de `periodo_inicio = '2026-08-11'` para `2026-02` se ejecuta fuera de `if (requiereBackfill)` para garantizar su aplicación incluso si el vault ya tenía fechas completas pero retenía `2026-08-15`.
   - Precisión de logging y trazabilidad (Fase 760): log del bloque `else` actualizado a `sin drop de triggers` y registro explícito de verificación de la regla V56.

3. **Registro en Pipeline y Compilación (`SqliteHelper.cs` y `.csproj`, Fase 746):**
   - Inclusión de `new MigrationV65_PeriodoVentanaCrdtRepair()` en `SqliteHelper.BuildMigrator()`.
   - Inclusión explícita de `PeriodoVentanaCrdtRepair.cs` y `MigrationV65_PeriodoVentanaCrdtRepair.cs` en `ConversorPDF.XLL.csproj`.

4. **Blindaje Retroactivo de V50 y V51 (`MigrationV50_AddPeriodoVentana.cs` y `MigrationV51_PeriodoVentanaRepair.cs`, Fases 747, 754, 759):**
   - Invocación condicionada de `PeriodoVentanaCrdtRepair.DropStaleCrsqlTriggers` posterior a DDL y únicamente si se requiere backfill.
   - Uso de `COALESCE(NULLIF(periodo_inicio, ''), @Ini)` y `COALESCE(NULLIF(periodo_fin, ''), @Fin)` en ambas migraciones para rellenar fechas nulas individuales sin descartar el registro (Fase 759).

5. **Detección, Purga y Auto-Restauración de Triggers en `SyncInitializer` (`SyncInitializer.cs`, Fases 755, 758):**
   - Incorporación del helper `TableHasCrdtTriggers(connection, tableName)` que valida la presencia de los 3 triggers CRDT.
   - Si la tabla ya es CRDT pero perdió sus triggers o quedaron triggers parciales, se eliminan remanentes con `DROP TRIGGER IF EXISTS` y se re-ejecuta `crsql_as_crr` antes de evaluar el short-circuit (Fase 758).

6. **Observabilidad y Diagnóstico de Huecos de Versión (`SchemaMigrator.cs`, Fase 748):**
   - Alerta con severidad `Error` indicando `HUECO DE VERSIÓN` cuando una migración tolerante aditiva falla y omite su registro.
   - Inclusión de acción correctiva en el mensaje de error de `VerificarColumnasRequeridas`.

7. **Guard de Coherencia en Repositorio (`AcademicRepository.cs`, Fase 749):**
   - `InsertPeriodo` valida si el esquema soporta ventana temporal antes de descartar fechas silenciosamente, logrando paridad con el guard de `UpdatePeriodo`.

8. **Pruebas Unitarias Automatizadas (`MigrationV65Tests.cs`, `MigrationV50Tests.cs`, `MigrationV51Tests.cs`, `AcademicRepositoryPeriodoTests.cs`, Fases 750, 752, 753, 757, 759):**
   - Cobertura completa de metadatos, migración sobre esquema legado con triggers stale, backfill sin mismatch, preservación de fechas existentes, idempotencia, protección de triggers en esquema completo, regla V56 incondicional, relleno de fin nulo en V50, V51 y V65, y guards en repositorio.

9. **Sincronización de Versión Oficial y Plantillas (`App.cs`, `ConversorPDF.XLL.csproj`, `InstaladorConversorPDF.wxs`, Fases 751, 756):**
   - Incremento a la versión oficial `1.6.3.58`.
   - `DefaultVersion="1.6.3.58"` en `filTemplate1` (`GF-FO-023.xlsx`) y `filTemplate2` (`FV-FO-033.xlsx`).

## How to use
1. Al arrancar el Add-In en Excel, `SchemaMigrator` detecta que el vault se encuentra en versión 64 y aplica la migración V65 (`PeriodoVentanaCrdtRepair`).
2. Abrir **Períodos** desde el menú contextual: las columnas *Inicio (ISO)* y *Fin (ISO)* cargan pobladas con los rangos canónicos.
3. El botón **Guardar** persiste modificaciones sin cuadros de error de esquema incompleto.
4. El motor de Algoritmo Genético (AG) resuelve la ventana semestral con **Prioridad 1 (PERIODOS)**.

## Test Coverage (Fases 750, 752, 753, 757, 759)
- `MigrationV65Tests.cs`:
  - `MigrationV65_Metadata_ShouldBeCorrect`
  - `Up_WhenLegacySchemaAndStaleCrsqlTriggers_ShouldAddColumnsBackfillAndDropTriggers`
  - `Up_WhenColumnsExistWithNullDatesAndStaleTriggers_ShouldBackfillWithoutArityError`
  - `Up_WhenDatesAlreadyPresent_ShouldPreserveThem`
  - `Up_WhenPeriodosTableDoesNotExist_ShouldReturnGracefully`
  - `Up_WhenSchemaAndDatesAreComplete_ShouldNotDropCrsqlTriggers` (Fase 752)
  - `Up_WhenInicioPresenteYFinNulo_ShouldFillOnlyFin` (Fase 753)
  - `Up_WhenInicioEs20260815_ShouldApplyV56Rule` (Fases 753, 757)
- `MigrationV50Tests.cs`:
  - `Up_WhenCrsqlTriggersAreStale_ShouldDropTriggersAndNotFailOnBackfill`
  - `Up_WhenInicioPresenteYFinNulo_ShouldFillOnlyFin` (Fase 759)
- `MigrationV51Tests.cs`:
  - `Up_WhenCrsqlTriggersAreStale_ShouldDropTriggersAndNotFailOnBackfill`
  - `Up_WhenInicioPresenteYFinNulo_ShouldFillOnlyFin` (Fase 759)
- `AcademicRepositoryPeriodoTests.cs`:
  - `InsertPeriodo_WhenEsquemaSinVentanaYConFechas_ShouldReturnFailure`
