# Aislamiento del Log de Pruebas y Descarte del Falso Positivo V66

**Fecha:** 2026-10-22  
**Módulo:** Core / Infraestructura / Pruebas (`ConversorPDF.XLL.Core`, `ConversorPDF.Tests`)  
**Ronda:** 207 (Fases 1066 a 1070)

## Propósito
Aislar la salida de registro de eventos (`LoggerHelper`) durante la ejecución de pruebas automatizadas y documentar el descarte del falso positivo observado en auditorías de logs asociado a la migración `MigrationV66_RenombrarModalidadTutoria`:
1. El error registrado `[V66] Bloque vinculaciones falló.` acompañado de `[TEST-V66] fallo intencional` proviene exclusivamente de la prueba unitaria `MigrationV66Tests.Up_WhenBlockFails_ShouldRollbackAndThrow()`.
2. Dicha prueba inyecta un trigger SQLite intencional con `RAISE(FAIL, '[TEST-V66] fallo intencional')` para asegurar el comportamiento de rollback transaccional ante excepciones.
3. Debido a que `LoggerHelper` compartía la misma ruta de archivo estática en `%LOCALAPPDATA%\ConversorPDF\log\log_yyyyMMdd.txt`, las ejecuciones de la suite de pruebas contaminaban el log que se audita en entornos de producción.

## Diagnóstico y Causa Raíz
1. **Evidencia en Log:** En `log_20261022.txt` se detectó un bloque de error generado desde `...\ConversorPDF.Tests\bin\Release\net48`.
2. **Inexistencia de Constraint:** La tabla `vinculaciones` en producción (`AcademicSchema.sql`) no posee `CHECK` ni `UNIQUE` sobre `modalidadVinculacion`. En bases de datos reales de producción, `MigrationV66` ejecutó exitosamente 4 veces consecutivas el mismo día sin huecos de versión.
3. **Causa Raíz:** El campo `_logDirectory` en `LoggerHelper` era `private static readonly string`, fijado irrevocablemente en el constructor estático a la carpeta local de la aplicación. No existía mecanismo de redirección para entornos de testing.

## Cambios Realizados
1. **Redirección configurable de la salida de log (`LoggerHelper.cs`, Fase 1066):**
   - Campo `_logDirectory` modificado a `private static string`.
   - Expuesta propiedad `internal static string DirectorioLog { get; set; }` para reasignar la carpeta destino de los logs. Si el directorio no existe, se crea dinámicamente. Al recalcularse la ruta en cada iteración de la cola asíncrona, el cambio toma efecto inmediatamente sin reiniciar tareas de fondo.
2. **Shim de `ModuleInitializerAttribute` para .NET Framework 4.8 (`ModuleInitializerAttribute.cs`, Fase 1067):**
   - Declarado atributo `[AttributeUsage(AttributeTargets.Method, Inherited = false)]` bajo el namespace `System.Runtime.CompilerServices` protegido por `#if !NET5_0_OR_GREATER`.
   - Permite al compilador C# emitir inicializadores de módulo (`.cctor`) en el ensamblado de pruebas `net48`.
3. **Bootstrap de pruebas con redirección a `%TEMP%` (`TestLogBootstrap.cs`, Fase 1068):**
   - Clase interna con método anotado `[ModuleInitializer]` que asigna `LoggerHelper.DirectorioLog = Path.Combine(Path.GetTempPath(), "ConversorPDF.Tests", "log")`.
   - Se ejecuta automáticamente al cargar el ensamblado de pruebas antes del primer test, garantizando que ninguna prueba escriba en el directorio de log de producción.
4. **Prueba de no regresión de la redirección (`LoggerHelperRedirectionTests.cs`, Fase 1069):**
   - Suite unitaria que valida la redirección temporal a una carpeta aislada, la emisión efectiva del log y la posterior restauración sin efectos colaterales.
5. **Trazabilidad y clarificación en código de prueba (`MigrationV66Tests.cs`, Fase 1070):**
   - Documentado en el cuerpo del test `Up_WhenBlockFails_ShouldRollbackAndThrow` que el error es intencional y que su salida se canaliza a `%TEMP%`.

## Conclusión Operativa
El mensaje de fallo en V66 **no representa un error en el código de producción ni en la base de datos**. Las Fases 1066 a 1070 garantizan la total separación de auditoría entre ejecuciones de pruebas y actividad productiva.
