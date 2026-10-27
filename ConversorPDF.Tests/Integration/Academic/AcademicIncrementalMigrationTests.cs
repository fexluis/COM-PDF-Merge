using System;
using System.IO;
using System.Linq;
using System.Data.SQLite;
using System.Threading;
using System.Threading.Tasks;
using Dapper;
using FluentAssertions;
using Xunit;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;

namespace ConversorPDF.Tests.Integration.Academic
{
    public class AcademicIncrementalMigrationTests
    {
        private readonly string _dbPath;

        public AcademicIncrementalMigrationTests()
        {
            // Usar un path específico para este test para no interferir con la BD real
            _dbPath = Path.Combine(Path.GetTempPath(), "AcademicMigrationTest_" + Guid.NewGuid().ToString("N") + ".db.enc");
            SqliteHelper.SetTestVaultPath(_dbPath);
        }

        [Fact]
        public void EnsureSchemaUpToDate_WhenV2Exists_ShouldMigrateToLatest()
        {
            // 1. Arrange: Crear una base de datos "huérfana" en versión 2
            // Obtener el password real vía Reflection para que PersistAndReopen (que usa GetSecurePassword) no rompa la cadena
            var getPasswordMethod = typeof(SqliteHelper).GetMethod("GetSecurePassword", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            var password = (System.Security.SecureString)getPasswordMethod.Invoke(null, null);

            try
            {
                using (var context = new EncryptedSQLiteContext(_dbPath, password))
                {
                    var connection = (SQLiteConnection)context.Open();
                    
                    // Crear tablas base mínimas para que las migraciones no fallen
                    connection.Execute(@"
                        CREATE TABLE IF NOT EXISTS __SchemaVersion (Version INTEGER PRIMARY KEY);
                        INSERT INTO __SchemaVersion (Version) VALUES (2);
                        
                        CREATE TABLE IF NOT EXISTS programas (id INTEGER PRIMARY KEY AUTOINCREMENT, nombreprograma TEXT, abreviatura TEXT, tipo TEXT);
                        CREATE TABLE IF NOT EXISTS terceros (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT, cedula TEXT, residencia TEXT, activo INTEGER DEFAULT 1);
                        CREATE TABLE IF NOT EXISTS sedes (id INTEGER PRIMARY KEY AUTOINCREMENT, ciudad TEXT);
                        CREATE TABLE IF NOT EXISTS asignaturas (id INTEGER PRIMARY KEY AUTOINCREMENT, idprograma INTEGER, nombreasignatura TEXT, abreviatura TEXT);
                        CREATE TABLE IF NOT EXISTS vinculaciones (id INTEGER PRIMARY KEY AUTOINCREMENT, idprograma INTEGER, idtercero INTEGER, idasignatura INTEGER, idsede INTEGER, semestre TEXT, fecha_inicio TEXT, fecha_fin TEXT, total_horas INTEGER);
                        CREATE TABLE IF NOT EXISTS sesiones (id INTEGER PRIMARY KEY AUTOINCREMENT, idvinculacion INTEGER, idtercero INTEGER, fechasesionInicial TEXT, fechasesionFinal TEXT, horas INTEGER, duracion TEXT, viatica INTEGER);
                    ");

                    context.CloseConnectionResiliently();
                    byte[] bytes = EncryptedSQLiteContext.LeerBytesConRetrySync(context.TempDbPath);
                    File.WriteAllBytes(_dbPath, DatabaseEncryptionEngine.Encrypt(bytes, password));
                }

                // 2. Act: Ejecutar EnsureSchemaUpToDate vía reflejo
                var tempContext = new EncryptedSQLiteContext(_dbPath, password);
                try
                {
                    var method = typeof(SqliteHelper).GetMethod("EnsureSchemaUpToDate", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                    method.Invoke(null, new object[] { tempContext });
                }
                finally
                {
                    tempContext.Dispose();
                }

                using (var context = new EncryptedSQLiteContext(_dbPath, password))
                {
                    var connection = (SQLiteConnection)context.Open();

                    // 3. Assert: Verificar Versión Actual (v45)
                    var version = connection.ExecuteScalar<int>("SELECT Version FROM __SchemaVersion ORDER BY Version DESC LIMIT 1;");
                    version.Should().BeGreaterThanOrEqualTo(46, "La base de datos debe haber migrado exitosamente a la versión 46 o superior.");

                    // 4. Assert: Verificar que la tabla 'periodos' existe (Migración v3)
                    var periodosExists = connection.ExecuteScalar<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='periodos';");
                    periodosExists.Should().Be(1, "La tabla 'periodos' debería haber sido creada en la migración v3.");

                    // Assert: Verificar que la tabla 'festivos' existe (Migración v13)
                    var festivosExists = connection.ExecuteScalar<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='festivos';");
                    festivosExists.Should().Be(1, "La tabla 'festivos' debería haber sido creada en la migración v13.");

                    var festivoColumns = connection.Query("PRAGMA table_info(festivos)").Select(c => (string)c.name).ToList();
                    festivoColumns.Should().Contain("tipofestivo", "La columna 'tipofestivo' debería haber sido añadida a 'festivos' en la migración v20.");

                    // 5. Assert: Verificar que la columna 'idperiodo' existe en vinculaciones (Migración v3)
                    var columns = connection.Query("PRAGMA table_info(vinculaciones)").Select(c => (string)c.name).ToList();
                    columns.Should().Contain("idperiodo", "La columna 'idperiodo' debería haber sido añadida a 'vinculaciones'.");

                    // 6. Assert: Verificar que vw_SesionesDetalle tiene 'CedulaTercero' (Migración v6) y 'AbreviaturaAsignatura' (Migración v7)
                    var viewColumns = connection.Query("PRAGMA table_info(vw_SesionesDetalle)").Select(c => (string)c.name).ToList();
                    viewColumns.Should().Contain("CedulaTercero", "La vista vw_SesionesDetalle debería tener la columna CedulaTercero tras migrar a v6.");
                    viewColumns.Should().Contain("AsignaturaAbreviatura", "La vista vw_SesionesDetalle debería tener la columna AsignaturaAbreviatura tras migrar a v8.");
                    viewColumns.Should().Contain("IdPrograma", "La vista vw_SesionesDetalle debería tener la columna IdPrograma tras migrar a v11.");
                    viewColumns.Should().Contain("SedeDepartamento", "La vista vw_SesionesDetalle debería tener la columna SedeDepartamento tras migrar a v12.");
                    viewColumns.Should().Contain("HorarioSesion", "La vista vw_SesionesDetalle debería tener la columna HorarioSesion tras migrar a v14.");
                    viewColumns.Should().Contain("FechaSesion", "La vista vw_SesionesDetalle debería tener la columna FechaSesion tras migrar a v14.");
                    viewColumns.Should().NotContain("FechaInicial", "La columna FechaInicial ya no debe existir desde v14.");
                    viewColumns.Should().NotContain("FechaFinal", "La columna FechaFinal ya no debe existir desde v14.");
                    
                    // 7. Assert: Verificar que la columna 'grupo' existe en vinculaciones (Migración v8)
                    columns.Should().Contain("grupo", "La columna 'grupo' debería haber sido añadida a 'vinculaciones' en la migración v8.");
                    
                    // 8. Assert: Verificar que la columna 'tipo' se renombró a 'modalidad' en programas (Migración v9)
                    var progCols = connection.Query("PRAGMA table_info(programas)").Select(c => (string)c.name).ToList();
                    progCols.Should().Contain("modalidad", "La columna 'tipo' debería haber sido renombrada a 'modalidad' en la migración v9.");
                    progCols.Should().NotContain("tipo", "La columna 'tipo' ya no debería existir.");
                    
                    // 9. Assert: Verificar que la columna 'modalidadVinculacion' existe (Migración v15)
                    columns.Should().Contain("modalidadVinculacion", "La columna 'modalidadVinculacion' debería haber sido añadida a 'vinculaciones' en la migración v15.");
                    viewColumns.Should().Contain("Modalidad", "La vista vw_SesionesDetalle debería tener la columna Modalidad tras migrar a v15.");
                    
                    // 10. Assert: Verificar tabla 'salvedades_conflictos' (Migración v16)
                    var salvedadesExists = connection.ExecuteScalar<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='salvedades_conflictos';");
                    salvedadesExists.Should().Be(1, "La tabla 'salvedades_conflictos' debería haber sido creada en la migración v16.");
                    
                    context.CloseConnectionResiliently();
                }
            }
            finally
            {
                if (File.Exists(_dbPath)) File.Delete(_dbPath);
            }
        }
    }
}
