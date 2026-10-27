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
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;

namespace ConversorPDF.Tests.Integration.Academic
{
    public class AcademicDbInitializationTests
    {
        private readonly string _dbPath;

        public AcademicDbInitializationTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"AcademicInitTest_{Guid.NewGuid():N}.db.enc");
            SqliteHelper.SetTestVaultPath(_dbPath);
        }

        [Fact]
        public async Task InitializeDatabase_WhenFresh_ShouldCreateLatestSchemaAndSeedData()
        {
            // 1. Arrange: Borrar base de datos existente si existe para forzar recreación
            if (File.Exists(_dbPath))
            {
                // Intentar borrar con reintentos por si Excel la tiene bloqueada (aunque en tests no debería)
                for (int i = 0; i < 3; i++)
                {
                    try { File.Delete(_dbPath); break; }
                    catch { Thread.Sleep(500); }
                }
            }

            // 2. Act: Forzar inicialización llamando a GetContext()
            // Esto dispara InitializeDatabase() internamente si el archivo no existe.
            using (var context = SqliteHelper.GetContext())
            {
                var connection = (SQLiteConnection)context.Open();

                // 3. Assert: Verificar Versión Actual
                int finalVersion = connection.QueryFirstOrDefault<int>("SELECT Version FROM __SchemaVersion ORDER BY Version DESC LIMIT 1");
                finalVersion.Should().BeGreaterThanOrEqualTo(46, "La base de datos nueva debe nacer directamente en la versión 46 o superior.");

                // 4. Assert: Verificar Seed Data (Asegurar que el script se ejecutó)
                var countTerceros = connection.ExecuteScalar<int>("SELECT COUNT(*) FROM terceros;");
                countTerceros.Should().BeGreaterThan(0, "El seed data de terceros debe haberse insertado exitosamente.");

                // Assert: Verificar correcta inicialización de permisos
                var permisosCount = connection.ExecuteScalar<int>("SELECT COUNT(*) FROM permisos;");
                permisosCount.Should().Be(5, "Deben existir 5 perfiles de permisos tras inicialización.");
                var adminPermisos = connection.ExecuteScalar<string>("SELECT opcionespermitidas FROM permisos WHERE perfil = 'Admin';");
                adminPermisos.Should().Contain("Permisos", "El perfil Admin debe tener permiso 'Permisos'.");

                // 5. Assert: Verificar que GetSesionesAsync no falla (Fix Error 1)
                var repo = new VinculacionQueryRepository(() => context.Open());
                var result = await repo.GetSesionesAsync(idSede: null, idTercero: null, idPeriodo: null, idPrograma: null, modalidadPrograma: null, subunidad: null, semestre: null, soloViaticas: false, fechaDesde: null, fechaHasta: null, limit: 10, offset: 0, ct: CancellationToken.None);
                
                if (result.IsFailure)
                {
                    Assert.Fail($"GetSesionesAsync debe funcionar correctamente tras la inicialización sin lanzar NRE. Error: {result.Error}");
                }
                
                context.CloseConnectionResiliently();
            }
        }
    }
}
