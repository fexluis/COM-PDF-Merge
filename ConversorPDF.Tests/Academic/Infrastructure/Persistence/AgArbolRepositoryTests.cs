using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AgArbolRepositoryTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;
        private readonly AgArbolRepository _repository;

        public AgArbolRepositoryTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AgArbolRepoTest_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";

            // Aplicar migraciones V53 y V62
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    var mig53 = new MigrationV53_AddAgArbolTables();
                    mig53.Up(conn, tx);
                    var mig62 = new MigrationV62_AddAgAulaVentanasYBloquesSemilla();
                    mig62.Up(conn, tx);
                    tx.Commit();
                }
            }

            _repository = new AgArbolRepository(() => new SQLiteConnection(_connectionString));
        }

        public void Dispose()
        {
            try
            {
                if (File.Exists(_dbPath))
                {
                    File.Delete(_dbPath);
                }
            }
            catch { }
        }

        [Fact]
        public async Task GuardarAsync_WhenValidProfileAndNodes_ShouldSaveAndRetrieveCorrectly()
        {
            // Arrange
            var perfil = new AgArbol
            {
                Nombre = "Perfil Ingeniería 2026-02",
                Descripcion = "Selección Semestre I y II",
                IdPeriodo = 2,
                FechaInicio = "2026-08-15",
                FechaFin = "2026-12-12",
                VentanaOrigen = "Periodos",
                ManualVentana = 0,
                Poblacion = 60,
                Generaciones = 250,
                TasaCruce = 0.88,
                TasaMutacion = 0.12,
                Elitismo = 3,
                Subunidad = "VALLE",
                Usuario = "admin"
            };

            var nodos = new List<AgArbolNodo>
            {
                new AgArbolNodo
                {
                    IdPeriodo = 2,
                    IdSede = 5,
                    IdPrograma = 10,
                    Semestre = "I",
                    Grupo = "A",
                    Modalidad = "Presencial",
                    AulaRef = "2|5|10|I|A|PRESENCIAL",
                    Capacidad = 35,
                    TipoAula = "ESTANDAR",
                    Subunidad = "VALLE"
                },
                new AgArbolNodo
                {
                    IdPeriodo = 2,
                    IdSede = 5,
                    IdPrograma = 10,
                    Semestre = "II",
                    Grupo = "B",
                    Modalidad = "Virtual",
                    AulaRef = "2|5|10|II|B|VIRTUAL",
                    Capacidad = 999,
                    TipoAula = "VIRTUAL",
                    EsVirtual = 1,
                    Subunidad = "VALLE"
                }
            };

            // Act
            var saveRes = await _repository.GuardarAsync(perfil, nodos, CancellationToken.None);

            // Assert
            saveRes.IsSuccess.Should().BeTrue();
            saveRes.Value.Should().BeGreaterThan(0);

            var getRes = await _repository.ObtenerAsync(saveRes.Value, CancellationToken.None);
            getRes.IsSuccess.Should().BeTrue();
            getRes.Value.Nombre.Should().Be("Perfil Ingeniería 2026-02");
            getRes.Value.Poblacion.Should().Be(60);
            getRes.Value.TasaCruce.Should().Be(0.88);

            var nodosRes = await _repository.ObtenerNodosAsync(saveRes.Value, CancellationToken.None);
            nodosRes.IsSuccess.Should().BeTrue();
            nodosRes.Value.Should().HaveCount(2);
            nodosRes.Value.First().AulaRef.Should().Be("2|5|10|I|A|PRESENCIAL");
            nodosRes.Value.Last().Modalidad.Should().Be("Virtual");
        }

        [Fact]
        public async Task GuardarAsync_WhenSameNameAndSubunidad_ShouldOverwriteProfileAndNodes()
        {
            // Arrange
            var perfil1 = new AgArbol
            {
                Nombre = "Perfil Sobrescritura",
                IdPeriodo = 1,
                Poblacion = 50,
                Subunidad = "VALLE"
            };
            var nodos1 = new List<AgArbolNodo>
            {
                new AgArbolNodo { IdPeriodo = 1, IdSede = 1, IdPrograma = 1, Semestre = "I", Grupo = "A", Modalidad = "Presencial", AulaRef = "1|1|1|I|A|PRESENCIAL", Subunidad = "VALLE" }
            };

            await _repository.GuardarAsync(perfil1, nodos1, CancellationToken.None);

            var perfil2 = new AgArbol
            {
                Nombre = "Perfil Sobrescritura",
                IdPeriodo = 2,
                Poblacion = 80,
                Subunidad = "VALLE"
            };
            var nodos2 = new List<AgArbolNodo>
            {
                new AgArbolNodo { IdPeriodo = 2, IdSede = 1, IdPrograma = 1, Semestre = "II", Grupo = "B", Modalidad = "Virtual", AulaRef = "2|1|1|II|B|VIRTUAL", Subunidad = "VALLE" },
                new AgArbolNodo { IdPeriodo = 2, IdSede = 1, IdPrograma = 1, Semestre = "III", Grupo = "C", Modalidad = "Presencial", AulaRef = "2|1|1|III|C|PRESENCIAL", Subunidad = "VALLE" }
            };

            // Act - guardar nuevamente con el mismo nombre y subunidad
            var overwriteRes = await _repository.GuardarAsync(perfil2, nodos2, CancellationToken.None);

            // Assert
            overwriteRes.IsSuccess.Should().BeTrue();

            var listRes = await _repository.ListarAsync("VALLE", CancellationToken.None);
            listRes.Value.Should().HaveCount(1);
            listRes.Value.First().Poblacion.Should().Be(80);

            var nodosRes = await _repository.ObtenerNodosAsync(overwriteRes.Value, CancellationToken.None);
            nodosRes.Value.Should().HaveCount(2);
            nodosRes.Value.Select(n => n.AulaRef).Should().Contain("2|1|1|II|B|VIRTUAL");
        }

        [Fact]
        public async Task EliminarAsync_WhenCalled_ShouldDeleteProfileAndCascadingNodes()
        {
            // Arrange
            var perfil = new AgArbol { Nombre = "Perfil Para Borrar", IdPeriodo = 1, Subunidad = "TEST" };
            var nodos = new List<AgArbolNodo>
            {
                new AgArbolNodo { IdPeriodo = 1, AulaRef = "REF1", Subunidad = "TEST" }
            };
            var saveRes = await _repository.GuardarAsync(perfil, nodos, CancellationToken.None);
            long id = saveRes.Value;

            // Act
            var deleteRes = await _repository.EliminarAsync(id, CancellationToken.None);

            // Assert
            deleteRes.IsSuccess.Should().BeTrue();

            var getRes = await _repository.ObtenerAsync(id, CancellationToken.None);
            getRes.IsFailure.Should().BeTrue();

            var nodosRes = await _repository.ObtenerNodosAsync(id, CancellationToken.None);
            nodosRes.Value.Should().BeEmpty();
        }

        [Fact]
        public async Task GuardarAsync_WhenNonExistentProfileId_ShouldFailAndNotCreateOrphanNodes()
        {
            // Arrange - Intento de actualizar un perfil con ID que no existe
            var perfil = new AgArbol { Id = 99999, Nombre = "Perfil Inexistente", IdPeriodo = 1, Subunidad = "TEST" };
            var nodos = new List<AgArbolNodo>
            {
                new AgArbolNodo { IdPeriodo = 1, AulaRef = "REF_INEXISTENTE", Subunidad = "TEST" }
            };

            // Act
            var saveRes = await _repository.GuardarAsync(perfil, nodos, CancellationToken.None);

            // Assert (Fase 247: debe retornar fallo y no dejar nodos huérfanos)
            saveRes.IsFailure.Should().BeTrue();
            saveRes.Error.Should().Contain("no existe");

            var nodosRes = await _repository.ObtenerNodosAsync(99999, CancellationToken.None);
            nodosRes.Value.Should().BeEmpty();
        }

        [Fact]
        public async Task GuardarAsync_WhenCaseAndWhitespaceVariationInName_ShouldOverwriteProfile()
        {
            // Arrange
            var perfil1 = new AgArbol { Nombre = "Perfil Exacto", IdPeriodo = 1, Subunidad = "VALLE", Poblacion = 40 };
            var nodos1 = new List<AgArbolNodo>
            {
                new AgArbolNodo { IdPeriodo = 1, AulaRef = "REF1", Subunidad = "VALLE" }
            };
            await _repository.GuardarAsync(perfil1, nodos1, CancellationToken.None);

            // Act - Guardar con espacios y variación de mayúsculas (Fase 248)
            var perfil2 = new AgArbol { Nombre = "  perfil exacto  ", IdPeriodo = 1, Subunidad = "VALLE", Poblacion = 95 };
            var nodos2 = new List<AgArbolNodo>
            {
                new AgArbolNodo { IdPeriodo = 1, AulaRef = "REF2", Subunidad = "VALLE" }
            };
            var overwriteRes = await _repository.GuardarAsync(perfil2, nodos2, CancellationToken.None);

            // Assert
            overwriteRes.IsSuccess.Should().BeTrue();

            var listRes = await _repository.ListarAsync("VALLE", CancellationToken.None);
            listRes.Value.Should().HaveCount(1);
            listRes.Value.First().Poblacion.Should().Be(95);
        }
    }
}
