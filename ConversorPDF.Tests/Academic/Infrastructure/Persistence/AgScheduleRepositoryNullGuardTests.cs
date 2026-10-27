using System;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AgScheduleRepositoryNullGuardTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public AgScheduleRepositoryNullGuardTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AgScheduleRepoNullGuardTest_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE ag_corrida (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idperiodo INTEGER NOT NULL,
                        idprograma INTEGER NOT NULL,
                        idaula INTEGER NOT NULL,
                        fecha_inicio_corrida TEXT NOT NULL DEFAULT '',
                        fecha_fin_corrida TEXT NOT NULL DEFAULT '',
                        estado INTEGER NOT NULL,
                        poblacion_tamano INTEGER NOT NULL,
                        generaciones_max INTEGER NOT NULL,
                        generaciones_ejecutadas INTEGER NOT NULL,
                        tasa_cruce REAL NOT NULL,
                        tasa_mutacion REAL NOT NULL,
                        elitismo INTEGER NOT NULL,
                        mejor_fitness REAL NOT NULL,
                        hard_violations INTEGER NOT NULL,
                        soft_violations INTEGER NOT NULL,
                        vacantes INTEGER NOT NULL DEFAULT 0,
                        id_individuo_optimo INTEGER NOT NULL,
                        parametros_json TEXT NOT NULL DEFAULT '{}',
                        subunidad TEXT NOT NULL DEFAULT ''
                    );
                ");
            }
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
            catch
            {
                // Best effort cleanup
            }
        }

        [Fact]
        public async Task InsertCorridaAsync_WhenDatesAndStringsAreNull_ShouldNotFailNotNullConstraint()
        {
            // Arrange
            var repo = new AgScheduleRepository(() => new SQLiteConnection(_connectionString));
            var corrida = new AgCorrida
            {
                IdPeriodo = 1,
                IdPrograma = 2,
                IdAula = 3,
                FechaInicioCorrida = null, // Debería normalizarse a ""
                FechaFinCorrida = null,    // Debería normalizarse a ""
                Estado = AgEstadoCorrida.EnProgreso,
                PoblacionTamano = 50,
                GeneracionesMax = 100,
                GeneracionesEjecutadas = 0,
                TasaCruce = 0.8,
                TasaMutacion = 0.1,
                Elitismo = 2,
                MejorFitness = 0,
                HardViolations = 0,
                SoftViolations = 0,
                IdIndividuoOptimo = 0,
                ParametrosJson = null,     // Debería normalizarse a "{}"
                Subunidad = null          // Debería normalizarse a ""
            };

            // Act
            var result = await repo.InsertCorridaAsync(corrida, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().BeGreaterThan(0);

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                var row = conn.QuerySingle("SELECT fecha_inicio_corrida, fecha_fin_corrida, parametros_json, subunidad FROM ag_corrida WHERE id = @id", new { id = result.Value });
                ((string)row.fecha_inicio_corrida).Should().Be(string.Empty);
                ((string)row.fecha_fin_corrida).Should().Be(string.Empty);
                ((string)row.parametros_json).Should().Be("{}");
                ((string)row.subunidad).Should().Be(string.Empty);
            }
        }
    }
}
