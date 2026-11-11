using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
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

                    CREATE TABLE ag_violacion (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        id_individuo INTEGER NOT NULL,
                        id_gen INTEGER NULL,
                        regla INTEGER NOT NULL,
                        es_dura INTEGER NOT NULL DEFAULT 1,
                        mensaje TEXT NOT NULL DEFAULT '',
                        detalle_json TEXT NOT NULL DEFAULT '{}',
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

        [Fact]
        public async Task GetViolacionesByIndividuoAsync_WhenViolationsExist_ReturnsOrderedByEsDuraDescAndRegla()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
                    VALUES (1, 10, 3, 0, 'Violación Soft 1', '{}', 'VALLE');

                    INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
                    VALUES (1, 11, 1, 1, 'Violación Hard 1', '{}', 'VALLE');

                    INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
                    VALUES (2, 20, 1, 1, 'Violación Otro Individuo', '{}', 'VALLE');
                ");
            }

            var repo = new AgScheduleRepository(() => new SQLiteConnection(_connectionString));

            // Act
            var result = await repo.GetViolacionesByIndividuoAsync(1, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().HaveCount(2);

            var lista = result.Value.ToList();
            lista[0].EsDura.Should().BeTrue();
            lista[0].Mensaje.Should().Be("Violación Hard 1");
            lista[0].Regla.Should().Be((AgReglaViolacion)1);

            lista[1].EsDura.Should().BeFalse();
            lista[1].Mensaje.Should().Be("Violación Soft 1");
            lista[1].Regla.Should().Be((AgReglaViolacion)3);
        }
    }
}
