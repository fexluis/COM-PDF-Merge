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
    public class AgScheduleRepositoryCorridaVacantesTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;
        private readonly AgScheduleRepository _sut;

        public AgScheduleRepositoryCorridaVacantesTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AgScheduleRepoCorridaVacantes_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
                        soft_violations REAL NOT NULL,
                        vacantes INTEGER NOT NULL DEFAULT 0,
                        id_individuo_optimo INTEGER NULL,
                        parametros_json TEXT NOT NULL DEFAULT '{}',
                        subunidad TEXT NOT NULL DEFAULT ''
                    );

                    CREATE TABLE ag_individuo (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        id_corrida INTEGER NOT NULL,
                        generacion INTEGER NOT NULL DEFAULT 0,
                        fitness REAL NOT NULL DEFAULT 0,
                        hard_violations INTEGER NOT NULL DEFAULT 0,
                        soft_violations REAL NOT NULL DEFAULT 0,
                        estado INTEGER NOT NULL DEFAULT 0,
                        subunidad TEXT NOT NULL DEFAULT ''
                    );

                    CREATE TABLE ag_gen (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        id_individuo INTEGER NOT NULL,
                        idperiodo INTEGER NOT NULL DEFAULT 0,
                        idasignatura INTEGER NOT NULL DEFAULT 0,
                        idprograma INTEGER NOT NULL DEFAULT 0,
                        idsede INTEGER NOT NULL DEFAULT 0,
                        semestre TEXT NOT NULL DEFAULT '',
                        grupo TEXT NOT NULL DEFAULT '',
                        idtercero INTEGER NOT NULL DEFAULT 0,
                        id_aula INTEGER NOT NULL DEFAULT 0,
                        id_bloque_horario INTEGER NOT NULL DEFAULT 0,
                        dia_semana INTEGER NOT NULL DEFAULT 0,
                        horas_semanales REAL NOT NULL DEFAULT 0,
                        total_horas_asignatura REAL NOT NULL DEFAULT 0,
                        fecha_inicio_estimada TEXT NOT NULL DEFAULT '',
                        fecha_fin_estimada TEXT NOT NULL DEFAULT '',
                        es_fijo INTEGER NOT NULL DEFAULT 0,
                        es_vacante INTEGER NOT NULL DEFAULT 0,
                        motivo_vacante TEXT NOT NULL DEFAULT '',
                        dentro_de_ventana INTEGER NOT NULL DEFAULT 0,
                        estado INTEGER NOT NULL DEFAULT 0,
                        subunidad TEXT NOT NULL DEFAULT '',
                        indice_dia_patron INTEGER NOT NULL DEFAULT 0,
                        patron_semanal TEXT NOT NULL DEFAULT '',
                        dias_patron TEXT NOT NULL DEFAULT '',
                        clave_grupo TEXT NOT NULL DEFAULT '',
                        horas_gen REAL NOT NULL DEFAULT 0
                    );

                    CREATE TABLE ag_violacion (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        id_individuo INTEGER NOT NULL,
                        id_gen INTEGER NULL,
                        regla TEXT NOT NULL DEFAULT '',
                        es_dura INTEGER NOT NULL DEFAULT 1,
                        mensaje TEXT NOT NULL DEFAULT '',
                        detalle_json TEXT NOT NULL DEFAULT '{}',
                        subunidad TEXT NOT NULL DEFAULT ''
                    );
                ");
            }

            _sut = new AgScheduleRepository(() => new SQLiteConnection(_connectionString));
        }

        [Fact]
        public async Task UpdateCorridaAsync_WhenVacantesUpdated_ShouldPersistVacantesInDatabase()
        {
            // Arrange
            var corrida = new AgCorrida
            {
                IdPeriodo = 20261,
                IdPrograma = 10,
                IdAula = 1,
                FechaInicioCorrida = "2026-11-15 10:00:00",
                FechaFinCorrida = "2026-11-15 10:05:00",
                Estado = AgEstadoCorrida.Completada,
                PoblacionTamano = 50,
                GeneracionesMax = 100,
                GeneracionesEjecutadas = 100,
                TasaCruce = 0.85,
                TasaMutacion = 0.1,
                Elitismo = 2,
                MejorFitness = 95.5,
                HardViolations = 0,
                SoftViolations = 2.0,
                Vacantes = 0,
                ParametrosJson = "{}",
                Subunidad = "FAC1"
            };

            var insertRes = await _sut.InsertCorridaAsync(corrida, CancellationToken.None);
            insertRes.IsSuccess.Should().BeTrue();
            long idCorrida = insertRes.Value;

            // Act - modificar vacantes y actualizar
            corrida.Vacantes = 7;
            corrida.GeneracionesEjecutadas = 120;
            corrida.MejorFitness = 99.0;
            var updateRes = await _sut.UpdateCorridaAsync(corrida, CancellationToken.None);

            // Assert
            updateRes.IsSuccess.Should().BeTrue();

            var getRes = await _sut.GetCorridaByIdAsync(idCorrida, CancellationToken.None);
            getRes.IsSuccess.Should().BeTrue();
            getRes.Value.Vacantes.Should().Be(7);
            getRes.Value.GeneracionesEjecutadas.Should().Be(120);
            getRes.Value.MejorFitness.Should().Be(99.0);

            // Verificación directa en base de datos
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                int vacantesDb = await conn.ExecuteScalarAsync<int>(
                    "SELECT vacantes FROM ag_corrida WHERE id = @idCorrida;",
                    new { idCorrida });
                vacantesDb.Should().Be(7);
            }
        }

        [Fact]
        public async Task SaveIndividuoAsync_WhenViolacionHasIndicesGenes_ShouldResolveIdGenAndIdGenesInDetalle()
        {
            // Arrange
            var corrida = new AgCorrida
            {
                IdPeriodo = 20261,
                IdPrograma = 10,
                IdAula = 1
            };
            var corridaRes = await _sut.InsertCorridaAsync(corrida, CancellationToken.None);
            corridaRes.IsSuccess.Should().BeTrue();

            var gen1 = new AgGen
            {
                IdPeriodo = 20261,
                IdAsignatura = 101,
                IdPrograma = 10,
                IdSede = 1,
                Semestre = "1",
                Grupo = "G1",
                IdTercero = 1001,
                IdAula = 1,
                IdBloqueHorario = 10,
                DiaSemana = 1,
                HorasSemanales = 4,
                TotalHorasAsignatura = 64
            };

            var gen2 = new AgGen
            {
                IdPeriodo = 20261,
                IdAsignatura = 102,
                IdPrograma = 10,
                IdSede = 1,
                Semestre = "2",
                Grupo = "G2",
                IdTercero = 1001,
                IdAula = 2,
                IdBloqueHorario = 10,
                DiaSemana = 1,
                HorasSemanales = 4,
                TotalHorasAsignatura = 64
            };

            var violacion = new AgViolacion
            {
                Regla = AgReglaViolacion.R1_SuperposicionDocente,
                EsDura = true,
                Mensaje = "Docente con cruce de horario",
                DetalleJson = "{\"motivo\":\"Cruce\",\"indicesGenes\":[0,1]}"
            };

            var individuo = new AgIndividuo
            {
                IdCorrida = corridaRes.Value,
                Generacion = 1,
                Fitness = 80.0,
                HardViolations = 1,
                SoftViolations = 0,
                Genes = new List<AgGen> { gen1, gen2 },
                Violaciones = new List<AgViolacion> { violacion }
            };

            // Act
            var saveRes = await _sut.SaveIndividuoAsync(individuo, CancellationToken.None);

            // Assert
            saveRes.IsSuccess.Should().BeTrue();
            long idIndividuo = saveRes.Value;

            gen1.Id.Should().BeGreaterThan(0);
            gen2.Id.Should().BeGreaterThan(0);

            // Verificar violacion en el modelo en memoria
            violacion.IdGen.Should().Be(gen1.Id);

            // Verificar violacion persistida en base de datos
            var violacionesRes = await _sut.GetViolacionesByIndividuoAsync(idIndividuo, CancellationToken.None);
            violacionesRes.IsSuccess.Should().BeTrue();
            violacionesRes.Value.Should().HaveCount(1);

            var violDb = violacionesRes.Value.First();
            violDb.IdGen.Should().Be(gen1.Id);
            violDb.DetalleJson.Should().Contain("\"idGenes\":[" + gen1.Id + "," + gen2.Id + "]");
        }

        [Fact]
        public async Task SaveIndividuoAsync_WhenViolacionHasNoIndicesGenes_ShouldNotOverwriteIdGenOrDetalle()
        {
            // Arrange
            var corrida = new AgCorrida { IdPeriodo = 20261, IdPrograma = 10, IdAula = 1 };
            var corridaRes = await _sut.InsertCorridaAsync(corrida, CancellationToken.None);

            var gen = new AgGen { IdPeriodo = 20261, IdAsignatura = 101, IdPrograma = 10, IdSede = 1 };
            var violacion = new AgViolacion
            {
                Regla = AgReglaViolacion.R4_CompatibilidadAula,
                EsDura = false,
                Mensaje = "Regla general sin genes asociados",
                DetalleJson = "{\"info\":\"sin genes\"}",
                IdGen = null
            };

            var individuo = new AgIndividuo
            {
                IdCorrida = corridaRes.Value,
                Generacion = 1,
                Genes = new List<AgGen> { gen },
                Violaciones = new List<AgViolacion> { violacion }
            };

            // Act
            var saveRes = await _sut.SaveIndividuoAsync(individuo, CancellationToken.None);

            // Assert
            saveRes.IsSuccess.Should().BeTrue();
            var violacionesRes = await _sut.GetViolacionesByIndividuoAsync(saveRes.Value, CancellationToken.None);
            violacionesRes.IsSuccess.Should().BeTrue();
            violacionesRes.Value.Should().HaveCount(1);

            var violDb = violacionesRes.Value.First();
            violDb.IdGen.Should().BeNull();
            violDb.DetalleJson.Should().Be("{\"info\":\"sin genes\"}");
        }

        [Fact]
        public async Task SaveIndividuoAsync_WhenIndicesGenesContainsInvalidIndex_ShouldIgnoreInvalidAndMapValid()
        {
            // Arrange
            var corrida = new AgCorrida { IdPeriodo = 20261, IdPrograma = 10, IdAula = 1 };
            var corridaRes = await _sut.InsertCorridaAsync(corrida, CancellationToken.None);

            var gen1 = new AgGen { IdPeriodo = 20261, IdAsignatura = 101, IdPrograma = 10, IdSede = 1 };
            var violacion = new AgViolacion
            {
                Regla = AgReglaViolacion.R1_SuperposicionDocente,
                EsDura = true,
                Mensaje = "Índice fuera de rango",
                DetalleJson = "{\"indicesGenes\":[0, 99, -1]}"
            };

            var individuo = new AgIndividuo
            {
                IdCorrida = corridaRes.Value,
                Generacion = 1,
                Genes = new List<AgGen> { gen1 },
                Violaciones = new List<AgViolacion> { violacion }
            };

            // Act
            var saveRes = await _sut.SaveIndividuoAsync(individuo, CancellationToken.None);

            // Assert
            saveRes.IsSuccess.Should().BeTrue();
            var violacionesRes = await _sut.GetViolacionesByIndividuoAsync(saveRes.Value, CancellationToken.None);
            var violDb = violacionesRes.Value.First();
            violDb.IdGen.Should().Be(gen1.Id);
            violDb.DetalleJson.Should().Contain("\"idGenes\":[" + gen1.Id + "]");
        }

        public void Dispose()
        {
            try
            {
                SQLiteConnection.ClearAllPools();
                if (File.Exists(_dbPath))
                {
                    File.Delete(_dbPath);
                }
            }
            catch
            {
                // Limpieza silenciosa
            }
        }
    }
}
