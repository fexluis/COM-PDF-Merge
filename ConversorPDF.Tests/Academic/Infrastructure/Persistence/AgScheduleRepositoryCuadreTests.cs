using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AgScheduleRepositoryCuadreTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;
        private readonly AgScheduleRepository _sut;

        public AgScheduleRepositoryCuadreTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AgScheduleRepoCuadre_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE ag_vinculaciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        id_corrida INTEGER NOT NULL DEFAULT 0,
                        idprograma INTEGER NOT NULL DEFAULT 0,
                        total_horas INTEGER NOT NULL DEFAULT 0
                    );

                    CREATE TABLE ag_sesiones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idvinculacion INTEGER NOT NULL,
                        horas REAL NOT NULL DEFAULT 0.0,
                        es_vacante INTEGER NOT NULL DEFAULT 0
                    );
                ");
            }

            _sut = new AgScheduleRepository(() => new SQLiteConnection(_connectionString));
        }

        public void Dispose()
        {
            try
            {
                if (File.Exists(_dbPath))
                    File.Delete(_dbPath);
            }
            catch { }
        }

        [Fact]
        public async Task GetCuadreCorridaAsync_WhenCorridaTieneDosProgramasYVacantes_ShouldCalcularTotalesYPorPrograma()
        {
            // Arrange
            long idCorrida = 104;
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                // Programa 1: 2 vinculaciones, 4 sesiones (16h), 1 vacante
                conn.Execute("INSERT INTO ag_vinculaciones (id, id_corrida, idprograma, total_horas) VALUES (1, @idCorrida, 1, 8);", new { idCorrida });
                conn.Execute("INSERT INTO ag_vinculaciones (id, id_corrida, idprograma, total_horas) VALUES (2, @idCorrida, 1, 8);", new { idCorrida });

                conn.Execute("INSERT INTO ag_sesiones (idvinculacion, horas, es_vacante) VALUES (1, 4.0, 0);");
                conn.Execute("INSERT INTO ag_sesiones (idvinculacion, horas, es_vacante) VALUES (1, 4.0, 0);");
                conn.Execute("INSERT INTO ag_sesiones (idvinculacion, horas, es_vacante) VALUES (2, 4.0, 1);");
                conn.Execute("INSERT INTO ag_sesiones (idvinculacion, horas, es_vacante) VALUES (2, 4.0, 0);");

                // Programa 2: 1 vinculación, 2 sesiones (8h), 1 vacante
                conn.Execute("INSERT INTO ag_vinculaciones (id, id_corrida, idprograma, total_horas) VALUES (3, @idCorrida, 2, 8);", new { idCorrida });
                conn.Execute("INSERT INTO ag_sesiones (idvinculacion, horas, es_vacante) VALUES (3, 4.0, 1);");
                conn.Execute("INSERT INTO ag_sesiones (idvinculacion, horas, es_vacante) VALUES (3, 4.0, 0);");
            }

            // Act
            var res = await _sut.GetCuadreCorridaAsync(idCorrida, CancellationToken.None);

            // Assert
            res.IsSuccess.Should().BeTrue();
            var cuadre = res.Value;
            cuadre.Sesiones.Should().Be(6);
            cuadre.Horas.Should().Be(24.0);
            cuadre.SesionesVacantes.Should().Be(2);
            cuadre.RequiereCompleto.Should().BeTrue();

            cuadre.PorPrograma.Should().HaveCount(2);
            var p1 = cuadre.PorPrograma.First(p => p.IdPrograma == 1);
            p1.Sesiones.Should().Be(4);
            p1.Horas.Should().Be(16.0);

            var p2 = cuadre.PorPrograma.First(p => p.IdPrograma == 2);
            p2.Sesiones.Should().Be(2);
            p2.Horas.Should().Be(8.0);
        }
    }
}
