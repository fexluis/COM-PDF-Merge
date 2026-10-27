using System;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AgAulaDimensionesSyncTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public AgAulaDimensionesSyncTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AgAulaDimensionesSyncTest_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
CREATE TABLE ag_aula (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aula_ref TEXT NOT NULL DEFAULT '',
    jornada TEXT NOT NULL DEFAULT '',
    horas_dia INTEGER DEFAULT 0,
    dias_ventana TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT ''
);
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES ('3|1|1|III|A|PRESENCIAL', 'MANANA', 4, 'MAR-VIE', 'CENTRAL');");
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
                // Limpieza defensiva
            }
        }

        private IDbConnection CreateConnection()
        {
            return new SQLiteConnection(_connectionString);
        }

        [Fact]
        public async Task SincronizarAsync_WhenValuesDiffer_ShouldUpdateRowAndReturnOne()
        {
            // Arrange
            var sut = new AgAulaDimensionesSync(CreateConnection);

            // Act
            int affected = await sut.SincronizarAsync("3|1|1|III|A|PRESENCIAL", "NOCHE", 4, "VIE-SAB", "CENTRAL", CancellationToken.None);

            // Assert
            affected.Should().Be(1);

            using (var conn = CreateConnection())
            {
                conn.Open();
                var row = conn.QueryFirstOrDefault<(string jornada, int horas_dia, string dias_ventana)>(
                    "SELECT jornada, horas_dia, dias_ventana FROM ag_aula WHERE aula_ref = '3|1|1|III|A|PRESENCIAL';");
                row.jornada.Should().Be("NOCHE");
                row.horas_dia.Should().Be(4);
                row.dias_ventana.Should().Be("VIE-SAB");
            }
        }

        [Fact]
        public async Task SincronizarAsync_WhenValuesAlreadyMatch_ShouldReturnZero()
        {
            // Arrange
            var sut = new AgAulaDimensionesSync(CreateConnection);

            // Act: Sincronizar con los mismos valores existentes
            int affected = await sut.SincronizarAsync("3|1|1|III|A|PRESENCIAL", "MANANA", 4, "MAR-VIE", "CENTRAL", CancellationToken.None);

            // Assert
            affected.Should().Be(0);
        }

        [Fact]
        public async Task SincronizarAsync_WhenSubunidadDiffers_ShouldNotUpdateAndReturnZero()
        {
            // Arrange
            var sut = new AgAulaDimensionesSync(CreateConnection);

            // Act: Otra subunidad distinta a CENTRAL
            int affected = await sut.SincronizarAsync("3|1|1|III|A|PRESENCIAL", "NOCHE", 8, "VIE-SAB", "SEDE_OTRA", CancellationToken.None);

            // Assert
            affected.Should().Be(0);
        }

        [Fact]
        public async Task SincronizarAsync_WhenAulaRefDoesNotExist_ShouldReturnZero()
        {
            // Arrange
            var sut = new AgAulaDimensionesSync(CreateConnection);

            // Act
            int affected = await sut.SincronizarAsync("INEXISTENTE|REF", "NOCHE", 8, "JUE-SAB", "CENTRAL", CancellationToken.None);

            // Assert
            affected.Should().Be(0);
        }

        [Fact]
        public async Task SincronizarAsync_WhenAulaRefIsEmpty_ShouldReturnZeroImmediately()
        {
            // Arrange
            var sut = new AgAulaDimensionesSync(CreateConnection);

            // Act
            int affected = await sut.SincronizarAsync(string.Empty, "NOCHE", 4, "VIE-SAB", "CENTRAL", CancellationToken.None);

            // Assert
            affected.Should().Be(0);
        }

        [Fact]
        public async Task SincronizarAsync_WhenConnectionFails_ShouldTolerateAndReturnZero()
        {
            // Arrange: Fábrica con cadena inválida
            var sut = new AgAulaDimensionesSync(() => new SQLiteConnection("Data Source=/invalid/path/test.db;FailIfMissing=True;"));

            // Act
            int affected = await sut.SincronizarAsync("3|1|1|III|A|PRESENCIAL", "NOCHE", 4, "VIE-SAB", "CENTRAL", CancellationToken.None);

            // Assert
            affected.Should().Be(0);
        }

        [Fact]
        public async Task SincronizarAsync_WhenSubunidadIsEmpty_ShouldNotSweepDifferentSubunidad()
        {
            // Arrange (Fase 1238: subunidad vacía no barre subunidad 'CENTRAL')
            var sut = new AgAulaDimensionesSync(CreateConnection);

            // Act
            int affected = await sut.SincronizarAsync("3|1|1|III|A|PRESENCIAL", "TARDE", 8, "LUN-JUE", string.Empty, CancellationToken.None);

            // Assert
            affected.Should().Be(0);
        }

        [Fact]
        public async Task SincronizarAsync_WhenHorasDiaIsNull_ShouldUpdateRowAndReturnOne()
        {
            // Arrange (Fase 1237: horas_dia NULL en BD)
            using (var conn = CreateConnection())
            {
                conn.Open();
                conn.Execute("INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad) VALUES ('REF|NULL|TEST', 'MANANA', NULL, 'MAR-VIE', 'CENTRAL');");
            }
            var sut = new AgAulaDimensionesSync(CreateConnection);

            // Act
            int affected = await sut.SincronizarAsync("REF|NULL|TEST", "MANANA", 4, "MAR-VIE", "CENTRAL", CancellationToken.None);

            // Assert
            affected.Should().Be(1);
        }
    }
}
