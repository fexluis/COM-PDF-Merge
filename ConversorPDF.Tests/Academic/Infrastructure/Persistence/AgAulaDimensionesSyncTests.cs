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

        [Fact]
        public async Task SincronizarLoteAsync_WhenAllValid_UpdatesAtomicallyAndReturnsTotal()
        {
            // Arrange
            using (var conn = CreateConnection())
            {
                conn.Open();
                conn.Execute("INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad) VALUES ('LOTE|1', 'MANANA', 4, 'MAR-VIE', 'CENTRAL');");
                conn.Execute("INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad) VALUES ('LOTE|2', 'MANANA', 4, 'MAR-VIE', 'CENTRAL');");
            }
            var sut = new AgAulaDimensionesSync(CreateConnection);
            var diffs = new[]
            {
                ("LOTE|1", "TARDE", 8, "LUN-VIE", "CENTRAL"),
                ("LOTE|2", "NOCHE", 4, "1,2,3,4,5", "CENTRAL")
            };

            // Act
            var res = await sut.SincronizarLoteAsync(diffs, CancellationToken.None);

            // Assert
            res.IsSuccess.Should().BeTrue();
            res.Value.Should().Be(2);

            using (var conn = CreateConnection())
            {
                conn.Open();
                var r1 = conn.QueryFirstOrDefault<(string jornada, int horas_dia, string dias_ventana)>("SELECT jornada, horas_dia, dias_ventana FROM ag_aula WHERE aula_ref = 'LOTE|1';");
                r1.jornada.Should().Be("TARDE");
                r1.horas_dia.Should().Be(8);
                r1.dias_ventana.Should().Be("LUN-VIE");

                var r2 = conn.QueryFirstOrDefault<(string jornada, int horas_dia, string dias_ventana)>("SELECT jornada, horas_dia, dias_ventana FROM ag_aula WHERE aula_ref = 'LOTE|2';");
                r2.jornada.Should().Be("NOCHE");
                r2.horas_dia.Should().Be(4);
                r2.dias_ventana.Should().Be("1,2,3,4,5");
            }
        }

        [Fact]
        public async Task SincronizarLoteAsync_WhenAulaDoesNotExist_FailsAndRollsBack()
        {
            // Arrange
            using (var conn = CreateConnection())
            {
                conn.Open();
                conn.Execute("INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad) VALUES ('LOTE|EXISTE', 'MANANA', 4, 'MAR-VIE', 'CENTRAL');");
            }
            var sut = new AgAulaDimensionesSync(CreateConnection);
            var diffs = new[]
            {
                ("LOTE|EXISTE", "TARDE", 8, "LUN-VIE", "CENTRAL"),
                ("LOTE|NO_EXISTE", "NOCHE", 4, "1,2,3,4,5", "CENTRAL")
            };

            // Act
            var res = await sut.SincronizarLoteAsync(diffs, CancellationToken.None);

            // Assert
            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("LOTE|NO_EXISTE");

            using (var conn = CreateConnection())
            {
                conn.Open();
                var r = conn.QueryFirstOrDefault<(string jornada, int horas_dia, string dias_ventana)>("SELECT jornada, horas_dia, dias_ventana FROM ag_aula WHERE aula_ref = 'LOTE|EXISTE';");
                r.jornada.Should().Be("MANANA");
                r.horas_dia.Should().Be(4);
                r.dias_ventana.Should().Be("MAR-VIE");
            }
        }

        [Fact]
        public async Task SincronizarLoteAsync_WhenTriggerFails_RollsBackAllAndReturnsFailure()
        {
            // Arrange
            using (var conn = CreateConnection())
            {
                conn.Open();
                conn.Execute("INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad) VALUES ('LOTE|OK', 'MANANA', 4, 'MAR-VIE', 'CENTRAL');");
                conn.Execute("INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad) VALUES ('LOTE|FAIL', 'MANANA', 4, 'MAR-VIE', 'CENTRAL');");
                conn.Execute(@"
CREATE TRIGGER trg_fail_update BEFORE UPDATE ON ag_aula
WHEN NEW.aula_ref = 'LOTE|FAIL'
BEGIN
    SELECT RAISE(ABORT, 'fallo simulado en trigger');
END;");
            }
            var sut = new AgAulaDimensionesSync(CreateConnection);
            var diffs = new[]
            {
                ("LOTE|OK", "TARDE", 8, "LUN-VIE", "CENTRAL"),
                ("LOTE|FAIL", "NOCHE", 4, "1,2,3,4,5", "CENTRAL")
            };

            // Act
            var res = await sut.SincronizarLoteAsync(diffs, CancellationToken.None);

            // Assert
            res.IsFailure.Should().BeTrue();

            using (var conn = CreateConnection())
            {
                conn.Open();
                var r = conn.QueryFirstOrDefault<(string jornada, int horas_dia, string dias_ventana)>("SELECT jornada, horas_dia, dias_ventana FROM ag_aula WHERE aula_ref = 'LOTE|OK';");
                r.jornada.Should().Be("MANANA");
                r.horas_dia.Should().Be(4);
                r.dias_ventana.Should().Be("MAR-VIE");
            }
        }
    }
}
