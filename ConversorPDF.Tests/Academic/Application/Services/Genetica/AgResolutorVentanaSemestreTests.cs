using System;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgResolutorVentanaSemestreTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public AgResolutorVentanaSemestreTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AgResolutorTest_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo TEXT,
                        periodo_inicio TEXT,
                        periodo_fin TEXT,
                        activo INTEGER DEFAULT 1
                    );

                    CREATE TABLE festivos (
                        idfestivo INTEGER PRIMARY KEY AUTOINCREMENT,
                        idperiodo INTEGER,
                        tipofestivo TEXT,
                        fechafestivoinicio TEXT,
                        fechafestivofinal TEXT
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
        public async Task ResolverVentanaAsync_WhenPeriodosHasDates_ShouldReturnPeriodosOriginWithDates()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                    VALUES (1, '2026-01', '2026-01-12', '2026-06-19', 1);

                    INSERT INTO festivos (idperiodo, tipofestivo, fechafestivoinicio, fechafestivofinal)
                    VALUES (1, 'PRESENCIAL', '2026-02-01', '2026-05-30');
                ");
            }

            var sut = new AgResolutorVentanaSemestre(() => new SQLiteConnection(_connectionString));

            // Act
            var result = await sut.ResolverVentanaAsync(1, null, null, CancellationToken.None);

            // Assert - Prioridad 1: PERIODOS tiene precedencia sobre FESTIVOS
            result.IsSuccess.Should().BeTrue();
            result.Value.Origen.Should().Be(AgOrigenVentana.Periodos);
            result.Value.FechaInicio.Should().Be(new DateTime(2026, 1, 12));
            result.Value.FechaFin.Should().Be(new DateTime(2026, 6, 19));
            result.Value.IdPeriodo.Should().Be(1);
        }

        [Fact]
        public async Task ResolverVentanaAsync_WhenPeriodosHasDatesAndManualDatesProvided_ShouldReturnManualOrigin()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                    VALUES (1, '2026-01', '2026-01-12', '2026-06-19', 1);
                ");
            }

            var sut = new AgResolutorVentanaSemestre(() => new SQLiteConnection(_connectionString));
            var manualIni = new DateTime(2026, 2, 1);
            var manualFin = new DateTime(2026, 5, 15);

            // Act
            var result = await sut.ResolverVentanaAsync(1, manualIni, manualFin, CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Origen.Should().Be(AgOrigenVentana.Manual);
            result.Value.FechaInicio.Should().Be(manualIni);
            result.Value.FechaFin.Should().Be(manualFin);
        }

        [Fact]
        public async Task ResolverVentanaAsync_WhenPeriodosHasNoDatesButFestivosHasPresenciales_ShouldReturnPresencialOrigin()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                    VALUES (2, '2026-02', NULL, NULL, 1);

                    INSERT INTO festivos (idperiodo, tipofestivo, fechafestivoinicio, fechafestivofinal)
                    VALUES (2, 'PRESENCIAL', '2026-08-15', '2026-12-12');
                ");
            }

            var sut = new AgResolutorVentanaSemestre(() => new SQLiteConnection(_connectionString));

            // Act
            var result = await sut.ResolverVentanaAsync(2, null, null, CancellationToken.None);

            // Assert - Prioridad 2: FESTIVOS PRESENCIAL
            result.IsSuccess.Should().BeTrue();
            result.Value.Origen.Should().Be(AgOrigenVentana.Presencial);
            result.Value.FechaInicio.Should().Be(new DateTime(2026, 8, 15));
            result.Value.FechaFin.Should().Be(new DateTime(2026, 12, 12));
        }

        [Fact]
        public async Task ResolverVentanaAsync_WhenNoConfiguredDatesButManualProvided_ShouldReturnManualOrigin()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                    VALUES (3, '2027-01', '', '', 1);
                ");
            }

            var sut = new AgResolutorVentanaSemestre(() => new SQLiteConnection(_connectionString));
            var manualIni = new DateTime(2027, 1, 15);
            var manualFin = new DateTime(2027, 6, 20);

            // Act
            var result = await sut.ResolverVentanaAsync(3, manualIni, manualFin, CancellationToken.None);

            // Assert - Prioridad 3: MANUAL
            result.IsSuccess.Should().BeTrue();
            result.Value.Origen.Should().Be(AgOrigenVentana.Manual);
            result.Value.FechaInicio.Should().Be(manualIni);
            result.Value.FechaFin.Should().Be(manualFin);
        }

        [Fact]
        public async Task ResolverVentanaAsync_WhenNoDatesConfiguredAndNoManualProvided_ShouldReturnFailure()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                    VALUES (4, '2027-02', NULL, NULL, 1);
                ");
            }

            var sut = new AgResolutorVentanaSemestre(() => new SQLiteConnection(_connectionString));

            // Act
            var result = await sut.ResolverVentanaAsync(4, null, null, CancellationToken.None);

            // Assert - Prioridad 4: FALLO
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("No se pudo determinar la ventana");
        }

        [Fact]
        public async Task ResolverVentanaAsync_WhenPeriodoNotFound_ShouldReturnFailure()
        {
            // Arrange
            var sut = new AgResolutorVentanaSemestre(() => new SQLiteConnection(_connectionString));

            // Act
            var result = await sut.ResolverVentanaAsync(999, null, null, CancellationToken.None);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("No se encontró ningún periodo");
        }

        [Theory]
        [InlineData("2027-01-01", "2027-05-30")] // inicia antes del período
        [InlineData("2027-02-05", "2027-07-01")] // termina después del período
        [InlineData("2027-01-01", "2027-07-01")] // ambos extremos fuera
        public async Task ResolverVentanaAsync_WhenManualExcedeElPeriodo_ShouldFail(string ini, string fin)
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                    VALUES (5, '2027-01', '2027-02-05', '2027-05-30', 1);
                ");
            }

            var sut = new AgResolutorVentanaSemestre(() => new SQLiteConnection(_connectionString));

            // Act
            var result = await sut.ResolverVentanaAsync(5,
                DateTime.ParseExact(ini, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture),
                DateTime.ParseExact(fin, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture),
                CancellationToken.None);

            // Assert - el período es el límite real, no una fecha fija ni una ventana ampliada
            result.IsFailure.Should().BeTrue();
            result.Error.Should().StartWith("VENTANA_FUERA_DE_PERIODO:");
            result.Error.Should().Contain("periodo=2027-02-05..2027-05-30");
        }

        [Fact]
        public async Task ResolverVentanaAsync_WhenManualCoincideConElPeriodo_ShouldReturn115Dias()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                    VALUES (6, '2027-01', '2027-02-05', '2027-05-30', 1);
                ");
            }

            var sut = new AgResolutorVentanaSemestre(() => new SQLiteConnection(_connectionString));

            // Act
            var result = await sut.ResolverVentanaAsync(6, new DateTime(2027, 2, 5), new DateTime(2027, 5, 30),
                CancellationToken.None);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Origen.Should().Be(AgOrigenVentana.Manual);
            result.Value.FechaInicio.Should().Be(new DateTime(2027, 2, 5));
            result.Value.FechaFin.Should().Be(new DateTime(2027, 5, 30));
            result.Value.DuracionDias.Should().Be(115);
        }

        [Fact]
        public async Task ResolverVentanaAsync_WhenPeriodoFinExtendido_ShouldPermitirVentanaManualLargaSinError()
        {
            // Arrange (Periodo 7 extendido a 2027-06-13)
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                    VALUES (7, '2027-01-EXT', '2027-02-05', '2027-06-13', 1);
                ");
            }

            var sut = new AgResolutorVentanaSemestre(() => new SQLiteConnection(_connectionString));

            // Act: ventana manual extendida a 2027-06-13
            var result = await sut.ResolverVentanaAsync(7, new DateTime(2027, 2, 5), new DateTime(2027, 6, 13),
                CancellationToken.None);

            // Assert: ya no devuelve VENTANA_FUERA_DE_PERIODO, sino éxito con ventana manual
            result.IsSuccess.Should().BeTrue();
            result.Value.Origen.Should().Be(AgOrigenVentana.Manual);
            result.Value.FechaInicio.Should().Be(new DateTime(2027, 2, 5));
            result.Value.FechaFin.Should().Be(new DateTime(2027, 6, 13));
            result.Value.DuracionDias.Should().Be(129);
        }
    }
}
