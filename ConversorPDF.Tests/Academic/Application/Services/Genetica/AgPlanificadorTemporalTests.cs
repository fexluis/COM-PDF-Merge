using System;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services.Genetica
{
    public class AgPlanificadorTemporalTests
    {
        [Fact]
        public void GenerarFechasValidas_WhenNoFestivos_ShouldProjectExactWeeklyDatesUntilHoursFulfilled()
        {
            // Arrange
            var sut = new AgPlanificadorTemporal();
            var fIni = new DateTime(2026, 2, 2); // Lunes
            var fFin = new DateTime(2026, 2, 28);
            double duracionBloque = 4.0;
            double totalHoras = 16.0; // 4 sesiones de 4 horas

            // Act
            var fechas = sut.GenerarFechasValidas(DayOfWeek.Monday, fIni, fFin, duracionBloque, totalHoras, null);

            // Assert
            fechas.Should().HaveCount(4);
            fechas.Should().ContainInOrder(
                new DateTime(2026, 2, 2),
                new DateTime(2026, 2, 9),
                new DateTime(2026, 2, 16),
                new DateTime(2026, 2, 23));
        }

        [Fact]
        public void GenerarFechasValidas_WhenFestivoCoincides_ShouldSkipFestivoAndAdvanceNextWeek()
        {
            // Arrange
            var sut = new AgPlanificadorTemporal();
            var fIni = new DateTime(2026, 2, 2); // Lunes
            var fFin = new DateTime(2026, 3, 10);
            double duracionBloque = 4.0;
            double totalHoras = 8.0; // 2 sesiones requeridas

            // Festivo el 9 de Febrero
            Func<DateTime, bool> esFestivo = d => d.Date == new DateTime(2026, 2, 9);

            // Act
            var fechas = sut.GenerarFechasValidas(DayOfWeek.Monday, fIni, fFin, duracionBloque, totalHoras, esFestivo);

            // Assert
            fechas.Should().HaveCount(2);
            fechas.Should().ContainInOrder(
                new DateTime(2026, 2, 2),
                new DateTime(2026, 2, 16));
        }
    }
}
