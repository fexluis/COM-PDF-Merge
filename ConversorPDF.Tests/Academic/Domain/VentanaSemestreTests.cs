using System;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class VentanaSemestreTests
    {
        [Fact]
        public void Crear_WhenFechaFinIsBeforeFechaInicio_ShouldReturnFailure()
        {
            // Arrange
            var fIni = new DateTime(2026, 6, 1);
            var fFin = new DateTime(2026, 5, 1);

            // Act
            var result = VentanaSemestre.Crear(1, "2026-1", fIni, fFin, AgOrigenVentana.Presencial);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("inválida");
        }

        [Fact]
        public void Contiene_WhenFechaIsInsideWindow_ShouldReturnTrue()
        {
            // Arrange
            var sut = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            // Act
            bool result = sut.Contiene(new DateTime(2026, 3, 15));

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public void Contiene_WhenFechaIsOutsideWindow_ShouldReturnFalse()
        {
            // Arrange
            var sut = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 6, 30), AgOrigenVentana.Presencial).Value;

            // Act
            bool result = sut.Contiene(new DateTime(2026, 7, 1));

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public void DuracionDias_WhenValidDates_ShouldReturnExactDaysCount()
        {
            // Arrange
            var sut = VentanaSemestre.Crear(1, "2026-1", new DateTime(2026, 2, 1), new DateTime(2026, 2, 10), AgOrigenVentana.Manual).Value;

            // Act
            int duracion = sut.DuracionDias;

            // Assert
            duracion.Should().Be(10);
        }
    }
}
