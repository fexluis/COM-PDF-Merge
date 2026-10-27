using System;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class PeriodoTests
    {
        [Fact]
        public void Create_WhenPeriodoNombreIsEmpty_ShouldReturnFailure()
        {
            // Act
            var result = Periodo.Create(1, "");

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("vacío");
        }

        [Fact]
        public void Create_WhenPeriodoNombreInvalidFormat_ShouldReturnFailure()
        {
            // Act
            var result = Periodo.Create(1, "2026-I");

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("YYYY-MM");
        }

        [Fact]
        public void Create_WhenValidPeriodo_ShouldReturnSuccess()
        {
            // Act
            var result = Periodo.Create(1, "2026-01", true);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.PeriodoNombre.Should().Be("2026-01");
            result.Value.Activo.Should().BeTrue();
        }

        [Fact]
        public void Create_WhenValidDatesProvided_ShouldParseDatesSuccessfully()
        {
            // Act
            var result = Periodo.Create(1, "2026-01", true, "2026-01-12", "2026-06-19");

            // Assert
            result.IsSuccess.Should().BeTrue();
            bool tieneVentana = result.Value.TryObtenerVentana(out DateTime ini, out DateTime fin);
            tieneVentana.Should().BeTrue();
            ini.Should().Be(new DateTime(2026, 1, 12));
            fin.Should().Be(new DateTime(2026, 6, 19));
        }

        [Fact]
        public void Create_WhenFechaFinIsBeforeFechaInicio_ShouldReturnFailure()
        {
            // Act
            var result = Periodo.Create(1, "2026-01", true, "2026-06-19", "2026-01-12");

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Contain("inválido");
        }
    }
}
