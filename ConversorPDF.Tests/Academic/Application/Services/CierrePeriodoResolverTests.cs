using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class CierrePeriodoResolverTests
    {
        private readonly CierrePeriodoResolver _sut;

        public CierrePeriodoResolverTests()
        {
            _sut = new CierrePeriodoResolver();
        }

        [Fact]
        public void Evaluar_WhenConVentanaYFechaDentro_ShouldReturnExcedeLimiteFalse()
        {
            // Arrange (Regresión del defecto reportado: período 2027-01 con fecha 02/02/2027)
            var periodo = new Periodo
            {
                IdPeriodo = 3,
                PeriodoNombre = "2027-01",
                PeriodoInicio = "2027-01-01",
                PeriodoFin = "2027-06-30",
                Activo = true
            };
            var fechaFinal = new DateTime(2027, 2, 2);
            var hoy = new DateTime(2026, 11, 19);

            // Act
            var result = _sut.Evaluar(periodo, fechaFinal, hoy);

            // Assert
            result.ExcedeLimite.Should().BeFalse();
            result.Limite.Should().Be(new DateTime(2027, 6, 30));
            result.LimiteEsVentanaPeriodo.Should().BeTrue();
            result.Mensaje.Should().BeEmpty();
        }

        [Fact]
        public void Evaluar_WhenConVentanaYFechaFuera_ShouldExcederConFinPeriodo()
        {
            // Arrange
            var periodo = new Periodo
            {
                IdPeriodo = 3,
                PeriodoNombre = "2027-01",
                PeriodoInicio = "2027-01-01",
                PeriodoFin = "2027-06-30",
                Activo = true
            };
            var fechaFinal = new DateTime(2027, 7, 1);
            var hoy = new DateTime(2026, 11, 19);

            // Act
            var result = _sut.Evaluar(periodo, fechaFinal, hoy);

            // Assert
            result.ExcedeLimite.Should().BeTrue();
            result.Limite.Should().Be(new DateTime(2027, 6, 30));
            result.LimiteEsVentanaPeriodo.Should().BeTrue();
            result.Mensaje.Should().Contain("2027-01");
            result.Mensaje.Should().Contain("30/06/2027");
        }

        [Fact]
        public void Evaluar_WhenSinVentanaPeriodo202602YFecha20Diciembre_ShouldExceder15Diciembre2026()
        {
            // Arrange
            var periodo = new Periodo
            {
                IdPeriodo = 2,
                PeriodoNombre = "2026-02",
                PeriodoInicio = null,
                PeriodoFin = null,
                Activo = true
            };
            var fechaFinal = new DateTime(2026, 12, 20);
            var hoy = new DateTime(2026, 11, 19);

            // Act
            var result = _sut.Evaluar(periodo, fechaFinal, hoy);

            // Assert
            result.ExcedeLimite.Should().BeTrue();
            result.Limite.Should().Be(new DateTime(2026, 12, 15));
            result.LimiteEsVentanaPeriodo.Should().BeFalse();
            result.Mensaje.Should().Contain("2026");
        }

        [Fact]
        public void Evaluar_WhenSinVentanaPeriodo202701YFecha20Diciembre_ShouldExceder15Diciembre2027()
        {
            // Arrange
            var periodo = new Periodo
            {
                IdPeriodo = 3,
                PeriodoNombre = "2027-01",
                PeriodoInicio = null,
                PeriodoFin = null,
                Activo = true
            };
            var fechaFinal = new DateTime(2027, 12, 20);
            var hoy = new DateTime(2026, 11, 19);

            // Act
            var result = _sut.Evaluar(periodo, fechaFinal, hoy);

            // Assert
            result.ExcedeLimite.Should().BeTrue();
            result.Limite.Should().Be(new DateTime(2027, 12, 15));
            result.LimiteEsVentanaPeriodo.Should().BeFalse();
            result.Mensaje.Should().Contain("2027");
        }

        [Fact]
        public void Evaluar_WhenPeriodoNulo_ShouldUsarAnioDeHoy()
        {
            // Arrange
            Periodo periodo = null;
            var fechaFinal = new DateTime(2028, 12, 20);
            var hoy = new DateTime(2028, 3, 1);

            // Act
            var result = _sut.Evaluar(periodo, fechaFinal, hoy);

            // Assert
            result.ExcedeLimite.Should().BeTrue();
            result.Limite.Should().Be(new DateTime(2028, 12, 15));
            result.LimiteEsVentanaPeriodo.Should().BeFalse();
            result.Mensaje.Should().Contain("2028");
        }

        [Fact]
        public void Evaluar_WhenNombreInvalido_ShouldUsarAnioDePeriodoFin()
        {
            // Arrange
            var periodo = new Periodo
            {
                IdPeriodo = 9,
                PeriodoNombre = "SEM-2029",
                PeriodoInicio = null,
                PeriodoFin = "2029-05-30",
                Activo = true
            };
            var fechaFinal = new DateTime(2029, 12, 20);
            var hoy = new DateTime(2026, 11, 19);

            // Act
            var result = _sut.Evaluar(periodo, fechaFinal, hoy);

            // Assert
            result.ExcedeLimite.Should().BeTrue();
            result.Limite.Should().Be(new DateTime(2029, 12, 15));
            result.LimiteEsVentanaPeriodo.Should().BeFalse();
            result.Mensaje.Should().Contain("2029");
        }

        [Fact]
        public void Evaluar_WhenExcedeConVentana_ShouldIncluirPeriodoYLimiteEnMensaje()
        {
            // Arrange
            var periodo = new Periodo
            {
                IdPeriodo = 3,
                PeriodoNombre = "2027-01",
                PeriodoInicio = "2027-01-01",
                PeriodoFin = "2027-06-30",
                Activo = true
            };
            var fechaFinal = new DateTime(2027, 7, 5);
            var hoy = new DateTime(2026, 11, 19);

            // Act
            var result = _sut.Evaluar(periodo, fechaFinal, hoy);

            // Assert
            result.ExcedeLimite.Should().BeTrue();
            result.Mensaje.Should().Contain("2027-01");
            result.Mensaje.Should().Contain("30/06/2027");
            result.Mensaje.Should().Contain("05/07/2027");
        }

        [Fact]
        public void Evaluar_WhenPeriodoNombreAnioCero_ShouldUsarAnioDeHoy()
        {
            // Arrange
            var periodo = new Periodo { PeriodoNombre = "0000-01" };
            var fechaFinal = new DateTime(2027, 12, 20);
            var hoy = new DateTime(2026, 11, 19);

            // Act
            var result = _sut.Evaluar(periodo, fechaFinal, hoy);

            // Assert
            result.Limite.Should().Be(new DateTime(2026, 12, 15));
        }
    }
}