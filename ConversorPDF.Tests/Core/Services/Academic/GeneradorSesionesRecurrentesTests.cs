using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Application.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class GeneradorSesionesRecurrentesTests
    {
        private readonly GeneradorSesionesRecurrentes _sut;

        public GeneradorSesionesRecurrentesTests()
        {
            _sut = new GeneradorSesionesRecurrentes();
        }

        [Fact]
        public void Generar_WhenViernesSabadoConsecutive_ShouldReturnEightSessionsOfSixHoursEach()
        {
            // Arrange
            var diasConfig = new List<ConfiguracionDia>
            {
                new ConfiguracionDia { DiaSemana = DayOfWeek.Friday, Horas = 4, Nombre = "VIERNES" },
                new ConfiguracionDia { DiaSemana = DayOfWeek.Saturday, Horas = 2, Nombre = "SABADO" }
            };

            DateTime fechaInicioGlobal = new DateTime(2026, 4, 16); // Jueves
            DateTime fechaFinGlobal = new DateTime(2026, 5, 9); // Sábado de la 4ta semana
            double horasTotales = 24;

            // Act
            var result = _sut.Generar(1, 100, fechaInicioGlobal, fechaFinGlobal, diasConfig, horasTotales, true);

            // Assert
            result.IsSuccess.Should().BeTrue();
            // 4 semanas × 2 días por semana = 8 sesiones individuales
            result.Value.Should().HaveCount(8);
            result.Value.Sum(s => s.Horas).Should().Be(24);

            // Primera sesión: Viernes 17/04/2026
            result.Value[0].FechaSesion.Should().Be(new DateTime(2026, 4, 17));
            result.Value[0].HorarioSesion.Should().Be("VIERNES|4");
            result.Value[0].Horas.Should().Be(4);

            // Segunda sesión: Sábado 18/04/2026
            result.Value[1].FechaSesion.Should().Be(new DateTime(2026, 4, 18));
            result.Value[1].HorarioSesion.Should().Be("SABADO|2");
            result.Value[1].Horas.Should().Be(2);

            // Última sesión: Sábado 09/05/2026
            result.Value[7].FechaSesion.Should().Be(new DateTime(2026, 5, 9));
        }

        [Fact]
        public void Generar_WhenTotalHorasIsUneven_ShouldAdjustLastSession()
        {
            // Arrange
            var diasConfig = new List<ConfiguracionDia>
            {
                new ConfiguracionDia { DiaSemana = DayOfWeek.Friday, Horas = 4, Nombre = "VIERNES" }
            };

            DateTime fechaInicioGlobal = new DateTime(2026, 4, 16);
            DateTime fechaFinGlobal = new DateTime(2026, 5, 30);
            double horasTotales = 10;

            // Act
            var result = _sut.Generar(1, 100, fechaInicioGlobal, fechaFinGlobal, diasConfig, horasTotales, true);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().HaveCount(3); // 4 + 4 + 2
            result.Value[0].Horas.Should().Be(4);
            result.Value[1].Horas.Should().Be(4);
            result.Value[2].Horas.Should().Be(2); // Adjusted
            result.Value.Sum(s => s.Horas).Should().Be(10);
        }

        [Fact]
        public void Generar_WhenTotalHorasIsZero_ShouldReturnEmpty()
        {
            // Arrange
            var diasConfig = new List<ConfiguracionDia>
            {
                new ConfiguracionDia { DiaSemana = DayOfWeek.Friday, Horas = 4 }
            };

            // Act
            var result = _sut.Generar(1, 100, DateTime.Today, DateTime.Today.AddDays(10), diasConfig, 0, true);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().BeEmpty();
        }

        [Fact]
        public void Generar_WhenDaysAreNotConsecutive_ShouldReturnFailureResult()
        {
            // Arrange
            var diasConfig = new List<ConfiguracionDia>
            {
                new ConfiguracionDia { DiaSemana = DayOfWeek.Monday, Horas = 4 },
                new ConfiguracionDia { DiaSemana = DayOfWeek.Wednesday, Horas = 4 }
            };

            // Act
            var result = _sut.Generar(1, 100, DateTime.Today, DateTime.Today.AddDays(30), diasConfig, 20, true);

            // Assert
            result.IsFailure.Should().BeTrue();
            result.Error.Should().Be("Los días configurados deben ser consecutivos para unir la sesión en una sola fila.");
        }

        [Fact]
        public void Generar_WhenRangeIsTooShortForAnyOccurrence_ShouldReturnEmpty()
        {
            // Arrange
            var diasConfig = new List<ConfiguracionDia>
            {
                new ConfiguracionDia { DiaSemana = DayOfWeek.Monday, Horas = 4 }
            };

            // 16/04/2026 is Thursday. Monday is 20/04/2026.
            DateTime inicio = new DateTime(2026, 4, 16);
            DateTime fin = new DateTime(2026, 4, 18); // Range ends before Monday

            // Act
            var result = _sut.Generar(1, 100, inicio, fin, diasConfig, 10, true);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().BeEmpty();
        }

        [Fact]
        public void Generar_WhenStartDateIsAfterEndDate_ShouldReturnEmpty()
        {
            // Arrange
            var diasConfig = new List<ConfiguracionDia>
            {
                new ConfiguracionDia { DiaSemana = DayOfWeek.Monday, Horas = 4 }
            };

            // Act
            var result = _sut.Generar(1, 100, DateTime.Today.AddDays(10), DateTime.Today, diasConfig, 10, true);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().BeEmpty();
        }
    }
}
