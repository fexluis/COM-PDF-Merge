using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Core.Models.Academic;
using ConversorPDF.XLL.Core.Services.Academic;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Core.Services.Academic
{
    public class GeneradorSesionesRecurrentesTests
    {
        private readonly GeneradorSesionesRecurrentes _sut;

        public GeneradorSesionesRecurrentesTests()
        {
            _sut = new GeneradorSesionesRecurrentes();
        }

        [Fact]
        public void Generar_WhenViernesSabadoConsecutive_ShouldReturnFourSessionsOfSixHours()
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
            result.Should().HaveCount(4);
            result.Sum(s => s.Horas).Should().Be(24);
            
            // First session check
            result[0].fechasesionInicial.Should().Be(new DateTime(2026, 4, 17)); // First Friday
            result[0].fechasesionFinal.Should().Be(new DateTime(2026, 4, 18)); // Next Saturday
            result[0].Duracion.Should().Be(1.5);
            result[0].Horas.Should().Be(6);

            // Last session check
            result[3].fechasesionInicial.Should().Be(new DateTime(2026, 5, 8));
            result[3].fechasesionFinal.Should().Be(new DateTime(2026, 5, 9));
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
            result.Should().HaveCount(3); // 4 + 4 + 2
            result[0].Horas.Should().Be(4);
            result[1].Horas.Should().Be(4);
            result[2].Horas.Should().Be(2); // Adjusted
            result.Sum(s => s.Horas).Should().Be(10);
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
            result.Should().BeEmpty();
        }

        [Fact]
        public void Generar_WhenDaysAreNotConsecutive_ShouldThrowArgumentException()
        {
            // Arrange
            var diasConfig = new List<ConfiguracionDia>
            {
                new ConfiguracionDia { DiaSemana = DayOfWeek.Monday, Horas = 4 },
                new ConfiguracionDia { DiaSemana = DayOfWeek.Wednesday, Horas = 4 }
            };

            // Act
            Action act = () => _sut.Generar(1, 100, DateTime.Today, DateTime.Today.AddDays(30), diasConfig, 20, true);

            // Assert
            act.Should().Throw<ArgumentException>()
                .WithMessage("Los días configurados deben ser consecutivos para unir la sesión en una sola fila.");
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
            result.Should().BeEmpty();
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
            result.Should().BeEmpty();
        }
    }
}
