using System;
using ConversorPDF.XLL.Academic.Application.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class DiaSemanaNormalizadorTests
    {
        [Theory]
        [InlineData("LUNES", "LUNES")]
        [InlineData("lunes", "LUNES")]
        [InlineData("MARTES", "MARTES")]
        [InlineData("MIERCOLES", "MIERCOLES")]
        [InlineData("MIÉRCOLES", "MIERCOLES")]
        [InlineData("miércoles", "MIERCOLES")]
        [InlineData("MIE\u0301RCOLES", "MIERCOLES")] // Acento combinante Unicode FormD
        [InlineData("JUEVES", "JUEVES")]
        [InlineData("VIERNES", "VIERNES")]
        [InlineData("SABADO", "SABADO")]
        [InlineData("SÁBADO", "SABADO")]
        [InlineData("sábado", "SABADO")]
        [InlineData("SA\u0301BADO", "SABADO")] // Acento combinante Unicode FormD
        [InlineData("DOMINGO", "DOMINGO")]
        [InlineData("  sábado  ", "SABADO")]
        [InlineData("MIERCOLES\u00A0", "MIERCOLES")] // NBSP
        [InlineData("  SÁBADO\u00A0  ", "SABADO")] // NBSP + espacios
        public void NormalizarDiaSemana_WhenInputHasAccentsOrSpaces_ShouldReturnCanonicalUppercaseWithoutAccents(string input, string expected)
        {
            // Arrange & Act
            string result = DiaSemanaNormalizador.NormalizarDiaSemana(input);

            // Assert
            result.Should().Be(expected);
        }

        [Fact]
        public void NormalizarDiaSemana_WhenNullOrEmpty_ShouldReturnEmptyString()
        {
            // Arrange & Act
            string resultNull = DiaSemanaNormalizador.NormalizarDiaSemana(null);
            string resultEmpty = DiaSemanaNormalizador.NormalizarDiaSemana("   ");

            // Assert
            resultNull.Should().BeEmpty();
            resultEmpty.Should().BeEmpty();
        }

        [Theory]
        [InlineData(DayOfWeek.Monday, "LUNES")]
        [InlineData(DayOfWeek.Tuesday, "MARTES")]
        [InlineData(DayOfWeek.Wednesday, "MIERCOLES")]
        [InlineData(DayOfWeek.Thursday, "JUEVES")]
        [InlineData(DayOfWeek.Friday, "VIERNES")]
        [InlineData(DayOfWeek.Saturday, "SABADO")]
        [InlineData(DayOfWeek.Sunday, "DOMINGO")]
        public void DesdeDayOfWeek_WhenValidDayOfWeek_ShouldReturnCanonicalSpanishName(DayOfWeek day, string expected)
        {
            // Arrange & Act
            string result = DiaSemanaNormalizador.DesdeDayOfWeek(day);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData("Lunes", DayOfWeek.Monday)]
        [InlineData("MIÉRCOLES", DayOfWeek.Wednesday)]
        [InlineData("MIE\u0301RCOLES", DayOfWeek.Wednesday)]
        [InlineData("miercoles", DayOfWeek.Wednesday)]
        [InlineData("Sábado", DayOfWeek.Saturday)]
        [InlineData("SÁBADO\u00A0", DayOfWeek.Saturday)]
        [InlineData("sabado", DayOfWeek.Saturday)]
        [InlineData("Domingo", DayOfWeek.Sunday)]
        public void TryParse_WhenValidSpanishName_ShouldReturnTrueAndCorrectDayOfWeek(string input, DayOfWeek expected)
        {
            // Arrange & Act
            bool success = DiaSemanaNormalizador.TryParse(input, out var result);

            // Assert
            success.Should().BeTrue();
            result.Should().Be(expected);
        }

        [Fact]
        public void TryParse_WhenInvalidString_ShouldReturnFalse()
        {
            // Arrange & Act
            bool success = DiaSemanaNormalizador.TryParse("Invalido", out _);

            // Assert
            success.Should().BeFalse();
        }

        [Theory]
        [InlineData(1, DayOfWeek.Monday)]
        [InlineData(2, DayOfWeek.Tuesday)]
        [InlineData(3, DayOfWeek.Wednesday)]
        [InlineData(4, DayOfWeek.Thursday)]
        [InlineData(5, DayOfWeek.Friday)]
        [InlineData(6, DayOfWeek.Saturday)]
        [InlineData(7, DayOfWeek.Sunday)]
        public void NumeroDiaADayOfWeek_WhenValidNumber1To7_ShouldReturnExpectedDayOfWeek(int number, DayOfWeek expected)
        {
            // Arrange & Act
            var result = DiaSemanaNormalizador.NumeroDiaADayOfWeek(number);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(8)]
        [InlineData(-1)]
        public void NumeroDiaADayOfWeek_WhenOutOfRange_ShouldReturnNull(int number)
        {
            // Arrange & Act
            var result = DiaSemanaNormalizador.NumeroDiaADayOfWeek(number);

            // Assert
            result.Should().BeNull();
        }

        [Theory]
        [InlineData(DayOfWeek.Monday, 1)]
        [InlineData(DayOfWeek.Tuesday, 2)]
        [InlineData(DayOfWeek.Wednesday, 3)]
        [InlineData(DayOfWeek.Thursday, 4)]
        [InlineData(DayOfWeek.Friday, 5)]
        [InlineData(DayOfWeek.Saturday, 6)]
        [InlineData(DayOfWeek.Sunday, 7)]
        public void DayOfWeekANumero_WhenGivenDayOfWeek_ShouldReturnNumber1To7(DayOfWeek day, int expected)
        {
            // Arrange & Act
            int result = DiaSemanaNormalizador.DayOfWeekANumero(day);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData("Lunes", 1)]
        [InlineData("MARTES", 2)]
        [InlineData("Miércoles", 3)]
        [InlineData("Miercoles", 3)]
        [InlineData("Jueves", 4)]
        [InlineData("Viernes", 5)]
        [InlineData("Sábado", 6)]
        [InlineData("Sabado", 6)]
        [InlineData("Domingo", 7)]
        public void ObtenerNumeroDia_WhenValidSpanishName_ShouldReturnNumber1To7(string input, int expected)
        {
            // Arrange & Act
            int result = DiaSemanaNormalizador.ObtenerNumeroDia(input);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData(1, "LUNES")]
        [InlineData(3, "MIERCOLES")]
        [InlineData(6, "SABADO")]
        [InlineData(7, "DOMINGO")]
        [InlineData(0, "")]
        [InlineData(8, "")]
        public void ObtenerNombreDia_WhenGivenNumber_ShouldReturnCanonicalNameOrEmpty(int number, string expected)
        {
            // Arrange & Act
            string result = DiaSemanaNormalizador.ObtenerNombreDia(number);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData("MIÉRCOLES", "MIERCOLES", true)]
        [InlineData("MIE\u0301RCOLES", "MIERCOLES", true)]
        [InlineData("miércoles", "miercoles", true)]
        [InlineData("SÁBADO\u00A0", "sabado", true)]
        [InlineData("Lunes", "Martes", false)]
        public void SonIguales_WhenComparingDaysWithOrWithoutAccents_ShouldReturnExpectedEquality(string a, string b, bool expected)
        {
            // Arrange & Act
            bool result = DiaSemanaNormalizador.SonIguales(a, b);

            // Assert
            result.Should().Be(expected);
        }
    }
}
