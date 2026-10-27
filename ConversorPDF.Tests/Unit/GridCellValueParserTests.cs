using System;
using ConversorPDF.XLL.Academic.Application.Mappers;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Unit
{
    public class GridCellValueParserTests
    {
        [Fact]
        public void TryParseHoras_WhenComaDecimal_ShouldParseCorrectlyWithoutMultiplyingByTen()
        {
            // Arrange
            string input = "4,5";

            // Act
            bool success = GridCellValueParser.TryParseHoras(input, out double horas);

            // Assert
            success.Should().BeTrue();
            horas.Should().Be(4.5);
        }

        [Fact]
        public void TryParseHoras_WhenPuntoDecimal_ShouldParseCorrectly()
        {
            // Arrange
            string input = "4.5";

            // Act
            bool success = GridCellValueParser.TryParseHoras(input, out double horas);

            // Assert
            success.Should().BeTrue();
            horas.Should().Be(4.5);
        }

        [Fact]
        public void TryParseHoras_WhenEntero_ShouldParseCorrectly()
        {
            // Arrange
            string input = "8";

            // Act
            bool success = GridCellValueParser.TryParseHoras(input, out double horas);

            // Assert
            success.Should().BeTrue();
            horas.Should().Be(8.0);
        }

        [Fact]
        public void TryParseHoras_WhenCeroONegativo_ShouldReturnFalse()
        {
            // Arrange & Act & Assert
            GridCellValueParser.TryParseHoras("0", out double h0).Should().BeFalse();
            GridCellValueParser.TryParseHoras("-4", out double hNeg).Should().BeFalse();
            GridCellValueParser.TryParseHoras("", out double hEmpty).Should().BeFalse();
            GridCellValueParser.TryParseHoras(null, out double hNull).Should().BeFalse();
            GridCellValueParser.TryParseHoras("abc", out double hTxt).Should().BeFalse();
        }

        [Fact]
        public void TryParseHoras_WhenGreaterThan24_ShouldReturnFalse()
        {
            // Act & Assert
            GridCellValueParser.TryParseHoras("24.1", out double h1).Should().BeFalse();
            GridCellValueParser.TryParseHoras("400", out double h2).Should().BeFalse();
            GridCellValueParser.TryParseHoras(25.0, out double h3).Should().BeFalse();
        }

        [Fact]
        public void TryParseHoras_WhenNativeDouble_ShouldParse()
        {
            GridCellValueParser.TryParseHoras(4.0, out double h).Should().BeTrue();
            h.Should().Be(4.0);
        }

        [Theory]
        [InlineData("4.000", 4.0)]
        [InlineData("4.000,5", 0.0)]
        public void TryParseHoras_LimitesDeSeparadores_Documentados(string input, double esperado)
        {
            bool ok = GridCellValueParser.TryParseHoras(input, out double h);
            if (esperado == 0.0)
            {
                ok.Should().BeFalse();
            }
            else
            {
                ok.Should().BeTrue();
                h.Should().Be(esperado);
            }
        }

        [Theory]
        [InlineData("28/09/2026", 2026, 9, 28)]
        [InlineData("1/5/2026", 2026, 5, 1)]
        [InlineData("28-09-2026", 2026, 9, 28)]
        [InlineData("2026-09-28", 2026, 9, 28)]
        [InlineData("10/10/26", 2026, 10, 10)]
        [InlineData("15/10/30", 2030, 10, 15)]
        [InlineData("28/09/2026 0:00:00", 2026, 9, 28)]
        [InlineData("28/09/2026\u00A00:00:00", 2026, 9, 28)]
        [InlineData("28/09/2026\t12:00:00", 2026, 9, 28)]
        [InlineData("2026/09/28", 2026, 9, 28)]
        [InlineData("20260928", 2026, 9, 28)]
        public void TryParseFecha_WhenValidFormat_ShouldParseCorrectDateTime(string input, int year, int month, int day)
        {
            // Act
            bool success = GridCellValueParser.TryParseFecha(input, out DateTime fecha);

            // Assert
            success.Should().BeTrue();
            fecha.Should().Be(new DateTime(year, month, day));
        }

        [Fact]
        public void TryParseFecha_WhenDateTimeObject_ShouldReturnTrueAndDateOnly()
        {
            // Arrange
            var dt = new DateTime(2026, 10, 10, 14, 30, 0);

            // Act
            bool success = GridCellValueParser.TryParseFecha(dt, out DateTime fecha);

            // Assert
            success.Should().BeTrue();
            fecha.Should().Be(new DateTime(2026, 10, 10));
        }

        [Theory]
        [InlineData("31/02/2026")]
        [InlineData("invalid-date")]
        [InlineData("")]
        [InlineData(null)]
        public void TryParseFecha_WhenInvalidFormatOrDate_ShouldReturnFalse(string input)
        {
            // Act
            bool success = GridCellValueParser.TryParseFecha(input, out DateTime fecha);

            // Assert
            success.Should().BeFalse();
            fecha.Should().Be(default(DateTime));
        }
    }
}
