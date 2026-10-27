using System;
using ConversorPDF.XLL.Academic.Application.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class HorarioSesionParserTests
    {
        [Fact]
        public void TryParse_WhenValidFormatWithPM_ShouldReturnSuccessWithCorrectTimeSpan()
        {
            // Arrange
            string text = "6:00 PM A 10:00 PM";

            // Act
            var result = HorarioSesionParser.TryParse(text);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Inicio.Should().Be(new TimeSpan(18, 0, 0));
            result.Value.Fin.Should().Be(new TimeSpan(22, 0, 0));
        }

        [Fact]
        public void TryParse_WhenValidFormatWithAM_ShouldReturnSuccessWithCorrectTimeSpan()
        {
            // Arrange
            string text = "8:00 AM A 12:00 PM";

            // Act
            var result = HorarioSesionParser.TryParse(text);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Inicio.Should().Be(new TimeSpan(8, 0, 0));
            result.Value.Fin.Should().Be(new TimeSpan(12, 0, 0));
        }

        [Fact]
        public void TryParse_WhenFormatIncludes12M_ShouldParseAs12PM()
        {
            // Arrange
            string text = "8:00 AM A 12:00 M";

            // Act
            var result = HorarioSesionParser.TryParse(text);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Inicio.Should().Be(new TimeSpan(8, 0, 0));
            result.Value.Fin.Should().Be(new TimeSpan(12, 0, 0));
        }

        [Fact]
        public void TryParse_WhenFormatIncludes12MSinCeros_ShouldParseAs12PM()
        {
            // Arrange
            string text = "8:00 AM A 12 M";

            // Act
            var result = HorarioSesionParser.TryParse(text);

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Inicio.Should().Be(new TimeSpan(8, 0, 0));
            result.Value.Fin.Should().Be(new TimeSpan(12, 0, 0));
        }

        [Fact]
        public void TryParse_WhenFormatIsInvalid_ShouldReturnFailure()
        {
            // Arrange
            string text = "Cualquier texto que no es hora";

            // Act
            var result = HorarioSesionParser.TryParse(text);

            // Assert
            result.IsFailure.Should().BeTrue();
        }

        [Fact]
        public void TryParse_WhenEmpty_ShouldReturnFailure()
        {
            // Arrange
            string text = "";

            // Act
            var result = HorarioSesionParser.TryParse(text);

            // Assert
            result.IsFailure.Should().BeTrue();
        }
    }
}
