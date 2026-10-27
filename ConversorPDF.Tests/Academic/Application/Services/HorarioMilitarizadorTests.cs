using ConversorPDF.XLL.Academic.Application.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Application.Services
{
    public class HorarioMilitarizadorTests
    {
        [Theory]
        [InlineData("1:00 PM A 5:00 PM", "13:00 A 17:00")]
        [InlineData("8:00 AM A 12:00 M", "08:00 A 12:00")]
        [InlineData("6:00 PM A 10:00 PM", "18:00 A 22:00")]
        [InlineData("2:00 PM A 10:00 PM", "14:00 A 22:00")]
        [InlineData("8:00 AM A 5:00 PM", "08:00 A 17:00")]
        [InlineData("13:00 A 17:00", "13:00 A 17:00")]
        [InlineData("08:00 A 12:00", "08:00 A 12:00")]
        [InlineData("14:00 A 22:00", "14:00 A 22:00")]
        public void NormalizarHorario_WhenInputIsValid_ShouldReturnMilitarFormat(string input, string expected)
        {
            // Act
            string result = HorarioMilitarizador.NormalizarHorario(input);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        public void NormalizarHorario_WhenInputIsNullOrWhiteSpace_ShouldReturnEmptyString(string input)
        {
            // Act
            string result = HorarioMilitarizador.NormalizarHorario(input);

            // Assert
            result.Should().BeEmpty();
        }
    }
}
