using Xunit;
using FluentAssertions;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.Tests.Unit
{
    public class ViaticosHelperSanitizarCedulaTests
    {
        [Theory]
        [InlineData("1.018.412.345", "1018412345")]
        [InlineData(" 72.170.937 ", "72170937")]
        [InlineData("1018-412-345", "1018412345")]
        [InlineData("72170937", "72170937")]
        [InlineData("", "")]
        [InlineData(null, "")]
        [InlineData("   ", "")]
        [InlineData("ABC-123.456", "123456")]
        public void SanitizarCedula_WhenInputHasFormattingOrSpaces_ShouldReturnCleanDigitsOnly(string input, string expected)
        {
            // Act
            string actual = ViaticosHelper.SanitizarCedula(input);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public void SanitizarCedula_WhenComparingFormattedAndUnformattedCedulas_ShouldMatch()
        {
            // Arrange
            string cedulaFormatted = "72.170.937";
            string cedulaRaw = "72170937";

            // Act
            string clean1 = ViaticosHelper.SanitizarCedula(cedulaFormatted);
            string clean2 = ViaticosHelper.SanitizarCedula(cedulaRaw);

            // Assert
            clean1.Should().Be(clean2);
        }
    }
}
