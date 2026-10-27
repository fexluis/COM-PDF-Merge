using Xunit;
using FluentAssertions;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.Tests.Unit
{
    public class ViaticosHelperEsTerminalAereaTests
    {
        [Theory]
        [InlineData("ANTIOQUIA (Rionegro)", true)]
        [InlineData("OTRO", true)]
        [InlineData("BOGOTÁ (El Dorado)", true)]
        [InlineData("Arauca", true)]
        [InlineData("NO APLICA", false)]
        [InlineData("Terrestre", false)]
        [InlineData("TERRESTRE", false)]
        [InlineData("Aéreo", false)]
        [InlineData("Aereo", false)]
        [InlineData("", false)]
        [InlineData("   ", false)]
        [InlineData(null, false)]
        public void EsTerminalAerea_WhenGivenTerminalString_ShouldDistinguishRealTerminalsFromSynthetic(string input, bool expected)
        {
            // Act
            bool actual = ViaticosHelper.EsTerminalAerea(input);

            // Assert
            actual.Should().Be(expected);
        }
    }
}
