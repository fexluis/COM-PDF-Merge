using Xunit;
using FluentAssertions;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.Tests.Unit
{
    public class ViaticosHelperResolverTopeTerminalTests
    {
        [Fact]
        public void ResolverTopeTerminal_WhenMatchingById_ShouldReturnExactCatalogRecord()
        {
            // Act
            bool res = ViaticosHelper.ResolverTopeTerminal(2, "", "Test", out int idRes, out string nomRes, out decimal maxVal);

            // Assert
            res.Should().BeTrue();
            idRes.Should().Be(2);
            nomRes.Should().Be("ANTIOQUIA (Rionegro)");
            maxVal.Should().Be(162634m);
        }

        [Fact]
        public void ResolverTopeTerminal_WhenMatchingByExactNormalizedName_ShouldResolveCorrectly()
        {
            // Act
            bool res = ViaticosHelper.ResolverTopeTerminal(0, "Antioquia (Rionegro)", "Test", out int idRes, out string nomRes, out decimal maxVal);

            // Assert
            res.Should().BeTrue();
            idRes.Should().Be(2);
            nomRes.Should().Be("ANTIOQUIA (Rionegro)");
            maxVal.Should().Be(162634m);
        }

        [Fact]
        public void ResolverTopeTerminal_WhenMatchingByContains_ShouldResolveHighestMaxValor()
        {
            // Act
            bool res = ViaticosHelper.ResolverTopeTerminal(0, "Rionegro", "Test", out int idRes, out string nomRes, out decimal maxVal);

            // Assert
            res.Should().BeTrue();
            idRes.Should().Be(2);
            maxVal.Should().Be(162634m);
        }

        [Fact]
        public void ResolverTopeTerminal_WhenUnmatchedRealAirportName_ShouldFallbackToOtroWithPositiveAmount()
        {
            // Act
            bool res = ViaticosHelper.ResolverTopeTerminal(0, "AEROPUERTO INTERGALACTICO", "Test", out int idRes, out string nomRes, out decimal maxVal);

            // Assert
            res.Should().BeTrue();
            idRes.Should().Be(13);
            nomRes.Should().Be("OTRO");
            maxVal.Should().Be(50689m);
        }

        [Theory]
        [InlineData("NO APLICA")]
        [InlineData("Terrestre")]
        [InlineData("Aéreo")]
        [InlineData("")]
        [InlineData(null)]
        public void ResolverTopeTerminal_WhenGivenSyntheticOrNonAirportName_ShouldReturnFalse(string syntheticName)
        {
            // Act
            bool res = ViaticosHelper.ResolverTopeTerminal(0, syntheticName, "Test", out int idRes, out string nomRes, out decimal maxVal);

            // Assert
            res.Should().BeFalse();
            idRes.Should().Be(1);
            maxVal.Should().Be(0m);
        }
    }
}
