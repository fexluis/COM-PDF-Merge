using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgDiasVentanaTests
    {
        [Theory]
        [InlineData("LUN-JUE", "1,2,3,4", new[] { 1, 2, 3, 4 })]
        [InlineData("MAR-VIE", "2,3,4,5", new[] { 2, 3, 4, 5 })]
        [InlineData("JUE-SAB", "4,5,6", new[] { 4, 5, 6 })]
        [InlineData("VIE-SAB", "5,6", new[] { 5, 6 })]
        [InlineData("VIE-DOM", "5,6,7", new[] { 5, 6, 7 })]
        [InlineData("LUN-VIE", "1,2,3,4,5", new[] { 1, 2, 3, 4, 5 })]
        [InlineData("JUE-DOM", "4,5,6,7", new[] { 4, 5, 6, 7 })]
        [InlineData("MIE-DOM", "3,4,5,6,7", new[] { 3, 4, 5, 6, 7 })]
        public void AIsoAndDias_WhenVentanaCanonica_ShouldReturnExpectedIsoAndDays(
            string ventana, string expectedIso, int[] expectedDays)
        {
            // Act
            string iso = AgDiasVentana.AIso(ventana);
            IReadOnlyList<int> dias = AgDiasVentana.Dias(ventana);

            // Assert
            iso.Should().Be(expectedIso);
            dias.Should().Equal(expectedDays);
        }

        [Theory]
        [InlineData("LUN-MAR", "1,2", new[] { 1, 2 })]
        [InlineData("SAB-DOM", "6,7", new[] { 6, 7 })]
        public void AIsoAndDias_WhenRangoGenericoValido_ShouldReturnExpectedIsoAndDays(
            string rango, string expectedIso, int[] expectedDays)
        {
            // Act
            string iso = AgDiasVentana.AIso(rango);
            IReadOnlyList<int> dias = AgDiasVentana.Dias(rango);

            // Assert
            iso.Should().Be(expectedIso);
            dias.Should().Equal(expectedDays);
        }

        [Theory]
        [InlineData("VIE-LUN")]
        [InlineData("DOM-LUN")]
        [InlineData("SAB-VIE")]
        public void AIsoAndDias_WhenRangoEsCircularOInvertido_ShouldReturnEmpty(string rangoCircular)
        {
            // Act
            string iso = AgDiasVentana.AIso(rangoCircular);
            IReadOnlyList<int> dias = AgDiasVentana.Dias(rangoCircular);

            // Assert
            iso.Should().BeEmpty();
            dias.Should().BeEmpty();
        }

        [Theory]
        [InlineData("6M")]
        [InlineData("6T")]
        public void AIsoAndDias_WhenTokenEsFranjaRechazadaR1_ShouldReturnEmpty(string tokenFranja)
        {
            // Act
            string iso = AgDiasVentana.AIso(tokenFranja);
            IReadOnlyList<int> dias = AgDiasVentana.Dias(tokenFranja);

            // Assert
            iso.Should().BeEmpty();
            dias.Should().BeEmpty();
        }

        [Theory]
        [InlineData("LUN-VIE", "1,2,3,4,5")]
        [InlineData("JUE-DOM", "4,5,6,7")]
        [InlineData("MIE-DOM", "3,4,5,6,7")]
        public void Opciones_WhenConsultadas_ShouldContainNuevasEtiquetasCanonicas(string ventana, string expectedIso)
        {
            // Act & Assert
            AgDiasVentana.Opciones.Should().ContainKey(ventana);
            AgDiasVentana.Opciones[ventana].Should().Be(expectedIso);
        }

        [Theory]
        [InlineData("LUN-JUE", "1,2,3,4")]
        [InlineData("MAR-VIE", "2,3,4,5")]
        [InlineData("JUE-SAB", "4,5,6")]
        [InlineData("VIE-SAB", "5,6")]
        [InlineData("VIE-DOM", "5,6,7")]
        public void Opciones_WhenConsultadas_ShouldContainEtiquetasCanonicasVigentes(string ventana, string expectedIso)
        {
            // Act & Assert
            AgDiasVentana.Opciones.Should().ContainKey(ventana);
            AgDiasVentana.Opciones[ventana].Should().Be(expectedIso);
        }

        [Theory]
        [InlineData("UNKNOWN")]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        public void AIsoAndDias_WhenVentanaNoEsValida_ShouldReturnEmpty(string ventanaInvalida)
        {
            // Act
            string iso = AgDiasVentana.AIso(ventanaInvalida);
            IReadOnlyList<int> dias = AgDiasVentana.Dias(ventanaInvalida);

            // Assert
            iso.Should().BeEmpty();
            dias.Should().BeEmpty();
        }

        [Theory]
        [InlineData("3,4,5,6M,6T,7")]
        [InlineData("1,2,6M")]
        [InlineData("6T,7")]
        [InlineData("LUN,6M,MIE")]
        public void AIsoAndDias_WhenListaContieneTokens6Mo6T_ShouldReturnEmpty(string ventanaConFranja)
        {
            // Act
            string iso = AgDiasVentana.AIso(ventanaConFranja);
            IReadOnlyList<int> dias = AgDiasVentana.Dias(ventanaConFranja);

            // Assert
            iso.Should().BeEmpty();
            dias.Should().BeEmpty();
        }
    }
}

