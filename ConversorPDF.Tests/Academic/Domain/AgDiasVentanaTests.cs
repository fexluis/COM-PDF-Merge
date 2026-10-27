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
        public void AIsoAndDias_WhenVentanaIsValid_ShouldReturnExpectedIsoAndDays(
            string ventana, string expectedIso, int[] expectedDays)
        {
            // Act
            string iso = AgDiasVentana.AIso(ventana);
            IReadOnlyList<int> dias = AgDiasVentana.Dias(ventana);

            // Assert
            iso.Should().Be(expectedIso);
            dias.Should().Equal(expectedDays);
        }

        [Fact]
        public void AIso_WhenVentanaIsUnknown_ShouldReturnEmptyString()
        {
            // Act
            string iso = AgDiasVentana.AIso("UNKNOWN");

            // Assert
            iso.Should().BeEmpty();
        }
    }
}
