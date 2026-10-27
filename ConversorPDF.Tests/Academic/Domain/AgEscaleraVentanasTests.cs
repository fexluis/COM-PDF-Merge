using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Domain
{
    public class AgEscaleraVentanasTests
    {
        [Theory]
        [InlineData("VIE-SAB", "JUE-SAB")]
        [InlineData("JUE-SAB", "4,5,6,7")]
        [InlineData("4,5,6,7", "")]
        [InlineData("MAR-VIE", "1,2,3,4,5")]
        [InlineData("LUN-JUE", "1,2,3,4,5")]
        [InlineData("VIE-DOM", "")]
        public void Siguiente_CadenaDeEscalones(string desde, string esperado)
        {
            AgEscaleraVentanas.Siguiente(desde).Should().Be(esperado);
        }

        [Fact]
        public void Escalera_NuncaSuperaDosPasos()
        {
            AgEscaleraVentanas.Escalera("VIE-SAB").Should().HaveCount(2);
        }

        [Fact]
        public void PuedeEscalar_CuandoVentanaEnTope_DevuelveFalse()
        {
            AgEscaleraVentanas.PuedeEscalar("4,5,6,7").Should().BeFalse();
        }
    }
}
