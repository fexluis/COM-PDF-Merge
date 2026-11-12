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
        [InlineData("LUN-JUE", "LUN-VIE")]
        [InlineData("LUN-VIE", "")]
        [InlineData("MAR-VIE", "LUN-VIE")]
        [InlineData("VIE-DOM", "JUE-DOM")]
        [InlineData("JUE-DOM", "MIE-DOM")]
        [InlineData("MIE-DOM", "")]
        [InlineData("VIE-SAB", "JUE-SAB")]
        [InlineData("JUE-SAB", "4,5,6,7")]
        [InlineData("4,5,6,7", "")]
        public void Siguiente_CadenaDeEscalones(string desde, string esperado)
        {
            AgEscaleraVentanas.Siguiente(desde).Should().Be(esperado);
        }

        [Fact]
        public void Escalera_VieDom_DebeTenerExactamenteDosPasosYNuncaTres()
        {
            var pasos = AgEscaleraVentanas.Escalera("VIE-DOM");

            pasos.Should().Equal(new[] { "JUE-DOM", "MIE-DOM" });
            pasos.Should().NotHaveCount(3);
            pasos.Count.Should().BeLessThanOrEqualTo(AgEscaleraVentanas.MaxPasos);
        }

        [Fact]
        public void Escalera_CadenaVieSab_DebeMantenerseIntacta()
        {
            var pasos = AgEscaleraVentanas.Escalera("VIE-SAB");

            pasos.Should().Equal(new[] { "JUE-SAB", "4,5,6,7" });
            pasos.Should().HaveCount(2);
        }

        [Theory]
        [InlineData("4,5,6,7")]
        [InlineData("MIE-DOM")]
        [InlineData("LUN-VIE")]
        public void PuedeEscalar_CuandoVentanaEnTope_DevuelveFalse(string ventanaTope)
        {
            AgEscaleraVentanas.PuedeEscalar(ventanaTope).Should().BeFalse();
        }

        [Theory]
        [InlineData("VIE-DOM")]
        [InlineData("VIE-SAB")]
        [InlineData("LUN-JUE")]
        [InlineData("MAR-VIE")]
        public void Escalera_Idempotencia_LlamadasMultiplesRetornanMismoResultadoSinRepeticion(string ventana)
        {
            var primera = AgEscaleraVentanas.Escalera(ventana);
            var segunda = AgEscaleraVentanas.Escalera(ventana);

            primera.Should().Equal(segunda);
            primera.Should().OnlyHaveUniqueItems();
        }

        [Fact]
        public void Rotular_MieDom_Unidad4h_DebeRotularCon6My6T()
        {
            var rotulo = AgEscaleraVentanas.Rotular("MIE-DOM", 4.0);

            rotulo.Should().Be("3,4,5,6M,6T,7");
        }

        [Fact]
        public void Rotular_MieDom_Unidad8h_DebeRotularConDiasPuros()
        {
            var rotulo = AgEscaleraVentanas.Rotular("MIE-DOM", 8.0);

            rotulo.Should().Be("3,4,5,6,7");
        }

        [Fact]
        public void Rotular_LunVie_Unidad4h_DebeRotularDiasCompletos1A5()
        {
            var rotulo = AgEscaleraVentanas.Rotular("LUN-VIE", 4.0);

            rotulo.Should().Be("1,2,3,4,5");
        }
    }
}

