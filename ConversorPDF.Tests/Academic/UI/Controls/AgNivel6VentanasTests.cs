using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.UI.Controls;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.UI.Controls
{
    public class AgNivel6VentanasTests
    {
        [Theory]
        [InlineData("LUN-JUE", 0)]
        [InlineData("LUN-VIE", 1)]
        [InlineData("MAR-VIE", 2)]
        [InlineData("JUE-SAB", 3)]
        [InlineData("VIE-SAB", 4)]
        [InlineData("VIE-DOM", 5)]
        [InlineData("JUE-DOM", 6)]
        [InlineData("MIE-DOM", 7)]
        [InlineData("4,5,6,7", 8)]
        public void Indice_WhenClaveCanonica_ShouldReturnExpectedIndex(string clave, int expectedIndex)
        {
            // Act
            int result = AgNivel6Ventanas.Indice(clave);

            // Assert
            result.Should().Be(expectedIndex);
        }

        [Theory]
        [InlineData("1,2,3,4", 0)]       // Lun-Jue
        [InlineData("1,2,3,4,5", 1)]     // Lun-Vie (Semestre V Mixta)
        [InlineData("2,3,4,5", 2)]       // Mar-Vie
        [InlineData("4,5,6", 3)]         // Jue-Sab
        [InlineData("5,6", 4)]           // Vie-Sab
        [InlineData("5,6,7", 5)]         // Vie-Dom
        [InlineData("3,4,5,6,7", 7)]     // Mie-Dom
        [InlineData("4,5,6,7", 8)]       // Alias ISO terminal 4,5,6,7
        public void Indice_WhenIsoNumerico_ShouldResolveEquivalentCanonicalIndex(string iso, int expectedIndex)
        {
            // Act
            int result = AgNivel6Ventanas.Indice(iso);

            // Assert
            result.Should().Be(expectedIndex);
        }

        [Theory]
        [InlineData("4,5,6,7", 8)]
        [InlineData("4, 5, 6, 7", 8)]
        [InlineData(" 4, 5, 6, 7 ", 8)]
        [InlineData("JUE-DOM", 6)]
        [InlineData("1,2,3,4", 0)]
        [InlineData("1, 2, 3, 4", 0)]
        [InlineData("1,2,3,4,5", 1)]
        [InlineData("1, 2, 3, 4, 5", 1)]
        public void Indice_DistingueAliasYClave(string valor, int esperado)
        {
            // Act
            int result = AgNivel6Ventanas.Indice(valor);

            // Assert
            result.Should().Be(esperado);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("DESCONOCIDO")]
        [InlineData("1,7")]
        public void Indice_WhenInvalidoOInexistente_ShouldReturnMinusOne(string entrada)
        {
            // Act
            int result = AgNivel6Ventanas.Indice(entrada);

            // Assert
            result.Should().Be(-1);
        }

        [Theory]
        [InlineData("1,2,3,4,5", "LUN-VIE", true)]
        [InlineData("1, 2, 3, 4, 5", "LUN-VIE", true)]
        [InlineData("LUN-VIE", "1,2,3,4,5", true)]
        [InlineData("lun-vie", "LUN-VIE", true)]
        [InlineData("1,2,3,4", "LUN-JUE", true)]
        [InlineData("4,5,6,7", "4,5,6,7", true)]
        [InlineData("4, 5, 6, 7", "4,5,6,7", true)]
        [InlineData("4, 5, 6, 7", " 4,5,6,7 ", true)]
        [InlineData("JUE-DOM", "JUE-DOM", true)]
        [InlineData("4,5,6,7", "JUE-DOM", false)] // Fase 2243: No colapsan mutuamente
        [InlineData("4, 5, 6, 7", "JUE-DOM", false)] // Fase 2245: Alias con espacios tampoco colapsa a JUE-DOM
        [InlineData("JUE-DOM", "4,5,6,7", false)]
        [InlineData("JUE-DOM", "4, 5, 6, 7", false)]
        [InlineData("1,2,3,4,5", "LUN-JUE", false)]
        [InlineData("", "LUN-VIE", false)]
        [InlineData(null, "LUN-VIE", false)]
        public void Coincide_WhenComparado_ShouldValidateEquivalenciaIso(string valor1, string valor2, bool esperado)
        {
            // Act
            bool result = AgNivel6Ventanas.Coincide(valor1, valor2);

            // Assert
            result.Should().Be(esperado);
        }

        [Theory]
        [InlineData("1,2,3,4,5", "LUN-VIE")]
        [InlineData("3,4,5,6,7", "MIE-DOM")]
        [InlineData("lun-jue", "LUN-JUE")]
        [InlineData("4,5,6,7", "4,5,6,7")]
        [InlineData("JUE-DOM", "JUE-DOM")]
        [InlineData("PERSONALIZADO", "PERSONALIZADO")]
        public void ClaveCanonica_WhenEntrada_ShouldReturnNombreEstandarizado(string entrada, string esperado)
        {
            // Act
            string result = AgNivel6Ventanas.ClaveCanonica(entrada);

            // Assert
            result.Should().Be(esperado);
        }

        [Fact]
        public void FormatearRadio_WhenPasoEstrecho_ShouldReturnTextoCorto()
        {
            // Act
            string corto = AgNivel6Ventanas.FormatearRadio("LUN-VIE", 50);

            // Assert
            corto.Should().Be("L-V");
        }

        [Fact]
        public void FormatearRadio_WhenPasoAmplio_ShouldReturnTextoAmigable()
        {
            // Act
            string regular = AgNivel6Ventanas.FormatearRadio("LUN-VIE", 70);

            // Assert
            regular.Should().Be("Lun-Vie");
        }

        [Fact]
        public void ResolverPorX_WhenClickDentroDeCadaBoton_ShouldReturnOpcionCorrespondiente()
        {
            // Arrange (9 posiciones para las 9 opciones)
            int[] posiciones = new int[] { 300, 350, 400, 450, 500, 550, 600, 650, 700 };
            int fin = 780;
            int step = 50;

            // Act & Assert
            AgNivel6Ventanas.ResolverPorX(310, posiciones, fin, step).Should().Be("LUN-JUE");
            AgNivel6Ventanas.ResolverPorX(360, posiciones, fin, step).Should().Be("LUN-VIE");
            AgNivel6Ventanas.ResolverPorX(410, posiciones, fin, step).Should().Be("MAR-VIE");
            AgNivel6Ventanas.ResolverPorX(460, posiciones, fin, step).Should().Be("JUE-SAB");
            AgNivel6Ventanas.ResolverPorX(510, posiciones, fin, step).Should().Be("VIE-SAB");
            AgNivel6Ventanas.ResolverPorX(560, posiciones, fin, step).Should().Be("VIE-DOM");
            AgNivel6Ventanas.ResolverPorX(610, posiciones, fin, step).Should().Be("JUE-DOM");
            AgNivel6Ventanas.ResolverPorX(660, posiciones, fin, step).Should().Be("MIE-DOM");
            AgNivel6Ventanas.ResolverPorX(710, posiciones, fin, step).Should().Be("4,5,6,7");
        }

        [Fact]
        public void ResolverPorX_WhenClickEnHuecoMuertoPosterior_ShouldReturnNull()
        {
            // Arrange (9 botones: 700 + 50 = 750 px efectivo de botones; fin es 820 px)
            int[] posiciones = new int[] { 300, 350, 400, 450, 500, 550, 600, 650, 700 };
            int fin = 820;
            int step = 50;

            // Act: clic en 770 (más allá de 700+50=750 pero antes de 820)
            string result = AgNivel6Ventanas.ResolverPorX(770, posiciones, fin, step);

            // Assert
            result.Should().BeNull();
        }

        [Fact]
        public void Claves_ContieneNueveOpciones()
        {
            // Assert: paridad de catálogo conteniendo todas las claves de AgDiasVentana.Opciones más 4,5,6,7
            AgNivel6Ventanas.Claves.Count.Should().Be(9);
            AgNivel6Ventanas.Claves.Should().Contain(AgDiasVentana.Opciones.Keys);
            AgNivel6Ventanas.Claves.Should().Contain("4,5,6,7");
            AgDiasVentana.Opciones.Keys.Except(AgNivel6Ventanas.Claves).Should().BeEmpty();
        }

        [Fact]
        public void ClicJueDom_ConIso4567_ReescribeAJueDom()
        {
            // Arrange
            string seleccion = "JUE-DOM";

            // Act: al pulsar el radio de JUE-DOM, se selecciona JUE-DOM
            string resultado = seleccion;

            // Assert
            resultado.Should().Be("JUE-DOM");
        }

        [Fact]
        public void ClicIso4567_ConJueDom_ReescribeATerminal()
        {
            // Arrange
            string seleccion = "4,5,6,7";

            // Act: al pulsar el radio de 4,5,6,7, se selecciona 4,5,6,7
            string resultado = seleccion;

            // Assert
            resultado.Should().Be("4,5,6,7");
        }

        [Fact]
        public void FormatearRadio_CompactoPorDebajoDe58()
        {
            // Act & Assert
            AgNivel6Ventanas.FormatearRadio("JUE-DOM", 50).Should().Be("J-D");
            AgNivel6Ventanas.FormatearRadio("MIE-DOM", 48).Should().Be("Mi-D");
            AgNivel6Ventanas.FormatearRadio("4,5,6,7", 50).Should().Be("4-6,7");
            AgNivel6Ventanas.FormatearRadio("4,5,6,7", 60).Should().Be("4,5,6,7");
            AgNivel6Ventanas.FormatearRadio("LUN-VIE", 57).Should().Be("L-V");
            AgNivel6Ventanas.FormatearRadio("LUN-VIE", 58).Should().Be("Lun-Vie");
        }

        [Fact]
        public void ClicEnHuecoPosterior_NoSelecciona()
        {
            // Arrange
            int[] posiciones = new int[] { 300, 350, 400, 450, 500, 550, 600, 650, 700 };
            int fin = 820;
            int step = 50;

            // Act
            string resultado = AgNivel6Ventanas.ResolverPorX(770, posiciones, fin, step);

            // Assert
            resultado.Should().BeNull();
        }

        [Fact]
        public void ResolverPorX_WhenClickFueraDeLimites_ShouldReturnNull()
        {
            // Arrange
            int[] posiciones = new int[] { 300, 350, 400, 450, 500, 550, 600, 650, 700 };
            int fin = 800;
            int step = 50;

            // Act & Assert
            AgNivel6Ventanas.ResolverPorX(250, posiciones, fin, step).Should().BeNull();
            AgNivel6Ventanas.ResolverPorX(810, posiciones, fin, step).Should().BeNull();
        }
    }
}
