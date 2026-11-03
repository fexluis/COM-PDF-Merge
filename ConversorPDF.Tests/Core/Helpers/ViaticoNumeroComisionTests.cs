using ConversorPDF.XLL.Core.Helpers;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Core.Helpers
{
    public class ViaticoNumeroComisionTests
    {
        private static int AnioVigenciaActual
            => int.TryParse(App.APP_VIGENCIA, out int vig) ? vig : 2026;

        [Fact]
        public void TryParse_WhenFormatoYYValido_ShouldExtraerAnioYConsecutivo()
        {
            int yy = AnioVigenciaActual % 100;
            string val1 = $"1{yy:D2}";
            string val45 = $"45{yy:D2}";
            string val145 = $"145{yy:D2}";

            ViaticoNumeroComision.TryParse(val1, out int a1, out int c1).Should().BeTrue();
            a1.Should().Be(AnioVigenciaActual);
            c1.Should().Be(1);

            ViaticoNumeroComision.TryParse(val45, out int a45, out int c45).Should().BeTrue();
            a45.Should().Be(AnioVigenciaActual);
            c45.Should().Be(45);

            ViaticoNumeroComision.TryParse(val145, out int a145, out int c145).Should().BeTrue();
            a145.Should().Be(AnioVigenciaActual);
            c145.Should().Be(145);

            // Año 2025 específico
            ViaticoNumeroComision.TryParse("12625", out int a25, out int c25).Should().BeTrue();
            a25.Should().Be(2025);
            c25.Should().Be(126);
        }

        [Theory]
        [InlineData("1", 1)]
        [InlineData("45", 45)]
        public void TryParse_WhenMenorA3Digitos_ShouldAsumirVigenciaActual(string valor, int expectedConsecutivo)
        {
            // Act
            bool result = ViaticoNumeroComision.TryParse(valor, out int anio, out int consecutivo);

            // Assert
            result.Should().BeTrue();
            anio.Should().Be(AnioVigenciaActual);
            consecutivo.Should().Be(expectedConsecutivo);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("abc")]
        [InlineData("COM-2026")]
        [InlineData("0")]
        public void TryParse_WhenValorInvalido_ShouldReturnFalse(string valor)
        {
            // Act
            bool result = ViaticoNumeroComision.TryParse(valor, out int anio, out int consecutivo);

            // Assert
            result.Should().BeFalse();
            anio.Should().Be(0);
            consecutivo.Should().Be(0);
        }

        [Fact]
        public void EnRango_WhenRangoNormalYY_ShouldFiltrarSegunAceptacion()
        {
            int yy = AnioVigenciaActual % 100;
            int yyAnt = (AnioVigenciaActual - 1) % 100;

            int ini = int.Parse($"1{yy:D2}");
            int f = int.Parse($"145{yy:D2}");

            ViaticoNumeroComision.EnRango($"1{yy:D2}", ini, f).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"45{yy:D2}", ini, f).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"145{yy:D2}", ini, f).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"146{yy:D2}", ini, f).Should().BeFalse();
            ViaticoNumeroComision.EnRango($"126{yyAnt:D2}", ini, f).Should().BeFalse(); // año anterior
            ViaticoNumeroComision.EnRango("", ini, f).Should().BeFalse();
            ViaticoNumeroComision.EnRango("abc", ini, f).Should().BeFalse();
        }

        [Fact]
        public void EnRango_WhenRangoAnioAnterior_ShouldFiltrarCorrectamente()
        {
            int yy = AnioVigenciaActual % 100;
            int yyAnt = (AnioVigenciaActual - 1) % 100;

            int inicio = int.Parse($"100{yyAnt:D2}");
            int fin = int.Parse($"150{yyAnt:D2}");

            ViaticoNumeroComision.EnRango($"126{yyAnt:D2}", inicio, fin).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"99{yyAnt:D2}", inicio, fin).Should().BeFalse();
            ViaticoNumeroComision.EnRango($"151{yyAnt:D2}", inicio, fin).Should().BeFalse();
            ViaticoNumeroComision.EnRango($"126{yy:D2}", inicio, fin).Should().BeFalse(); // año vigente fuera de rango año anterior
        }

        [Fact]
        public void EnRango_WhenSinFiltro_ShouldRetornarTrue()
        {
            int yy = AnioVigenciaActual % 100;
            int yyAnt = (AnioVigenciaActual - 1) % 100;

            // (0, 0) representa sin filtro
            int inicio = 0;
            int fin = 0;

            ViaticoNumeroComision.EnRango($"1{yy:D2}", inicio, fin).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"145{yy:D2}", inicio, fin).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"146{yy:D2}", inicio, fin).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"126{yyAnt:D2}", inicio, fin).Should().BeTrue();
        }

        [Fact]
        public void EnRango_WhenLimitesConsecutivosPuros_ShouldInterpretarComoVigenciaActual()
        {
            int yy = AnioVigenciaActual % 100;
            int yyAnt = (AnioVigenciaActual - 1) % 100;

            int inicio = 1;
            int fin = 145;

            ViaticoNumeroComision.EnRango($"1{yy:D2}", inicio, fin).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"45{yy:D2}", inicio, fin).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"145{yy:D2}", inicio, fin).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"146{yy:D2}", inicio, fin).Should().BeFalse();
            ViaticoNumeroComision.EnRango($"126{yyAnt:D2}", inicio, fin).Should().BeFalse();
        }

        [Fact]
        public void EnRango_WhenRangoInvertido_ShouldNormalizarYEvaluarCorrectamente()
        {
            int yy = AnioVigenciaActual % 100;
            int inicio = int.Parse($"145{yy:D2}");
            int fin = int.Parse($"1{yy:D2}");

            ViaticoNumeroComision.EnRango($"45{yy:D2}", inicio, fin).Should().BeTrue();
            ViaticoNumeroComision.EnRango($"146{yy:D2}", inicio, fin).Should().BeFalse();
        }
    }
}
