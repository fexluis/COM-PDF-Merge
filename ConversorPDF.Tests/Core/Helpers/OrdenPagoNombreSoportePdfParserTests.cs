using ConversorPDF.XLL.Core.Helpers;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Core.Helpers
{
    /// <summary>
    /// Pruebas unitarias para <see cref="OrdenPagoNombreSoportePdfParser"/> y métodos auxiliares de <see cref="UnificarOrdenPagoHelper"/>.
    /// Sigue las convenciones AAA, nombres explicativos {TestAction}_When_{Condition(s)}_Should_{ExpectedBehavior}
    /// y aserciones mediante BeEquivalentTo según test-convention.md.
    /// </summary>
    public class OrdenPagoNombreSoportePdfParserTests
    {
        [Theory]
        [InlineData("39436225_OB-6125-NOM-02-OCASIONALES.pdf", "39436225", "6125", "NOM", "02-OCASIONALES")]
        [InlineData("37898125_OB-5725-SPU-02-EMCALI-ENERGIA-ppse.pdf", "37898125", "5725", "SPU", "02-EMCALI-ENERGIA-ppse")]
        [InlineData("25446325_OB-1225-VIA-YEIMI VIVIANA AGUDO RODRIGUEZ AUTORIZACION 825.pdf",
                    "25446325", "1225", "VIA", "YEIMI VIVIANA AGUDO RODRIGUEZ AUTORIZACION 825")]
        [InlineData("FACTURA_20260205_39436225_OB-6125-NOM-02-X.pdf", "39436225", "6125", "NOM", "02-X")]
        public void TryParse_WhenPatronCompleto_ShouldReturnTodosLosCampos(
            string nombreArchivo, string orden, string obligacion, string categoria, string tercero)
        {
            // Arrange
            var esperado = new OrdenPagoNombreSoportePdf
            {
                Orden = orden,
                Obligacion = obligacion,
                Categoria = categoria,
                Tercero = tercero
            };

            // Act
            bool ok = OrdenPagoNombreSoportePdfParser.TryParse(nombreArchivo, out OrdenPagoNombreSoportePdf result);

            // Assert
            ok.Should().BeTrue();
            result.Should().BeEquivalentTo(esperado);
        }

        [Fact]
        public void TryParse_WhenCategoriaFueraDeCatalogo_ShouldUsarFallbackSinCategoria()
        {
            // Arrange
            var esperado = new OrdenPagoNombreSoportePdf
            {
                Orden = "39436225",
                Obligacion = "6125",
                Categoria = null,
                Tercero = "ZZZ-02-X"
            };

            // Act
            bool ok = OrdenPagoNombreSoportePdfParser.TryParse("39436225_OB-6125-ZZZ-02-X.pdf", out OrdenPagoNombreSoportePdf result);

            // Assert
            ok.Should().BeTrue();
            result.Should().BeEquivalentTo(esperado);
        }

        [Fact]
        public void TryParse_WhenSinCategoria_ShouldDevolverSoloOrdenObligacionYTercero()
        {
            // Arrange
            var esperado = new OrdenPagoNombreSoportePdf
            {
                Orden = "39436225",
                Obligacion = "6125",
                Categoria = null,
                Tercero = "OCASIONALES"
            };

            // Act
            bool ok = OrdenPagoNombreSoportePdfParser.TryParse("39436225_OB-6125-OCASIONALES.pdf", out OrdenPagoNombreSoportePdf result);

            // Assert
            ok.Should().BeTrue();
            result.Should().BeEquivalentTo(esperado);
        }

        [Theory]
        [InlineData("factura-123.pdf")]
        [InlineData("")]
        [InlineData("39436225_OC-6125-NOM-02-X.pdf")]
        public void TryParse_WhenNombreNoValido_ShouldDevolverFalse(string nombreArchivo)
        {
            // Arrange & Act
            bool ok = OrdenPagoNombreSoportePdfParser.TryParse(nombreArchivo, out OrdenPagoNombreSoportePdf result);

            // Assert
            ok.Should().BeFalse();
            result.Should().BeNull();
        }

        [Fact]
        public void SeleccionarSoportes_WhenOrdenCoincide_ShouldIgnorarOtrosNumeros()
        {
            // Arrange
            string[] raiz =
            {
                @"C:\x\39436225_OB-6125-NOM-02-OCASIONALES.pdf",
                @"C:\x\39436299_OB-6125-NOM-02-OTRO.pdf",
                @"C:\x\39436225_OB-1225-VIA-YEIMI 825.pdf"
            };

            var soportes = new System.Collections.Generic.Dictionary<string, OrdenPagoNombreSoportePdf>(System.StringComparer.OrdinalIgnoreCase);
            foreach (string ruta in raiz)
            {
                OrdenPagoNombreSoportePdf d;
                if (OrdenPagoNombreSoportePdfParser.TryParse(System.IO.Path.GetFileName(ruta), out d))
                {
                    soportes[ruta] = d;
                }
            }

            // Act
            var result = UnificarOrdenPagoHelper.SeleccionarSoportes(soportes, "OP", "39436225", "6125");

            // Assert
            result.Should().HaveCount(2);
        }

        [Theory]
        [InlineData("20250204", "51225", "XXX", "Yeimi Viviana", "20250204_51225_YeimiViviana")]
        [InlineData("20250204", "51225", "", "Yeimi Viviana", "20250204_51225_YeimiViviana")]
        [InlineData("20250204", "51225", null, "Yeimi Viviana", "20250204_51225_YeimiViviana")]
        [InlineData("20250204", "51225", "VIA", "Yeimi Viviana", "20250204_51225_YeimiViviana")]
        public void ConstruirNombreSalida_WhenTerceroCabeEnPresupuesto_ShouldNoAgregarHash(
            string fecha, string obligacion, string categoria, string tercero, string esperado)
        {
            // Act
            string result = UnificarOrdenPagoHelper.ConstruirNombreSalida(fecha, obligacion, categoria, tercero);

            // Assert
            result.Should().Be(esperado);
            result.Should().NotMatchRegex(@"_([0-9A-F]{6})$");
            result.Length.Should().BeLessOrEqualTo(UnificarOrdenPagoHelper.MAX_LARGO_NOMBRE);
        }

        [Theory]
        [InlineData("20250204", "51225", "NOM", "CONSTRUCTORAMARQUEZYCIASAS1234", "20250204_51225_Nom_Constructor")]
        public void ConstruirNombreSalida_WhenTerceroEsRecortado_ShouldTruncarSinHash(
            string fecha, string obligacion, string categoria, string tercero, string esperado)
        {
            // Act
            string result = UnificarOrdenPagoHelper.ConstruirNombreSalida(fecha, obligacion, categoria, tercero);

            // Assert
            result.Should().Be(esperado);
            result.Length.Should().BeLessOrEqualTo(UnificarOrdenPagoHelper.MAX_LARGO_NOMBRE);
            result.Should().NotMatchRegex(@"_([0-9A-F]{6})$");
        }

        [Fact]
        public void ConstruirNombreSalida_WhenTerceroEsVacio_ShouldUsarSinTerceroGarantizandoIdempotencia()
        {
            // Act
            string result = UnificarOrdenPagoHelper.ConstruirNombreSalida("20250204", "51225", "NOM", "");

            // Assert
            result.Should().Be("20250204_51225_Nom_Sintercero");
            result.Should().Contain("Nom");
            result.Length.Should().BeLessOrEqualTo(UnificarOrdenPagoHelper.MAX_LARGO_NOMBRE);
        }

        [Theory]
        [InlineData("")]
        [InlineData("99/99/9999")]
        public void ConstruirNombreSalida_WhenFechaInvalida_ShouldUsarSinFecha(string fecha)
        {
            // Act
            string result = UnificarOrdenPagoHelper.ConstruirNombreSalida(fecha, "51225", "NOM", "Yeimi Viviana");

            // Assert
            result.Should().StartWith("SINFECHA_51225_");
            result.Length.Should().BeLessOrEqualTo(UnificarOrdenPagoHelper.MAX_LARGO_NOMBRE);
        }

        [Theory]
        [InlineData("04/02/2025", "20250204")]
        [InlineData("2025-02-04", "20250204")]
        [InlineData("20250204", "20250204")]
        [InlineData("45692", "20250204")]
        [InlineData("", "")]
        [InlineData("texto", "")]
        public void LeerFechaPagoDesdeTexto_WhenFormatoConocido_ShouldNormalizar(string entrada, string esperado)
        {
            OrdenPagoSheetReader.LeerFechaPagoDesdeTexto(entrada).Should().Be(esperado);
        }

        [Fact]
        public void ConstruirNombreSalida_WhenObligacionConPrefijo_ShouldNormalizarADigitos()
        {
            string result = UnificarOrdenPagoHelper.ConstruirNombreSalida("20250204", "OB-51225", "NOM", "Yeimi Viviana");
            result.Should().StartWith("20250204_51225_");
            result.Length.Should().BeLessOrEqualTo(UnificarOrdenPagoHelper.MAX_LARGO_NOMBRE);
        }

        [Fact]
        public void ConstruirNombreSalida_WhenTercerosCompartenPrefijo_ShouldGenerarMismoNombreBaseAlTruncar()
        {
            string a = "ANDERSON CAMACHO ZAPATA";
            string b = "ANDERSON CAMACHO ZAPATA JR";

            string resA = UnificarOrdenPagoHelper.ConstruirNombreSalida("20250528", "68925", null, a);
            string resB = UnificarOrdenPagoHelper.ConstruirNombreSalida("20250528", "68925", null, b);

            resA.Should().Be("20250528_68925_AndersonCamacho");
            resB.Should().Be("20250528_68925_AndersonCamacho");
            resA.Should().NotMatchRegex(@"_([0-9A-F]{6})$");
        }

        [Fact]
        public void ConstruirNombreSalida_WhenObligacionYLargoTerceroExceden30_ShouldTruncarDirectoA30SinHash()
        {
            string obLarga = "OB-2025-00051225"; // 12 dígitos 
            string terceroLargo = new string('A', 40); // > 30 chars

            string result = UnificarOrdenPagoHelper.ConstruirNombreSalida("20250204", obLarga, "CATEGORIALARGA", terceroLargo);

            result.Length.Should().BeLessOrEqualTo(UnificarOrdenPagoHelper.MAX_LARGO_NOMBRE);
            result.Should().NotMatchRegex(@"_[0-9A-F]{6}$");
        }

        [Fact]
        public void ConstruirNombreSalida_WhenMaxLargoReducidoPorPresupuesto_ShouldRespetarLimiteSinHash()
        {
            string terceroLargo = new string('A', 40);
            int presupuesto = 25;

            string result = UnificarOrdenPagoHelper.ConstruirNombreSalida("20250204", "51225", "NOM", terceroLargo, presupuesto);

            result.Length.Should().BeLessOrEqualTo(presupuesto);
            result.Should().NotMatchRegex(@"_[0-9A-F]{6}$");
        }

        
        [Fact]
        public void ConstruirNombreSalida_WhenCategoriaReducePresupuestoTercero_ShouldOmitirCategoria()
        {
            string result = UnificarOrdenPagoHelper.ConstruirNombreSalida("20250204", "51225", "CATEGORIALARGA", "Juan Perez");

            result.Should().StartWith("20250204_51225_JuanPere");
            result.Should().NotContain("Categoria");
            result.Length.Should().BeLessOrEqualTo(UnificarOrdenPagoHelper.MAX_LARGO_NOMBRE);
        }
    }
}
