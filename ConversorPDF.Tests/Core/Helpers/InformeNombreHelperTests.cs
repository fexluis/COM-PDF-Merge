using System.IO;
using ConversorPDF.Tests.TestData;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Core.Helpers
{
    /// <summary>
    /// Pruebas del estándar de nombres {solicitud}-{cedula}-{tipo}-{nombre}.
    /// </summary>
    public class InformeNombreHelperTests
    {
        [Fact]
        public void TryParse_ConNombreEstandar_DevuelveLosSegmentos()
        {
            InformePdf info;

            bool ok = InformeNombreHelper.TryParse("126-12345678-INF-JUANCARLOSPE.pdf", out info);

            ok.Should().BeTrue();
            info.NumeroSolicitud.Should().Be("126");
            info.Cedula.Should().Be("12345678");
            info.Tipo.Should().Be(InformeTipo.Inf);
            info.NombreComisionado.Should().Be("JUANCARLOSPE");
            info.EsParValido.Should().BeTrue();
        }

        [Fact]
        public void TryParse_ConGuionesBajos_NormalizaElSeparador()
        {
            InformePdf info;

            InformeNombreHelper.TryParse("126_12345678_CUM_JUANCARLOSPE.pdf", out info).Should().BeTrue();
            info.Tipo.Should().Be(InformeTipo.Cum);
        }

        [Fact]
        public void TryParse_ConMenosDeCuatroSegmentos_DevuelveFalse()
        {
            InformePdf info;

            InformeNombreHelper.TryParse("126-12345678.pdf", out info).Should().BeFalse();
            info.Should().BeNull();
        }

        [Fact]
        public void TryParse_ConTipoDesconocido_DevuelveFalsePeroLooseLoAcepta()
        {
            InformePdf info;
            string solicitud, cedula, tipo, nombre;

            InformeNombreHelper.TryParse("126-12345678-XXX-JUAN.pdf", out info).Should().BeFalse();
            InformeNombreHelper.TryParseLoose("126-12345678-XXX-JUAN.pdf", out solicitud, out cedula, out tipo, out nombre)
                .Should().BeTrue();
            tipo.Should().Be("XXX");
            nombre.Should().Be("JUAN");
        }

        [Fact]
        public void Construir_LimpiaTildesYTruncaANoDoceCaracteres()
        {
            string nombre = InformeNombreHelper.Construir("126", "12.345.678", "inf", "José Pérez Núñez");

            nombre.Should().Be("126-12345678-INF-JOSEPEREZNUN");
            InformeNombreHelper.LimpiarNombre("José Pérez Núñez").Length.Should().Be(12);
        }

        [Fact]
        public void Construir_ConTipoInvalido_UsaErr()
        {
            InformeNombreHelper.Construir("126", "12345678", "XXX", "JUAN")
                .Should().Be("126-12345678-ERR-JUAN");
        }

        [Fact]
        public void Construir_SinCedula_UsaCeros()
        {
            InformeNombreHelper.Construir("126", "", "ERR", "")
                .Should().Be("126-000000-ERR-DESCONOCIDO");
        }

        [Theory]
        [InlineData("126-informe.pdf", "126")]
        [InlineData("ACTA-126-2026.pdf", "126")]
        [InlineData("acta.pdf", "")]
        public void ExtraerNumeroSolicitud_TomaElPrimerBloqueNumerico(string archivo, string esperado)
        {
            InformeNombreHelper.ExtraerNumeroSolicitud(archivo).Should().Be(esperado);
        }

        [Fact]
        public void NormalizarDocumento_QuitaPuntosEspaciosYComas()
        {
            InformeNombreHelper.NormalizarDocumento("12.345, 678").Should().Be("12345678");
            InformeNombreHelper.SoloDigitos("CC 12.345.678").Should().Be("12345678");
        }

        [Fact]
        public void NombreUnico_AgregaSufijoSiElArchivoExiste()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("nombre_unico");
            try
            {
                File.WriteAllText(Path.Combine(carpeta, "126-12345678-INF-JUAN.pdf"), "x");

                string resultado = InformeNombreHelper.NombreUnico(carpeta, "126-12345678-INF-JUAN.pdf");

                Path.GetFileName(resultado).Should().Be("126-12345678-INF-JUAN_1.pdf");
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }
    }
}
