using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Models;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Core.Helpers
{
    /// <summary>
    /// Pruebas de la detección de tipo de informe y de la verificación de la cédula
    /// dentro del texto extraído de un PDF.
    /// </summary>
    public class InformeTextoHelperTests
    {
        [Theory]
        [InlineData("CERTIFICADO DE CUMPLIMIENTO", InformeTipo.Cum)]
        [InlineData("INFORME DE CUMPLIMIENTO", InformeTipo.Inf)]
        [InlineData("Informe de Cumplimiento\nCERTIFICADO DE CUMPLIMIENTO", InformeTipo.Cum)]
        [InlineData("CERTIFICADO DE CUMPLIMIENTO\nINFORME DE CUMPLIMIENTO", InformeTipo.Cum)]
        [InlineData("OTRA COSA", InformeTipo.Err)]
        [InlineData("", InformeTipo.Err)]
        [InlineData("ERROR: no se pudo abrir el PDF", InformeTipo.Err)]
        public void DetectarTipo_DevuelveInfCumOErr(string texto, string esperado)
        {
            InformeTextoHelper.DetectarTipo(texto).Should().Be(esperado);
        }

        [Fact]
        public void BuscarCedula_ConPatronCedulaYCorreo_DevuelveAlta()
        {
            const string texto = "INFORME DE CUMPLIMIENTO\nNo. Identificación Persona 12.345.678 juan@esap.edu.co";

            var resultado = InformeTextoHelper.BuscarCedula(texto, new[] { "12345678" });

            resultado.Hallada.Should().BeTrue();
            resultado.Cedula.Should().Be("12345678");
            resultado.Confianza.Should().Be(ConfianzaCedula.Alta);
            resultado.EsVerificacionFiable.Should().BeTrue();
            resultado.Evidencia.Should().Be("cedula+correo");
        }

        [Fact]
        public void BuscarCedula_ConEtiquetaCcYFormatoDeMiles_DevuelveAlta()
        {
            const string texto = "CERTIFICADO DE CUMPLIMIENTO\nC.C. 1.020.456.789";

            var resultado = InformeTextoHelper.BuscarCedula(texto, new[] { "1020456789" });

            resultado.Hallada.Should().BeTrue();
            resultado.Cedula.Should().Be("1020456789");
            resultado.Confianza.Should().Be(ConfianzaCedula.Alta);
        }

        [Fact]
        public void BuscarCedula_ConDosCedulasCandidatasDistintas_EsAmbigua()
        {
            const string texto = "C.C. 12345678\nC.C. 87654321";

            var resultado = InformeTextoHelper.BuscarCedula(texto, new[] { "12345678", "87654321" });

            resultado.Hallada.Should().BeFalse();
            resultado.Evidencia.Should().Be(MotivoInforme.CedulaAmbigua);
        }

        [Fact]
        public void BuscarCedula_CuandoLaCedulaEsperadaNoAparece_ReportaNoHallada()
        {
            const string texto = "INFORME DE CUMPLIMIENTO\nC.C. 11111111";

            var resultado = InformeTextoHelper.BuscarCedula(texto, new[] { "12345678" });

            resultado.Hallada.Should().BeFalse();
            resultado.Evidencia.Should().Be(MotivoInforme.CedulaNoHallada);
        }

        [Fact]
        public void BuscarCedula_SinTexto_ReportaSinTexto()
        {
            InformeTextoHelper.BuscarCedula(string.Empty, new[] { "12345678" })
                .Evidencia.Should().Be(MotivoInforme.SinTexto);

            InformeTextoHelper.BuscarCedula("ERROR: archivo ilegible", new[] { "12345678" })
                .Evidencia.Should().Be(MotivoInforme.SinTexto);
        }

        [Fact]
        public void BuscarCedula_SinCandidatasYConEtiqueta_DevuelveAlta()
        {
            var resultado = InformeTextoHelper.BuscarCedula("Documento 12.345.678", new string[0]);

            resultado.Hallada.Should().BeTrue();
            resultado.Cedula.Should().Be("12345678");
            resultado.Confianza.Should().Be(ConfianzaCedula.Alta);
        }

        [Fact]
        public void BuscarCedula_NoConfundeRubroNiValoresConCedula()
        {
            const string texto = "Rubro presupuestal 2.02.02.12.001\nValor total 1.500.000";

            var resultado = InformeTextoHelper.BuscarCedula(texto, new[] { "12345678" });

            resultado.Hallada.Should().BeFalse();
        }

        [Fact]
        public void BuscarCedula_CedulaSinEtiquetaPeroCoincidente_DevuelveMedia()
        {
            const string texto = "INFORME\nSe certifica la asistencia del señor 12345678 a la comisión programada.";

            var resultado = InformeTextoHelper.BuscarCedula(texto, new[] { "12345678" });

            resultado.Hallada.Should().BeTrue();
            resultado.Confianza.Should().Be(ConfianzaCedula.Media);
        }
    }
}
