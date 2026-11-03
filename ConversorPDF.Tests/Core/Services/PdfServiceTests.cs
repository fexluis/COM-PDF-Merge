using System;
using System.IO;
using ConversorPDF.Tests.TestData;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Core.Services
{
    /// <summary>
    /// Pruebas de integración reales de PdfSharp (unificación) y PdfPig (lectura y conteo
    /// de páginas) usando PDFs generados en una carpeta temporal.
    /// </summary>
    public class PdfServiceTests
    {
        [Fact]
        public void CombinarPDFs_GeneraUnPdfConLasPaginasEnOrden()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("pdf_service");
            try
            {
                string cum = PdfFixtureFactory.CrearPdf(carpeta, "cum.pdf", "CERTIFICADO DE CUMPLIMIENTO", "C.C. 1020456789");
                string inf = PdfFixtureFactory.CrearPdf(carpeta, "inf.pdf", "INFORME DE CUMPLIMIENTO", "C.C. 1020456789");
                string salida = Path.Combine(carpeta, "uni.pdf");

                new PdfService().CombinarPDFs(new[] { cum, inf }, salida);

                var lector = new PdfToTextoHelper();

                File.Exists(salida).Should().BeTrue();
                lector.ContarPaginas(salida).Should().Be(2);
                lector.ExtraerTextoPagina(salida, 1).Should().Contain("CERTIFICADO");
                lector.ExtraerTextoPagina(salida, 2).Should().Contain("INFORME");
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }

        [Fact]
        public void CombinarPDFs_SinArchivosValidos_LanzaInvalidOperationException()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("pdf_service_vacio");
            try
            {
                string salida = Path.Combine(carpeta, "uni.pdf");
                string inexistente = Path.Combine(carpeta, "noexiste.pdf");

                Action accion = () => new PdfService().CombinarPDFs(new[] { inexistente }, salida);

                accion.Should().Throw<InvalidOperationException>();
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }
    }
}
