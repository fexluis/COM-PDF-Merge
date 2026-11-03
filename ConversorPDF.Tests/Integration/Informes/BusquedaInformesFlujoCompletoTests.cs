using System;
using System.IO;
using ConversorPDF.Tests.TestData;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Services;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Integration.Informes
{
    /// <summary>
    /// Flujo completo "Buscar Informes" con PDFs reales: PdfSharp genera los archivos,
    /// PdfPig extrae el texto y PDFsharp los unifica. Sin binarios externos.
    /// </summary>
    public class BusquedaInformesFlujoCompletoTests
    {
        [Fact]
        public void Flujo_RenombrarUnificarYAuditar_MarcaVerificado()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("informes_flujo");
            try
            {
                string especial = Path.Combine(carpeta, "Academica");
                PdfFixtureFactory.CrearPdf(especial, "126 informe.pdf",
                    "INFORME DE CUMPLIMIENTO",
                    "No. Identificación Persona 12.345.678 juan@esap.edu.co");
                PdfFixtureFactory.CrearPdf(especial, "126 certificado.pdf",
                    "CERTIFICADO DE CUMPLIMIENTO",
                    "C.C. 12.345.678");

                var data = Array.CreateInstance(
                    typeof(object),
                    new[] { 2, ConstantsHelper.CONSULTAS_COL_MAX },
                    new[] { 1, 1 }) as object[,];
                data[2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "126";
                data[2, ConstantsHelper.CONSULTAS_CEDULA] = "12345678";
                data[2, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = "JUAN CARLOS PEREZ";
                data[2, ConstantsHelper.CONSULTAS_INFORME] = string.Empty;

                var excel = new ExcelRepositoryFalsa { Consultas = data, Hoja = new HojaExcelFalsa(data) };
                var sut = new BusquedaInformesProcessor(excel, new PdfToTextoHelper(), new PdfService());

                var renombrado = sut.RenombrarYUnificar(carpeta, 0, 0, true, false);

                renombrado.HayError.Should().BeFalse();
                renombrado.Renombrados.Should().Be(2);
                renombrado.Unificados.Should().Be(1);
                renombrado.MergeFallidos.Should().Be(0);
                File.Exists(Path.Combine(carpeta, "126-12345678-UNI-JUANCARLOSPE.pdf")).Should().BeTrue();
                Directory.GetFiles(especial, "*.pdf").Should().BeEmpty("los originales se eliminan tras unificar");

                var auditoria = sut.AuditarConsultas(carpeta, 0, 0, false);

                auditoria.HayError.Should().BeFalse();
                data[2, ConstantsHelper.CONSULTAS_INFORME].Should().Be(EstadoComision.Verificado);
                File.Exists(Path.Combine(carpeta, App.FolderNames.INFORMES_VERIFICADOS, "126-12345678-UNI-JUANCARLOSPE.pdf"))
                      .Should().BeTrue();
                excel.VecesGuardado.Should().Be(1);
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }
    }
}
