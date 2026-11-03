using System;
using System.IO;
using ConversorPDF.Tests.TestData;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Processors;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Core.Processors
{
    /// <summary>
    /// Pruebas del renombrado al estándar {solicitud}-{cedula}-{tipo}-{nombre} del proceso
    /// "Buscar Informes" (PdfPig simulados y hoja CONSULTAS falsa).
    /// </summary>
    public class BusquedaInformesRenombradoTests
    {
        private readonly Mock<IPdfToTexto> _pdfToTexto;
        private readonly Mock<IPdfService> _pdfService;
        private readonly ExcelRepositoryFalsa _excel;
        private readonly BusquedaInformesProcessor _sut;

        public BusquedaInformesRenombradoTests()
        {
            _pdfToTexto = new Mock<IPdfToTexto>();
            _pdfService = new Mock<IPdfService>();
            _excel = new ExcelRepositoryFalsa();

            _pdfToTexto.Setup(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>())).Returns(string.Empty);
            _pdfToTexto.Setup(x => x.ContarPaginas(It.IsAny<string>())).Returns(1);

            _sut = new BusquedaInformesProcessor(_excel, _pdfToTexto.Object, _pdfService.Object);
        }

        private static object[,] CrearConsultas()
        {
            return Array.CreateInstance(
                typeof(object),
                new[] { 2, ConstantsHelper.CONSULTAS_COL_MAX },
                new[] { 1, 1 }) as object[,];
        }

        private void ConfigurarFila(object[,] data, string solicitud, string cedula, string nombre, string informe = "")
        {
            data[2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = solicitud;
            data[2, ConstantsHelper.CONSULTAS_CEDULA] = cedula;
            data[2, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = nombre;
            data[2, ConstantsHelper.CONSULTAS_INFORME] = informe;

            _excel.Consultas = data;
            _excel.Hoja = new HojaExcelFalsa(data);
        }

        [Fact]
        public void RenombrarYUnificar_CarpetaInexistente_DevuelveError()
        {
            var resultado = _sut.RenombrarYUnificar(
                Path.Combine(Path.GetTempPath(), "no_existe_" + Guid.NewGuid()), 0, 0, true, false);

            resultado.HayError.Should().BeTrue();
            resultado.PdfsExaminados.Should().Be(0);
        }

        [Fact]
        public void RenombrarYUnificar_NombreSinPatronConCedulaEnPdf_RenombraAlEstandar()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("renombrar_ok");
            try
            {
                ConfigurarFila(CrearConsultas(), "126", "12345678", "JUAN CARLOS PEREZ");
                File.WriteAllText(Path.Combine(carpeta, "126-escaneo.pdf"), "x");

                _pdfToTexto.Setup(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()))
                           .Returns("INFORME DE CUMPLIMIENTO\nNo. Identificación Persona 12.345.678");

                var resultado = _sut.RenombrarYUnificar(carpeta, 0, 0, true, false);

                resultado.HayError.Should().BeFalse();
                resultado.Renombrados.Should().Be(1);
                resultado.PdfsExaminados.Should().Be(1);
                File.Exists(Path.Combine(carpeta, "126-12345678-INF-JUANCARLOSPE.pdf")).Should().BeTrue();
                File.Exists(Path.Combine(carpeta, "126-escaneo.pdf")).Should().BeFalse();
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }

        [Fact]
        public void RenombrarYUnificar_CedulaDelPdfNoCoincide_GeneraErr()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("renombrar_err");
            try
            {
                ConfigurarFila(CrearConsultas(), "126", "12345678", "JUAN CARLOS PEREZ");
                File.WriteAllText(Path.Combine(carpeta, "126-escaneo.pdf"), "x");

                _pdfToTexto.Setup(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()))
                           .Returns("INFORME DE CUMPLIMIENTO\nC.C. 87654321");

                var resultado = _sut.RenombrarYUnificar(carpeta, 0, 0, true, false);

                resultado.Renombrados.Should().Be(1);
                File.Exists(Path.Combine(carpeta, "126-12345678-ERR-JUANCARLOSPE.pdf")).Should().BeTrue();
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }

        [Fact]
        public void RenombrarYUnificar_SinTexto_GeneraErrYCuentaSinTexto()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("renombrar_sin_texto");
            try
            {
                ConfigurarFila(CrearConsultas(), "126", "12345678", "JUAN CARLOS PEREZ");
                File.WriteAllText(Path.Combine(carpeta, "126-escaneado.pdf"), "x");

                var resultado = _sut.RenombrarYUnificar(carpeta, 0, 0, true, false);

                resultado.SinTexto.Should().Be(1);
                File.Exists(Path.Combine(carpeta, "126-12345678-ERR-JUANCARLOSPE.pdf")).Should().BeTrue();
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }
    }
}
