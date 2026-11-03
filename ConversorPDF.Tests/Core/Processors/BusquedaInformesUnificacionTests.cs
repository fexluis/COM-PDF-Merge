using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
    /// Pruebas de la unificación de parejas INF + CUM del proceso "Buscar Informes".
    /// </summary>
    public class BusquedaInformesUnificacionTests
    {
        private readonly Mock<IPdfToTexto> _pdfToTexto;
        private readonly Mock<IPdfService> _pdfService;
        private readonly ExcelRepositoryFalsa _excel;
        private readonly BusquedaInformesProcessor _sut;
        private List<string> _archivosUnificados;

        public BusquedaInformesUnificacionTests()
        {
            _pdfToTexto = new Mock<IPdfToTexto>();
            _pdfService = new Mock<IPdfService>();
            _excel = new ExcelRepositoryFalsa();

            _pdfToTexto.Setup(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>())).Returns(string.Empty);
            _pdfToTexto.Setup(x => x.ContarPaginas(It.IsAny<string>())).Returns(1);

            _pdfService.Setup(x => x.CombinarPDFs(It.IsAny<IEnumerable<string>>(), It.IsAny<string>()))
                       .Callback<IEnumerable<string>, string>((archivos, salida) =>
                       {
                           _archivosUnificados = archivos.ToList();
                           File.WriteAllText(salida, "PDF");
                       });

            _sut = new BusquedaInformesProcessor(_excel, _pdfToTexto.Object, _pdfService.Object);
        }

        private static object[,] CrearConsultas()
        {
            return Array.CreateInstance(
                typeof(object),
                new[] { 2, ConstantsHelper.CONSULTAS_COL_MAX },
                new[] { 1, 1 }) as object[,];
        }

        private void ConfigurarFila(string solicitud, string cedula, string nombre, string informe = "")
        {
            var data = CrearConsultas();
            data[2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = solicitud;
            data[2, ConstantsHelper.CONSULTAS_CEDULA] = cedula;
            data[2, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = nombre;
            data[2, ConstantsHelper.CONSULTAS_INFORME] = informe;

            _excel.Consultas = data;
            _excel.Hoja = new HojaExcelFalsa(data);
        }

        [Fact]
        public void RenombrarYUnificar_ParejaInfCum_UnificaEnLaRaizConCumPrimero()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("unificar_pareja");
            try
            {
                ConfigurarFila("126", "12345678", "JUAN CARLOS PEREZ");

                string especial = Path.Combine(carpeta, "Academica");
                Directory.CreateDirectory(especial);
                string inf = Path.Combine(especial, "126-12345678-INF-JUANCARLOSPE.pdf");
                string cum = Path.Combine(especial, "126-12345678-CUM-JUANCARLOSPE.pdf");
                File.WriteAllText(inf, "inf");
                File.WriteAllText(cum, "cum");

                _pdfToTexto.Setup(x => x.ContarPaginas(It.Is<string>(p => p.EndsWith(".tmp.pdf")))).Returns(2);

                var resultado = _sut.RenombrarYUnificar(carpeta, 0, 0, true, false);

                resultado.YaCorrectos.Should().Be(2);
                resultado.Unificados.Should().Be(1);
                resultado.MergeFallidos.Should().Be(0);

                File.Exists(Path.Combine(carpeta, "126-12345678-UNI-JUANCARLOSPE.pdf")).Should().BeTrue();
                File.Exists(inf).Should().BeFalse();
                File.Exists(cum).Should().BeFalse();

                _archivosUnificados.Should().HaveCount(2);
                Path.GetFileName(_archivosUnificados[0]).Should().Contain("-CUM-");
                Path.GetFileName(_archivosUnificados[1]).Should().Contain("-INF-");
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }

        [Fact]
        public void RenombrarYUnificar_SoloInf_NoUnificaYCuentaSinPareja()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("unificar_sin_pareja");
            try
            {
                ConfigurarFila("126", "12345678", "JUAN CARLOS PEREZ");

                string especial = Path.Combine(carpeta, "Direccion");
                Directory.CreateDirectory(especial);
                string inf = Path.Combine(especial, "126-12345678-INF-JUANCARLOSPE.pdf");
                File.WriteAllText(inf, "inf");

                var resultado = _sut.RenombrarYUnificar(carpeta, 0, 0, true, false);

                resultado.Unificados.Should().Be(0);
                resultado.RechazadosSinPareja.Should().Be(1);
                File.Exists(inf).Should().BeTrue();
                File.Exists(Path.Combine(carpeta, "126-12345678-UNI-JUANCARLOSPE.pdf")).Should().BeFalse();
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }
        [Fact]
        public void RenombrarYUnificar_EnSimulacion_NoTocaElDisco()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("unificar_simulado");
            try
            {
                ConfigurarFila("126", "12345678", "JUAN CARLOS PEREZ");
                string original = Path.Combine(carpeta, "126-escaneo.pdf");
                File.WriteAllText(original, "x");

                _pdfToTexto.Setup(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()))
                           .Returns("INFORME DE CUMPLIMIENTO\nNo. Identificación Persona 12.345.678");

                var resultado = _sut.RenombrarYUnificar(carpeta, 0, 0, true, true);

                resultado.Simulado.Should().BeTrue();
                resultado.Renombrados.Should().Be(1);
                File.Exists(original).Should().BeTrue();
                File.Exists(Path.Combine(carpeta, "126-12345678-INF-JUANCARLOSPE.pdf")).Should().BeFalse();
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }

        [Fact]
        public void RenombrarYUnificar_ExcluyeVerificadosYDuplicados()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("unificar_excluye");
            try
            {
                ConfigurarFila("126", "12345678", "JUAN CARLOS PEREZ");

                string verificados = Path.Combine(carpeta, App.FolderNames.INFORMES_VERIFICADOS);
                string duplicados = Path.Combine(carpeta, App.FolderNames.INFORMES_DUPLICADOS);
                Directory.CreateDirectory(verificados);
                Directory.CreateDirectory(duplicados);
                File.WriteAllText(Path.Combine(verificados, "126-12345678-INF-JUANCARLOSPE.pdf"), "x");
                File.WriteAllText(Path.Combine(duplicados, "126-12345678-INF-JUANCARLOSPE.pdf"), "x");

                var resultado = _sut.RenombrarYUnificar(carpeta, 0, 0, true, false);

                resultado.PdfsExaminados.Should().Be(0);
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }
    }
}
