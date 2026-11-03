using System;
using System.IO;
using ConversorPDF.Tests.TestData;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Processors;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Core.Processors
{
    /// <summary>
    /// Pruebas de la auditoría de la columna INFORME (fase 2 del proceso "Buscar Informes").
    /// </summary>
    public class BusquedaInformesAuditoriaTests
    {
        private readonly Mock<IPdfToTexto> _pdfToTexto;
        private readonly Mock<IPdfService> _pdfService;
        private readonly ExcelRepositoryFalsa _excel;
        private readonly BusquedaInformesProcessor _sut;

        public BusquedaInformesAuditoriaTests()
        {
            _pdfToTexto = new Mock<IPdfToTexto>();
            _pdfService = new Mock<IPdfService>();
            _excel = new ExcelRepositoryFalsa();

            _sut = new BusquedaInformesProcessor(_excel, _pdfToTexto.Object, _pdfService.Object);
        }

        private object[,] ConfigurarConsultas(string informe = "")
        {
            var data = Array.CreateInstance(
                typeof(object),
                new[] { 2, ConstantsHelper.CONSULTAS_COL_MAX },
                new[] { 1, 1 }) as object[,];

            data[2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "126";
            data[2, ConstantsHelper.CONSULTAS_CEDULA] = "12345678";
            data[2, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = "JUAN CARLOS PEREZ";
            data[2, ConstantsHelper.CONSULTAS_INFORME] = informe;

            _excel.Consultas = data;
            _excel.Hoja = new HojaExcelFalsa(data);
            return data;
        }

        [Fact]
        public void AuditarConsultas_ConSoloInf_MarcaFaltaCum()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("auditar_falta_cum");
            try
            {
                string especial = Path.Combine(carpeta, "Direccion");
                Directory.CreateDirectory(especial);
                File.WriteAllText(Path.Combine(especial, "126-12345678-INF-JUANCARLOSPE.pdf"), "x");

                object[,] data = ConfigurarConsultas();

                var resultado = _sut.AuditarConsultas(carpeta, 0, 0, false);

                resultado.HayError.Should().BeFalse();
                data[2, ConstantsHelper.CONSULTAS_INFORME].Should().Be(EstadoComision.FaltaCum);
                resultado.FaltaCum.Should().Be(1);
                resultado.CeldasEscritas.Should().Be(1);
                resultado.LibroGuardado.Should().BeTrue();
                _excel.VecesGuardado.Should().Be(1);
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }

        [Fact]
        public void AuditarConsultas_SinArchivos_MarcaSinInforme()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("auditar_sin_informe");
            try
            {
                object[,] data = ConfigurarConsultas();

                var resultado = _sut.AuditarConsultas(carpeta, 0, 0, false);

                data[2, ConstantsHelper.CONSULTAS_INFORME].Should().Be(EstadoComision.SinInforme);
                resultado.SinInforme.Should().Be(1);
                resultado.CeldasEscritas.Should().Be(1);
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }

        [Fact]
        public void AuditarConsultas_EstadoProtegido_NoSobrescribe()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("auditar_protegido");
            try
            {
                object[,] data = ConfigurarConsultas(EstadoComision.Verificado);

                var resultado = _sut.AuditarConsultas(carpeta, 0, 0, false);

                data[2, ConstantsHelper.CONSULTAS_INFORME].Should().Be(EstadoComision.Verificado);
                resultado.OmitidasProtegidas.Should().Be(1);
                resultado.CeldasEscritas.Should().Be(0);
                _excel.VecesGuardado.Should().Be(0);
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }

        [Fact]
        public void AuditarConsultas_UniEnLaRaiz_MueveAVerificadosYMarcaVerificado()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("auditar_uni");
            try
            {
                string uni = Path.Combine(carpeta, "126-12345678-UNI-JUANCARLOSPE.pdf");
                File.WriteAllText(uni, "x");

                object[,] data = ConfigurarConsultas();

                var resultado = _sut.AuditarConsultas(carpeta, 0, 0, false);

                data[2, ConstantsHelper.CONSULTAS_INFORME].Should().Be(EstadoComision.Verificado);
                resultado.Verificados.Should().Be(1);
                File.Exists(uni).Should().BeFalse();
                File.Exists(Path.Combine(carpeta, App.FolderNames.INFORMES_VERIFICADOS, "126-12345678-UNI-JUANCARLOSPE.pdf"))
                      .Should().BeTrue();
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }

        [Fact]
        public void AuditarConsultas_EnSimulacion_NoEscribeNiMueveArchivos()
        {
            string carpeta = PdfFixtureFactory.CrearCarpetaTemporal("auditar_simulado");
            try
            {
                string especial = Path.Combine(carpeta, "Direccion");
                Directory.CreateDirectory(especial);
                File.WriteAllText(Path.Combine(especial, "126-12345678-INF-JUANCARLOSPE.pdf"), "x");

                object[,] data = ConfigurarConsultas();

                var resultado = _sut.AuditarConsultas(carpeta, 0, 0, true);

                resultado.Simulado.Should().BeTrue();
                data[2, ConstantsHelper.CONSULTAS_INFORME].Should().Be(string.Empty);
                resultado.CeldasEscritas.Should().Be(1);
                resultado.LibroGuardado.Should().BeFalse();
                _excel.VecesGuardado.Should().Be(0);
            }
            finally
            {
                PdfFixtureFactory.EliminarCarpeta(carpeta);
            }
        }
    }
}
