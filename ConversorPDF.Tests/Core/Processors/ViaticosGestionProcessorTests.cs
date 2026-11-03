using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Repositories;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Core.Processors
{
    public class ViaticosGestionProcessorTests
    {
        private readonly Mock<IExcelRepository> _mockExcel;
        private readonly Mock<IViaticosRepository> _mockViaticos;
        private readonly Mock<IPdfToTexto> _mockPdfToTexto;
        private readonly Mock<IPdfMetadataService> _mockMetadata;
        private readonly ViaticosGestionProcessor _sut;

        public ViaticosGestionProcessorTests()
        {
            _mockExcel = new Mock<IExcelRepository>();
            _mockViaticos = new Mock<IViaticosRepository>();
            _mockPdfToTexto = new Mock<IPdfToTexto>();
            _mockMetadata = new Mock<IPdfMetadataService>();

            _sut = new ViaticosGestionProcessor(
                _mockExcel.Object,
                _mockViaticos.Object,
                _mockPdfToTexto.Object,
                _mockMetadata.Object);
        }

        [Fact]
        public void GetSolicitudes_WhenNombreLargoConEspacios_ShouldLimpiarSinLanzarExcepcionYTruncarA18()
        {
            // Arrange
            // Matriz 1-based (Excel interop simulation): 2 filas (fila 1 encabezado, fila 2 datos), 35 columnas
            var data = Array.CreateInstance(typeof(object), new int[] { 2, 35 }, new int[] { 1, 1 }) as object[,];
            data[2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "126";
            data[2, ConstantsHelper.CONSULTAS_CEDULA] = "12345678";

            // Cadena original de 219 caracteres con múltiples espacios que antes causaba ArgumentOutOfRangeException
            string nombreLargo = "JUAN CARLOS DE LOS SANTOS PEREZ RODRIGUEZ HERNANDEZ MARTINEZ GOMEZ GONZALEZ LOPEZ DIAZ CASTILLO " +
                                 "MORALES SANCHEZ ALVAREZ ROMERO CONTRERAS VARGAS TORRES RAMIREZ FLORES CASTRO ORTIZ SILVA " +
                                 "NUNEZ MEDINA DEL VALLE MONDRAGON Y ASOCIADOS";
            data[2, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = nombreLargo;

            _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS)).Returns(data);

            // Act
            var solicitudes = _sut.GetSolicitudes(0, 0);

            // Assert
            solicitudes.Should().HaveCount(1);
            solicitudes[0].Nom.Length.Should().BeLessOrEqualTo(18);
            solicitudes[0].Nom.Should().NotContain(" ");
        }

        [Fact]
        public void GetSolicitudes_WhenFiltroRangoActivo_ShouldRetornarSoloCoincidentes()
        {
            // Arrange
            // 6 filas (1 encabezado + 5 datos: 126, 4526, 14526, 14626, 12625)
            var data = Array.CreateInstance(typeof(object), new int[] { 6, 35 }, new int[] { 1, 1 }) as object[,];
            
            data[2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "126";
            data[2, ConstantsHelper.CONSULTAS_CEDULA] = "1001";
            data[2, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = "Funcionario Uno";

            data[3, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "4526";
            data[3, ConstantsHelper.CONSULTAS_CEDULA] = "1002";
            data[3, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = "Funcionario Dos";

            data[4, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "14526";
            data[4, ConstantsHelper.CONSULTAS_CEDULA] = "1003";
            data[4, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = "Funcionario Tres";

            data[5, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "14626"; // Fuera de rango
            data[5, ConstantsHelper.CONSULTAS_CEDULA] = "1004";
            data[5, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = "Funcionario Cuatro";

            data[6, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "12625"; // Año 2025 (fuera de rango)
            data[6, ConstantsHelper.CONSULTAS_CEDULA] = "1005";
            data[6, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = "Funcionario Cinco";

            _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS)).Returns(data);

            // Act: Rango 126 a 14526 (inclusivo en 2026)
            var solicitudes = _sut.GetSolicitudes(126, 14526);

            // Assert
            solicitudes.Should().HaveCount(3);
            solicitudes.Select(s => s.Num).Should().Contain(new[] { "126", "4526", "14526" });
            solicitudes.Select(s => s.Num).Should().NotContain("14626");
            solicitudes.Select(s => s.Num).Should().NotContain("12625");
        }

        [Fact]
        public void GetSolicitudes_WhenSinFiltro_ShouldRetornarTodasLasFilasValidas()
        {
            // Arrange
            var data = Array.CreateInstance(typeof(object), new int[] { 4, 35 }, new int[] { 1, 1 }) as object[,];
            data[2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "126";
            data[2, ConstantsHelper.CONSULTAS_CEDULA] = "1001";
            data[2, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = "Funcionario Uno";

            data[3, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "12625";
            data[3, ConstantsHelper.CONSULTAS_CEDULA] = "1002";
            data[3, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = "Funcionario Dos";

            data[4, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = ""; // Fila vacía
            data[4, ConstantsHelper.CONSULTAS_CEDULA] = "1003";
            data[4, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = "Funcionario Tres";

            _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS)).Returns(data);

            // Act: (0, 0) sin filtro
            var solicitudes = _sut.GetSolicitudes(0, 0);

            // Assert
            solicitudes.Should().HaveCount(2);
            solicitudes.Select(s => s.Num).Should().Contain(new[] { "126", "12625" });
        }

        private static object[,] CrearConsultas(params (string Num, string Cc, string Nom)[] filas)
        {
            var data = Array.CreateInstance(typeof(object), new int[] { filas.Length + 1, 35 }, new int[] { 1, 1 }) as object[,];
            for (int i = 0; i < filas.Length; i++)
            {
                data[i + 2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = filas[i].Num;
                data[i + 2, ConstantsHelper.CONSULTAS_CEDULA] = filas[i].Cc;
                data[i + 2, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = filas[i].Nom;
            }
            return data;
        }

        private static string CrearCarpetaTemporal()
        {
            string root = Path.Combine(Path.GetTempPath(), "ConversorPDF.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            return root;
        }

        [Fact]
        public void RenombrarSolicitudes_WhenPdfEnSubcarpetaSolicitudes_ShouldRenombrar()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "17526-Obert-Alejandro-OKEY.pdf"), "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(("17526", "79633812", "OBERT ALEJANDRO ORTIZ")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.HayError.Should().BeFalse();
                r.PdfsExaminados.Should().Be(1);
                r.Renombrados.Should().Be(1);
                File.Exists(Path.Combine(pSoli, "17526-79633812-ObertAlejandroOrti.pdf")).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenComisionDuplicadaConMismaCedula_ShouldRenombrarSinOcr()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "21626-VIATICOS -YENY RAYO.pdf"), "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(
                              ("21626", "52371279", "YENY PATRICIA RAYO PRADA"),
                              ("21626", "52371279", "YENY PATRICIA RAYO PRADA")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Renombrados.Should().Be(1);
                _mockPdfToTexto.Verify(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()), Times.Never());
                File.Exists(Path.Combine(pSoli, "21626-52371279-YenyPatriciaRayoPr.pdf")).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenComisionDuplicadaConCedulasDistintas_ShouldUsarOcr()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "14526-Solicitud.pdf"), "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(
                              ("14526", "12345678", "JUAN PEREZ"),
                              ("14526", "87654321", "ANA GOMEZ")));
                _mockPdfToTexto.Setup(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()))
                               .Returns("C.C. 12.345.678  juan.perez@esap.edu.co  Rubro C-0505-1000-4-20306D-0505023-02");

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Renombrados.Should().Be(1);
                _mockPdfToTexto.Verify(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()), Times.AtLeastOnce());
                File.Exists(Path.Combine(pSoli, "14526-12345678-JuanPerez.pdf")).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenOcrDevuelveNumeroAjeno_ShouldMarcarAmbiguo()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                string origen = Path.Combine(pSoli, "14526-Solicitud.pdf");
                File.WriteAllText(origen, "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(
                              ("14526", "12345678", "JUAN PEREZ"),
                              ("14526", "87654321", "ANA GOMEZ")));
                _mockPdfToTexto.Setup(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()))
                               .Returns("Rubro C-0505-1000-4-20306D-0505023-02");

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Renombrados.Should().Be(0);
                r.Ambiguos.Should().Be(1);
                File.Exists(origen).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenDestinoYaExiste_ShouldMarcarColisionYNoSobrescribir()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                string origen = Path.Combine(pSoli, "14526-juan.pdf");
                string destino = Path.Combine(pSoli, "14526-12345678-JuanPerez.pdf");
                File.WriteAllText(origen, "origen");
                File.WriteAllText(destino, "destino");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(("14526", "12345678", "JUAN PEREZ")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Colisiones.Should().Be(1);
                r.Renombrados.Should().Be(0);
                File.ReadAllText(destino).Should().Be("destino");
                File.Exists(origen).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenNombreSinPrefijoNumerico_ShouldMarcarPrefijoInvalido()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "ADOLFO-LEON-GUERRERO.pdf"), "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(("14526", "12345678", "JUAN PEREZ")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.PrefijoInvalido.Should().Be(1);
                r.Renombrados.Should().Be(0);
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenCarpetaNoExiste_ShouldRetornarError()
        {
            string inexistente = Path.Combine(Path.GetTempPath(), "ConversorPDF.Tests", Guid.NewGuid().ToString("N"));

            var r = _sut.RenombrarSolicitudesConDetalle(inexistente, 0, 0);

            r.HayError.Should().BeTrue();
            _sut.RenombrarSolicitudes(inexistente, 0, 0).Should().Be(-1);
        }

        private static object[,] CrearConsultasDetalle(params (string Num, string Cc, string Nom, string Estado, string Informe)[] filas)
        {
            var data = Array.CreateInstance(typeof(object), new int[] { filas.Length + 1, 35 }, new int[] { 1, 1 }) as object[,];
            for (int i = 0; i < filas.Length; i++)
            {
                data[i + 2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = filas[i].Num;
                data[i + 2, ConstantsHelper.CONSULTAS_CEDULA] = filas[i].Cc;
                data[i + 2, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = filas[i].Nom;
                data[i + 2, ConstantsHelper.CONSULTAS_ESTADO] = filas[i].Estado;
                data[i + 2, ConstantsHelper.CONSULTAS_INFORME] = filas[i].Informe;
            }
            return data;
        }

        [Fact]
        public void RenombrarSolicitudes_WhenDestinoIdentico_ShouldMoverADuplicados()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                string destino = Path.Combine(pSoli, "14526-12345678-JuanPerez.pdf");
                string origen = Path.Combine(pSoli, "14526-juan-copia.pdf");
                File.WriteAllText(destino, "contenido-identico");
                File.WriteAllText(origen, "contenido-identico");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(("14526", "12345678", "JUAN PEREZ")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Duplicados.Should().Be(1);
                r.Colisiones.Should().Be(0);
                File.Exists(origen).Should().BeFalse();
                File.Exists(Path.Combine(pSoli, App.FolderNames.DUPLICADOS, "14526-juan-copia.pdf")).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenDestinoDistinto_ShouldMarcarColisionSinMover()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                string destino = Path.Combine(pSoli, "14526-12345678-JuanPerez.pdf");
                string origen = Path.Combine(pSoli, "14526-juan-copia.pdf");
                File.WriteAllText(destino, "AAAA");
                File.WriteAllText(origen, "BBBB");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(("14526", "12345678", "JUAN PEREZ")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Colisiones.Should().Be(1);
                r.Duplicados.Should().Be(0);
                File.Exists(origen).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenPrefijoDe6DigitosCoincideNombre_ShouldRemapear()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "202626-GERMAN-LUIS-ROLDAN-N-FILA-153.pdf"), "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(("20626", "16796878", "GERMAN LUIS ROLDAN NOGUERA")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Renombrados.Should().Be(1);
                File.Exists(Path.Combine(pSoli, "20626-16796878-GermanLuisRoldanNo.pdf")).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenPrefijoDe6DigitosNoCoincide_ShouldReportarSinComision()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "202426-OTRA-PERSONA-FILA-001.pdf"), "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(("20626", "16796878", "GERMAN LUIS ROLDAN NOGUERA")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.SinComision.Should().Be(1);
                r.Renombrados.Should().Be(0);
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void ObtenerRangoSolicitudes_ShouldDevolverPrimerYUltimoRegistro()
        {
            var data = Array.CreateInstance(typeof(object), new int[] { 5, 35 }, new int[] { 1, 1 }) as object[,];
            data[2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "126";
            data[3, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "";
            data[4, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "41326";
            data[5, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = "20626";
            _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS)).Returns(data);

            var rango = _sut.ObtenerRangoSolicitudes();

            rango.HayDatos.Should().BeTrue();
            rango.Inicio.Should().Be(126);
            rango.Fin.Should().Be(20626);
        }

        [Fact]
        public void VerificarSolicitudes_WhenSinArchivo_ShouldMarcarSinSolicitudYGuardar()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                Directory.CreateDirectory(Path.Combine(root, App.FolderNames.SOLICITUDES));
                var hoja = new object();
                _mockExcel.Setup(x => x.GetWorksheet(ConstantsHelper.SHEET_CONSULTAS)).Returns(hoja);
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultasDetalle(("14526", "12345678", "JUAN PEREZ", "LEGALIZADA", "")));

                var r = _sut.VerificarSolicitudesPresentes(root, 0, 0);

                r.HayError.Should().BeFalse();
                r.CeldasEscritas.Should().Be(1);
                r.SinArchivo.Should().Be(1);
                r.LibroGuardado.Should().BeTrue();
                _mockExcel.Verify(x => x.SetCellValue(hoja, "H2", "SIN SOLICITUD"), Times.Once);
                _mockExcel.Verify(x => x.SaveActiveWorkbook(), Times.Once);
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void VerificarSolicitudes_WhenExisteArchivoYMarcaInexistente_ShouldLimpiarCelda()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "14526-12345678-JuanPerez.pdf"), "x");
                var hoja = new object();
                _mockExcel.Setup(x => x.GetWorksheet(ConstantsHelper.SHEET_CONSULTAS)).Returns(hoja);
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultasDetalle(("14526", "12345678", "JUAN PEREZ", "LEGALIZADA", "SIN SOLICITUD")));

                var r = _sut.VerificarSolicitudesPresentes(root, 0, 0);

                r.CeldasLimpiadas.Should().Be(1);
                r.CeldasEscritas.Should().Be(0);
                _mockExcel.Verify(x => x.SetCellValue(hoja, "H2", string.Empty), Times.Once);
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void VerificarSolicitudes_WhenVerificadoOReintegro_ShouldNoEscribirNiConsultarDisco()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                var hoja = new object();
                _mockExcel.Setup(x => x.GetWorksheet(ConstantsHelper.SHEET_CONSULTAS)).Returns(hoja);
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultasDetalle(
                              ("14526", "12345678", "JUAN PEREZ", "LEGALIZADA", "VERIFICADO"),
                              ("14527", "87654321", "ANA GOMEZ", "LEGALIZADA", "REINTEGRO")));

                var r = _sut.VerificarSolicitudesPresentes(root, 0, 0);

                r.OmitidasProtegidas.Should().Be(2);
                r.CeldasEscritas.Should().Be(0);
                r.CeldasLimpiadas.Should().Be(0);
                r.LibroGuardado.Should().BeFalse();
                _mockExcel.Verify(x => x.SetCellValue(It.IsAny<object>(), It.IsAny<string>(), It.IsAny<object>()), Times.Never);
                _mockExcel.Verify(x => x.SaveActiveWorkbook(), Times.Never);
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void VerificarSolicitudes_WhenEstadoTerminal_ShouldCopiarEstadoSinConsultarDisco()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                var hoja = new object();
                _mockExcel.Setup(x => x.GetWorksheet(ConstantsHelper.SHEET_CONSULTAS)).Returns(hoja);
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultasDetalle(
                              ("14526", "12345678", "JUAN PEREZ", "REVOCADA", ""),
                              ("14527", "87654321", "ANA GOMEZ", "ANULADA", "")));

                var r = _sut.VerificarSolicitudesPresentes(root, 0, 0);

                r.CopiadasEstado.Should().Be(2);
                _mockExcel.Verify(x => x.SetCellValue(hoja, "H2", "REVOCADA"), Times.Once);
                _mockExcel.Verify(x => x.SetCellValue(hoja, "H3", "ANULADA"), Times.Once);
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void VerificarSolicitudes_WhenOtraMarcaOperativa_ShouldPreservarla()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                var hoja = new object();
                _mockExcel.Setup(x => x.GetWorksheet(ConstantsHelper.SHEET_CONSULTAS)).Returns(hoja);
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultasDetalle(("14526", "12345678", "JUAN PEREZ", "LEGALIZADA", "ERROR")));

                var r = _sut.VerificarSolicitudesPresentes(root, 0, 0);

                r.OmitidasOtrasMarcas.Should().Be(1);
                r.CeldasEscritas.Should().Be(0);
                r.LibroGuardado.Should().BeFalse();
                _mockExcel.Verify(x => x.SaveActiveWorkbook(), Times.Never);
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void VerificarSolicitudes_WhenYaMarcada_ShouldNoReescribirNiGuardar()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                var hoja = new object();
                _mockExcel.Setup(x => x.GetWorksheet(ConstantsHelper.SHEET_CONSULTAS)).Returns(hoja);
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultasDetalle(("14526", "12345678", "JUAN PEREZ", "LEGALIZADA", "SIN SOLICITUD")));

                var r = _sut.VerificarSolicitudesPresentes(root, 0, 0);

                r.YaMarcadas.Should().Be(1);
                r.CeldasEscritas.Should().Be(0);
                r.LibroGuardado.Should().BeFalse();
            }
            finally { Directory.Delete(root, true); }
        }
    }
}

