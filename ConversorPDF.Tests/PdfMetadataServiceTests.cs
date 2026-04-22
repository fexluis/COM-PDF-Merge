using Xunit;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Enums;
using System;
using System.IO;

namespace ConversorPDF.Tests
{
    public class PdfMetadataServiceTests : IDisposable
    {
        private readonly string _testPdfPath;
        private readonly PdfMetadataService _sut;

        public PdfMetadataServiceTests()
        {
            _testPdfPath = Path.Combine(Path.GetTempPath(), $"test_metadata_{Guid.NewGuid()}.pdf");
            _sut = new PdfMetadataService();
        }

        public void Dispose()
        {
            if (File.Exists(_testPdfPath))
            {
                try { File.Delete(_testPdfPath); } catch { }
            }
        }

        [Fact]
        public void Save_And_Read_Metadata_Should_Match()
        {
            // Arrange
            var expected = new Metadata
            {
                Titulo = "Test Document",
                Autor = "Antigravity Agent",
                Asunto = "Unit Testing",
                PalabrasClave = "PDF, Metadata, XLL",
                Fecha = new DateTime(2026, 4, 4),
                MetaDocumento = TipoMetadata.VIA_SOLICITUD,
                CDP = "CDP-123",
                RP = "RP-456",
                OB = "OB-789",
                OP = "OP-012",
                Tipo = "VIATICOS",
                Comision = "COM-2026"
            };

            // Act
            _sut.GuardarMetadatos(_testPdfPath, expected);
            var actual = _sut.LeerMetadatos(_testPdfPath);

            // Assert
            Assert.Equal(expected.Titulo, actual.Titulo);
            Assert.Equal(expected.Autor, actual.Autor);
            Assert.Equal(expected.Asunto, actual.Asunto);
            Assert.Equal(expected.PalabrasClave, actual.PalabrasClave);
            Assert.Equal(expected.Fecha, actual.Fecha);
            Assert.Equal(expected.MetaDocumento, actual.MetaDocumento);
            Assert.Equal(expected.CDP, actual.CDP);
            Assert.Equal(expected.RP, actual.RP);
            Assert.Equal(expected.OB, actual.OB);
            Assert.Equal(expected.OP, actual.OP);
            Assert.Equal(expected.Tipo, actual.Tipo);
            Assert.Equal(expected.Comision, actual.Comision);
        }

        [Fact]
        public void GetMetadataSummary_Should_Return_Formatted_String()
        {
            // Arrange
            var metadata = new Metadata
            {
                Titulo = "Summary Test",
                CDP = "CDP-999"
            };
            _sut.GuardarMetadatos(_testPdfPath, metadata);

            // Act
            var summary = _sut.ObtenerResumenMetadatos(_testPdfPath);

            // Assert
            Assert.NotNull(summary);
            Assert.Contains("METADATOS DEL DOCUMENTO", summary);
            Assert.Contains("Summary Test", summary);
            Assert.Contains("CDP-999", summary);
        }

        [Fact]
        public void LeerMetadatos_WhenFileNotExists_ShouldThrowFileNotFoundException()
        {
            // Arrange
            string nonExistentPath = "not_found.pdf";

            // Act & Assert
            Assert.Throws<FileNotFoundException>(() => _sut.LeerMetadatos(nonExistentPath));
        }
    }
}
