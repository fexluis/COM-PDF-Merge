using System;
using System.IO;
using System.Threading;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using Xunit;
using FluentAssertions;

namespace ConversorPDF.Tests.Integration.Academic
{
    [Collection("SQLite Integration")] // Serializar para evitar interferencia en archivos físicos
    public class EncryptedSQLiteContextTests : IDisposable
    {
        private readonly string _testId;

        public EncryptedSQLiteContextTests()
        {
            _testId = Guid.NewGuid().ToString("N").Substring(0, 8);
        }

        [Fact]
        [Trait("Category", "Hotfix_1_2_51_25")]
        [Trait("Priority", "P0")]
        [Trait("Type", "Integration")]
        public void CleanupOrphanedTempFiles_NoDebeBorrarArchivosRecientes()
        {
            // Arrange
            string tempFileName = $"esap_acad_tmp_unit_test_recent_{_testId}.db";
            string tempFilePath = Path.Combine(Path.GetTempPath(), tempFileName);
            
            File.WriteAllText(tempFilePath, "test content");
            File.SetLastWriteTime(tempFilePath, DateTime.Now.AddMinutes(-5));

            try
            {
                // Act: Limpieza con umbral de 2 horas
                EncryptedSQLiteContext.CleanupOrphanedTempFiles(TimeSpan.FromHours(2));

                // Assert
                File.Exists(tempFilePath).Should().BeTrue("El archivo reciente no debe ser borrado (Hotfix P1)");
            }
            finally
            {
                if (File.Exists(tempFilePath)) File.Delete(tempFilePath);
            }
        }

        [Fact]
        [Trait("Category", "Hotfix_1_2_51_25")]
        [Trait("Priority", "P0")]
        [Trait("Type", "Integration")]
        public void CleanupOrphanedTempFiles_DebeBorrarArchivosAntiguos()
        {
            // Arrange
            string tempFileName = $"esap_acad_tmp_unit_test_old_{_testId}.db";
            string tempFilePath = Path.Combine(Path.GetTempPath(), tempFileName);
            
            File.WriteAllText(tempFilePath, "test content");
            // Manipular fecha a hace 3 horas (umbral es 2h)
            File.SetLastWriteTime(tempFilePath, DateTime.Now.AddHours(-3));

            // Act
            EncryptedSQLiteContext.CleanupOrphanedTempFiles(TimeSpan.FromHours(2));

            // Assert
            File.Exists(tempFilePath).Should().BeFalse("El archivo antiguo debe ser borrado");
        }

        public void Dispose()
        {
            // No hay estado global que limpiar aparte de los archivos creados en cada test
        }
    }
}
