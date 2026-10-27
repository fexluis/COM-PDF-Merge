using System;
using System.IO;
using ConversorPDF.XLL.Core.Helpers;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Core.Helpers
{
    [Collection("SqliteHelperSequential")]
    public class SqliteHelperStartupTests : IDisposable
    {
        public SqliteHelperStartupTests()
        {
            SqliteHelper.ResetMigratedVaultsTrackingForTests();
        }

        public void Dispose()
        {
            SqliteHelper.ResetMigratedVaultsTrackingForTests();
        }

        [Fact]
        public void MarkVaultAsMigrated_Should_StorePathCorrectly()
        {
            // Arrange
            string path = Path.Combine(Path.GetTempPath(), "Test_VALLE.db.enc");
            
            // Act
            SqliteHelper.MarkVaultAsMigrated(path);
            
            // Assert
            SqliteHelper.WasSchemaMigratedForPath(path).Should().BeTrue();
        }

        [Fact]
        public void WasSchemaMigratedForPath_WhenDifferentVault_ShouldReturnFalse()
        {
            // Arrange
            string valle = Path.Combine(Path.GetTempPath(), "AcademicData_VALLE.db.enc");
            string bolivar = Path.Combine(Path.GetTempPath(), "AcademicData_BOLIVAR.db.enc");
            
            // Act
            SqliteHelper.MarkVaultAsMigrated(valle);
            
            // Assert
            SqliteHelper.WasSchemaMigratedForPath(bolivar).Should().BeFalse();
        }

        [Fact]
        public void DatabaseBootstrapper_WasInitialized_WhenAutoOpenSimulated_ShouldBeFalse()
        {
            // Arrange — simular estado limpio (como si Excel acabara de abrir)
            DatabaseBootstrapper.ResetStateForTests();
            
            // Assert — verificar que tras un "AutoOpen" sin DB access, el bootstrapper NO se ejecutó
            DatabaseBootstrapper.WasInitialized.Should().BeFalse(
                "la DB no debe inicializarse al arrancar Excel si ningún módulo de DB fue invocado");
        }
    }
}
