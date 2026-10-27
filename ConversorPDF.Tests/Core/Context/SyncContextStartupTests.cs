using System;
using System.IO;
using ConversorPDF.XLL.Core.Sync;
using ConversorPDF.XLL.Core.Helpers;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Core.Context
{
    public class SyncContextStartupTests : IDisposable
    {
        private readonly string _activeVaultPath;
        private readonly string _originalContent;
        private readonly string _testVault;

        public SyncContextStartupTests()
        {
            _activeVaultPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ConversorPDF", "Data", ".active_vault");
            
            _originalContent = File.Exists(_activeVaultPath) ? File.ReadAllText(_activeVaultPath) : null;
            _testVault = Path.Combine(Path.GetTempPath(), $"AcademicData_TEST_{Guid.NewGuid():N}.db.enc");
            
            File.WriteAllText(_testVault, ""); // existe para que no falle validación
            Directory.CreateDirectory(Path.GetDirectoryName(_activeVaultPath));
            File.WriteAllText(_activeVaultPath, _testVault);
        }

        [Fact]
        public void GetActiveVaultPath_WhenActiveVaultFileExists_ShouldReturnCustomVaultPath()
        {
            // Act
            string result = SyncContext.GetActiveVaultPath();

            // Assert
            result.Should().Be(_testVault);
        }

        public void Dispose()
        {
            if (_originalContent != null) 
            {
                File.WriteAllText(_activeVaultPath, _originalContent);
            }
            else if (File.Exists(_activeVaultPath)) 
            {
                File.Delete(_activeVaultPath);
            }
            
            if (File.Exists(_testVault)) 
            {
                File.Delete(_testVault);
            }
        }
    }
}
