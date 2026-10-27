using System;
using System.IO;
using System.Security;
using System.Data.SQLite;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Sync;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Core.Sync
{
    public class SyncInitializerIntegrationTests : IDisposable
    {
        private readonly string _tempDir;
        private readonly string _vaultPath;
        private readonly SecureString _password;

        public SyncInitializerIntegrationTests()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "ConversorPDF_SyncTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
            _vaultPath = Path.Combine(_tempDir, $"test_{Guid.NewGuid():N}.vault");
            _password = new SecureString();
            foreach (var c in ConversorPDF.XLL.Core.Helpers.ConstantsHelper.DB_PASSWORD) _password.AppendChar(c);
            CrSqliteLoader.Initialize();
            SyncInitializer.ResetStateForTests();
        }

        private void CreateMinimalVault()
        {
            using (var ctx = new EncryptedSQLiteContext(_vaultPath, _password))
            {
                var conn = ctx.Open();
                conn.Execute("CREATE TABLE terceros (id INTEGER PRIMARY KEY NOT NULL, nombre TEXT);");
                conn.Execute("INSERT INTO terceros VALUES (1, 'Test');");
                
                ctx.CloseConnectionResiliently();
                byte[] plainBytes = EncryptedSQLiteContext.LeerBytesConRetrySync(ctx.TempDbPath);
                EncryptedSQLiteContext.ReEncryptAndPersistStatic(plainBytes, _vaultPath, _password);
                ctx.MarkChangesPersisted();
            }
            SqliteHelper.SetTestVaultPath(_vaultPath);
            SqliteHelper.MarkVaultAsMigrated(_vaultPath); // FASE 55: Evitar que GetContextForPath intente aplicar las migraciones V1 a V32 en una BD mínima que no tiene las tablas esperadas.
        }

        [Fact]
        public void EnableAllCrdts_WhenNewSchema_ShouldCreateClock()
        {
            // Arrange
            CreateMinimalVault();

            // Act & Assert
            using (var ctx = SqliteHelper.GetContextForPath(_vaultPath))
            {
                var conn = ctx.Open();
                
                // primero verificar que NO existe reloj
                int before = conn.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE name='terceros__crsql_clock';");
                before.Should().Be(0);

                SyncInitializer.EnableAllCrdts();

                int after = conn.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE name='terceros__crsql_clock';");
                after.Should().Be(1);
            }
        }

        [Fact]
        public void EnableAllCrdts_WhenAlreadyCrdt_ShouldShortCircuitAndNotPersist()
        {
            // Arrange
            CreateMinimalVault();
            SyncInitializer.EnableAllCrdts(); // primera vez, se vuelve CRDT
            
            var before = new FileInfo(_vaultPath);
            long sizeBefore = before.Length;
            DateTime lastWriteBefore = before.LastWriteTimeUtc;

            // Act
            SyncInitializer.EnableAllCrdts(); // segunda vez, debería cortocircuitar

            // Assert
            var after = new FileInfo(_vaultPath);
            after.Length.Should().Be(sizeBefore);
            after.LastWriteTimeUtc.Should().Be(lastWriteBefore); // no se re-escribió
        }

        public void Dispose()
        {
            _password?.Dispose();
            try { if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, true); } catch { }
        }
    }
}
