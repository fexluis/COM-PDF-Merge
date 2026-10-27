using Xunit;
using FluentAssertions;
using System;
using System.IO;
using System.Security;
using System.Data.SQLite;
using Dapper;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence.Integration
{
    public class EncryptedSQLiteContextTests : IDisposable
    {
        private readonly string _tempDir;
        private readonly SecureString _password;

        public EncryptedSQLiteContextTests()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "ConversorPDF_Tests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
            _password = new SecureString();
            foreach (var c in "testpass2026") _password.AppendChar(c);
        }

        private static byte[] CreateMinimalSQLiteDb()
        {
            string tmpPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".db");
            try
            {
                SQLiteConnection.CreateFile(tmpPath);
                using (var conn = new SQLiteConnection($"Data Source={tmpPath};Version=3;"))
                {
                    conn.Open();
                    conn.Execute("CREATE TABLE sesiones (id INTEGER PRIMARY KEY, horas REAL)");
                    conn.Execute("INSERT INTO sesiones VALUES (1, 8.0)");
                }
                return File.ReadAllBytes(tmpPath);
            }
            finally
            {
                if (File.Exists(tmpPath)) File.Delete(tmpPath);
            }
        }

        [Fact]
        public void ReEncryptAndPersistStatic_VaultEscrito_PuedeDesencriptarseConHeaderValido()
        {
            string vaultPath = Path.Combine(_tempDir, "test.vault");
            byte[] plainDb = CreateMinimalSQLiteDb();

            EncryptedSQLiteContext.ReEncryptAndPersistStatic(plainDb, vaultPath, _password);

            File.Exists(vaultPath).Should().BeTrue();
            
            byte[] decrypted = DatabaseEncryptionEngine.Decrypt(File.ReadAllBytes(vaultPath), _password);
            ((char)decrypted[0]).Should().Be('S');
            ((char)decrypted[1]).Should().Be('Q');
            ((char)decrypted[2]).Should().Be('L');
            ((char)decrypted[3]).Should().Be('i');
        }

        [Fact]
        public void LeerBytesConRetrySync_ArchivoPresente_RetornaBytes()
        {
            string tempFile = Path.Combine(_tempDir, "readable.db");
            byte[] expected = { 0x53, 0x51, 0x4C, 0x69 };
            File.WriteAllBytes(tempFile, expected);

            byte[] result = EncryptedSQLiteContext.LeerBytesConRetrySync(tempFile);

            result.Should().NotBeNull();
            result.Should().Equal(expected);
        }

        public void Dispose()
        {
            _password?.Dispose();
            try { if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, true); } catch { }
        }
    }
}
