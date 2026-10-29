using System;
using System.IO;
using System.Security;
using System.Data.SQLite;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Integration.Usuarios
{
    public class MoverUsuarioEntreVaultsE2ETests : IDisposable
    {
        private readonly string _tempDir;
        private readonly string _origenVaultPath;
        private readonly string _destinoVaultPath;
        private readonly SecureString _password;

        public MoverUsuarioEntreVaultsE2ETests()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "MUV_E2E_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);

            _origenVaultPath = Path.Combine(_tempDir, "AcademicData_ANTIOQUIA.db.enc");
            _destinoVaultPath = Path.Combine(_tempDir, "AcademicData_VALLE.db.enc");

            _password = new SecureString();
            foreach (var c in "TestPassword123!")
            {
                _password.AppendChar(c);
            }
            _password.MakeReadOnly();

            CreateInitialVault(_origenVaultPath, "ANTIOQUIA");
            CreateInitialVault(_destinoVaultPath, "VALLE");
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_tempDir))
                {
                    Directory.Delete(_tempDir, true);
                }
            }
            catch
            {
                // Ignorar locks en cleanup de pruebas
            }
        }

        private void CreateInitialVault(string vaultPath, string subunidad)
        {
            string tmpPlain = Path.Combine(_tempDir, $"plain_{subunidad}.db");
            try
            {
                SQLiteConnection.CreateFile(tmpPlain);
                using (var conn = new SQLiteConnection($"Data Source={tmpPlain};Version=3;"))
                {
                    conn.Open();
                    conn.Execute(@"
                        CREATE TABLE usuarios (
                            id TEXT PRIMARY KEY,
                            nombre TEXT NOT NULL,
                            nombreusuario TEXT NOT NULL,
                            subunidad TEXT NOT NULL,
                            perfil TEXT DEFAULT 'Admin',
                            activo INTEGER DEFAULT 1,
                            firma TEXT DEFAULT NULL,
                            created_at TEXT DEFAULT (datetime('now')),
                            updated_at TEXT DEFAULT (datetime('now'))
                        );

                        CREATE TABLE config (
                            id INTEGER PRIMARY KEY,
                            subunidad TEXT NOT NULL
                        );

                        INSERT INTO config (id, subunidad) VALUES (1, @subunidad);
                    ", new { subunidad });
                }

                byte[] plainBytes = File.ReadAllBytes(tmpPlain);
                EncryptedSQLiteContext.ReEncryptAndPersistStatic(plainBytes, vaultPath, _password);
            }
            finally
            {
                if (File.Exists(tmpPlain))
                {
                    File.Delete(tmpPlain);
                }
            }
        }

        [Fact]
        public void Mover_WhenCambiaSubunidad_ShouldPersistirFilaEnVaultDestinoYReabrirla()
        {
            // Arrange
            var usuario = new Usuario
            {
                Id = "usr_sofia",
                Nombre = "Sofia Ramirez",
                NombreUsuario = "sofia",
                Subunidad = "ANTIOQUIA",
                Perfil = "Academica",
                Activo = 1,
                Firma = "firma_base64_ejemplo"
            };

            var sut = new UsuarioVaultMover(
                resolvePath: sub => Path.Combine(_tempDir, $"AcademicData_{sub.ToUpperInvariant().Trim()}.db.enc"),
                openVault: p =>
                {
                    var ctx = new EncryptedSQLiteContext(p, _password);
                    ctx.Open();
                    return ctx;
                },
                persist: (bytes, p) => EncryptedSQLiteContext.ReEncryptAndPersistStatic(bytes, p, _password)
            );

            // Act
            var result = sut.Mover(usuario, "VALLE");

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Insertado.Should().BeTrue();
            result.Value.SubunidadDestino.Should().Be("VALLE");

            // Reabrir el vault destino de forma independiente para certificar la persistencia real
            using (var ctx = new EncryptedSQLiteContext(_destinoVaultPath, _password))
            {
                var conn = ctx.Open();
                var row = conn.QuerySingle(@"
                    SELECT id, nombre, nombreusuario, subunidad, perfil, activo, firma
                    FROM usuarios
                    WHERE id = 'usr_sofia';
                ");

                ((string)row.id).Should().Be("usr_sofia");
                ((string)row.nombre).Should().Be("Sofia Ramirez");
                ((string)row.nombreusuario).Should().Be("sofia");
                ((string)row.subunidad).Should().Be("VALLE");
                ((string)row.perfil).Should().Be("Academica");
                ((int)row.activo).Should().Be(1);
                ((string)row.firma).Should().Be("firma_base64_ejemplo");
            }
        }

        [Fact]
        public void Mover_WhenRutaDestinoIgualAOrigen_ShouldSerNoOpSinPerdida()
        {
            // Arrange
            var usuario = new Usuario
            {
                Id = "usr_sofia",
                Nombre = "Sofia Ramirez",
                NombreUsuario = "sofia",
                Subunidad = "ANTIOQUIA",
                Perfil = "Academica",
                Activo = 1
            };

            var sut = new UsuarioVaultMover(
                resolvePath: sub => _origenVaultPath,
                openVault: p =>
                {
                    var ctx = new EncryptedSQLiteContext(p, _password);
                    ctx.Open();
                    return ctx;
                },
                persist: (bytes, p) => EncryptedSQLiteContext.ReEncryptAndPersistStatic(bytes, p, _password)
            );

            // Act
            var result = sut.Mover(usuario, "ANTIOQUIA");

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Insertado.Should().BeFalse();
            result.Value.Actualizado.Should().BeFalse();
        }
    }
}
