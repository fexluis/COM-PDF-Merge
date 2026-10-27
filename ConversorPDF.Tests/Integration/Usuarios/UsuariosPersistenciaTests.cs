using System;
using System.IO;
using System.Security;
using System.Data.SQLite;
using Dapper;
using FluentAssertions;
using Xunit;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.Tests.Integration.Usuarios
{
    public class UsuariosPersistenciaTests : IDisposable
    {
        private readonly string _tempDir;
        private readonly string _vaultPath;
        private readonly SecureString _password;

        public UsuariosPersistenciaTests()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "ConversorPDF_PersistTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
            _vaultPath = Path.Combine(_tempDir, "AcademicData_TEST.db.enc");

            _password = new SecureString();
            foreach (var c in "TestPassword123!")
            {
                _password.AppendChar(c);
            }

            CreateInitialVault();
        }

        private void CreateInitialVault()
        {
            string tmpPlain = Path.Combine(_tempDir, "plain_init.db");
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
                            subunidad TEXT NOT NULL,
                            id_territorial INTEGER DEFAULT 1,
                            id_periodo INTEGER DEFAULT 1,
                            fecha_inicio TEXT,
                            fecha_fin TEXT,
                            cdp TEXT,
                            salario_decreto_613_2025 REAL,
                            viatico_diario_tarifa_a REAL,
                            viatico_diario_tarifa_b REAL,
                            porcentaje_decreto_a REAL,
                            porcentaje_decreto_b REAL,
                            dt_nombre TEXT,
                            dt_cargo TEXT,
                            corporativo_cedula TEXT,
                            corporativo_nombre TEXT,
                            corporativo_cargo TEXT,
                            departamento_base TEXT,
                            ciudad_base TEXT,
                            aeropuerto_base TEXT,
                            created_at TEXT DEFAULT (datetime('now')),
                            updated_at TEXT DEFAULT (datetime('now'))
                        );

                        INSERT INTO config (id, subunidad) VALUES (1, 'ANTIOQUIA');
                        INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, activo)
                        VALUES ('usr_init', 'Test User', 'testuser', 'ANTIOQUIA', 1);
                    ");
                }

                byte[] plainBytes = File.ReadAllBytes(tmpPlain);
                EncryptedSQLiteContext.ReEncryptAndPersistStatic(plainBytes, _vaultPath, _password);
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
        public void Persist_WhenSubunidadChanged_ShouldPersistVaultAndSurviveReload()
        {
            // Arrange: Abrir sesión sobre el vault
            using (var ctx = new EncryptedSQLiteContext(_vaultPath, _password))
            using (var session = new AcademicDbSession(ctx))
            {
                var conn = session.Connection;
                conn.Execute("UPDATE config SET subunidad = 'VALLE' WHERE id = 1;");
                conn.Execute("UPDATE usuarios SET subunidad = 'VALLE' WHERE nombreusuario = 'testuser';");

                // Act: Llamar a Persist() en la sesión
                var persistResult = session.Persist();
                persistResult.IsSuccess.Should().BeTrue("la persistencia del vault en disco debe completarse con éxito");
            }

            // Assert: Reabrir el vault desde disco con un nuevo contexto para simular cerrar y abrir Excel
            using (var reloadedCtx = new EncryptedSQLiteContext(_vaultPath, _password))
            {
                var reloadedConn = reloadedCtx.Open();
                string subConfig = reloadedConn.QuerySingle<string>("SELECT subunidad FROM config WHERE id = 1;");
                string subUsuario = reloadedConn.QuerySingle<string>("SELECT subunidad FROM usuarios WHERE nombreusuario = 'testuser';");

                subConfig.Should().Be("VALLE");
                subUsuario.Should().Be("VALLE");
            }
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
            catch { }
        }
    }
}
