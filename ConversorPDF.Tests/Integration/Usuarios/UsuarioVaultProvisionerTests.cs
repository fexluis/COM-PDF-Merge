using System;
using System.Data.SQLite;
using System.Linq;
using Dapper;
using FluentAssertions;
using Xunit;
using ConversorPDF.XLL.Core.Sync;

namespace ConversorPDF.Tests.Integration.Usuarios
{
    /// <summary>
    /// Regresion del hallazgo H1/H4 de la auditoria 2026-10-27: al conmutar de vault por
    /// cambio de subunidad, la fila del usuario debe existir en el vault destino con la
    /// subunidad nueva y con identidad determinista (usr_&lt;nombreusuario&gt;).
    /// </summary>
    public class UsuarioVaultProvisionerTests : IDisposable
    {
        private readonly SQLiteConnection _connection;

        public UsuarioVaultProvisionerTests()
        {
            _connection = new SQLiteConnection("Data Source=:memory:;Version=3;");
            _connection.Open();

            _connection.Execute(@"
                CREATE TABLE usuarios (
                    id TEXT PRIMARY KEY DEFAULT (lower(hex(randomblob(16)))),
                    nombre TEXT NOT NULL,
                    nombreusuario TEXT NOT NULL,
                    subunidad TEXT NOT NULL,
                    perfil TEXT DEFAULT 'Admin',
                    activo INTEGER DEFAULT 1,
                    firma TEXT DEFAULT NULL,
                    created_at TEXT DEFAULT (datetime('now')),
                    updated_at TEXT DEFAULT (datetime('now'))
                );
            ");
        }

        [Fact]
        public void EnsureUsuarioEnVault_WhenDestinoTieneFilaSemillaConSubunidadVieja_ShouldActualizarSubunidad()
        {
            // Arrange: el vault destino ya trae la semilla con la subunidad anterior
            _connection.Execute(@"
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, activo)
                VALUES ('usr_lucs', 'Luis', 'lucs', 'VALLE', 1);
            ");

            // Act
            UsuarioVaultProvisioner.EnsureUsuarioEnVault(_connection, "lucs", "Luis", "ANTIOQUIA");

            // Assert
            _connection.QuerySingle<string>(
                "SELECT subunidad FROM usuarios WHERE LOWER(nombreusuario) = 'lucs';")
                .Should().Be("ANTIOQUIA", "el vault destino debe quedar con la subunidad resuelta");
        }

        [Fact]
        public void EnsureUsuarioEnVault_WhenExisteFilaConIdLegacy_ShouldNormalizarIdYDejarUnaSolaFila()
        {
            // Arrange: fila con GUID legacy (sobreviviente de un merge CRDT)
            _connection.Execute(@"
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, activo)
                VALUES ('9f2c1b7a4d', 'Luis', 'lucs', 'VALLE', 1);
            ");

            // Act
            UsuarioVaultProvisioner.EnsureUsuarioEnVault(_connection, "lucs", "Luis", "ANTIOQUIA");

            // Assert
            var rows = _connection.Query("SELECT id, subunidad FROM usuarios WHERE LOWER(nombreusuario) = 'lucs';").ToList();
            rows.Should().HaveCount(1, "no debe quedar duplicado del mismo usuario");
            ((string)rows[0].id).Should().Be("usr_lucs");
            ((string)rows[0].subunidad).Should().Be("ANTIOQUIA");
        }

        [Fact]
        public void EnsureUsuarioEnVault_WhenUsuarioNoExiste_ShouldInsertarConIdDeterminista()
        {
            // Act
            UsuarioVaultProvisioner.EnsureUsuarioEnVault(_connection, "Nuevo.User", "Nuevo Usuario", "BOLIVAR");

            // Assert
            _connection.QuerySingle<string>(
                "SELECT id FROM usuarios WHERE LOWER(nombreusuario) = 'nuevo.user';")
                .Should().Be("usr_nuevo.user");
            _connection.QuerySingle<string>(
                "SELECT subunidad FROM usuarios WHERE LOWER(nombreusuario) = 'nuevo.user';")
                .Should().Be("BOLIVAR");
        }

        [Fact]
        public void EnsureUsuarioEnVault_WhenSeInvocaDosVeces_ShouldSerIdempotente()
        {
            // Act
            UsuarioVaultProvisioner.EnsureUsuarioEnVault(_connection, "lucs", "Luis", "ANTIOQUIA");
            UsuarioVaultProvisioner.EnsureUsuarioEnVault(_connection, "lucs", "Luis", "ANTIOQUIA");

            // Assert
            _connection.QuerySingle<int>(
                "SELECT COUNT(1) FROM usuarios WHERE LOWER(nombreusuario) = 'lucs';")
                .Should().Be(1);
        }

        public void Dispose()
        {
            _connection.Dispose();
        }
    }
}
