using System;
using System.Data.SQLite;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Integration.Usuarios
{
    public class UsuarioVaultMoverTests : IDisposable
    {
        private readonly SQLiteConnection _connection;

        public UsuarioVaultMoverTests()
        {
            _connection = new SQLiteConnection("Data Source=:memory:;Version=3;");
            _connection.Open();

            _connection.Execute(@"
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
            ");
        }

        public void Dispose()
        {
            _connection?.Dispose();
        }

        [Fact]
        public void Mover_WhenUsuarioExisteEnOrigen_ShouldPreservePerfilActivoYFirma()
        {
            // Arrange
            var usuario = new Usuario
            {
                Id = "usr_luis.docente",
                Nombre = "Luis Docente",
                NombreUsuario = "luis.docente",
                Subunidad = "ANTIOQUIA",
                Perfil = "Docente",
                Activo = 0,
                Firma = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAA"
            };

            // Act
            var result = UsuarioVaultMover.UpsertEnVault(_connection, usuario, "VALLE");

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Insertado.Should().BeTrue();

            var row = _connection.QuerySingle(@"
                SELECT id, nombre, nombreusuario, subunidad, perfil, activo, firma, updated_at
                FROM usuarios
                WHERE id = 'usr_luis.docente';
            ");

            ((string)row.id).Should().Be("usr_luis.docente");
            ((string)row.nombre).Should().Be("Luis Docente");
            ((string)row.nombreusuario).Should().Be("luis.docente");
            ((string)row.subunidad).Should().Be("VALLE");
            ((string)row.perfil).Should().Be("Docente", "debe preservar el perfil original (R-MUV-03)");
            ((int)row.activo).Should().Be(0, "debe preservar el estado de activo original (R-MUV-03)");
            ((string)row.firma).Should().Be("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAA", "debe preservar la firma (R-MUV-03)");
            ((string)row.updated_at).Should().NotBeNullOrEmpty();
        }

        [Fact]
        public void Mover_WhenSeInvocaDosVeces_ShouldSerIdempotente()
        {
            // Arrange
            var usuario = new Usuario
            {
                Id = "usr_carlos",
                Nombre = "Carlos Lopez",
                NombreUsuario = "carlos",
                Subunidad = "ANTIOQUIA",
                Perfil = "Academica",
                Activo = 1
            };

            // Act
            var res1 = UsuarioVaultMover.UpsertEnVault(_connection, usuario, "VALLE");
            var res2 = UsuarioVaultMover.UpsertEnVault(_connection, usuario, "VALLE");

            // Assert
            res1.IsSuccess.Should().BeTrue();
            res1.Value.Insertado.Should().BeTrue();

            res2.IsSuccess.Should().BeTrue();
            res2.Value.Actualizado.Should().BeTrue();

            var count = _connection.QuerySingle<int>("SELECT COUNT(1) FROM usuarios WHERE id = 'usr_carlos';");
            count.Should().Be(1, "el movimiento debe ser idempotente y no duplicar registros");
        }

        [Fact]
        public void Mover_WhenUsuarioNoExisteEnDestino_ShouldInsertarConIdDeterminista()
        {
            // Arrange
            var usuario = new Usuario
            {
                Id = null,
                Nombre = "Maria Paz",
                NombreUsuario = "Maria.Paz",
                Subunidad = "ANTIOQUIA",
                Perfil = "Admin",
                Activo = 1
            };

            // Act
            var result = UsuarioVaultMover.UpsertEnVault(_connection, usuario, "CALDAS");

            // Assert
            result.IsSuccess.Should().BeTrue();
            result.Value.Insertado.Should().BeTrue();

            var savedId = _connection.QuerySingle<string>("SELECT id FROM usuarios WHERE LOWER(nombreusuario) = 'maria.paz';");
            savedId.Should().Be("usr_maria.paz", "debe generarse un id determinista usr_<nombreusuario>");
        }
    }
}
