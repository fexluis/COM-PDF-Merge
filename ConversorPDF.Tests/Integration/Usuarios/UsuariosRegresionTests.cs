using System;
using System.Data.SQLite;
using System.Linq;
using Dapper;
using FluentAssertions;
using Xunit;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;

namespace ConversorPDF.Tests.Integration.Usuarios
{
    /// <summary>
    /// Regresion de los hallazgos H2 (updated_at no se actualiza) y H3 (el dedupe
    /// puede resucitar un id legacy mas "reciente") de la auditoria 2026-10-27.
    /// </summary>
    public class UsuariosRegresionTests : IDisposable
    {
        private readonly SQLiteConnection _connection;

        public UsuariosRegresionTests()
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

        [Fact]
        public void UpdateUsuario_WhenSeGuardaCambio_ShouldActualizarUpdatedAt()
        {
            // Arrange
            _connection.Execute(@"
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, activo, updated_at)
                VALUES ('usr_lucs', 'Luis', 'lucs', 'VALLE', 1, '2020-01-01 00:00:00');
            ");
            var repo = new AcademicRepository(_connection);
            var usuario = new Usuario
            {
                Id = "usr_lucs",
                Nombre = "Luis",
                NombreUsuario = "lucs",
                Subunidad = "ANTIOQUIA",
                Activo = 1,
                Perfil = "Admin",
                Firma = null
            };

            // Act
            var result = repo.UpdateUsuario(usuario);

            // Assert
            result.IsSuccess.Should().BeTrue();
            string updatedAt = _connection.QuerySingle<string>("SELECT updated_at FROM usuarios WHERE id = 'usr_lucs';");
            updatedAt.Should().NotBe("2020-01-01 00:00:00", "una edicion debe marcar la fila como reciente");
        }

        [Fact]
        public void GetUsuarios_WhenDuplicadoLegacyEsMasReciente_ShouldPreferirIdDeterminista()
        {
            // Arrange: misma cuenta, un id determinista (antiguo) y un id legacy (mas "reciente")
            _connection.Execute(@"
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, activo, updated_at)
                VALUES ('usr_lucs', 'Luis', 'lucs', 'VALLE', 1, '2020-01-01 00:00:00');
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, activo, updated_at)
                VALUES ('9f2c1b7a4d', 'Luis', 'lucs', 'VALLE', 1, '2030-01-01 00:00:00');
            ");
            var repo = new AcademicRepository(_connection);

            // Act
            var result = repo.GetUsuarios();

            // Assert
            result.IsSuccess.Should().BeTrue();
            var list = result.Value.ToList();
            list.Should().HaveCount(1);
            list[0].Id.Should().Be("usr_lucs", "la identidad determinista debe prevalecer sobre un id legacy");
        }

        public void Dispose()
        {
            _connection.Dispose();
        }
    }
}
