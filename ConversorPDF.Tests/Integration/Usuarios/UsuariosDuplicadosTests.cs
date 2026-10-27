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
    public class UsuariosDuplicadosTests : IDisposable
    {
        private readonly SQLiteConnection _connection;

        public UsuariosDuplicadosTests()
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
        public void GetUsuarios_WhenDuplicateUsersExist_ShouldReturnSingleUserWithRichestData()
        {
            // Arrange: insertar dos filas con el mismo nombreusuario (simula merge CRDT)
            // Fila 1 (más antigua, sin firma, perfil genérico):
            _connection.Execute(@"
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, perfil, activo, firma, updated_at)
                VALUES ('usr_old_1', 'Luis Bechara', 'lbechara', 'VALLE', 'Academica', 1, NULL, '2026-01-01 10:00:00');
            ");

            // Fila 2 (más reciente, con firma y perfil Admin):
            _connection.Execute(@"
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, perfil, activo, firma, updated_at)
                VALUES ('usr_new_2', 'Luis Bechara', 'lbechara', 'VALLE', 'Admin', 1, 'base64firmadata...', '2026-02-01 12:00:00');
            ");

            var repo = new AcademicRepository(_connection);

            // Act
            var result = repo.GetUsuarios();

            // Assert: debe retornar exactamente UNA fila para 'lbechara' y debe ser la que tiene firma/perfil Admin
            result.IsSuccess.Should().BeTrue();
            var list = result.Value.ToList();
            list.Should().HaveCount(1, "el listado de usuarios no debe mostrar duplicados por nombre de usuario");
            list[0].NombreUsuario.Should().Be("lbechara");
            list[0].Perfil.Should().Be("Admin");
            list[0].Firma.Should().Be("base64firmadata...");
        }

        [Fact]
        public void InsertUsuario_WhenDeterministicId_ShouldUseUsrPrefixAndLowerUsername()
        {
            // Arrange
            var repo = new AcademicRepository(_connection);
            var nuevo = new Usuario
            {
                Nombre = "Pedro Perez",
                NombreUsuario = "PedroPerez",
                Subunidad = "VALLE",
                Activo = 1,
                Perfil = "Admin"
            };

            // Act
            var result = repo.InsertUsuario(nuevo);

            // Assert: el ID generado debe ser determinista: usr_pedroperez
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be("usr_pedroperez", "el ID del usuario debe ser determinista para evitar duplicaciones en CRDT");
        }

        [Fact]
        public void MigrationV79_WhenExecuted_ShouldDeduplicateAndNormalizeIds()
        {
            // Arrange: 2 filas para lbechara con GUIDs aleatorios, una sin firma y otra con firma
            _connection.Execute(@"
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, perfil, activo, firma, updated_at)
                VALUES ('guid_old_1111', 'Luis Bechara', 'lbechara', 'VALLE', 'Academica', 1, NULL, '2026-01-01 10:00:00');
            ");
            _connection.Execute(@"
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, perfil, activo, firma, updated_at)
                VALUES ('guid_old_2222', 'Luis Bechara', 'lbechara', 'VALLE', 'Admin', 1, 'base64firmadata...', '2026-02-01 12:00:00');
            ");

            var migration = new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations.MigrationV79_NormalizeUsuariosIds();

            // Act
            migration.Up(_connection);

            // Assert
            var rows = _connection.Query<Usuario>("SELECT * FROM usuarios;").ToList();
            rows.Should().HaveCount(1, "la migración V79 debe haber eliminado el duplicado");
            rows[0].Id.Should().Be("usr_lbechara", "la migración V79 debe haber normalizado el ID al formato usr_{nombreusuario}");
            rows[0].Perfil.Should().Be("Admin");
            rows[0].Firma.Should().Be("base64firmadata...");
        }

        public void Dispose()
        {
            _connection?.Dispose();
        }
    }
}
