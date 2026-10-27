using System;
using System.Data.SQLite;
using Dapper;
using FluentAssertions;
using Xunit;
using ConversorPDF.XLL.Core.Sync;

namespace ConversorPDF.Tests.Unit.Usuarios
{
    public class SubunidadResolverTests : IDisposable
    {
        private readonly SQLiteConnection _connection;
        private readonly SubunidadResolver _resolver;

        public SubunidadResolverTests()
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

                CREATE TABLE config (
                    id INTEGER PRIMARY KEY,
                    subunidad TEXT NOT NULL
                );
            ");

            _resolver = new SubunidadResolver();
        }

        [Fact]
        public void Resolve_WhenDuplicateUsersExist_ShouldSelectActiveWithMostRecentUpdatedAt()
        {
            string currentUser = Environment.UserName;

            // Arrange: Fila 1 antigua con ANTIOQUIA
            _connection.Execute(@"
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, activo, updated_at)
                VALUES ('usr_ant', 'Usuario Test', @User, 'ANTIOQUIA', 1, '2026-01-01 10:00:00');
            ", new { User = currentUser });

            // Fila 2 más reciente con VALLE
            _connection.Execute(@"
                INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, activo, updated_at)
                VALUES ('usr_val', 'Usuario Test', @User, 'VALLE', 1, '2026-02-01 12:00:00');
            ", new { User = currentUser });

            // Act
            var result = _resolver.Resolve(_connection);

            // Assert: debe seleccionar VALLE (la más reciente) de forma determinista
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be("VALLE", "debe resolver la subunidad de la fila activa con updated_at más reciente");
        }

        public void Dispose()
        {
            _connection?.Dispose();
        }
    }
}
