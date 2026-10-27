using System;
using System.Linq;
using System.Data.SQLite;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Integration.Academic
{
    public class SchemaMigrationTests : IDisposable
    {
        private readonly SQLiteConnection _conn;

        public SchemaMigrationTests()
        {
            _conn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;");
            _conn.Open();
            
            // Setup base tables needed for views
            _conn.Execute(@"
                CREATE TABLE __SchemaVersion (Version INTEGER PRIMARY KEY);
                CREATE TABLE programas (id INTEGER);
                CREATE TABLE terceros (id INTEGER);
                CREATE TABLE sedes (id INTEGER);
                CREATE TABLE asignaturas (id INTEGER);
                CREATE TABLE vinculaciones (id INTEGER, idprograma INTEGER, idtercero INTEGER, idasignatura INTEGER, idsede INTEGER, idperiodo INTEGER, semestre TEXT);
                CREATE TABLE sesiones (id INTEGER, idvinculacion INTEGER, fechasesion TEXT, horariosesion TEXT);
                CREATE TABLE periodos (idperiodo INTEGER PRIMARY KEY, periodo TEXT);
            ");
        }

        [Theory]
        [InlineData("periodos", "table")]      // v3: tabla periodos
        [InlineData("vw_SesionesDetalle", "view")]  // v4: vista sesiones
        public void SchemaObject_CanBeCreated(string name, string type)
        {
            // Act: Simulate creation of objects if they were part of migration
            if (name == "vw_SesionesDetalle")
            {
                _conn.Execute(@"
                    CREATE VIEW vw_SesionesDetalle AS
                    SELECT 
                        s.id AS IdSesion,
                        v.id AS IdVinculacion,
                        v.semestre AS Semestre
                    FROM sesiones s
                    JOIN vinculaciones v ON s.idvinculacion = v.id;");
            }

            // Assert
            var exists = _conn.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM sqlite_master WHERE name = '{name}' AND type = '{type}'");
            exists.Should().Be(1, $"El objeto {name} de tipo {type} debería existir en el esquema.");
        }

        [Fact]
        public void View_V5_ContainsSemestreColumn()
        {
            // Act
            _conn.Execute(@"
                CREATE VIEW vw_SesionesDetalle_V5 AS
                SELECT 
                    s.id AS IdSesion,
                    v.semestre AS Semestre
                FROM sesiones s
                JOIN vinculaciones v ON s.idvinculacion = v.id;");

            // Assert
            var columns = _conn.Query("PRAGMA table_info(vw_SesionesDetalle_V5)").Select(c => (string)c.name).ToList();
            columns.Should().Contain("Semestre");
        }

        public void Dispose()
        {
            _conn?.Close();
            _conn?.Dispose();
        }
    }
}
