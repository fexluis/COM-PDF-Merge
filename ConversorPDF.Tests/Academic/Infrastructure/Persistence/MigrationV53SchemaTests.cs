using System;
using System.Data.SQLite;
using System.IO;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class MigrationV53SchemaTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV53SchemaTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV53Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";
        }

        public void Dispose()
        {
            try
            {
                if (File.Exists(_dbPath))
                {
                    File.Delete(_dbPath);
                }
            }
            catch { }
        }

        [Fact]
        public void MigrationV53_Metadata_ShouldBeCorrect()
        {
            var mig = new MigrationV53_AddAgArbolTables();
            mig.Version.Should().Be(53);
            mig.Name.Should().Be("AddAgArbolTables");
            mig.Aditiva.Should().BeFalse();
        }

        [Fact]
        public void MigrationV53_Up_WhenExecutedTwice_ShouldBeIdempotent()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                var mig = new MigrationV53_AddAgArbolTables();

                // Primera ejecución
                using (var tx1 = conn.BeginTransaction())
                {
                    mig.Up(conn, tx1);
                    tx1.Commit();
                }

                // Segunda ejecución (idempotencia)
                using (var tx2 = conn.BeginTransaction())
                {
                    Action act = () => mig.Up(conn, tx2);
                    act.Should().NotThrow();
                    tx2.Commit();
                }

                // Verificar existencia de tablas y columnas
                int countArbol = conn.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_arbol';");
                int countNodo = conn.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_arbol_nodo';");

                countArbol.Should().Be(1);
                countNodo.Should().Be(1);

                var colsNodo = conn.Query<string>("SELECT name FROM pragma_table_info('ag_arbol_nodo');");
                colsNodo.Should().Contain(new[] { "idarbol", "aula_ref", "modalidad", "capacidad", "tipo_aula", "caracteristicas_json", "es_virtual" });
            }
        }
    }
}
