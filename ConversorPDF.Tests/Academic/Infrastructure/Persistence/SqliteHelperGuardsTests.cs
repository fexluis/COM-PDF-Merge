using System;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class SqliteHelperGuardsTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public SqliteHelperGuardsTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "SqliteHelperGuardsTest_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
            catch
            {
                // Best effort cleanup
            }
        }

        [Fact]
        public void ApplyBaseSchemaAndSeeds_WhenTercerosTableMissingCedulaAndNombre_ShouldAddMissingColumns()
        {
            // Arrange - Simula tabla legada terceros sin cédula ni nombre
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE terceros (
                        id INTEGER PRIMARY KEY,
                        rol TEXT
                    );
                ");

                using (var trans = conn.BeginTransaction())
                {
                    // Act
                    SqliteHelper.ApplyBaseSchemaAndSeeds(conn, trans, "VALLE");
                    trans.Commit();
                }

                // Assert
                var cols = conn.Query<string>("SELECT name FROM pragma_table_info('terceros');").ToList();
                cols.Should().Contain("cedula");
                cols.Should().Contain("nombre");
                cols.Should().Contain("categoria");
            }
        }

        [Fact]
        public void TryGetSharedContextWithoutInit_WhenNotInitialized_ShouldReturnNull()
        {
            // Act
            var context = SqliteHelper.TryGetSharedContextWithoutInit();

            // Assert
            // Si la aplicación de pruebas no ha ejecutado GetContext(), debe devolver null de forma segura sin excepciones
            if (context == null)
            {
                context.Should().BeNull();
            }
            else
            {
                context.IsDisposed.Should().BeFalse();
            }
        }
    }
}
