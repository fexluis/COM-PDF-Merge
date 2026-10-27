using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class MigrationV77Tests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV77Tests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV77Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
                // Limpieza defensiva
            }
        }

        [Fact]
        public void MigrationV77_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV77_AddPernoctaAgVinculaciones();

            // Assert
            mig.Version.Should().Be(77);
            mig.Name.Should().Be("AddPernoctaAgVinculaciones");
            mig.Aditiva.Should().BeTrue();
            mig.Should().BeAssignableTo<IMigration>();
            mig.Should().BeAssignableTo<IMigrationTolerante>();
        }

        [Fact]
        public void Up_WhenTablesExistWithoutColumn_ShouldAddColumnToBothTablesWithDefaultZero()
        {
            // Arrange
            var mig = new MigrationV77_AddPernoctaAgVinculaciones();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                conn.Execute(@"
                    CREATE TABLE vinculaciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idtercero INTEGER DEFAULT 1
                    );
                    CREATE TABLE ag_vinculaciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idtercero INTEGER DEFAULT 1
                    );
                    INSERT INTO vinculaciones (idtercero) VALUES (10);
                    INSERT INTO ag_vinculaciones (idtercero) VALUES (20);
                ");

                // Act
                mig.Up(conn, null);

                // Assert
                var colsVinc = ObtenerNombresColumnas(conn, "vinculaciones");
                var colsAgVinc = ObtenerNombresColumnas(conn, "ag_vinculaciones");

                colsVinc.Should().Contain("isPernoctaUltimoDia");
                colsAgVinc.Should().Contain("isPernoctaUltimoDia");

                int pernoctaVinc = conn.QuerySingle<int>("SELECT isPernoctaUltimoDia FROM vinculaciones WHERE id = 1;");
                int pernoctaAgVinc = conn.QuerySingle<int>("SELECT isPernoctaUltimoDia FROM ag_vinculaciones WHERE id = 1;");

                pernoctaVinc.Should().Be(0);
                pernoctaAgVinc.Should().Be(0);
            }
        }

        [Fact]
        public void Up_WhenExecutedMultipleTimes_ShouldBeIdempotent()
        {
            // Arrange
            var mig = new MigrationV77_AddPernoctaAgVinculaciones();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                conn.Execute(@"
                    CREATE TABLE vinculaciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT
                    );
                    CREATE TABLE ag_vinculaciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT
                    );
                ");

                // Act
                mig.Up(conn, null);
                Action act = () => mig.Up(conn, null);

                // Assert
                act.Should().NotThrow();

                var colsAgVinc = ObtenerNombresColumnas(conn, "ag_vinculaciones");
                colsAgVinc.Should().Contain("isPernoctaUltimoDia");
            }
        }

        [Fact]
        public void Up_WhenTablesDoNotExist_ShouldNotFail()
        {
            // Arrange
            var mig = new MigrationV77_AddPernoctaAgVinculaciones();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                // Act
                Action act = () => mig.Up(conn, null);

                // Assert
                act.Should().NotThrow();
            }
        }

        private static List<string> ObtenerNombresColumnas(SQLiteConnection conn, string tabla)
        {
            var columnas = new List<string>();
            var info = conn.Query($"PRAGMA table_info({tabla});");
            foreach (var row in info)
            {
                var dict = (IDictionary<string, object>)row;
                if (dict.TryGetValue("name", out var val) && val != null)
                {
                    columnas.Add(val.ToString());
                }
            }
            return columnas;
        }
    }
}
