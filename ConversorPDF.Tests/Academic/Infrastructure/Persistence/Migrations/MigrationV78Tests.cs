using System.Data.SQLite;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV78Tests
    {
        [Fact]
        public void Up_WhenExecutedOnCleanDatabase_ShouldCreateTableAndIndex()
        {
            // Arrange
            using (var conn = new SQLiteConnection("Data Source=:memory:;Version=3;"))
            {
                conn.Open();
                var migration = new MigrationV78_AddAgFiltroSeleccion();

                // Act
                migration.Up(conn);

                // Assert
                int tableCount = conn.ExecuteScalar<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_filtro_seleccion';");
                tableCount.Should().Be(1);

                int indexCount = conn.ExecuteScalar<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND name='ux_ag_filtro_seleccion';");
                indexCount.Should().Be(1);
            }
        }

        [Fact]
        public void Up_WhenExecutedTwice_ShouldBeIdempotentWithoutError()
        {
            // Arrange
            using (var conn = new SQLiteConnection("Data Source=:memory:;Version=3;"))
            {
                conn.Open();
                var migration = new MigrationV78_AddAgFiltroSeleccion();

                // Act
                migration.Up(conn);
                migration.Up(conn);

                // Assert
                int tableCount = conn.ExecuteScalar<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_filtro_seleccion';");
                tableCount.Should().Be(1);
            }
        }
    }
}
