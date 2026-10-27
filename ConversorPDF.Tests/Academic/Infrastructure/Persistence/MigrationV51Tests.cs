using System;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class MigrationV51Tests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV51Tests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV51Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void MigrationV51_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV51_PeriodoVentanaRepair();

            // Assert
            mig.Version.Should().Be(51);
            mig.Name.Should().Be("PeriodoVentanaRepair");
            mig.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void Up_WhenPeriodosTableExistsWithoutDateColumns_ShouldAddColumnsAndApplyBackfill()
        {
            // Arrange - Simula vault en v50 donde las columnas no se crearon físicamente
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo TEXT,
                        activo INTEGER DEFAULT 1
                    );

                    INSERT INTO periodos (idperiodo, periodo, activo) VALUES
                    (1, '2026-01', 0),
                    (2, '2026-02', 1),
                    (3, '2027-01', 0);
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV51_PeriodoVentanaRepair();

                    // Act
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Assert: Columnas reparadas
                var cols = conn.Query<string>("SELECT name FROM pragma_table_info('periodos');").ToList();
                cols.Should().Contain("periodo_inicio");
                cols.Should().Contain("periodo_fin");

                // Assert: Backfill aplicado
                var p1 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 1;");
                ((string)p1.periodo_inicio).Should().Be("2026-01-12");
                ((string)p1.periodo_fin).Should().Be("2026-06-19");

                var p2 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 2;");
                ((string)p2.periodo_inicio).Should().Be("2026-08-15");
                ((string)p2.periodo_fin).Should().Be("2026-12-12");

                var p3 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 3;");
                ((string)p3.periodo_inicio).Should().Be("2027-01-01");
                ((string)p3.periodo_fin).Should().Be("2027-06-30");
            }
        }

        [Fact]
        public void Up_WhenPeriodosAlreadyHasColumnsWithNullValues_ShouldApplyBackfillWithoutAltering()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo TEXT,
                        periodo_inicio TEXT,
                        periodo_fin TEXT,
                        activo INTEGER DEFAULT 1
                    );

                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo) VALUES
                    (2, '2026-02', NULL, NULL, 1);
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV51_PeriodoVentanaRepair();

                    // Act
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Assert
                var p2 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 2;");
                ((string)p2.periodo_inicio).Should().Be("2026-08-15");
                ((string)p2.periodo_fin).Should().Be("2026-12-12");
            }
        }

        [Fact]
        public void Up_WhenPeriodosTableDoesNotExist_ShouldReturnGracefully()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV51_PeriodoVentanaRepair();

                    // Act & Assert
                    Action act = () => mig.Up(conn, trans);
                    act.Should().NotThrow();
                    trans.Commit();
                }
            }
        }

        [Fact]
        public void Up_WhenCrsqlTriggersAreStale_ShouldDropTriggersAndNotFailOnBackfill()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo TEXT,
                        activo INTEGER DEFAULT 1,
                        subunidad TEXT
                    );

                    INSERT INTO periodos (idperiodo, periodo, activo, subunidad) VALUES
                    (2, '2026-02', 1, 'VALLE');

                    CREATE TRIGGER periodos__crsql_utrig AFTER UPDATE ON periodos BEGIN
                        SELECT RAISE(ABORT, 'SQL logic error expected 13 values, got 9');
                    END;
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV51_PeriodoVentanaRepair();

                    // Act & Assert
                    Action act = () => mig.Up(conn, trans);
                    act.Should().NotThrow();
                    trans.Commit();
                }

                // Assert
                var p2 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 2;");
                ((string)p2.periodo_inicio).Should().Be("2026-08-15");
                ((string)p2.periodo_fin).Should().Be("2026-12-12");
            }
        }

        [Fact]
        public void Up_WhenInicioPresenteYFinNulo_ShouldFillOnlyFin()
        {
            // Arrange: Fila con periodo_inicio ya establecido y periodo_fin NULL
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo TEXT,
                        periodo_inicio TEXT,
                        periodo_fin TEXT,
                        activo INTEGER DEFAULT 1
                    );

                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo) VALUES
                    (2, '2026-02', '2026-08-01', NULL, 1);
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV51_PeriodoVentanaRepair();

                    // Act
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Assert: periodo_inicio se preserva ('2026-08-01') y periodo_fin se rellena ('2026-12-12')
                var p2 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 2;");
                ((string)p2.periodo_inicio).Should().Be("2026-08-01");
                ((string)p2.periodo_fin).Should().Be("2026-12-12");
            }
        }
    }
}
