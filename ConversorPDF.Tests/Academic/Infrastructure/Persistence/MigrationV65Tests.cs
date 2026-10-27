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
    public class MigrationV65Tests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV65Tests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV65Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void MigrationV65_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV65_PeriodoVentanaCrdtRepair();

            // Assert
            mig.Version.Should().Be(65);
            mig.Name.Should().Be("PeriodoVentanaCrdtRepair");
            mig.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void Up_WhenLegacySchemaAndStaleCrsqlTriggers_ShouldAddColumnsBackfillAndDropTriggers()
        {
            // Arrange: Simula vault donde las columnas nunca se agregaron y los triggers viejos provocan 'expected 13 values, got 9'
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
                    (1, '2026-01', 0, 'VALLE'),
                    (2, '2026-02', 1, 'VALLE'),
                    (3, '2027-01', 0, 'VALLE');

                    -- Triggers centinela simulando mismatch de aridad CRDT previo
                    CREATE TRIGGER periodos__crsql_utrig AFTER UPDATE ON periodos BEGIN
                        SELECT RAISE(ABORT, 'SQL logic error expected 13 values, got 9');
                    END;
                    CREATE TRIGGER periodos__crsql_itrig AFTER INSERT ON periodos BEGIN
                        SELECT RAISE(ABORT, 'SQL logic error expected 13 values, got 9');
                    END;
                    CREATE TRIGGER periodos__crsql_dtrig AFTER DELETE ON periodos BEGIN
                        SELECT RAISE(ABORT, 'SQL logic error expected 13 values, got 9');
                    END;
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV65_PeriodoVentanaCrdtRepair();

                    // Act
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Assert: Columnas agregadas físicamente
                var cols = conn.Query<string>("SELECT name FROM pragma_table_info('periodos');").ToList();
                cols.Should().Contain("periodo_inicio");
                cols.Should().Contain("periodo_fin");

                // Assert: Triggers stale eliminados
                var triggers = conn.Query<string>("SELECT name FROM sqlite_master WHERE type='trigger' AND tbl_name='periodos';").ToList();
                triggers.Should().NotContain("periodos__crsql_utrig");
                triggers.Should().NotContain("periodos__crsql_itrig");
                triggers.Should().NotContain("periodos__crsql_dtrig");

                // Assert: Backfill canónico aplicado (con regla V56 para 2026-02)
                var p1 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 1;");
                ((string)p1.periodo_inicio).Should().Be("2026-01-12");
                ((string)p1.periodo_fin).Should().Be("2026-06-19");

                var p2 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 2;");
                ((string)p2.periodo_inicio).Should().Be("2026-08-11");
                ((string)p2.periodo_fin).Should().Be("2026-12-12");

                var p3 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 3;");
                ((string)p3.periodo_inicio).Should().Be("2027-01-01");
                ((string)p3.periodo_fin).Should().Be("2027-06-30");

                // Assert: Un UPDATE posterior no revienta con el error de aridad
                Action actUpdate = () => conn.Execute("UPDATE periodos SET periodo = '2026-02-MOD' WHERE idperiodo = 2;");
                actUpdate.Should().NotThrow();
            }
        }

        [Fact]
        public void Up_WhenColumnsExistWithNullDatesAndStaleTriggers_ShouldBackfillWithoutArityError()
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
                        activo INTEGER DEFAULT 1,
                        subunidad TEXT
                    );

                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo, subunidad) VALUES
                    (2, '2026-02', NULL, NULL, 1, 'VALLE');

                    CREATE TRIGGER periodos__crsql_utrig AFTER UPDATE ON periodos BEGIN
                        SELECT RAISE(ABORT, 'SQL logic error expected 13 values, got 9');
                    END;
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV65_PeriodoVentanaCrdtRepair();

                    // Act
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Assert
                var p2 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 2;");
                ((string)p2.periodo_inicio).Should().Be("2026-08-11");
                ((string)p2.periodo_fin).Should().Be("2026-12-12");
            }
        }

        [Fact]
        public void Up_WhenDatesAlreadyPresent_ShouldPreserveThem()
        {
            // Arrange: Fechas personalizadas existentes no deben ser sobrescritas
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo TEXT,
                        periodo_inicio TEXT,
                        periodo_fin TEXT,
                        activo INTEGER DEFAULT 1,
                        subunidad TEXT
                    );

                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo, subunidad) VALUES
                    (1, '2026-01', '2026-02-01', '2026-05-30', 0, 'VALLE');
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV65_PeriodoVentanaCrdtRepair();

                    // Act
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Assert: Fechas preservadas intactas
                var p1 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 1;");
                ((string)p1.periodo_inicio).Should().Be("2026-02-01");
                ((string)p1.periodo_fin).Should().Be("2026-05-30");
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
                    var mig = new MigrationV65_PeriodoVentanaCrdtRepair();

                    // Act & Assert
                    Action act = () => mig.Up(conn, trans);
                    act.Should().NotThrow();
                    trans.Commit();
                }
            }
        }

        [Fact]
        public void Up_WhenSchemaAndDatesAreComplete_ShouldNotDropCrsqlTriggers()
        {
            // Arrange: FASE 752: Vault con columnas y fechas completas no debe soltar los triggers
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo TEXT,
                        periodo_inicio TEXT,
                        periodo_fin TEXT,
                        activo INTEGER DEFAULT 1,
                        subunidad TEXT
                    );

                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo, subunidad) VALUES
                    (1, '2026-01', '2026-01-12', '2026-06-19', 0, 'VALLE'),
                    (2, '2026-02', '2026-08-11', '2026-12-12', 1, 'VALLE'),
                    (3, '2027-01', '2027-01-01', '2027-06-30', 0, 'VALLE');

                    CREATE TRIGGER periodos__crsql_utrig AFTER UPDATE ON periodos BEGIN SELECT 1; END;
                    CREATE TRIGGER periodos__crsql_itrig AFTER INSERT ON periodos BEGIN SELECT 1; END;
                    CREATE TRIGGER periodos__crsql_dtrig AFTER DELETE ON periodos BEGIN SELECT 1; END;
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV65_PeriodoVentanaCrdtRepair();

                    // Act
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Assert: Los 3 triggers permanecen intactos
                int triggerCount = conn.QuerySingle<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='trigger' AND tbl_name='periodos';");
                triggerCount.Should().Be(3);
            }
        }

        [Fact]
        public void Up_WhenInicioPresenteYFinNulo_ShouldFillOnlyFin()
        {
            // Arrange: FASE 753: id 1 tiene inicio pero fin es nulo
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo TEXT,
                        periodo_inicio TEXT,
                        periodo_fin TEXT,
                        activo INTEGER DEFAULT 1,
                        subunidad TEXT
                    );

                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo, subunidad) VALUES
                    (1, '2026-01', '2026-01-12', NULL, 0, 'VALLE');
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV65_PeriodoVentanaCrdtRepair();

                    // Act
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Assert: periodo_inicio se conserva, periodo_fin se rellena con 2026-06-19
                var p1 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 1;");
                ((string)p1.periodo_inicio).Should().Be("2026-01-12");
                ((string)p1.periodo_fin).Should().Be("2026-06-19");
            }
        }

        [Fact]
        public void Up_WhenInicioEs20260815_ShouldApplyV56Rule()
        {
            // Arrange: FASE 753: id 2 retenía 2026-08-15 de V50 pero no alcanzó a ejecutar V56
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo TEXT,
                        periodo_inicio TEXT,
                        periodo_fin TEXT,
                        activo INTEGER DEFAULT 1,
                        subunidad TEXT
                    );

                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo, subunidad) VALUES
                    (2, '2026-02', '2026-08-15', '2026-12-12', 1, 'VALLE');
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV65_PeriodoVentanaCrdtRepair();

                    // Act
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Assert: periodo_inicio se actualiza a 2026-08-11 según regla V56
                var p2 = conn.QuerySingle("SELECT periodo_inicio, periodo_fin FROM periodos WHERE idperiodo = 2;");
                ((string)p2.periodo_inicio).Should().Be("2026-08-11");
                ((string)p2.periodo_fin).Should().Be("2026-12-12");
            }
        }
    }
}
