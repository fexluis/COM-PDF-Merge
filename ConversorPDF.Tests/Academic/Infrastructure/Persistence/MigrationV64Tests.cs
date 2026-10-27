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
    public class MigrationV64Tests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV64Tests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV64Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void MigrationV64_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV64_AddPernoctaUltimoDia();

            // Assert
            mig.Version.Should().Be(64);
            mig.Name.Should().Be("AddPernoctaUltimoDia");
        }

        [Fact]
        public void Up_WhenOldDatabaseWithoutColumn_ShouldAddColumnAndDefaultToZero()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE vinculaciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idsede INTEGER,
                        idtercero INTEGER,
                        idprograma INTEGER,
                        idperiodo INTEGER,
                        grupo TEXT,
                        semestre TEXT,
                        idasignatura INTEGER,
                        total_horas INTEGER,
                        dia_horas_semanales TEXT,
                        fecha_inicio TEXT,
                        fecha_fin TEXT,
                        modalidadVinculacion TEXT,
                        subunidad TEXT DEFAULT '',
                        isTerminalAereo INTEGER DEFAULT 0,
                        idTerminales INTEGER DEFAULT 0,
                        gastoTerminalAerea REAL DEFAULT 0
                    );

                    CREATE TABLE sesiones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idvinculacion INTEGER,
                        idtercero INTEGER,
                        fechasesion TEXT,
                        horas REAL,
                        horariosesion TEXT,
                        viatica INTEGER DEFAULT 0
                    );

                    CREATE TABLE programas (id INTEGER PRIMARY KEY, nombreprograma TEXT, modalidad TEXT, abreviatura TEXT);
                    CREATE TABLE asignaturas (id INTEGER PRIMARY KEY, nombre TEXT, creditos INTEGER, metodologia TEXT, cod_materia TEXT, num_nivel INTEGER, abreviatura TEXT);
                    CREATE TABLE sedes (id INTEGER PRIMARY KEY, ciudad TEXT, subsede TEXT, departamento TEXT, transporte REAL);
                    CREATE TABLE terceros (id INTEGER PRIMARY KEY, nombre TEXT, cedula TEXT);
                    CREATE TABLE terminales (id INTEGER PRIMARY KEY, terminal TEXT, maxvalor REAL);

                    INSERT INTO vinculaciones (id, idsede, idtercero, idprograma, idperiodo, idasignatura, total_horas)
                    VALUES (1, 1, 1, 1, 1, 1, 32);
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV64_AddPernoctaUltimoDia();

                    // Act
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Assert
                var cols = conn.Query<string>("SELECT name FROM pragma_table_info('vinculaciones');").ToList();
                cols.Should().Contain("isPernoctaUltimoDia");

                int pernocta = conn.QuerySingle<int>("SELECT isPernoctaUltimoDia FROM vinculaciones WHERE id = 1;");
                pernocta.Should().Be(0);
            }
        }

        [Fact]
        public void Up_WhenColumnAlreadyExists_ShouldBeIdempotent()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE vinculaciones (
                        id INTEGER PRIMARY KEY,
                        idsede INTEGER,
                        idtercero INTEGER,
                        idprograma INTEGER,
                        idperiodo INTEGER,
                        grupo TEXT,
                        semestre TEXT,
                        idasignatura INTEGER,
                        total_horas INTEGER,
                        dia_horas_semanales TEXT,
                        fecha_inicio TEXT,
                        fecha_fin TEXT,
                        modalidadVinculacion TEXT,
                        subunidad TEXT DEFAULT '',
                        isPernoctaUltimoDia INTEGER DEFAULT 0,
                        isTerminalAereo INTEGER DEFAULT 0,
                        idTerminales INTEGER DEFAULT 0,
                        gastoTerminalAerea REAL DEFAULT 0
                    );

                    CREATE TABLE sesiones (
                        id INTEGER PRIMARY KEY,
                        idvinculacion INTEGER,
                        idtercero INTEGER,
                        fechasesion TEXT,
                        horas REAL,
                        horariosesion TEXT,
                        viatica INTEGER DEFAULT 0
                    );

                    CREATE TABLE programas (id INTEGER PRIMARY KEY, nombreprograma TEXT, modalidad TEXT, abreviatura TEXT);
                    CREATE TABLE asignaturas (id INTEGER PRIMARY KEY, nombre TEXT, creditos INTEGER, metodologia TEXT, cod_materia TEXT, num_nivel INTEGER, abreviatura TEXT);
                    CREATE TABLE sedes (id INTEGER PRIMARY KEY, ciudad TEXT, subsede TEXT, departamento TEXT, transporte REAL);
                    CREATE TABLE terceros (id INTEGER PRIMARY KEY, nombre TEXT, cedula TEXT);
                    CREATE TABLE terminales (id INTEGER PRIMARY KEY, terminal TEXT, maxvalor REAL);
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV64_AddPernoctaUltimoDia();

                    // Act: Should not throw on second run
                    Action act = () => mig.Up(conn, trans);
                    act.Should().NotThrow();
                    trans.Commit();
                }
            }
        }

        [Fact]
        public void Up_WhenExecuted_ShouldExposeIsPernoctaUltimoDiaInView()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE vinculaciones (
                        id INTEGER PRIMARY KEY,
                        idsede INTEGER,
                        idtercero INTEGER,
                        idprograma INTEGER,
                        idperiodo INTEGER,
                        grupo TEXT,
                        semestre TEXT,
                        idasignatura INTEGER,
                        total_horas INTEGER,
                        dia_horas_semanales TEXT,
                        fecha_inicio TEXT,
                        fecha_fin TEXT,
                        modalidadVinculacion TEXT,
                        subunidad TEXT DEFAULT '',
                        isPernoctaUltimoDia INTEGER DEFAULT 1,
                        isTerminalAereo INTEGER DEFAULT 0,
                        idTerminales INTEGER DEFAULT 0,
                        gastoTerminalAerea REAL DEFAULT 0
                    );

                    CREATE TABLE sesiones (
                        id INTEGER PRIMARY KEY,
                        idvinculacion INTEGER,
                        idtercero INTEGER,
                        fechasesion TEXT,
                        horas REAL,
                        horariosesion TEXT,
                        viatica INTEGER DEFAULT 0
                    );

                    CREATE TABLE programas (id INTEGER PRIMARY KEY, nombreprograma TEXT, modalidad TEXT, abreviatura TEXT);
                    CREATE TABLE asignaturas (id INTEGER PRIMARY KEY, nombre TEXT, creditos INTEGER, metodologia TEXT, cod_materia TEXT, num_nivel INTEGER, abreviatura TEXT);
                    CREATE TABLE sedes (id INTEGER PRIMARY KEY, ciudad TEXT, subsede TEXT, departamento TEXT, transporte REAL);
                    CREATE TABLE terceros (id INTEGER PRIMARY KEY, nombre TEXT, cedula TEXT);
                    CREATE TABLE terminales (id INTEGER PRIMARY KEY, terminal TEXT, maxvalor REAL);

                    INSERT INTO programas VALUES (1, 'Prog', 'Presencial', 'P');
                    INSERT INTO asignaturas VALUES (1, 'Asig', 4, 'Presencial', 'COD', 1, 'A');
                    INSERT INTO sedes VALUES (1, 'Cali', 'Sede 1', 'Valle', 100);
                    INSERT INTO terceros VALUES (1, 'Docente', '12345');
                    INSERT INTO vinculaciones VALUES (1, 1, 1, 1, 1, 'A', 'I', 1, 32, 'LUN|4', '2026-08-01', '2026-08-30', 'Presencial', '', 1, 0, 0, 0);
                    INSERT INTO sesiones VALUES (1, 1, 1, '2026-08-10', 4, '08:00 A 12:00', 1);
                ");

                using (var trans = conn.BeginTransaction())
                {
                    var mig = new MigrationV64_AddPernoctaUltimoDia();
                    mig.Up(conn, trans);
                    trans.Commit();
                }

                // Act
                var viewCols = conn.Query<string>("SELECT name FROM pragma_table_info('vw_SesionesDetalle');").ToList();
                viewCols.Should().Contain("IsPernoctaUltimoDia");

                int isPernocta = conn.QuerySingle<int>("SELECT IsPernoctaUltimoDia FROM vw_SesionesDetalle WHERE IdSesion = 1;");
                isPernocta.Should().Be(1);
            }
        }
    }
}
