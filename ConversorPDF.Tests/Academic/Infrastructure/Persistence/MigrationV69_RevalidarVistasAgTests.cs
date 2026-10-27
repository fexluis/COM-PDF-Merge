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
    public class MigrationV69_RevalidarVistasAgTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV69_RevalidarVistasAgTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV69Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void MigrationV69_Metadata_ShouldBeCorrect()
        {
            var v69 = new MigrationV69_RevalidarVistasAg();
            v69.Version.Should().Be(69);
            v69.Name.Should().Be("RevalidarVistasAg");
            v69.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void MigrationV69_Up_WhenSchemaFromV49_ShouldRecreateAndValidateQueryableViews()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                // DDL canónico de V49 con id_individuo
                conn.Execute(@"
                    CREATE TABLE ag_vinculaciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        id_corrida INTEGER NOT NULL DEFAULT 0,
                        id_individuo INTEGER NOT NULL DEFAULT 0,
                        idperiodo INTEGER NOT NULL DEFAULT 0,
                        idsede INTEGER NOT NULL DEFAULT 0,
                        idprograma INTEGER NOT NULL DEFAULT 0,
                        idasignatura INTEGER NOT NULL DEFAULT 0,
                        idtercero INTEGER NOT NULL DEFAULT 0,
                        semestre TEXT,
                        grupo TEXT,
                        modalidadVinculacion TEXT,
                        dia_horas_semanales TEXT,
                        total_horas REAL DEFAULT 0,
                        fecha_inicio TEXT,
                        fecha_fin TEXT,
                        estado INTEGER DEFAULT 0,
                        es_vacante INTEGER DEFAULT 0,
                        motivo_vacante TEXT,
                        candidatos_pool INTEGER DEFAULT 0,
                        subunidad TEXT,
                        isTerminalAereo INTEGER DEFAULT 0,
                        idTerminales INTEGER DEFAULT 0,
                        gastoTerminalAerea REAL DEFAULT 0
                    );
                    CREATE TABLE ag_sesiones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idvinculacion INTEGER NOT NULL,
                        idtercero INTEGER NOT NULL,
                        id_aula INTEGER DEFAULT 0,
                        fechasesion TEXT,
                        horariosesion TEXT,
                        horas REAL DEFAULT 0,
                        viatica INTEGER DEFAULT 0,
                        subunidad TEXT,
                        idindividuo INTEGER DEFAULT 0,
                        idgen INTEGER DEFAULT 0,
                        es_vacante INTEGER DEFAULT 0,
                        motivo_vacante TEXT,
                        candidatos_pool INTEGER DEFAULT 0
                    );
                    CREATE TABLE programas (id INTEGER PRIMARY KEY, modalidad TEXT, abreviatura TEXT, nombreprograma TEXT);
                    CREATE TABLE terceros (id INTEGER PRIMARY KEY, nombre TEXT, cedula TEXT);
                    CREATE TABLE asignaturas (id INTEGER PRIMARY KEY, nombre TEXT, abreviatura TEXT, creditos INTEGER, num_nivel INTEGER, metodologia TEXT, cod_materia TEXT);
                    CREATE TABLE sedes (id INTEGER PRIMARY KEY, ciudad TEXT, subsede TEXT, departamento TEXT, transporte TEXT);
                    CREATE TABLE periodos (idperiodo INTEGER PRIMARY KEY, periodo TEXT);
                    CREATE TABLE terminales (id INTEGER PRIMARY KEY, terminal TEXT);
                ");

                // Semillas de prueba
                conn.Execute("INSERT INTO ag_vinculaciones (id, id_corrida, id_individuo, estado) VALUES (1, 10, 5, 0);");
                conn.Execute("INSERT INTO ag_sesiones (id, idvinculacion, idtercero, fechasesion) VALUES (1, 1, 1, '2026-03-03');");

                var v69 = new MigrationV69_RevalidarVistasAg();
                using (var tx = conn.BeginTransaction())
                {
                    v69.Up(conn, tx);
                    tx.Commit();
                }

                // Aserción: las vistas existen y son consultables sin error de columnas
                var sesiones = conn.Query("SELECT * FROM vw_AgSesionesDetalle;").ToList();
                sesiones.Should().NotBeEmpty();

                var vinculaciones = conn.Query("SELECT * FROM vw_AgVinculacionesDetalle;").ToList();
                vinculaciones.Should().NotBeEmpty();

                var cols = conn.Query<string>("SELECT name FROM pragma_table_info('vw_AgVinculacionesDetalle');").ToList();
                cols.Should().Contain("IdIndividuo");
            }
        }
    }
}
