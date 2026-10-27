using System;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class MigrationV57V58V59Tests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV57V58V59Tests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV575859Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void Migrations_Metadata_ShouldBeCorrect()
        {
            var v57 = new MigrationV57_RepairSalvedadesConflictos();
            v57.Version.Should().Be(57);
            v57.Name.Should().Be("RepairSalvedadesConflictos");
            v57.Aditiva.Should().BeTrue();

            var v58 = new MigrationV58_EnsureTerceroVacante();
            v58.Version.Should().Be(58);
            v58.Name.Should().Be("EnsureTerceroVacante");
            v58.Aditiva.Should().BeTrue();

            var v59 = new MigrationV59_RecreateAgViews();
            v59.Version.Should().Be(59);
            v59.Name.Should().Be("RecreateAgViews");
            v59.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void MigrationV58_Up_WhenExecutedTwice_ShouldBeIdempotentAndActivateSentinel()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                conn.Execute(@"
                    CREATE TABLE terceros (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        nombre TEXT,
                        cedula TEXT,
                        residencia TEXT,
                        correo TEXT,
                        telefono TEXT,
                        activo INTEGER DEFAULT 1,
                        subunidad TEXT,
                        tipo TEXT,
                        categoria TEXT
                    );");

                var v58 = new MigrationV58_EnsureTerceroVacante();

                // Primera ejecución
                using (var tx1 = conn.BeginTransaction())
                {
                    v58.Up(conn, tx1);
                    tx1.Commit();
                }

                int count1 = conn.ExecuteScalar<int>("SELECT COUNT(*) FROM terceros WHERE cedula = '1111111111';");
                count1.Should().Be(1);

                // Segunda ejecución (idempotencia)
                using (var tx2 = conn.BeginTransaction())
                {
                    v58.Up(conn, tx2);
                    tx2.Commit();
                }

                int count2 = conn.ExecuteScalar<int>("SELECT COUNT(*) FROM terceros WHERE cedula = '1111111111';");
                count2.Should().Be(1);
            }
        }

        [Fact]
        public void MigrationV58_Up_WhenStaleCrsqlTriggersExistOnTerceros_ShouldDropTriggersAndInsertVacanteSuccessfully()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                conn.Execute(@"
                    CREATE TABLE terceros (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        nombre TEXT,
                        cedula TEXT,
                        residencia TEXT,
                        correo TEXT,
                        telefono TEXT,
                        activo INTEGER DEFAULT 1,
                        subunidad TEXT,
                        tipo TEXT,
                        categoria TEXT
                    );

                    -- Simular triggers stale de CR-SQLite que lanzan 'expected 43 values, got 5'
                    CREATE TRIGGER [terceros__crsql_itrig] BEFORE INSERT ON terceros
                    BEGIN
                        SELECT RAISE(FAIL, 'SQL logic error: expected 43 values, got 5');
                    END;
                    CREATE TRIGGER [terceros__crsql_utrig] BEFORE UPDATE ON terceros
                    BEGIN
                        SELECT RAISE(FAIL, 'SQL logic error: expected 43 values, got 5');
                    END;
                    CREATE TRIGGER [terceros__crsql_dtrig] BEFORE DELETE ON terceros
                    BEGIN
                        SELECT RAISE(FAIL, 'SQL logic error: expected 43 values, got 5');
                    END;
                ");

                var v58 = new MigrationV58_EnsureTerceroVacante();

                using (var tx = conn.BeginTransaction())
                {
                    v58.Up(conn, tx);
                    tx.Commit();
                }

                int count = conn.ExecuteScalar<int>("SELECT COUNT(*) FROM terceros WHERE cedula = '1111111111' AND activo = 1;");
                count.Should().Be(1);

                // Verificar que los triggers stale fueron efectivamente eliminados
                int triggerCount = conn.ExecuteScalar<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='trigger' AND name LIKE 'terceros__crsql_%';");
                triggerCount.Should().Be(0);
            }
        }

        [Fact]
        public void MigrationV58_Up_WhenStaleCrsqlTriggerFailsWithNoSuchColumn_ShouldDropTriggersAndInsertVacanteSuccessfully()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                conn.Execute(@"
                    CREATE TABLE terceros (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        nombre TEXT,
                        cedula TEXT,
                        residencia TEXT,
                        correo TEXT,
                        telefono TEXT,
                        activo INTEGER DEFAULT 1,
                        subunidad TEXT,
                        tipo TEXT,
                        categoria TEXT
                    );

                    -- Simular trigger stale que referencia una columna eliminada
                    CREATE TRIGGER [terceros__crsql_itrig] BEFORE INSERT ON terceros
                    BEGIN
                        SELECT RAISE(FAIL, 'SQL logic error: no such column: columna_eliminada');
                    END;
                ");

                var v58 = new MigrationV58_EnsureTerceroVacante();

                using (var tx = conn.BeginTransaction())
                {
                    v58.Up(conn, tx);
                    tx.Commit();
                }

                int count = conn.ExecuteScalar<int>("SELECT COUNT(*) FROM terceros WHERE cedula = '1111111111' AND activo = 1;");
                count.Should().Be(1);
            }
        }

        [Fact]
        public async Task MigrationV59_Up_ShouldCreateViewsWithFullContract()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

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

                var v59 = new MigrationV59_RecreateAgViews();

                using (var tx = conn.BeginTransaction())
                {
                    v59.Up(conn, tx);
                    tx.Commit();
                }

                var cols = (await conn.QueryAsync<string>("SELECT name FROM pragma_table_info('vw_AgSesionesDetalle');")).ToList();
                cols.Should().Contain(new[] { "IdSesion", "IdCorrida", "IdTercero", "EsVacante", "MotivoVacante", "CandidatosPool", "IdAula", "IdGen" });

                int tableSalvedades = conn.ExecuteScalar<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_salvedades_conflictos';");
                tableSalvedades.Should().Be(1);
            }
        }
    }
}
