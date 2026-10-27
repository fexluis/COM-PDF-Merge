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
    public class MigrationV67_RecrearVistasAgCanonicasTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV67_RecrearVistasAgCanonicasTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV67Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void MigrationV67_Metadata_ShouldBeCorrect()
        {
            var v67 = new MigrationV67_RecrearVistasAgCanonicas();
            v67.Version.Should().Be(67);
            v67.Name.Should().Be("RecrearVistasAgCanonicas");
            v67.Aditiva.Should().BeTrue();
        }

        [Fact]
        public async Task MigrationV67_Up_ShouldIncludeDiaSemanaAndFilterOnlyDiscarded()
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

                // Datos de prueba:
                // Vinculación 2: estado = 3 (descartada) -> NO DEBE APARECER en vw_AgSesionesDetalle
                // Vinculación 3: estado = 1 (aprobada/otra) -> DEBE APARECER en vw_AgSesionesDetalle
                // vw_AgVinculacionesDetalle incluye ambas (no filtra por estado)
                conn.Execute(@"
                    INSERT INTO ag_vinculaciones (id, estado) VALUES (2, 3);
                    INSERT INTO ag_vinculaciones (id, estado) VALUES (3, 1);

                    INSERT INTO ag_sesiones (id, idvinculacion, idtercero, fechasesion) 
                    VALUES (102, 2, 20, '2026-03-03');

                    -- Fases 866 y 869: Probar normalización de fecha legacy dd/MM/yyyy a ISO yyyy-MM-dd
                    INSERT INTO ag_sesiones (id, idvinculacion, idtercero, fechasesion)
                    VALUES (103, 3, 30, '03/03/2026');
                ");

                var v67 = new MigrationV67_RecrearVistasAgCanonicas();
                var v68 = new MigrationV68_NormalizarFechasAgSesiones();

                using (var tx = conn.BeginTransaction())
                {
                    v67.Up(conn, tx);
                    v68.Up(conn, tx);
                    tx.Commit();
                }

                // 1. Verificar columnas de la vista vw_AgSesionesDetalle
                var cols = (await conn.QueryAsync<string>("SELECT name FROM pragma_table_info('vw_AgSesionesDetalle');")).ToList();
                cols.Should().Contain(new[] { "DiaSemana", "IdAula", "IdGen", "EsVacante", "MotivoVacante", "CandidatosPool", "DiaHorasSemanales" });

                // 2. Fases 866, 869 y 875: Verificar normalización de fecha legacy en ag_sesiones y proyección en vw_AgSesionesDetalle
                var legacy = await conn.QuerySingleAsync<string>(
                    "SELECT fechasesion FROM ag_sesiones WHERE id = 103;");
                legacy.Should().Be("2026-03-03");

                var diaLegacy = await conn.QuerySingleAsync<string>(
                    "SELECT DiaSemana FROM vw_AgSesionesDetalle WHERE IdSesion = 103;");
                diaLegacy.Should().Be("MARTES");

                // Fase 877: Aserción determinista con ORDER BY IdSesion evitando dependencia de orden físico
                var sesiones = (await conn.QueryAsync<dynamic>(
                    "SELECT IdSesion, DiaSemana FROM vw_AgSesionesDetalle ORDER BY IdSesion;")).ToList();

                sesiones.Select(s => (int)s.IdSesion).Should().BeEquivalentTo(new[] { 103 });
                sesiones.Select(s => (string)s.DiaSemana).Should().OnlyContain(d => d == "MARTES");

                // 3. Fases 865 y 871: Verificar columnas de la vista vw_AgVinculacionesDetalle
                var colsVinc = (await conn.QueryAsync<string>("SELECT name FROM pragma_table_info('vw_AgVinculacionesDetalle');")).ToList();
                colsVinc.Should().Contain(new[]
                {
                    "IdVinculacion", "IdCorrida", "IdIndividuo", "IdPrograma", "IdAsignatura",
                    "Grupo", "Estado", "DiaHorasSemanales", "TotalHoras",
                    "EsVacante", "MotivoVacante", "CandidatosPool"
                });

                // Fase 869: Reemplazar la aserción débil (SELECT * sin LIMIT 1 y verificar conteo exacto de vinculaciones)
                var vinc = (await conn.QueryAsync<dynamic>("SELECT * FROM vw_AgVinculacionesDetalle;")).ToList();
                vinc.Should().HaveCount(2); // estado 1 y estado 3 (la vista de vinculaciones no filtra)
            }
        }
    }
}
