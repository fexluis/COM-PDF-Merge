using System;
using System.Data.SQLite;
using System.IO;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class MigrationV68_NormalizarFechasAgSesionesTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV68_NormalizarFechasAgSesionesTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV68Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void MigrationV68_Metadata_ShouldBeCorrect()
        {
            var v68 = new MigrationV68_NormalizarFechasAgSesiones();
            v68.Version.Should().Be(68);
            v68.Name.Should().Be("NormalizarFechasAgSesiones");
            v68.Aditiva.Should().BeTrue();
        }

        [Fact]
        public async Task MigrationV68_Up_ShouldNormalizeLegacyFormatsAndIgnoreInvalidCalendar()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                conn.Execute(@"
                    CREATE TABLE ag_sesiones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idvinculacion INTEGER NOT NULL DEFAULT 0,
                        idtercero INTEGER NOT NULL DEFAULT 0,
                        fechasesion TEXT
                    );
                    CREATE TABLE ag_vinculaciones (
                        id INTEGER PRIMARY KEY,
                        fecha_inicio TEXT,
                        fecha_fin TEXT
                    );
                    CREATE TABLE ag_gen (
                        id INTEGER PRIMARY KEY,
                        fecha_inicio_estimada TEXT,
                        fecha_fin_estimada TEXT
                    );
                ");

                // Datos con diversos formatos legacy y casos de borde
                conn.Execute(@"
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (1, '15/04/2026');         -- dd/MM/yyyy
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (2, '3/4/2026');           -- d/M/yyyy
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (3, '05-06-2026');         -- dd-MM-yyyy
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (4, '10/05/2026 14:30');   -- dd/MM/yyyy HH:mm
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (5, '2026-07-08');         -- Ya ISO
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (6, '31/02/2026');         -- Fecha de calendario inválida
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (7, '2026/09/10');         -- yyyy/MM/dd
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (8,  '12/3/2026');    -- dd/M/yyyy
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (9,  '3/12/2026');    -- d/MM/yyyy
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (10, '2026-02-31');   -- ISO imposible
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (11, '2026-05-10 14:30'); -- ISO con hora
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (12, '2026-4-5');          -- yyyy-M-d
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (13, '2026-05-10 00:00:00'); -- ISO con segundos
                    INSERT INTO ag_sesiones (id, fechasesion) VALUES (14, '10/05/2026 14:30:55'); -- dd/MM/yyyy con segundos

                    INSERT INTO ag_vinculaciones (id, fecha_inicio, fecha_fin)
                    VALUES (1, '10/09/2026', '20/12/2026');
                    INSERT INTO ag_vinculaciones (id, fecha_inicio, fecha_fin)
                    VALUES (2, '31/02/2026', '--');
                    INSERT INTO ag_vinculaciones (id, fecha_inicio, fecha_fin)
                    VALUES (3, '2026-05-10 00:00:00', '20/12/2026');

                    INSERT INTO ag_gen (id, fecha_inicio_estimada, fecha_fin_estimada)
                    VALUES (1, '10/05/2026 14:30', '2026/11/15');
                    INSERT INTO ag_gen (id, fecha_inicio_estimada, fecha_fin_estimada)
                    VALUES (2, '2026-05-10 00:00:00', '2026-12-20');
                ");

                var v68 = new MigrationV68_NormalizarFechasAgSesiones();

                using (var tx = conn.BeginTransaction())
                {
                    v68.Up(conn, tx);
                    tx.Commit();
                }

                var f1 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 1;");
                f1.Should().Be("2026-04-15");

                var f2 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 2;");
                f2.Should().Be("2026-04-03");

                var f3 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 3;");
                f3.Should().Be("2026-06-05");

                var f4 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 4;");
                f4.Should().Be("2026-05-10");

                var f5 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 5;");
                f5.Should().Be("2026-07-08");

                // La fecha de calendario inválida (31 de febrero) no debe haber mutado
                var f6 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 6;");
                f6.Should().Be("31/02/2026");

                // Fase 880: Formato yyyy/MM/dd
                var f7 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 7;");
                f7.Should().Be("2026-09-10");

                // Fase 882: Variantes de 9 caracteres y fecha ISO imposible
                var f8 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 8;");
                f8.Should().Be("2026-03-12");

                var f9 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 9;");
                f9.Should().Be("2026-12-03");

                var f10 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 10;");
                f10.Should().Be("2026-02-31");   // intacta

                // Fase 884: Formatos ISO con hora y componentes de 1 dígito
                var f11 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 11;");
                f11.Should().Be("2026-05-10");

                var f12 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 12;");
                f12.Should().Be("2026-04-05");

                // Fases 879 y 880: Normalización de ag_vinculaciones
                var vinc = await conn.QuerySingleAsync<(string, string)>(
                    "SELECT fecha_inicio, fecha_fin FROM ag_vinculaciones WHERE id = 1;");
                vinc.Item1.Should().Be("2026-09-10");
                vinc.Item2.Should().Be("2026-12-20");

                // Fase 884: Vinculación con fecha imposible y marcador '--'
                var v2 = await conn.QuerySingleAsync<(string, string)>(
                    "SELECT fecha_inicio, fecha_fin FROM ag_vinculaciones WHERE id = 2;");
                v2.Item1.Should().Be("31/02/2026");  // imposible, intacta
                v2.Item2.Should().Be("--");          // marcador preservado

                // Fase 885: Normalización de ag_gen
                var g1 = await conn.QuerySingleAsync<(string, string)>(
                    "SELECT fecha_inicio_estimada, fecha_fin_estimada FROM ag_gen WHERE id = 1;");
                g1.Item1.Should().Be("2026-05-10");
                g1.Item2.Should().Be("2026-11-15");

                // Fase 888: Formatos con segundos y normalización de vinculaciones y genoma con tiempo
                var f13 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 13;");
                f13.Should().Be("2026-05-10");

                var f14 = await conn.QuerySingleAsync<string>("SELECT fechasesion FROM ag_sesiones WHERE id = 14;");
                f14.Should().Be("2026-05-10");

                var v3 = await conn.QuerySingleAsync<(string, string)>(
                    "SELECT fecha_inicio, fecha_fin FROM ag_vinculaciones WHERE id = 3;");
                v3.Item1.Should().Be("2026-05-10");
                v3.Item2.Should().Be("2026-12-20");

                var g2 = await conn.QuerySingleAsync<(string, string)>(
                    "SELECT fecha_inicio_estimada, fecha_fin_estimada FROM ag_gen WHERE id = 2;");
                g2.Item1.Should().Be("2026-05-10");
                g2.Item2.Should().Be("2026-12-20");

                // Fases 881, 882, 884 y 886: Residual esperado: exactamente id 6 ('31/02/2026') e id 10 ('2026-02-31'); id 11 a 14 normalizados con éxito
                var residuales = await conn.QuerySingleAsync<int>(@"
                    SELECT COUNT(*) FROM ag_sesiones
                     WHERE TRIM(IFNULL(fechasesion,'')) <> ''
                       AND (length(fechasesion) <> 10
                            OR date(fechasesion) IS NULL
                            OR date(fechasesion) <> fechasesion);");
                residuales.Should().Be(2);
            }
        }
    }
}
