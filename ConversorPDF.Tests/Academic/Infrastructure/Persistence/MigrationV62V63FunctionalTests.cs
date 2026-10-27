using System;
using System.Data.SQLite;
using System.IO;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class MigrationV62V63FunctionalTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV62V63FunctionalTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV62V63Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void Up_WhenV62Executes_ShouldAddColumnsAndSeed42Bloques()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE ag_aula (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        codigo TEXT DEFAULT ''
                    );
                    CREATE TABLE ag_arbol_nodo (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        aula_ref TEXT DEFAULT ''
                    );
                    CREATE TABLE ag_bloque_horario (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        dia_semana INTEGER NOT NULL DEFAULT 1,
                        hora_inicio TEXT NOT NULL DEFAULT '08:00',
                        hora_fin TEXT NOT NULL DEFAULT '12:00',
                        duracion_horas REAL NOT NULL DEFAULT 4.0,
                        franja TEXT NOT NULL DEFAULT 'MANANA',
                        activo INTEGER NOT NULL DEFAULT 1,
                        subunidad TEXT NOT NULL DEFAULT ''
                    );
                ");

                var v62 = new MigrationV62_AddAgAulaVentanasYBloquesSemilla();
                using (var tx = conn.BeginTransaction())
                {
                    v62.Up(conn, tx);
                    tx.Commit();
                }

                // Verificar columnas
                int colJornadaAula = conn.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM pragma_table_info('ag_aula') WHERE name='jornada';");
                int colHorasDiaAula = conn.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM pragma_table_info('ag_aula') WHERE name='horas_dia';");
                int colDiasVentanaAula = conn.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM pragma_table_info('ag_aula') WHERE name='dias_ventana';");

                colJornadaAula.Should().Be(1);
                colHorasDiaAula.Should().Be(1);
                colDiasVentanaAula.Should().Be(1);

                // Verificar 30 bloques (Fase 1006: 5 plantillas x 6 días, sin domingo)
                int totalBloques = conn.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM ag_bloque_horario;");
                totalBloques.Should().Be(30);

                // Idempotencia: ejecutar segunda vez
                using (var tx = conn.BeginTransaction())
                {
                    v62.Up(conn, tx);
                    tx.Commit();
                }
                int totalBloques2 = conn.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM ag_bloque_horario;");
                totalBloques2.Should().Be(30);
            }
        }

        [Fact]
        public void Up_WhenV63Executes_ShouldUnifyHorariosToMilitar()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE sesiones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idvinculacion INTEGER,
                        horariosesion TEXT
                    );
                    CREATE TABLE vinculaciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        dia_horas_semanales TEXT
                    );
                    INSERT INTO sesiones (idvinculacion, horariosesion) VALUES (1, '1:00 PM A 5:00 PM');
                    INSERT INTO vinculaciones (id, dia_horas_semanales) VALUES (1, 'MARTES|4|1:00 PM A 5:00 PM');
                ");

                var v63 = new MigrationV63_UnificarHorariosMilitar();
                using (var tx = conn.BeginTransaction())
                {
                    v63.Up(conn, tx);
                    tx.Commit();
                }

                string horarioSesion = conn.QueryFirstOrDefault<string>("SELECT horariosesion FROM sesiones WHERE id = 1;");
                string diaHoras = conn.QueryFirstOrDefault<string>("SELECT dia_horas_semanales FROM vinculaciones WHERE id = 1;");

                horarioSesion.Should().Be("13:00 A 17:00");
                diaHoras.Should().Be("MARTES|4|13:00 A 17:00");
            }
        }
    }
}
