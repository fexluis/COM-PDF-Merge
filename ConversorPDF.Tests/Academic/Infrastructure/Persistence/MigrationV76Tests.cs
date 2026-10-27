using System;
using System.Data.SQLite;
using System.IO;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class MigrationV76Tests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV76Tests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV76Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void MigrationV76_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV76_SanearDimensionesAgAula();

            // Assert
            mig.Version.Should().Be(76);
            mig.Name.Should().Be("SanearDimensionesAgAula");
            mig.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void Up_WhenAulaLacksDimensionsAndNodoHasDimensions_ShouldSanitizeAgAulaAndBeIdempotent()
        {
            // Arrange
            var mig = new MigrationV76_SanearDimensionesAgAula();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                conn.Execute(@"
CREATE TABLE ag_aula (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aula_ref TEXT NOT NULL DEFAULT '',
    jornada TEXT NOT NULL DEFAULT '',
    horas_dia INTEGER NOT NULL DEFAULT 0,
    dias_ventana TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT ''
);

CREATE TABLE ag_arbol_nodo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idarbol INTEGER NOT NULL DEFAULT 1,
    aula_ref TEXT NOT NULL DEFAULT '',
    jornada TEXT NOT NULL DEFAULT '',
    horas_dia INTEGER NOT NULL DEFAULT 0,
    dias_ventana TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT ''
);");

                // Aula 1: Sin dimensiones, nodo con configuración
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES ('3|1|1|III|A|PRESENCIAL', '', 0, '', 'SEDE1');

INSERT INTO ag_arbol_nodo (idarbol, aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (1, '3|1|1|III|A|PRESENCIAL', 'NOCHE', 4, 'VIE-SAB', 'SEDE1');");

                // Aula 2: Sin nodo coincidente
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES ('3|1|2|IV|B|PRESENCIAL', '', 0, '', 'SEDE1');");

                // Aula 3: Subunidad distinta a la del nodo
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES ('3|1|3|V|A|PRESENCIAL', '', 0, '', 'SUB_A');

INSERT INTO ag_arbol_nodo (idarbol, aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (1, '3|1|3|V|A|PRESENCIAL', 'MANANA', 8, 'MAR-VIE', 'SUB_B');");

                // Act: Primera ejecución
                mig.Up(conn, null);

                // Assert Aula 1 (Saneada)
                var aula1 = conn.QueryFirstOrDefault<(string jornada, int horas_dia, string dias_ventana)>(
                    "SELECT jornada, horas_dia, dias_ventana FROM ag_aula WHERE aula_ref = '3|1|1|III|A|PRESENCIAL';");
                aula1.jornada.Should().Be("NOCHE");
                aula1.horas_dia.Should().Be(4);
                aula1.dias_ventana.Should().Be("VIE-SAB");

                // Assert Aula 2 (Sin nodo: no inventa nada)
                var aula2 = conn.QueryFirstOrDefault<(string jornada, int horas_dia, string dias_ventana)>(
                    "SELECT jornada, horas_dia, dias_ventana FROM ag_aula WHERE aula_ref = '3|1|2|IV|B|PRESENCIAL';");
                aula2.jornada.Should().Be(string.Empty);
                aula2.horas_dia.Should().Be(0);
                aula2.dias_ventana.Should().Be(string.Empty);

                // Assert Aula 3 (Subunidad diferente: no cruza datos)
                var aula3 = conn.QueryFirstOrDefault<(string jornada, int horas_dia, string dias_ventana)>(
                    "SELECT jornada, horas_dia, dias_ventana FROM ag_aula WHERE aula_ref = '3|1|3|V|A|PRESENCIAL';");
                aula3.jornada.Should().Be(string.Empty);
                aula3.horas_dia.Should().Be(0);
                aula3.dias_ventana.Should().Be(string.Empty);

                // Act: Segunda ejecución (idempotencia)
                mig.Up(conn, null);

                var aula1Post = conn.QueryFirstOrDefault<(string jornada, int horas_dia, string dias_ventana)>(
                    "SELECT jornada, horas_dia, dias_ventana FROM ag_aula WHERE aula_ref = '3|1|1|III|A|PRESENCIAL';");
                aula1Post.Should().BeEquivalentTo(aula1);
            }
        }

        [Fact]
        public void Up_WhenAulaHasEmptySubunidadAndMultipleArboles_ShouldHealFromMaxIdarbolNodo()
        {
            // Arrange (Fases 1196 y 1197)
            var mig = new MigrationV76_SanearDimensionesAgAula();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                conn.Execute(@"
CREATE TABLE ag_aula (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    aula_ref TEXT NOT NULL DEFAULT '',
    jornada TEXT NOT NULL DEFAULT '',
    horas_dia INTEGER NOT NULL DEFAULT 0,
    dias_ventana TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT ''
);

CREATE TABLE ag_arbol_nodo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idarbol INTEGER NOT NULL DEFAULT 1,
    aula_ref TEXT NOT NULL DEFAULT '',
    jornada TEXT NOT NULL DEFAULT '',
    horas_dia INTEGER NOT NULL DEFAULT 0,
    dias_ventana TEXT NOT NULL DEFAULT '',
    subunidad TEXT NOT NULL DEFAULT ''
);");

                // ag_aula legacy con subunidad vacía
                conn.Execute(@"
INSERT INTO ag_aula (aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES ('3|1|4|I|A|PRESENCIAL', '', 0, '', '');");

                // Dos nodos en distintos árboles: idarbol=1 (perfil viejo) e idarbol=5 (perfil vigente)
                conn.Execute(@"
INSERT INTO ag_arbol_nodo (idarbol, aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (1, '3|1|4|I|A|PRESENCIAL', 'MANANA', 4, 'LUN-MIE', 'SEDE_CENTRAL');

INSERT INTO ag_arbol_nodo (idarbol, aula_ref, jornada, horas_dia, dias_ventana, subunidad)
VALUES (5, '3|1|4|I|A|PRESENCIAL', 'NOCHE', 4, 'JUE-SAB', 'SEDE_CENTRAL');");

                // Act
                mig.Up(conn, null);

                // Assert: Debe curar la fila legacy eligiendo el nodo de idarbol=5 (NOCHE, JUE-SAB)
                var aula = conn.QueryFirstOrDefault<(string jornada, int horas_dia, string dias_ventana)>(
                    "SELECT jornada, horas_dia, dias_ventana FROM ag_aula WHERE aula_ref = '3|1|4|I|A|PRESENCIAL';");
                aula.jornada.Should().Be("NOCHE");
                aula.horas_dia.Should().Be(4);
                aula.dias_ventana.Should().Be("JUE-SAB");
            }
        }
    }
}
