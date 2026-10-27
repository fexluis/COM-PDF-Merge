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
    public class MigrationV66Tests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV66Tests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV66Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void MigrationV66_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV66_RenombrarModalidadTutoria();

            // Assert
            mig.Version.Should().Be(66);
            mig.Name.Should().Be("RenombrarModalidadTutoria");
            mig.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void Up_WhenTablesMissing_ShouldCompleteWithoutThrowing()
        {
            // Arrange: base de datos vacía sin tablas
            var mig = new MigrationV66_RenombrarModalidadTutoria();

            // Act & Assert
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    Action act = () => mig.Up(conn, tx);
                    act.Should().NotThrow();
                    tx.Commit();
                }
            }
        }

        [Fact]
        public void Up_WhenLegacyValuesPresent_ShouldNormalizeToPlataformaAndMixta()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE vinculaciones (
                        id INTEGER PRIMARY KEY,
                        modalidadVinculacion TEXT
                    );
                    CREATE TABLE ag_vinculaciones (
                        id INTEGER PRIMARY KEY,
                        modalidadVinculacion TEXT
                    );
                    CREATE TABLE ag_aula (
                        id INTEGER PRIMARY KEY,
                        codigo TEXT,
                        nombre TEXT,
                        tipo_aula TEXT,
                        es_virtual INTEGER,
                        modalidadVinculacion TEXT,
                        aula_ref TEXT
                    );
                    CREATE TABLE ag_arbol_nodo (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idarbol INTEGER,
                        modalidad TEXT,
                        aula_ref TEXT,
                        codigo TEXT,
                        nombre TEXT
                    );
                    CREATE UNIQUE INDEX uq_ag_arbol_nodo ON ag_arbol_nodo(idarbol, aula_ref);

                    -- vinculaciones
                    INSERT INTO vinculaciones (id, modalidadVinculacion) VALUES (1, 'VIRTUAL');
                    INSERT INTO vinculaciones (id, modalidadVinculacion) VALUES (2, 'Semipresencial');
                    INSERT INTO vinculaciones (id, modalidadVinculacion) VALUES (3, 'A Distancia');
                    INSERT INTO vinculaciones (id, modalidadVinculacion) VALUES (4, 'Presencial');
                    INSERT INTO vinculaciones (id, modalidadVinculacion) VALUES (5, '');

                    -- ag_vinculaciones
                    INSERT INTO ag_vinculaciones (id, modalidadVinculacion) VALUES (10, 'virtual');
                    INSERT INTO ag_vinculaciones (id, modalidadVinculacion) VALUES (20, 'SEMIPRESENCIAL');

                    -- ag_aula
                    INSERT INTO ag_aula (id, codigo, nombre, tipo_aula, es_virtual, modalidadVinculacion, aula_ref)
                    VALUES (100, '2|1|5|III|A|VIRTUAL', '2|1|5|III|A|VIRTUAL', 'VIRTUAL', 1, 'Virtual', '2|1|5|III|A|VIRTUAL');
                    INSERT INTO ag_aula (id, codigo, nombre, tipo_aula, es_virtual, modalidadVinculacion, aula_ref)
                    VALUES (101, '2|1|5|III|A|SEMIPRESENCIAL', '2|1|5|III|A|SEMIPRESENCIAL', 'ESTANDAR', 0, 'Semipresencial', '2|1|5|III|A|SEMIPRESENCIAL');

                    -- ag_arbol_nodo (caso normal)
                    INSERT INTO ag_arbol_nodo (id, idarbol, modalidad, aula_ref, codigo, nombre)
                    VALUES (200, 1, 'Virtual', '2|1|5|III|A|VIRTUAL', '2|1|5|III|A|VIRTUAL', '2|1|5|III|A|VIRTUAL');
                    INSERT INTO ag_arbol_nodo (id, idarbol, modalidad, aula_ref, codigo, nombre)
                    VALUES (201, 1, 'Semipresencial', '2|1|5|III|A|SEMIPRESENCIAL', '2|1|5|III|A|SEMIPRESENCIAL', '2|1|5|III|A|SEMIPRESENCIAL');
                ");

                var mig = new MigrationV66_RenombrarModalidadTutoria();

                // Act
                using (var tx = conn.BeginTransaction())
                {
                    mig.Up(conn, tx);
                    tx.Commit();
                }

                // Assert vinculaciones
                var vincs = conn.Query<(int Id, string Mod)>("SELECT id, modalidadVinculacion FROM vinculaciones ORDER BY id;").ToList();
                vincs.First(x => x.Id == 1).Mod.Should().Be("Plataforma");
                vincs.First(x => x.Id == 2).Mod.Should().Be("Mixta");
                vincs.First(x => x.Id == 3).Mod.Should().Be("Mixta");
                vincs.First(x => x.Id == 4).Mod.Should().Be("Presencial");
                vincs.First(x => x.Id == 5).Mod.Should().Be("");

                // Assert ag_vinculaciones
                var agVincs = conn.Query<(int Id, string Mod)>("SELECT id, modalidadVinculacion FROM ag_vinculaciones ORDER BY id;").ToList();
                agVincs.First(x => x.Id == 10).Mod.Should().Be("Plataforma");
                agVincs.First(x => x.Id == 20).Mod.Should().Be("Mixta");

                // Assert ag_aula
                var aula100 = conn.QuerySingle("SELECT * FROM ag_aula WHERE id = 100;");
                ((string)aula100.modalidadVinculacion).Should().Be("Plataforma");
                ((string)aula100.aula_ref).Should().Be("2|1|5|III|A|PLATAFORMA");
                ((string)aula100.codigo).Should().Be("2|1|5|III|A|PLATAFORMA");
                ((string)aula100.nombre).Should().Be("2|1|5|III|A|PLATAFORMA");
                ((string)aula100.tipo_aula).Should().Be("VIRTUAL"); // Intacto
                ((long)aula100.es_virtual).Should().Be(1); // Intacto

                var aula101 = conn.QuerySingle("SELECT * FROM ag_aula WHERE id = 101;");
                ((string)aula101.modalidadVinculacion).Should().Be("Mixta");
                ((string)aula101.aula_ref).Should().Be("2|1|5|III|A|MIXTA");
                ((string)aula101.codigo).Should().Be("2|1|5|III|A|MIXTA");
                ((string)aula101.nombre).Should().Be("2|1|5|III|A|MIXTA");

                // Assert ag_arbol_nodo
                var nodo200 = conn.QuerySingle("SELECT * FROM ag_arbol_nodo WHERE id = 200;");
                ((string)nodo200.modalidad).Should().Be("Plataforma");
                ((string)nodo200.aula_ref).Should().Be("2|1|5|III|A|PLATAFORMA");

                var nodo201 = conn.QuerySingle("SELECT * FROM ag_arbol_nodo WHERE id = 201;");
                ((string)nodo201.modalidad).Should().Be("Mixta");
                ((string)nodo201.aula_ref).Should().Be("2|1|5|III|A|MIXTA");

                // Idempotencia: segunda corrida no altera datos ni falla
                using (var tx2 = conn.BeginTransaction())
                {
                    Action actIdempotente = () => mig.Up(conn, tx2);
                    actIdempotente.Should().NotThrow();
                    tx2.Commit();
                }

                int residuales = conn.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM vinculaciones WHERE UPPER(modalidadVinculacion) IN ('VIRTUAL','SEMIPRESENCIAL');");
                residuales.Should().Be(0);
            }
        }

        [Fact]
        public void Up_WhenCollisionInAgArbolNodoUniqueIndex_ShouldDeduplicateLegacyRowGracefully()
        {
            // Arrange: simula caso donde en el mismo idarbol ya existe un nodo canónico y uno legado
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE ag_arbol_nodo (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idarbol INTEGER,
                        modalidad TEXT,
                        aula_ref TEXT,
                        codigo TEXT,
                        nombre TEXT
                    );
                    CREATE UNIQUE INDEX uq_ag_arbol_nodo ON ag_arbol_nodo(idarbol, aula_ref);

                    -- Fila canónica ya existente
                    INSERT INTO ag_arbol_nodo (id, idarbol, modalidad, aula_ref, codigo, nombre)
                    VALUES (1, 10, 'Plataforma', '1|2|3|I|A|PLATAFORMA', '1|2|3|I|A|PLATAFORMA', '1|2|3|I|A|PLATAFORMA');

                    -- Fila legada duplicada en el mismo perfil
                    INSERT INTO ag_arbol_nodo (id, idarbol, modalidad, aula_ref, codigo, nombre)
                    VALUES (2, 10, 'Virtual', '1|2|3|I|A|VIRTUAL', '1|2|3|I|A|VIRTUAL', '1|2|3|I|A|VIRTUAL');
                ");

                var mig = new MigrationV66_RenombrarModalidadTutoria();

                // Act
                using (var tx = conn.BeginTransaction())
                {
                    Action act = () => mig.Up(conn, tx);
                    act.Should().NotThrow();
                    tx.Commit();
                }

                // Assert: La fila legada 2 debe haberse eliminado para evitar violación de uq_ag_arbol_nodo
                var filas = conn.Query<long>("SELECT id FROM ag_arbol_nodo WHERE idarbol = 10;").ToList();
                filas.Should().HaveCount(1);
                filas.First().Should().Be(1);
            }
        }

        [Fact]
        public void Up_WhenMultipleLegacyRowsForSameCanonicAulaRef_ShouldDeduplicatePreservingLowestId()
        {
            // Arrange: simula caso donde en el mismo idarbol existen dos filas legadas (|VIRTUAL y |Virtual)
            // que normalizan al mismo string pero no colisionaban antes de la migración
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE ag_arbol_nodo (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idarbol INTEGER,
                        modalidad TEXT,
                        aula_ref TEXT,
                        codigo TEXT,
                        nombre TEXT
                    );
                    CREATE UNIQUE INDEX uq_ag_arbol_nodo ON ag_arbol_nodo(idarbol, aula_ref);

                    -- Dos filas legadas con diferente capitalización en el mismo idarbol
                    INSERT INTO ag_arbol_nodo (id, idarbol, modalidad, aula_ref, codigo, nombre)
                    VALUES (5, 20, 'Virtual', '2|2|3|I|A|VIRTUAL', '2|2|3|I|A|VIRTUAL', '2|2|3|I|A|VIRTUAL');

                    INSERT INTO ag_arbol_nodo (id, idarbol, modalidad, aula_ref, codigo, nombre)
                    VALUES (8, 20, 'Virtual', '2|2|3|I|A|Virtual', '2|2|3|I|A|Virtual', '2|2|3|I|A|Virtual');
                ");

                var mig = new MigrationV66_RenombrarModalidadTutoria();

                // Act
                using (var tx = conn.BeginTransaction())
                {
                    Action act = () => mig.Up(conn, tx);
                    act.Should().NotThrow();
                    tx.Commit();
                }

                // Assert: Se debe conservar la de menor id (5) y eliminar la de mayor id (8)
                var filas = conn.Query<long>("SELECT id FROM ag_arbol_nodo WHERE idarbol = 20;").ToList();
                filas.Should().HaveCount(1);
                filas.First().Should().Be(5);

                var nodo5 = conn.QuerySingle("SELECT aula_ref, modalidad FROM ag_arbol_nodo WHERE id = 5;");
                ((string)nodo5.aula_ref).Should().Be("2|2|3|I|A|PLATAFORMA");
                ((string)nodo5.modalidad).Should().Be("Plataforma");
            }
        }

        [Fact]
        public void Up_WhenAgAulaHasPartialColumns_ShouldUpdateExistingColumnsWithoutThrowing()
        {
            // Arrange: tabla ag_aula con solo aula_ref y modalidadVinculacion (sin columnas codigo ni nombre)
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE ag_aula (
                        id INTEGER PRIMARY KEY,
                        modalidadVinculacion TEXT,
                        aula_ref TEXT
                    );
                    INSERT INTO ag_aula (id, modalidadVinculacion, aula_ref)
                    VALUES (1, 'Virtual', '1|1|1|I|A|VIRTUAL');
                ");

                var mig = new MigrationV66_RenombrarModalidadTutoria();

                // Act
                using (var tx = conn.BeginTransaction())
                {
                    Action act = () => mig.Up(conn, tx);
                    act.Should().NotThrow();
                    tx.Commit();
                }

                // Assert
                var fila = conn.QuerySingle("SELECT modalidadVinculacion, aula_ref FROM ag_aula WHERE id = 1;");
                ((string)fila.modalidadVinculacion).Should().Be("Plataforma");
                ((string)fila.aula_ref).Should().Be("1|1|1|I|A|PLATAFORMA");
            }
        }

        [Fact]
        public void Up_WhenAgArbolNodoLacksModalidadColumn_ShouldNotThrowAndStillRenameAulaRef()
        {
            // Arrange: tabla ag_arbol_nodo con solo idarbol, aula_ref, codigo, nombre (sin modalidad)
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE ag_arbol_nodo (
                        id INTEGER PRIMARY KEY,
                        idarbol INTEGER,
                        aula_ref TEXT,
                        codigo TEXT,
                        nombre TEXT
                    );
                    INSERT INTO ag_arbol_nodo (id, idarbol, aula_ref, codigo, nombre)
                    VALUES (1, 10, '1|2|3|I|A|VIRTUAL', '1|2|3|I|A|VIRTUAL', '1|2|3|I|A|VIRTUAL');
                ");

                var mig = new MigrationV66_RenombrarModalidadTutoria();

                // Act
                using (var tx = conn.BeginTransaction())
                {
                    Action act = () => mig.Up(conn, tx);
                    act.Should().NotThrow();
                    tx.Commit();
                }

                // Assert
                var fila = conn.QuerySingle("SELECT aula_ref, codigo, nombre FROM ag_arbol_nodo WHERE id = 1;");
                ((string)fila.aula_ref).Should().Be("1|2|3|I|A|PLATAFORMA");
                ((string)fila.codigo).Should().Be("1|2|3|I|A|PLATAFORMA");
                ((string)fila.nombre).Should().Be("1|2|3|I|A|PLATAFORMA");
            }
        }

        [Fact]
        public void Up_WhenBlockFails_ShouldThrowInvalidOperationExceptionWithInnerException()
        {
            // Arrange: crear tabla con trigger que aborta intencionalmente ante UPDATE
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE vinculaciones (
                        id INTEGER PRIMARY KEY,
                        modalidadVinculacion TEXT
                    );
                    INSERT INTO vinculaciones (id, modalidadVinculacion) VALUES (1, 'VIRTUAL');

                    -- Trigger intencional para verificar rollback ante fallo en bloque (Fase 821).
                    -- Este error es esperado en pruebas y su log se redirige a %TEMP% sin contaminar producción (Fases 1066-1068).
                    CREATE TRIGGER trg_fail BEFORE UPDATE ON vinculaciones
                    BEGIN
                        SELECT RAISE(FAIL, '[TEST-V66] fallo intencional');
                    END;
                ");

                var mig = new MigrationV66_RenombrarModalidadTutoria();

                // Act & Assert
                using (var tx = conn.BeginTransaction())
                {
                    Action act = () => mig.Up(conn, tx);
                    var ex = act.Should().Throw<InvalidOperationException>().Which;
                    ex.Message.Should().Contain("bloque(s) fallaron");
                    ex.InnerException.Should().NotBeNull();
                    ex.InnerException.Message.Should().Contain("[TEST-V66] fallo intencional");
                }
            }
        }

        [Fact]
        public void Up_WhenNoModalidadColumnAndCanonicRowHasHigherId_ShouldKeepCanonicRow()
        {
            // Arrange: ag_arbol_nodo sin columna modalidad pero con índice único (idarbol, aula_ref)
            // Fila 1 (legacy, id=1 menor) vs Fila 2 (canónica, id=2 mayor)
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE ag_arbol_nodo (
                        id INTEGER PRIMARY KEY,
                        idarbol INTEGER,
                        aula_ref TEXT,
                        codigo TEXT,
                        nombre TEXT
                    );
                    CREATE UNIQUE INDEX uq_ag_arbol_nodo ON ag_arbol_nodo(idarbol, aula_ref);
                    INSERT INTO ag_arbol_nodo (id, idarbol, aula_ref, codigo, nombre)
                    VALUES (1, 10, '1|2|3|I|A|VIRTUAL', '1|2|3|I|A|VIRTUAL', '1|2|3|I|A|VIRTUAL');
                    INSERT INTO ag_arbol_nodo (id, idarbol, aula_ref, codigo, nombre)
                    VALUES (2, 10, '1|2|3|I|A|PLATAFORMA', '1|2|3|I|A|PLATAFORMA', '1|2|3|I|A|PLATAFORMA');
                ");

                var mig = new MigrationV66_RenombrarModalidadTutoria();

                // Act
                using (var tx = conn.BeginTransaction())
                {
                    Action act = () => mig.Up(conn, tx);
                    act.Should().NotThrow();
                    tx.Commit();
                }

                // Assert: la deduplicación debe conservar la fila canónica (id=2) y eliminar la legacy (id=1)
                var ids = conn.Query<long>("SELECT id FROM ag_arbol_nodo WHERE idarbol = 10;").ToList();
                ids.Should().HaveCount(1);
                ids[0].Should().Be(2);

                var aulaRef = conn.QuerySingle<string>("SELECT aula_ref FROM ag_arbol_nodo WHERE id = 2;");
                aulaRef.Should().Be("1|2|3|I|A|PLATAFORMA");
            }
        }
    }
}
