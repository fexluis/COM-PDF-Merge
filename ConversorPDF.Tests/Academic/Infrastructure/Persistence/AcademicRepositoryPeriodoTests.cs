using System;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AcademicRepositoryPeriodoTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public AcademicRepositoryPeriodoTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AcademicRepoPeriodoTest_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";

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
                    CREATE TABLE vinculaciones (
                        id INTEGER PRIMARY KEY,
                        idperiodo INTEGER
                    );
                    CREATE TABLE festivos (
                        idfestivo INTEGER PRIMARY KEY,
                        idperiodo INTEGER
                    );
                    CREATE TABLE config (
                        id INTEGER PRIMARY KEY,
                        id_periodo INTEGER
                    );
                    CREATE TABLE ag_periodo_ventana_auditoria (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idperiodo INTEGER NOT NULL,
                        periodo TEXT,
                        inicio_antes TEXT,
                        fin_antes TEXT,
                        inicio_despues TEXT,
                        fin_despues TEXT,
                        actor TEXT NOT NULL,
                        origen TEXT NOT NULL,
                        motivo TEXT,
                        fecha TEXT NOT NULL
                    );
                    CREATE INDEX idx_ag_periodo_ventana_aud ON ag_periodo_ventana_auditoria(idperiodo);
                ");
            }
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
        public void InsertAndUpdatePeriodo_WhenValid_ShouldPersistChanges()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                var repo = new AcademicRepository(conn, subunidad: "TEST", suppressMarkDirty: true);

                // Act 1: Insert
                var nuevo = new Periodo
                {
                    IdPeriodo = 10,
                    PeriodoNombre = "2028-01",
                    PeriodoInicio = "2028-01-15",
                    PeriodoFin = "2028-06-30",
                    Activo = true
                };
                var insRes = repo.InsertPeriodo(nuevo);
                insRes.IsSuccess.Should().BeTrue();

                // Act 2: Update
                nuevo.PeriodoNombre = "2028-01-REV";
                nuevo.PeriodoFin = "2028-07-15";
                var updRes = repo.UpdatePeriodo(nuevo);
                updRes.IsSuccess.Should().BeTrue();

                // Assert
                var getRes = repo.GetPeriodoById(10);
                getRes.IsSuccess.Should().BeTrue();
                getRes.Value.PeriodoNombre.Should().Be("2028-01-REV");
                getRes.Value.PeriodoFin.Should().Be("2028-07-15");
            }
        }

        [Fact]
        public void DeletePeriodo_WhenReferencedByVinculaciones_ShouldPreventDeletion()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute("INSERT INTO periodos (idperiodo, periodo, activo) VALUES (5, '2025-01', 1);");
                conn.Execute("INSERT INTO vinculaciones (id, idperiodo) VALUES (1, 5);");

                var repo = new AcademicRepository(conn, subunidad: "TEST", suppressMarkDirty: true);

                // Act
                var result = repo.DeletePeriodo(5);

                // Assert
                result.IsSuccess.Should().BeFalse();
                result.Error.Should().Contain("vinculaciones académicas asociadas");
            }
        }

        [Fact]
        public void DeletePeriodo_WhenReferencedByFestivos_ShouldPreventDeletion()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute("INSERT INTO periodos (idperiodo, periodo, activo) VALUES (6, '2025-02', 1);");
                conn.Execute("INSERT INTO festivos (idfestivo, idperiodo) VALUES (1, 6);");

                var repo = new AcademicRepository(conn, subunidad: "TEST", suppressMarkDirty: true);

                // Act
                var result = repo.DeletePeriodo(6);

                // Assert
                result.IsSuccess.Should().BeFalse();
                result.Error.Should().Contain("días festivos configurados");
            }
        }

        [Fact]
        public void DeletePeriodo_WhenReferencedByConfig_ShouldPreventDeletion()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute("INSERT INTO periodos (idperiodo, periodo, activo) VALUES (7, '2025-03', 1);");
                conn.Execute("INSERT INTO config (id, id_periodo) VALUES (1, 7);");

                var repo = new AcademicRepository(conn, subunidad: "TEST", suppressMarkDirty: true);

                // Act
                var result = repo.DeletePeriodo(7);

                // Assert
                result.IsSuccess.Should().BeFalse();
                result.Error.Should().Contain("período activo en el sistema");
            }
        }

        [Fact]
        public void UpdatePeriodo_WhenEsquemaSinVentanaYConFechas_ShouldReturnFailure()
        {
            string dbSinVentanaPath = Path.Combine(Path.GetTempPath(), "SinVentana_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(dbSinVentanaPath);
            try
            {
                using (var conn = new SQLiteConnection($"Data Source={dbSinVentanaPath};Version=3;"))
                {
                    conn.Open();
                    conn.Execute(@"
                        CREATE TABLE periodos (
                            idperiodo INTEGER PRIMARY KEY,
                            periodo TEXT,
                            activo INTEGER DEFAULT 1
                        );
                    ");

                    var repo = new AcademicRepository(conn, subunidad: "TEST", suppressMarkDirty: true);
                    var periodo = new Periodo
                    {
                        IdPeriodo = 1,
                        PeriodoNombre = "2025-01",
                        PeriodoInicio = "2025-01-10",
                        PeriodoFin = "2025-06-15",
                        Activo = true
                    };

                    // Act
                    var result = repo.UpdatePeriodo(periodo);

                    // Assert
                    result.IsSuccess.Should().BeFalse();
                    result.Error.Should().Contain("no soporta ventana temporal");
                }
            }
            finally
            {
                try { if (File.Exists(dbSinVentanaPath)) File.Delete(dbSinVentanaPath); } catch { }
            }
        }

        [Fact]
        public void InsertPeriodo_WhenEsquemaSinVentanaYConFechas_ShouldReturnFailure()
        {
            string dbSinVentanaPath = Path.Combine(Path.GetTempPath(), "SinVentanaIns_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(dbSinVentanaPath);
            try
            {
                using (var conn = new SQLiteConnection($"Data Source={dbSinVentanaPath};Version=3;"))
                {
                    conn.Open();
                    conn.Execute(@"
                        CREATE TABLE periodos (
                            idperiodo INTEGER PRIMARY KEY,
                            periodo TEXT,
                            activo INTEGER DEFAULT 1
                        );
                    ");

                    var repo = new AcademicRepository(conn, subunidad: "TEST", suppressMarkDirty: true);
                    var periodo = new Periodo
                    {
                        IdPeriodo = 20,
                        PeriodoNombre = "2029-01",
                        PeriodoInicio = "2029-01-10",
                        PeriodoFin = "2029-06-15",
                        Activo = true
                    };

                    // Act
                    var result = repo.InsertPeriodo(periodo);

                    // Assert
                    result.IsSuccess.Should().BeFalse();
                    result.Error.Should().Contain("no soporta ventana temporal");
                }
            }
            finally
            {
                try { if (File.Exists(dbSinVentanaPath)) File.Delete(dbSinVentanaPath); } catch { }
            }
        }

        [Fact]
        public void UpdatePeriodoVentana_WhenExtiendeFin_ShouldUpdateDatesAndInsertAudit()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                    VALUES (10, '2027-01', '2027-01-15', '2027-05-30', 1);
                ");

                var repo = new AcademicRepository(conn, subunidad: "TEST", suppressMarkDirty: true);

                // Act
                var result = repo.UpdatePeriodoVentana(
                    idPeriodo: 10,
                    nuevoInicio: "2027-01-15",
                    nuevoFin: "2027-06-13",
                    inicioAntes: "2027-01-15",
                    finAntes: "2027-05-30",
                    actor: "TEST_USER",
                    origen: "AG_VENTANA_MANUAL",
                    motivo: "Extensión manual para alcanzar 64 ranuras");

                // Assert
                result.IsSuccess.Should().BeTrue();
                result.Value.Should().Be(1);

                // Verificar que periodos se actualizó sin tocar nombre ni activo
                var p = conn.QueryFirstOrDefault<Periodo>("SELECT idperiodo as IdPeriodo, periodo as PeriodoNombre, periodo_inicio as PeriodoInicio, periodo_fin as PeriodoFin, activo as Activo FROM periodos WHERE idperiodo = 10;");
                p.Should().NotBeNull();
                p.PeriodoInicio.Should().Be("2027-01-15");
                p.PeriodoFin.Should().Be("2027-06-13");
                p.PeriodoNombre.Should().Be("2027-01");
                p.Activo.Should().BeTrue();

                // Verificar registro de auditoría en ag_periodo_ventana_auditoria
                var aud = conn.QueryFirstOrDefault<dynamic>("SELECT * FROM ag_periodo_ventana_auditoria WHERE idperiodo = 10;");
                ((object)aud).Should().NotBeNull();
                ((string)aud.periodo).Should().Be("2027-01");
                ((string)aud.inicio_antes).Should().Be("2027-01-15");
                ((string)aud.fin_antes).Should().Be("2027-05-30");
                ((string)aud.inicio_despues).Should().Be("2027-01-15");
                ((string)aud.fin_despues).Should().Be("2027-06-13");
                ((string)aud.actor).Should().Be("TEST_USER");
                ((string)aud.origen).Should().Be("AG_VENTANA_MANUAL");
                ((string)aud.motivo).Should().Be("Extensión manual para alcanzar 64 ranuras");
            }
        }

        [Fact]
        public void UpdatePeriodoVentana_WhenOtroUsuarioModificoPeriodo_ShouldReturnFailureAndNotUpdate()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                    VALUES (11, '2027-02', '2027-08-01', '2027-11-30', 1);
                ");

                var repo = new AcademicRepository(conn, subunidad: "TEST", suppressMarkDirty: true);

                // Act: finAntes difiere de lo que está en BD (ej. 2027-11-15 vs 2027-11-30)
                var result = repo.UpdatePeriodoVentana(
                    idPeriodo: 11,
                    nuevoInicio: "2027-08-01",
                    nuevoFin: "2027-12-15",
                    inicioAntes: "2027-08-01",
                    finAntes: "2027-11-15",
                    actor: "USER_B",
                    origen: "AG_VENTANA_MANUAL",
                    motivo: "Intento concurrente desfasado");

                // Assert
                result.IsSuccess.Should().BeFalse();
                result.Error.Should().Be("PERIODO_MODIFICADO_POR_OTRO_USUARIO");

                // La BD no debe haber sido tocada
                var p = conn.QueryFirstOrDefault<Periodo>("SELECT idperiodo as IdPeriodo, periodo as PeriodoNombre, periodo_inicio as PeriodoInicio, periodo_fin as PeriodoFin, activo as Activo FROM periodos WHERE idperiodo = 11;");
                p.PeriodoFin.Should().Be("2027-11-30");

                // No debe haber fila en auditoría para el id 11
                int countAud = conn.QueryFirstOrDefault<int>("SELECT COUNT(1) FROM ag_periodo_ventana_auditoria WHERE idperiodo = 11;");
                countAud.Should().Be(0);
            }
        }

        [Fact]
        public void UpdatePeriodoVentana_WhenTablaAuditoriaNoExiste_ShouldReturnFailureAuditoriaNoMigrada()
        {
            string dbSinAudPath = Path.Combine(Path.GetTempPath(), "SinAud_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(dbSinAudPath);
            try
            {
                using (var conn = new SQLiteConnection($"Data Source={dbSinAudPath};Version=3;"))
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
                        INSERT INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo)
                        VALUES (12, '2027-01', '2027-01-15', '2027-05-30', 1);
                    ");

                    var repo = new AcademicRepository(conn, subunidad: "TEST", suppressMarkDirty: true);

                    // Act
                    var result = repo.UpdatePeriodoVentana(
                        idPeriodo: 12,
                        nuevoInicio: "2027-01-15",
                        nuevoFin: "2027-06-13",
                        inicioAntes: "2027-01-15",
                        finAntes: "2027-05-30",
                        actor: "TEST_USER",
                        origen: "AG_VENTANA_MANUAL",
                        motivo: "Prueba sin tabla de auditoría");

                    // Assert
                    result.IsSuccess.Should().BeFalse();
                    result.Error.Should().Be("AUDITORIA_VENTANA_NO_MIGRADA");
                }
            }
            finally
            {
                try { if (File.Exists(dbSinAudPath)) File.Delete(dbSinAudPath); } catch { }
            }
        }
    }
}
