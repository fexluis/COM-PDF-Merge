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
    }
}
