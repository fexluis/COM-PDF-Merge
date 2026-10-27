using System;
using System.Data.SQLite;
using System.IO;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AcademicRepositoryAgModoPernoctaTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public AcademicRepositoryAgModoPernoctaTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AcademicRepoAgPernoctaTest_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE ag_vinculaciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idsede INTEGER DEFAULT 1,
                        idperiodo INTEGER DEFAULT 1,
                        idtercero INTEGER DEFAULT 1,
                        idprograma INTEGER DEFAULT 1,
                        grupo TEXT DEFAULT 'A',
                        semestre TEXT DEFAULT '1',
                        idasignatura INTEGER DEFAULT 1,
                        total_horas INTEGER DEFAULT 0,
                        dia_horas_semanales TEXT DEFAULT '',
                        fecha_inicio TEXT,
                        fecha_fin TEXT,
                        modalidadVinculacion TEXT DEFAULT 'Presencial',
                        subunidad TEXT DEFAULT '',
                        isTerminalAereo INTEGER DEFAULT 0,
                        idTerminales INTEGER DEFAULT 0,
                        gastoTerminalAerea REAL DEFAULT 0,
                        es_vacante INTEGER DEFAULT 0,
                        motivo_vacante TEXT DEFAULT ''
                    );");
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
                // Limpieza defensiva
            }
        }

        [Fact]
        public void GetVinculacionById_WhenAgVinculacionesLacksPernoctaColumn_ShouldReturnFailure()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute("INSERT INTO ag_vinculaciones (idtercero) VALUES (1);");

                var repo = new AcademicRepository(conn, suppressMarkDirty: true, modo: AgModoDatos.Ag);

                // Act
                var result = repo.GetVinculacionById(1);

                // Assert
                result.IsFailure.Should().BeTrue();
                result.Error.Should().Contain("no such column: isPernoctaUltimoDia");
            }
        }

        [Fact]
        public void GetVinculacionById_AfterMigrationV77_ShouldReturnSuccessWithDefaultZero()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute("INSERT INTO ag_vinculaciones (idtercero) VALUES (1);");

                var mig = new MigrationV77_AddPernoctaAgVinculaciones();
                mig.Up(conn, null);

                var repo = new AcademicRepository(conn, suppressMarkDirty: true, modo: AgModoDatos.Ag);

                // Act
                var result = repo.GetVinculacionById(1);

                // Assert
                result.IsSuccess.Should().BeTrue();
                result.Value.Should().NotBeNull();
                result.Value.IsPernoctaUltimoDia.Should().Be(0);
            }
        }

        [Fact]
        public void InsertAndUpdateVinculacion_InAgModo_ShouldPersistIsPernoctaUltimoDia()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                var mig = new MigrationV77_AddPernoctaAgVinculaciones();
                mig.Up(conn, null);

                var repo = new AcademicRepository(conn, suppressMarkDirty: true, modo: AgModoDatos.Ag);

                var nuevaVinc = new Vinculacion
                {
                    IdSede = 1,
                    IdPeriodo = 2,
                    IdTercero = 3,
                    IdPrograma = 4,
                    Grupo = "G1",
                    Semestre = "5",
                    IdAsignatura = 6,
                    TotalHoras = 32,
                    DiaHorasSemanales = "SABADO|8",
                    Modalidad = "Presencial",
                    IsPernoctaUltimoDia = 1
                };

                // Act - Inserción
                var resInsert = repo.InsertVinculacion(nuevaVinc);

                // Assert - Inserción
                resInsert.IsSuccess.Should().BeTrue();
                int idGenerado = resInsert.Value;
                idGenerado.Should().BeGreaterThan(0);

                var vincCargada = repo.GetVinculacionById(idGenerado);
                vincCargada.IsSuccess.Should().BeTrue();
                vincCargada.Value.IsPernoctaUltimoDia.Should().Be(1);

                // Act - Actualización
                var vincParaActualizar = vincCargada.Value;
                vincParaActualizar.IsPernoctaUltimoDia = 0;
                var resUpdate = repo.UpdateVinculacion(vincParaActualizar);

                // Assert - Actualización
                resUpdate.IsSuccess.Should().BeTrue();
                int valorEnDb = conn.ExecuteScalar<int>(
                    "SELECT isPernoctaUltimoDia FROM ag_vinculaciones WHERE id = @id", new { id = idGenerado });
                valorEnDb.Should().Be(0);

                // Act - SetPernoctaUltimoDia
                var resSet = repo.SetPernoctaUltimoDia(idGenerado, 1);
                resSet.IsSuccess.Should().BeTrue();
                int valorFinal = conn.ExecuteScalar<int>(
                    "SELECT isPernoctaUltimoDia FROM ag_vinculaciones WHERE id = @id", new { id = idGenerado });
                valorFinal.Should().Be(1);
            }
        }
    }
}
