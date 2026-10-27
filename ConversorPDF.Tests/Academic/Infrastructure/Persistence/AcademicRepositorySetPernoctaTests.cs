using System;
using System.Data.SQLite;
using System.IO;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class AcademicRepositorySetPernoctaTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public AcademicRepositorySetPernoctaTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AcademicRepoPernoctaTest_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE vinculaciones (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        id_periodo INTEGER DEFAULT 1,
                        id_tercero INTEGER DEFAULT 1,
                        id_programa INTEGER DEFAULT 1,
                        grupo TEXT DEFAULT 'A',
                        semestre INTEGER DEFAULT 1,
                        id_asignatura INTEGER DEFAULT 1,
                        total_horas REAL DEFAULT 0,
                        dia_horas_semanales TEXT DEFAULT '',
                        fecha_inicio TEXT,
                        fecha_fin TEXT,
                        modalidad TEXT DEFAULT 'Presencial',
                        isPernoctaUltimoDia INTEGER DEFAULT 0,
                        is_terminal_aereo INTEGER DEFAULT 0,
                        id_terminales INTEGER DEFAULT 0,
                        gasto_terminal_aerea REAL DEFAULT 0,
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
                // Best effort cleanup
            }
        }

        [Fact]
        public void SetPernoctaUltimoDia_WhenActivar_ShouldUpdateToOneAndReturnOneRow()
        {
            // Arrange
            int idVinculacion;
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                idVinculacion = conn.ExecuteScalar<int>(
                    "INSERT INTO vinculaciones (isPernoctaUltimoDia) VALUES (0); SELECT last_insert_rowid();");

                var repo = new AcademicRepository(conn, suppressMarkDirty: true, modo: AgModoDatos.Determinista);

                // Act
                var result = repo.SetPernoctaUltimoDia(idVinculacion, 1);

                // Assert
                result.IsSuccess.Should().BeTrue();
                result.Value.Should().Be(1);

                int valorEnDb = conn.ExecuteScalar<int>(
                    "SELECT isPernoctaUltimoDia FROM vinculaciones WHERE id = @id", new { id = idVinculacion });
                valorEnDb.Should().Be(1);
            }
        }

        [Fact]
        public void SetPernoctaUltimoDia_WhenDesactivar_ShouldUpdateToZeroAndReturnOneRow()
        {
            // Arrange
            int idVinculacion;
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                idVinculacion = conn.ExecuteScalar<int>(
                    "INSERT INTO vinculaciones (isPernoctaUltimoDia) VALUES (1); SELECT last_insert_rowid();");

                var repo = new AcademicRepository(conn, suppressMarkDirty: true, modo: AgModoDatos.Determinista);

                // Act
                var result = repo.SetPernoctaUltimoDia(idVinculacion, 0);

                // Assert
                result.IsSuccess.Should().BeTrue();
                result.Value.Should().Be(1);

                int valorEnDb = conn.ExecuteScalar<int>(
                    "SELECT isPernoctaUltimoDia FROM vinculaciones WHERE id = @id", new { id = idVinculacion });
                valorEnDb.Should().Be(0);
            }
        }

        [Fact]
        public void SetPernoctaUltimoDia_WhenIdInexistente_ShouldReturnSuccessWithZeroRows()
        {
            // Arrange
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                var repo = new AcademicRepository(conn, suppressMarkDirty: true, modo: AgModoDatos.Determinista);

                // Act
                var result = repo.SetPernoctaUltimoDia(9999, 1);

                // Assert
                result.IsSuccess.Should().BeTrue();
                result.Value.Should().Be(0);
            }
        }

        [Fact]
        public void SetPernoctaUltimoDia_WhenValorDistintoDeUno_ShouldNormalizeToZero()
        {
            // Arrange
            int idVinculacion;
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                idVinculacion = conn.ExecuteScalar<int>(
                    "INSERT INTO vinculaciones (isPernoctaUltimoDia) VALUES (1); SELECT last_insert_rowid();");

                var repo = new AcademicRepository(conn, suppressMarkDirty: true, modo: AgModoDatos.Determinista);

                // Act: pasar un valor arbitrario != 1
                var result = repo.SetPernoctaUltimoDia(idVinculacion, 2);

                // Assert
                result.IsSuccess.Should().BeTrue();
                result.Value.Should().Be(1);

                int valorEnDb = conn.ExecuteScalar<int>(
                    "SELECT isPernoctaUltimoDia FROM vinculaciones WHERE id = @id", new { id = idVinculacion });
                valorEnDb.Should().Be(0);
            }
        }
    }
}
