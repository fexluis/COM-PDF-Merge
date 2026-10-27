using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using Dapper;
using FluentAssertions;
using Moq;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence
{
    public class SchemaColumnsAssertTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public SchemaColumnsAssertTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "SchemaColumnsAssertTest_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void VerificarColumnasRequeridas_WhenTableMissingRequiredColumns_ShouldLogErrorMessage()
        {
            // Arrange - Simula tabla periodos y terceros sin columnas requeridas
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo TEXT
                    );
                    CREATE TABLE terceros (
                        id INTEGER PRIMARY KEY,
                        categoria TEXT
                    );
                ");

                var mockLogger = new Mock<ILogger>();
                var mockVaultPersister = new Mock<IVaultPersister>();
                var migrator = new SchemaMigrator(new List<IMigration>(), mockVaultPersister.Object, mockLogger.Object);

                // Act
                migrator.VerificarColumnasRequeridas(conn);

                // Assert: Debe registrar error para periodos y terceros sin lanzar excepción
                mockLogger.Verify(l => l.Error(It.Is<string>(msg => msg.Contains("ESQUEMA INCOMPLETO en 'periodos'") && msg.Contains("periodo_inicio"))), Times.Once);
                mockLogger.Verify(l => l.Error(It.Is<string>(msg => msg.Contains("ESQUEMA INCOMPLETO en 'terceros'") && msg.Contains("cedula"))), Times.Once);
            }
        }

        [Fact]
        public void VerificarColumnasRequeridas_WhenSchemaIsComplete_ShouldNotLogError()
        {
            // Arrange - Crea tablas completas con todas las columnas esperadas
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                conn.Execute(@"
                    CREATE TABLE periodos (
                        idperiodo INTEGER PRIMARY KEY,
                        periodo_inicio TEXT,
                        periodo_fin TEXT
                    );
                    CREATE TABLE terceros (
                        id INTEGER PRIMARY KEY,
                        cedula TEXT,
                        nombre TEXT,
                        subunidad TEXT
                    );
                    CREATE TABLE asignaturas (
                        id INTEGER PRIMARY KEY,
                        cod_materia TEXT,
                        int_horaria INTEGER,
                        nombre TEXT
                    );
                    CREATE TABLE programas (
                        id INTEGER PRIMARY KEY,
                        nombreprograma TEXT,
                        abreviatura TEXT,
                        activo INTEGER
                    );
                    CREATE TABLE sedes (
                        id INTEGER PRIMARY KEY,
                        ciudad TEXT,
                        activo INTEGER
                    );
                ");

                var mockLogger = new Mock<ILogger>();
                var mockVaultPersister = new Mock<IVaultPersister>();
                var migrator = new SchemaMigrator(new List<IMigration>(), mockVaultPersister.Object, mockLogger.Object);

                // Act
                migrator.VerificarColumnasRequeridas(conn);

                // Assert: No se debe emitir ningún error de esquema incompleto
                mockLogger.Verify(l => l.Error(It.Is<string>(msg => msg.Contains("ESQUEMA INCOMPLETO"))), Times.Never);
            }
        }
    }
}
