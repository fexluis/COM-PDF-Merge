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
    public class MigrationV74Tests : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        public MigrationV74Tests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "MigrationV74Test_" + Guid.NewGuid().ToString("N") + ".sqlite");
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
        public void MigrationV74_Metadata_ShouldBeCorrect()
        {
            // Arrange
            var mig = new MigrationV74_SembrarBloques4hFaltantes();

            // Assert
            mig.Version.Should().Be(74);
            mig.Name.Should().Be("SembrarBloques4hFaltantes");
            mig.Aditiva.Should().BeTrue();
        }

        [Fact]
        public void Up_WhenTableAgBloqueHorarioExists_ShouldSeedMissing4hBlocksAndBeIdempotent()
        {
            // Arrange
            var mig = new MigrationV74_SembrarBloques4hFaltantes();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                conn.Execute(@"
CREATE TABLE ag_bloque_horario (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dia_semana INTEGER,
    hora_inicio TEXT,
    hora_fin TEXT,
    duracion_horas REAL,
    franja TEXT,
    activo INTEGER,
    subunidad TEXT
);");

                // Act: Primera ejecución
                mig.Up(conn, null);

                // Assert: Para días 1 a 6 debe existir al menos un bloque de 4h
                var conteo4h = conn.ExecuteScalar<int>(
                    "SELECT COUNT(DISTINCT dia_semana) FROM ag_bloque_horario WHERE duracion_horas = 4.0 AND dia_semana BETWEEN 1 AND 6;");
                conteo4h.Should().Be(6);

                // Fase 1079: Verificar que la franja TARDE tenga el intervalo canónico 13:00-17:00
                var tarde1 = conn.QueryFirstOrDefault<(string Ini, string Fin)>(
                    "SELECT hora_inicio, hora_fin FROM ag_bloque_horario WHERE dia_semana = 1 AND franja = 'TARDE' AND duracion_horas = 4.0;");
                tarde1.Ini.Should().Be("13:00");
                tarde1.Fin.Should().Be("17:00");

                int totalBloques = conn.ExecuteScalar<int>("SELECT COUNT(1) FROM ag_bloque_horario;");

                // Act: Segunda ejecución (idempotencia)
                mig.Up(conn, null);

                int totalBloquesPost = conn.ExecuteScalar<int>("SELECT COUNT(1) FROM ag_bloque_horario;");
                totalBloquesPost.Should().Be(totalBloques);
            }
        }
    }
}
