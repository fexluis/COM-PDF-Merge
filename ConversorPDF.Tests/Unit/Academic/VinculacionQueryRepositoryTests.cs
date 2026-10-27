using System;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Moq;
using Xunit;
using FluentAssertions;
using Dapper;


namespace ConversorPDF.Tests.Unit.Academic
{
    public class VinculacionQueryRepositoryTests
    {
        [Fact]
        [Trait("Category", "Hotfix_1_2_51_25")]
        [Trait("Priority", "P0")]
        [Trait("Type", "Unit")]
        public async Task GetSesionesAsync_NoDebeDisponerLaConexion()
        {
            // Arrange
            // Dapper requiere una conexión real para realizar llamadas internas de reflexión.
            // Usamos SQLite en memoria para máxima velocidad y fidelidad.
            using (var conn = new System.Data.SQLite.SQLiteConnection("Data Source=:memory:;Version=3;New=True;"))
            {
                conn.Open();
                // Crear tabla mínima para que Dapper no falle por falta de esquema
                conn.Execute("CREATE TABLE sesiones (id INTEGER PRIMARY KEY);");
                conn.Execute("CREATE VIEW vw_SesionesDetalle AS SELECT 1 as IdSesion, '2026-01-01' as FechaInicial;");

                var factory = new Func<IDbConnection>(() => conn);
                var repository = new VinculacionQueryRepository(factory);

                // Act
                await repository.GetSesionesAsync(
                    idSede: null, 
                    idTercero: null, 
                    idPeriodo: null, 
                    idPrograma: null, 
                    modalidadPrograma: null, 
                    subunidad: null, 
                    semestre: null, 
                    soloViaticas: false, 
                    fechaDesde: null, 
                    fechaHasta: null, 
                    limit: 10, 
                    offset: 0, 
                    ct: CancellationToken.None);

                // Assert
                conn.State.Should().Be(ConnectionState.Open, 
                    "FAIL: El repositorio cerró o dispuso la conexión que no le pertenecía.");
            }
        }
    }
}
