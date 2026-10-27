using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Academic.Infrastructure.Persistence.Repositories
{
    public class AgFiltroSeleccionRepositoryTests
    {
        [Fact]
        public async Task GuardarYObtenerFiltros_WhenExecuted_ShouldPersistAndRetrieveData()
        {
            string dbPath = Path.Combine(Path.GetTempPath(), $"ag_filtro_{Guid.NewGuid():N}.db");
            string connStr = $"Data Source={dbPath};Version=3;";

            try
            {
                using (var setupConn = new SQLiteConnection(connStr))
                {
                    setupConn.Open();
                    new MigrationV78_AddAgFiltroSeleccion().Up(setupConn);
                }

                var repo = new AgFiltroSeleccionRepository(() => new SQLiteConnection(connStr));
                var filtros = new AgFiltrosSeleccion
                {
                    PeriodosIds = new List<int> { 1, 2 },
                    SedesIds = new List<int> { 10 },
                    ProgramasIds = new List<int> { 101, 102 },
                    TipoPrograma = "PREGRADO",
                    TercerosIds = new List<int> { 500, 501 }
                };

                // Act
                var saveRes = await repo.GuardarFiltrosAsync(filtros, "SUB_TEST", CancellationToken.None);
                var getRes = await repo.ObtenerFiltrosAsync("SUB_TEST", CancellationToken.None);

                // Assert
                saveRes.IsSuccess.Should().BeTrue();
                getRes.IsSuccess.Should().BeTrue();
                var loaded = getRes.Value;
                loaded.PeriodosIds.Should().BeEquivalentTo(new[] { 1, 2 });
                loaded.SedesIds.Should().BeEquivalentTo(new[] { 10 });
                loaded.ProgramasIds.Should().BeEquivalentTo(new[] { 101, 102 });
                loaded.TipoPrograma.Should().Be("PREGRADO");
                loaded.TercerosIds.Should().BeEquivalentTo(new[] { 500, 501 });
            }
            finally
            {
                try
                {
                    if (File.Exists(dbPath)) File.Delete(dbPath);
                }
                catch { }
            }
        }
    }
}
