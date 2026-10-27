using Xunit;
using FluentAssertions;
using System;
using System.Data.SQLite;
using Dapper;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.Tests.Integration.Academic
{
    public class FestivoRepositoryTests : IDisposable
    {
        private readonly SQLiteConnection _conn;
        private readonly AcademicRepository _repo;

        public FestivoRepositoryTests()
        {
            _conn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;");
            _conn.Open();
            
            _conn.Execute(@"
                CREATE TABLE periodos (idperiodo INTEGER PRIMARY KEY, periodo TEXT, activo INTEGER);
                CREATE TABLE festivos (
                    idfestivo INTEGER PRIMARY KEY AUTOINCREMENT,
                    idperiodo INTEGER,
                    fechafestivoinicio TEXT,
                    fechafestivofinal TEXT,
                    detallefestivo TEXT,
                    tipofestivo TEXT DEFAULT 'FESTIVO',
                    subunidad TEXT,
                    FOREIGN KEY(idperiodo) REFERENCES periodos(idperiodo)
                );
                INSERT INTO periodos VALUES (1, '2026-01', 1);
            ");
            
            _repo = new AcademicRepository(_conn);
        }

        [Fact]
        public void CRUD_FestivoFlow_ShouldInsertUpdateAndDeleteCorrectly()
        {
            var festivo = new Festivo
            {
                IdPeriodo = 1,
                FechaFestivoInicio = new DateTime(2026, 5, 1),
                FechaFestivoFinal = new DateTime(2026, 5, 1),
                DetalleFestivo = "Día del Trabajo",
                TipoFestivo = "PRESENCIALES"
            };

            // 1. Insert
            var insertResult = _repo.InsertFestivo(festivo);
            insertResult.IsSuccess.Should().BeTrue();
            int newId = insertResult.Value;
            newId.Should().BeGreaterThan(0);

            // 2. Read
            var getResult = _repo.GetFestivoById(newId);
            getResult.IsSuccess.Should().BeTrue();
            getResult.Value.DetalleFestivo.Should().Be("Día del Trabajo");
            getResult.Value.TipoFestivo.Should().Be("PRESENCIALES");
            getResult.Value.FechaFestivoInicio.Date.Should().Be(new DateTime(2026, 5, 1));

            // 3. Update
            getResult.Value.DetalleFestivo = "Día del Trabajo Modificado";
            getResult.Value.TipoFestivo = "FESTIVO";
            var updateResult = _repo.UpdateFestivo(getResult.Value);
            updateResult.IsSuccess.Should().BeTrue();

            var getUpdatedResult = _repo.GetFestivoById(newId);
            getUpdatedResult.Value.DetalleFestivo.Should().Be("Día del Trabajo Modificado");
            getUpdatedResult.Value.TipoFestivo.Should().Be("FESTIVO");

            // 4. Delete
            var deleteResult = _repo.DeleteFestivo(newId);
            deleteResult.IsSuccess.Should().BeTrue();

            var getDeletedResult = _repo.GetFestivoById(newId);
            getDeletedResult.IsFailure.Should().BeTrue();
        }

        public void Dispose()
        {
            _conn?.Close();
            _conn?.Dispose();
        }
    }
}
