using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Integration.Academic
{
    public class VinculacionQueryRepositoryTests : IDisposable
    {
        private readonly SQLiteConnection _conn;
        private readonly VinculacionQueryRepository _repo;

        public VinculacionQueryRepositoryTests()
        {
            _conn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;");
            _conn.Open();
            CrearEsquemaYDatos();
            _repo = new VinculacionQueryRepository(() => _conn);
        }

        private void CrearEsquemaYDatos()
        {
            _conn.Execute(@"
                CREATE TABLE programas (id INTEGER PRIMARY KEY, modalidad TEXT, abreviatura TEXT);
                CREATE TABLE terceros (id INTEGER PRIMARY KEY, nombre TEXT);
                CREATE TABLE sedes (id INTEGER PRIMARY KEY, ciudad TEXT);
                CREATE TABLE asignaturas (id INTEGER PRIMARY KEY, nombre TEXT);
                CREATE TABLE periodos (id INTEGER PRIMARY KEY, nombre TEXT);
                CREATE TABLE vinculaciones (
                    id INTEGER PRIMARY KEY, idsede INTEGER, idtercero INTEGER,
                    idprograma INTEGER, semestre TEXT, idasignatura INTEGER,
                    idperiodo INTEGER, total_horas INTEGER
                );
                CREATE TABLE sesiones (
                    id INTEGER PRIMARY KEY, idvinculacion INTEGER, idtercero INTEGER,
                    fechasesion TEXT, horariosesion TEXT,
                    horas REAL, viatica INTEGER
                );

                -- Vista simplificada para tests que emula la real
                CREATE VIEW vw_SesionesDetalle AS
                SELECT 
                    s.id AS IdSesion,
                    s.idvinculacion AS IdVinculacion,
                    t.nombre AS DocenteNombre,
                    p.modalidad AS ProgramaTipo,
                    p.abreviatura AS ProgramaAbreviatura,
                    v.semestre AS Semestre,
                    a.nombre AS AsignaturaNombre,
                    sd.ciudad AS SedeCiudad,
                    s.horariosesion AS HorarioSesion,
                    s.horas AS Horas,
                    NULLIF(s.fechasesion, '') AS FechaSesion,
                    s.viatica AS Viatica,
                    v.idsede AS IdSede,
                    v.idtercero AS IdTercero,
                    v.idperiodo AS IdPeriodo,
                    v.idprograma AS IdPrograma
                FROM sesiones s
                JOIN vinculaciones v ON s.idvinculacion = v.id
                JOIN terceros t ON v.idtercero = t.id
                JOIN programas p ON v.idprograma = p.id
                JOIN sedes sd ON v.idsede = sd.id
                JOIN asignaturas a ON v.idasignatura = a.id;

                INSERT INTO programas VALUES (1, 'PREGRADO', 'APT');
                INSERT INTO terceros VALUES (1, 'Docente A');
                INSERT INTO sedes VALUES (1, 'Cali');
                INSERT INTO asignaturas VALUES (1, 'Matemáticas');
                INSERT INTO vinculaciones VALUES (1, 1, 1, 1, 'I', 1, 1, 24);
                INSERT INTO sesiones VALUES 
                    (1, 1, 1, '2026-04-23', 'JUEVES|8', 8, 1),
                    (2, 1, 1, '2026-05-07', 'JUEVES|8', 8, 0);
            ");
        }

        [Fact]
        public async Task GetSesionesAsync_SinFiltros_RetornaTodas()
        {
            var result = await _repo.GetSesionesAsync(idSede: null, idTercero: null, idPeriodo: null, idPrograma: null, modalidadPrograma: null, subunidad: null, semestre: null, soloViaticas: false, fechaDesde: null, fechaHasta: null, limit: 100, offset: 0, ct: CancellationToken.None);
            result.IsSuccess.Should().BeTrue();
            result.Value.Should().HaveCount(2);
        }

        [Fact]
        public async Task GetSesionesAsync_FiltroPorFechaDesde_RetornaSoloPosteriores()
        {
            var result = await _repo.GetSesionesAsync(idSede: null, idTercero: null, idPeriodo: null, idPrograma: null, modalidadPrograma: null, subunidad: null, semestre: null, soloViaticas: false, 
                fechaDesde: new DateTime(2026, 5, 1), fechaHasta: null, limit: 100, offset: 0, ct: CancellationToken.None);
            result.Value.Should().HaveCount(1);
            result.Value.First().FechaSesion.Should().Be(new DateTime(2026, 5, 7));
        }

        [Fact]
        public async Task GetSesionesAsync_FiltroPorRangoFechas_RetornaDentroDelRango()
        {
            var result = await _repo.GetSesionesAsync(idSede: null, idTercero: null, idPeriodo: null, idPrograma: null, modalidadPrograma: null, subunidad: null, semestre: null, soloViaticas: false,
                fechaDesde: new DateTime(2026, 4, 20), fechaHasta: new DateTime(2026, 4, 25), limit: 100, offset: 0, ct: CancellationToken.None);
            result.Value.Should().HaveCount(1);
            result.Value.First().FechaSesion.Should().Be(new DateTime(2026, 4, 23));
        }

        [Fact]
        public async Task GetSesionesAsync_FiltroSoloViaticas_RetornaSoloConViatica()
        {
            var result = await _repo.GetSesionesAsync(idSede: null, idTercero: null, idPeriodo: null, idPrograma: null, modalidadPrograma: null, subunidad: null, semestre: null, soloViaticas: true, fechaDesde: null, fechaHasta: null, limit: 100, offset: 0, ct: CancellationToken.None);
            result.Value.Should().HaveCount(1);
            result.Value.First().Viatica.Should().BeTrue();
        }

        public void Dispose()
        {
            _conn?.Dispose();
        }
    }
}
