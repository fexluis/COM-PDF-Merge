using System;
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
    /// <summary>
    /// Pruebas de integración del staging AG: selector de corridas y resumen de diagnóstico
    /// (Fases 1489/1499). El fixture reproduce el vault real del 2026-11-09:
    /// #46 BOLIVAR/Error sin staging, #44 VALLE 3 sesiones, #37 sin staging, #1 período anterior
    /// y la corrida huérfana #45 (staging sin cabecera en ag_corrida).
    /// </summary>
    public class VinculacionQueryRepositoryAgStagingTests : IDisposable
    {
        private readonly SQLiteConnection _conn;
        private readonly VinculacionQueryRepository _repo;

        public VinculacionQueryRepositoryAgStagingTests()
        {
            _conn = new SQLiteConnection("Data Source=:memory:;Version=3;New=True;");
            _conn.Open();
            CrearEsquemaYDatos();
            _repo = new VinculacionQueryRepository(() => _conn, AgModoDatos.Ag);
        }

        private void CrearEsquemaYDatos()
        {
            _conn.Execute(@"
                CREATE TABLE periodos (idperiodo INTEGER PRIMARY KEY, periodo TEXT, nombre TEXT);
                CREATE TABLE programas (id INTEGER PRIMARY KEY, nombreprograma TEXT, abreviatura TEXT);
                CREATE TABLE sedes (id INTEGER PRIMARY KEY, ciudad TEXT);
                CREATE TABLE terceros (id INTEGER PRIMARY KEY, nombre TEXT);
                CREATE TABLE asignaturas (id INTEGER PRIMARY KEY, nombre TEXT);

                CREATE TABLE ag_corrida (
                    id INTEGER PRIMARY KEY, idperiodo INTEGER, idprograma INTEGER, subunidad TEXT,
                    fecha_inicio_corrida TEXT, estado INTEGER, id_individuo_optimo INTEGER);

                CREATE TABLE ag_vinculaciones (
                    id INTEGER PRIMARY KEY, id_corrida INTEGER, idperiodo INTEGER, idprograma INTEGER,
                    idsede INTEGER, idasignatura INTEGER, idtercero INTEGER, semestre TEXT, grupo TEXT,
                    dia_horas_semanales TEXT, total_horas INTEGER, fecha_inicio TEXT, fecha_fin TEXT,
                    estado INTEGER, es_vacante INTEGER, motivo_vacante TEXT, candidatos_pool INTEGER,
                    subunidad TEXT);

                CREATE TABLE ag_sesiones (
                    id INTEGER PRIMARY KEY, idvinculacion INTEGER, idtercero INTEGER, id_aula INTEGER,
                    fechasesion TEXT, horariosesion TEXT, horas REAL, viatica INTEGER, subunidad TEXT,
                    idindividuo INTEGER, idgen INTEGER, es_vacante INTEGER, motivo_vacante TEXT,
                    candidatos_pool INTEGER);

                CREATE VIEW vw_AgSesionesDetalle AS
                SELECT s.id AS IdSesion,
                       s.idvinculacion AS IdVinculacion,
                       v.id_corrida AS IdCorrida,
                       v.idperiodo AS IdPeriodo,
                       v.idsede AS IdSede,
                       v.idprograma AS IdPrograma,
                       s.idtercero AS IdTercero,
                       CASE WHEN v.idprograma = 20 THEN 'POSGRADO' ELSE 'PREGRADO' END AS ProgramaTipo,
                       p.abreviatura AS ProgramaAbreviatura,
                       v.semestre AS Semestre,
                       t.nombre AS DocenteNombre,
                       a.nombre AS AsignaturaNombre,
                       sd.ciudad AS SedeCiudad,
                       s.horariosesion AS HorarioSesion,
                       s.horas AS Horas,
                       NULLIF(s.fechasesion, '') AS FechaSesion,
                       s.viatica AS Viatica,
                       v.grupo AS Grupo,
                       v.subunidad AS Subunidad
                FROM ag_sesiones s
                JOIN ag_vinculaciones v ON s.idvinculacion = v.id AND IFNULL(v.estado, 0) <> 3
                LEFT JOIN programas p ON v.idprograma = p.id
                LEFT JOIN terceros t ON s.idtercero = t.id
                LEFT JOIN sedes sd ON v.idsede = sd.id
                LEFT JOIN asignaturas a ON v.idasignatura = a.id;
            ");

            _conn.Execute(@"
                INSERT INTO periodos (idperiodo, periodo, nombre) VALUES
                    (2, '2026-02', '2026-02'), (3, '2027-01', '2027-01');
                INSERT INTO programas (id, nombreprograma, abreviatura) VALUES
                    (1, 'ADMINISTRACION PUBLICA TERRITORIAL', 'APU'),
                    (20, 'MAESTRIA EN ADMINISTRACION PUBLICA', 'MAP');
                INSERT INTO sedes (id, ciudad) VALUES (1, 'CALI');
                INSERT INTO terceros (id, nombre) VALUES (1, 'Docente A');
                INSERT INTO asignaturas (id, nombre) VALUES (1, 'TEORIA ADMINISTRATIVA');

                INSERT INTO ag_corrida (id, idperiodo, idprograma, subunidad, fecha_inicio_corrida, estado) VALUES
                    (44, 3, 1,  'VALLE',   '2026-10-26 14:00:29', 2),
                    (46, 3, 20, 'BOLIVAR', '2026-10-29 12:21:44', 4),
                    (37, 3, 2,  'VALLE',   '2026-10-24 10:58:15', 2),
                    (1,  2, 1,  'VALLE',   '2026-10-16 11:45:01', 2);

                INSERT INTO ag_vinculaciones
                    (id, id_corrida, idperiodo, idprograma, idsede, idasignatura, idtercero, semestre, grupo,
                     dia_horas_semanales, total_horas, fecha_inicio, fecha_fin, estado, es_vacante,
                     motivo_vacante, candidatos_pool, subunidad) VALUES
                    (100, 44, 3, 1, 1, 1, 1, 'I', 'D', 'LUNES|4', 8, '2027-02-01', '2027-03-01', 0, 0, '', 0, 'VALLE'),
                    (101, 44, 3, 1, 1, 1, 1, 'I', 'D', 'MARTES|4', 4, '2027-02-01', '2027-03-01', 0, 0, '', 0, 'VALLE'),
                    (102, 45, 3, 1, 1, 1, 1, 'I', 'D', 'MIERCOLES|4', 4, '2027-02-01', '2027-03-01', 0, 0, '', 0, 'BOLIVAR'),
                    (103, 1,  2, 1, 1, 1, 1, 'I', 'D', 'JUEVES|4', 4, '2026-03-01', '2026-04-01', 0, 0, '', 0, 'VALLE');

                INSERT INTO ag_sesiones
                    (id, idvinculacion, idtercero, id_aula, fechasesion, horariosesion, horas, viatica,
                     subunidad, idindividuo, idgen, es_vacante, motivo_vacante, candidatos_pool) VALUES
                    (1000, 100, 1, 1, '2027-02-08', '08:00 A 12:00', 4, 0, 'VALLE',   1, 1, 0, '', 0),
                    (1001, 100, 1, 1, '2027-02-15', '08:00 A 12:00', 4, 0, 'VALLE',   1, 1, 0, '', 0),
                    (1002, 101, 1, 1, '2027-02-09', '08:00 A 12:00', 4, 0, 'VALLE',   1, 2, 0, '', 0),
                    (1003, 102, 1, 1, '2027-02-10', '08:00 A 12:00', 4, 0, 'BOLIVAR', 1, 3, 0, '', 0),
                    (1004, 103, 1, 1, '2026-03-02', '08:00 A 12:00', 4, 0, 'VALLE',   1, 4, 0, '', 0);
            ");
        }

        [Fact]
        public async Task GetCorridasStagingAsync_SinFiltros_IncluyeHuerfanaYMarcaLasVacias()
        {
            var res = await _repo.GetCorridasStagingAsync(null, null, null, CancellationToken.None);

            res.IsSuccess.Should().BeTrue(res.IsFailure ? res.Error : string.Empty);
            var items = res.Value.ToList();

            items.Select(c => c.IdCorrida).Should().Equal(new long[] { 46, 45, 44, 37, 1 });

            var c44 = items.Single(c => c.IdCorrida == 44);
            c44.TotalSesiones.Should().Be(3);
            c44.TieneStaging.Should().BeTrue();
            c44.PeriodoNombre.Should().Be("2027-01");
            c44.Subunidad.Should().Be("VALLE");

            // #45: staging sin cabecera en ag_corrida (huérfana) debe ser visible
            items.Single(c => c.IdCorrida == 45).TotalSesiones.Should().Be(1);

            // #46 y #37: registradas sin staging
            items.Single(c => c.IdCorrida == 46).TieneStaging.Should().BeFalse();
            items.Single(c => c.IdCorrida == 46).EstadoCorrida.Should().Be(4);
            items.Single(c => c.IdCorrida == 37).TieneStaging.Should().BeFalse();
        }

        [Fact]
        public async Task GetCorridasStagingAsync_FiltroSubunidad_EsToleranteAMayusculasYEspacios()
        {
            var res = await _repo.GetCorridasStagingAsync(null, null, "  valle  ", CancellationToken.None);

            res.IsSuccess.Should().BeTrue(res.IsFailure ? res.Error : string.Empty);
            res.Value.Select(c => c.IdCorrida).Should().BeEquivalentTo(new long[] { 44, 37, 1 });
        }

        [Fact]
        public async Task GetResumenStagingAsync_PeriodoDistinto_ReportaDescriptoresRealesY0ConFiltros()
        {
            // Escenario real del bug: config.id_periodo = 2 y la corrida #44 es del período 3.
            var res = await _repo.GetResumenStagingAsync(
                idCorrida: 44, idSede: null, idTercero: null, idPeriodo: 2, idPrograma: null,
                modalidadPrograma: null, subunidad: null, semestre: null, soloViaticas: false,
                fechaDesde: null, fechaHasta: null, ct: CancellationToken.None);

            res.IsSuccess.Should().BeTrue(res.IsFailure ? res.Error : string.Empty);
            var r = res.Value;

            r.SesionesCorrida.Should().Be(3);
            r.VincCorrida.Should().Be(2);
            r.SesionesConFiltros.Should().Be(0, "el período 2 no coincide con el período real de la corrida");
            r.PeriodosCorrida.Should().Be("2027-01");
            r.SubunidadesCorrida.Should().Be("VALLE");
            r.ProgramasCorrida.Should().Be("ADMINISTRACION PUBLICA TERRITORIAL");
        }

        [Fact]
        public async Task GetResumenStagingAsync_SinFiltrosDeUi_LasSesionesCoinciden()
        {
            var res = await _repo.GetResumenStagingAsync(
                idCorrida: 44, idSede: null, idTercero: null, idPeriodo: null, idPrograma: null,
                modalidadPrograma: null, subunidad: null, semestre: null, soloViaticas: false,
                fechaDesde: null, fechaHasta: null, ct: CancellationToken.None);

            res.IsSuccess.Should().BeTrue(res.IsFailure ? res.Error : string.Empty);
            res.Value.SesionesConFiltros.Should().Be(3);
        }

        [Fact]
        public async Task GetResumenStagingAsync_CorridaSinStaging_ReportaCero()
        {
            var res = await _repo.GetResumenStagingAsync(
                idCorrida: 46, idSede: null, idTercero: null, idPeriodo: null, idPrograma: null,
                modalidadPrograma: null, subunidad: null, semestre: null, soloViaticas: false,
                fechaDesde: null, fechaHasta: null, ct: CancellationToken.None);

            res.IsSuccess.Should().BeTrue(res.IsFailure ? res.Error : string.Empty);
            res.Value.SesionesCorrida.Should().Be(0);
            res.Value.SesionesConFiltros.Should().Be(0);
        }

        [Fact]
        public async Task GetSesionesAsync_ConCorridaYPeriodoDistinto_DevuelveCeroYConPeriodoCorrectoDevuelveFilas()
        {
            var conPeriodoIncorrecto = await _repo.GetSesionesAsync(
                idSede: null, idTercero: null, idPeriodo: 2, idPrograma: null, modalidadPrograma: null,
                subunidad: null, semestre: null, soloViaticas: false, fechaDesde: null, fechaHasta: null,
                limit: -1, offset: 0, ct: CancellationToken.None, idCorrida: 44);

            conPeriodoIncorrecto.IsSuccess.Should().BeTrue(conPeriodoIncorrecto.IsFailure ? conPeriodoIncorrecto.Error : string.Empty);
            conPeriodoIncorrecto.Value.Should().BeEmpty();

            var sinPeriodo = await _repo.GetSesionesAsync(
                idSede: null, idTercero: null, idPeriodo: null, idPrograma: null, modalidadPrograma: null,
                subunidad: null, semestre: null, soloViaticas: false, fechaDesde: null, fechaHasta: null,
                limit: -1, offset: 0, ct: CancellationToken.None, idCorrida: 44);

            sinPeriodo.IsSuccess.Should().BeTrue(sinPeriodo.IsFailure ? sinPeriodo.Error : string.Empty);
            sinPeriodo.Value.Should().HaveCount(3);
        }

        [Fact]
        public async Task EliminarCorridaStagingAsync_BorraSoloLaCorridaIndicada()
        {
            var res = await _repo.EliminarCorridaStagingAsync(44, CancellationToken.None);

            res.IsSuccess.Should().BeTrue(res.IsFailure ? res.Error : string.Empty);
            res.Value.Vinculaciones.Should().Be(2);
            res.Value.Sesiones.Should().Be(3);

            // La corrida #44 desaparece de staging y de su cabecera
            _conn.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_vinculaciones WHERE id_corrida = 44").Should().Be(0);
            _conn.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_corrida WHERE id = 44").Should().Be(0);
            _conn.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_sesiones WHERE idvinculacion IN (100, 101)").Should().Be(0);

            // El resto de corridas permanece intacto
            _conn.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_corrida WHERE id IN (46, 37, 1)").Should().Be(3);
            _conn.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_sesiones WHERE idvinculacion IN (102, 103)").Should().Be(2);
        }

        [Fact]
        public async Task EliminarCorridaStagingAsync_CorridaInexistente_DevuelveFallo()
        {
            var res = await _repo.EliminarCorridaStagingAsync(9999, CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("no existe");
        }

        [Fact]
        public async Task EliminarCorridaStagingAsync_CorridaSinStaging_BorraSoloLaCabecera()
        {
            var res = await _repo.EliminarCorridaStagingAsync(46, CancellationToken.None);

            res.IsSuccess.Should().BeTrue(res.IsFailure ? res.Error : string.Empty);
            res.Value.Sesiones.Should().Be(0);
            res.Value.Vinculaciones.Should().Be(0);
            _conn.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_corrida WHERE id = 46").Should().Be(0);
        }

        public void Dispose()
        {
            _conn?.Dispose();
        }
    }
}
