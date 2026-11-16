using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using Dapper;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Integration.Academic
{
    /// <summary>
    /// Fixture del vault AG mínimo para probar la purga anclada a id_corrida y la reparación de corridas
    /// huérfanas (Fases 1560 y 1563). Reproduce el caso real de la corrida #48: cabecera con un programa
    /// (APT-53) y staging materializado con el programa de la asignatura (APT-52).
    /// Usa un archivo temporal porque <see cref="AgScheduleRepository"/> abre y cierra la conexión en cada
    /// operación (una base ':memory:' se perdería entre conexiones).
    /// </summary>
    public abstract class AgScheduleRepoHuerfanaFixture : IDisposable
    {
        private readonly string _dbPath;
        private readonly string _connectionString;

        protected readonly AgScheduleRepository Repo;

        protected AgScheduleRepoHuerfanaFixture()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), "AgScheduleRepoHuerfanas_" + Guid.NewGuid().ToString("N") + ".sqlite");
            SQLiteConnection.CreateFile(_dbPath);
            _connectionString = $"Data Source={_dbPath};Version=3;";

            using (var c = Abrir())
            {
                c.Open();
                CrearEsquemaCategorias(c);
                CrearEsquemaGenoma(c);
            }

            Repo = new AgScheduleRepository(() => new SQLiteConnection(_connectionString));
        }

        protected SQLiteConnection Abrir() => new SQLiteConnection(_connectionString);

        private void CrearEsquemaCategorias(SQLiteConnection c)
        {
            c.Execute(@"
                CREATE TABLE ag_corrida (
                    id INTEGER PRIMARY KEY,
                    idperiodo INTEGER NOT NULL,
                    idprograma INTEGER NOT NULL,
                    idaula INTEGER NOT NULL DEFAULT 0,
                    fecha_inicio_corrida TEXT NOT NULL DEFAULT '',
                    fecha_fin_corrida TEXT NOT NULL DEFAULT '',
                    estado INTEGER NOT NULL DEFAULT 0,
                    poblacion_tamano INTEGER NOT NULL DEFAULT 0,
                    generaciones_max INTEGER NOT NULL DEFAULT 0,
                    generaciones_ejecutadas INTEGER NOT NULL DEFAULT 0,
                    tasa_cruce REAL NOT NULL DEFAULT 0.85,
                    tasa_mutacion REAL NOT NULL DEFAULT 0.1,
                    elitismo INTEGER NOT NULL DEFAULT 2,
                    mejor_fitness REAL NOT NULL DEFAULT 0,
                    hard_violations INTEGER NOT NULL DEFAULT 0,
                    soft_violations REAL NOT NULL DEFAULT 0,
                    vacantes INTEGER NOT NULL DEFAULT 0,
                    id_individuo_optimo INTEGER NOT NULL DEFAULT 0,
                    parametros_json TEXT NOT NULL DEFAULT '{}',
                    subunidad TEXT NOT NULL DEFAULT ''
                );

                CREATE TABLE ag_individuo (
                    id INTEGER PRIMARY KEY,
                    id_corrida INTEGER NOT NULL,
                    generacion INTEGER NOT NULL DEFAULT 0,
                    fitness REAL NOT NULL DEFAULT 0,
                    hard_violations INTEGER NOT NULL DEFAULT 0,
                    soft_violations REAL NOT NULL DEFAULT 0,
                    estado INTEGER NOT NULL DEFAULT 0,
                    subunidad TEXT NOT NULL DEFAULT ''
                );

                CREATE TABLE ag_vinculaciones (
                    id INTEGER PRIMARY KEY,
                    id_corrida INTEGER NOT NULL,
                    id_individuo INTEGER NOT NULL DEFAULT 0,
                    idperiodo INTEGER NOT NULL DEFAULT 0,
                    idprograma INTEGER NOT NULL DEFAULT 0,
                    idasignatura INTEGER NOT NULL DEFAULT 0,
                    subunidad TEXT NOT NULL DEFAULT ''
                );

                CREATE TABLE ag_sesiones (
                    id INTEGER PRIMARY KEY,
                    idvinculacion INTEGER NOT NULL,
                    idindividuo INTEGER NOT NULL DEFAULT 0,
                    id_aula INTEGER NOT NULL DEFAULT 0,
                    fechasesion TEXT NOT NULL DEFAULT '',
                    horariosesion TEXT NOT NULL DEFAULT '',
                    horas REAL NOT NULL DEFAULT 0
                );

                CREATE TABLE ag_violacion (
                    id INTEGER PRIMARY KEY,
                    id_individuo INTEGER NOT NULL,
                    id_gen INTEGER NULL,
                    regla INTEGER NOT NULL DEFAULT 0,
                    es_dura INTEGER NOT NULL DEFAULT 1,
                    mensaje TEXT NOT NULL DEFAULT '',
                    detalle_json TEXT NOT NULL DEFAULT '{}',
                    subunidad TEXT NOT NULL DEFAULT ''
                );

                CREATE TABLE ag_bitacora (id INTEGER PRIMARY KEY, id_corrida INTEGER NOT NULL);
                CREATE TABLE ag_promocion (id INTEGER PRIMARY KEY, id_corrida INTEGER NOT NULL);
                CREATE TABLE ag_salvedades_conflictos (id INTEGER PRIMARY KEY, id_corrida INTEGER NOT NULL);
                CREATE TABLE ag_vinculacion_gen (id INTEGER PRIMARY KEY, idagvinculacion INTEGER NOT NULL);
            ");
        }

        private void CrearEsquemaGenoma(SQLiteConnection c)
        {
            c.Execute(@"
                CREATE TABLE ag_gen (
                    id INTEGER PRIMARY KEY,
                    id_individuo INTEGER NOT NULL,
                    idperiodo INTEGER NOT NULL DEFAULT 0,
                    idasignatura INTEGER NOT NULL DEFAULT 0,
                    idprograma INTEGER NOT NULL DEFAULT 0,
                    idsede INTEGER NOT NULL DEFAULT 0,
                    semestre TEXT NOT NULL DEFAULT '',
                    grupo TEXT NOT NULL DEFAULT '',
                    idtercero INTEGER NOT NULL DEFAULT 0,
                    id_aula INTEGER NOT NULL DEFAULT 0,
                    id_bloque_horario INTEGER NOT NULL DEFAULT 0,
                    dia_semana INTEGER NOT NULL DEFAULT 0,
                    horas_semanales REAL NOT NULL DEFAULT 0,
                    total_horas_asignatura REAL NOT NULL DEFAULT 0,
                    fecha_inicio_estimada TEXT NOT NULL DEFAULT '',
                    fecha_fin_estimada TEXT NOT NULL DEFAULT '',
                    es_fijo INTEGER NOT NULL DEFAULT 0,
                    es_vacante INTEGER NOT NULL DEFAULT 0,
                    motivo_vacante TEXT NOT NULL DEFAULT '',
                    dentro_de_ventana INTEGER NOT NULL DEFAULT 0,
                    estado INTEGER NOT NULL DEFAULT 0,
                    subunidad TEXT NOT NULL DEFAULT '',
                    indice_dia_patron INTEGER NOT NULL DEFAULT 0,
                    patron_semanal TEXT NOT NULL DEFAULT '',
                    dias_patron TEXT NOT NULL DEFAULT '',
                    clave_grupo TEXT NOT NULL DEFAULT '',
                    horas_gen REAL NOT NULL DEFAULT 0,
                    fechas_gen TEXT NOT NULL DEFAULT ''
                );
            ");
        }

        public void Dispose()
        {
            try
            {
                SQLiteConnection.ClearAllPools();
                if (File.Exists(_dbPath)) File.Delete(_dbPath);
            }
            catch
            {
                // Best effort cleanup (archivo bloqueado por OneDrive/antivirus)
            }
        }

        protected void Ejecutar(string sql, object param = null)
        {
            using (var c = Abrir())
            {
                c.Open();
                c.Execute(sql, param);
            }
        }

        protected void InsertarCorrida(long id, int idPeriodo, int idPrograma, long idIndividuoOptimo = 0)
            => Ejecutar(
                "INSERT INTO ag_corrida (id, idperiodo, idprograma, estado, id_individuo_optimo, subunidad) " +
                "VALUES (@id, @idPeriodo, @idPrograma, 2, @idIndividuoOptimo, 'VALLE');",
                new { id, idPeriodo, idPrograma, idIndividuoOptimo });

        protected void InsertarIndividuo(long id, long idCorrida)
            => Ejecutar(
                "INSERT INTO ag_individuo (id, id_corrida, estado, subunidad) VALUES (@id, @idCorrida, 2, 'VALLE');",
                new { id, idCorrida });

        protected void InsertarVinculacion(int id, long idCorrida, long idIndividuo, int idPeriodo, int idPrograma)
            => Ejecutar(
                "INSERT INTO ag_vinculaciones (id, id_corrida, id_individuo, idperiodo, idprograma, subunidad) " +
                "VALUES (@id, @idCorrida, @idIndividuo, @idPeriodo, @idPrograma, 'VALLE');",
                new { id, idCorrida, idIndividuo, idPeriodo, idPrograma });

        protected void InsertarSesion(int id, int idVinculacion, long idIndividuo)
            => Ejecutar(
                "INSERT INTO ag_sesiones (id, idvinculacion, idindividuo, id_aula, fechasesion, horariosesion, horas) " +
                "VALUES (@id, @idVinculacion, @idIndividuo, 31, '2027-02-04', '08:00 - 12:00', 4);",
                new { id, idVinculacion, idIndividuo });

        protected void InsertarGen(int id, long idIndividuo, int idAula, int idAsignatura)
            => Ejecutar(
                "INSERT INTO ag_gen (id, id_individuo, idperiodo, idasignatura, idprograma, id_aula, horas_gen, total_horas_asignatura, subunidad) " +
                "VALUES (@id, @idIndividuo, 3, @idAsignatura, 2, @idAula, 24, 48, 'VALLE');",
                new { id, idIndividuo, idAula, idAsignatura });

        protected void InsertarViolacion(long idIndividuo)
            => Ejecutar(
                "INSERT INTO ag_violacion (id_individuo, regla, es_dura, mensaje) VALUES (@idIndividuo, 'R1_SuperposicionDocente', 1, 'Solape');",
                new { idIndividuo });

        protected long Contar(string tabla, string where = "1=1")
        {
            using (var c = Abrir())
            {
                c.Open();
                return c.ExecuteScalar<long>($"SELECT COUNT(*) FROM {tabla} WHERE {where};");
            }
        }
    }

    public class AgScheduleRepositoryPurgaTests : AgScheduleRepoHuerfanaFixture
    {
        [Fact]
        public async Task LimpiarStaging_PorPeriodoYPrograma_NoDebeDejarStagingHuerfano()
        {
            // Caso real corrida #48: cabecera (período 3, programa 1 = APT-53) y staging con el programa
            // de la asignatura (programa 2 = APT-52). Antes de la Fase 1563 se borraba la cabecera y el
            // staging quedaba huérfano (sin violaciones y sin promoción).
            InsertarCorrida(48, 3, 1, 33);
            InsertarIndividuo(33, 48);
            InsertarVinculacion(700, 48, 33, 3, 2);
            InsertarSesion(1, 700, 33);
            InsertarGen(1, 33, 31, 148);
            InsertarViolacion(33);

            var res = await Repo.LimpiarStagingAsync(3, 1, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            Contar("ag_corrida").Should().Be(0);
            Contar("ag_individuo").Should().Be(0);
            Contar("ag_gen").Should().Be(0);
            Contar("ag_violacion").Should().Be(0);
            Contar("ag_vinculaciones").Should().Be(0);
            Contar("ag_sesiones").Should().Be(0);
            Contar("ag_salvedades_conflictos").Should().Be(0);
        }

        [Fact]
        public async Task LimpiarStaging_DeStagingSinCabecera_TambienDebePurgarlo()
        {
            // Staging huérfano preexistente que sí coincide con el par solicitado: la purga debe recogerlo
            // aunque no haya fila en ag_corrida (idsDel se alimenta también desde ag_vinculaciones).
            InsertarVinculacion(701, 77, 40, 3, 2);
            InsertarSesion(2, 701, 40);
            InsertarGen(2, 40, 31, 148);

            var res = await Repo.LimpiarStagingAsync(3, 2, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            Contar("ag_vinculaciones").Should().Be(0);
            Contar("ag_sesiones").Should().Be(0);
            Contar("ag_gen", "id_individuo = 40").Should().Be(0);
        }

        [Fact]
        public async Task UpdateCorrida_DeUnaCorridaInexistente_DebeFallarEnVezDeSerNoOp()
        {
            var corrida = new AgCorrida
            {
                Id = 999,
                IdPeriodo = 3,
                IdPrograma = 1,
                Estado = AgEstadoCorrida.Aprobada
            };

            var res = await Repo.UpdateCorridaAsync(corrida, CancellationToken.None);

            res.IsFailure.Should().BeTrue();
            res.Error.Should().Contain("ya no existe en ag_corrida");
        }
    }

    public class AgScheduleRepositoryReparacionTests : AgScheduleRepoHuerfanaFixture
    {
        [Fact]
        public async Task DetectarCorridaHuerfana_DerivaPeriodoProgramaEIndividuoDesdeElStaging()
        {
            InsertarVinculacion(700, 48, 33, 3, 2);
            InsertarSesion(1, 700, 33);
            InsertarGen(1, 33, 31, 148);

            var res = await Repo.DetectarCorridaHuerfanaAsync(48, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            var h = res.Value;
            h.IdPeriodo.Should().Be(3);
            h.IdPrograma.Should().Be(2);
            h.IdIndividuo.Should().Be(33);
            h.TotalVinculaciones.Should().Be(1);
            h.TotalSesiones.Should().Be(1);
            h.ExisteCabecera.Should().BeFalse();
            h.ExisteIndividuo.Should().BeFalse();
            h.PuedeRepararse.Should().BeTrue();
        }

        [Fact]
        public async Task GetGenesByIndividuo_DevuelveGenesAunqueNoExistaElIndividuo()
        {
            InsertarVinculacion(700, 48, 33, 3, 2);
            InsertarGen(1, 33, 31, 148);
            InsertarGen(2, 33, 12, 149);

            var res = await Repo.GetGenesByIndividuoAsync(33, CancellationToken.None);

            res.IsSuccess.Should().BeTrue();
            res.Value.Should().HaveCount(2);
        }

        [Fact]
        public async Task ReinsertarCabecera_ReconstruyeCorridaEIndividuoConLosIdsDelStaging()
        {
            InsertarVinculacion(700, 48, 33, 3, 2);
            InsertarSesion(1, 700, 33);
            InsertarGen(1, 33, 31, 148);

            var individuo = new AgIndividuo
            {
                Id = 33,
                IdCorrida = 48,
                Fitness = -1234.5,
                HardViolations = 2,
                SoftViolations = 1.0,
                Estado = AgEstadoIndividuo.Infactible,
                Subunidad = "VALLE"
            };
            var corrida = new AgCorrida
            {
                Id = 48,
                IdPeriodo = 3,
                IdPrograma = 2,
                Estado = AgEstadoCorrida.Completada,
                MejorFitness = -1234.5,
                HardViolations = 2,
                SoftViolations = 1.0,
                IdIndividuoOptimo = 33,
                Subunidad = "VALLE"
            };

            var ins = await Repo.ReinsertarCabeceraHuerfanaAsync(corrida, 48, individuo, CancellationToken.None);
            ins.IsSuccess.Should().BeTrue();

            Contar("ag_corrida", "id = 48").Should().Be(1);
            Contar("ag_individuo", "id = 33").Should().Be(1);

            var leida = await Repo.GetCorridaByIdAsync(48, CancellationToken.None);
            leida.IsSuccess.Should().BeTrue();
            leida.Value.IdPeriodo.Should().Be(3);
            leida.Value.IdPrograma.Should().Be(2);
            leida.Value.IdIndividuoOptimo.Should().Be(33);
            leida.Value.Estado.Should().Be(AgEstadoCorrida.Completada);

            // Idempotencia: una segunda reparación no debe duplicar ni sobrescribir.
            var segunda = await Repo.ReinsertarCabeceraHuerfanaAsync(corrida, 48, individuo, CancellationToken.None);
            segunda.IsFailure.Should().BeTrue();
            Contar("ag_corrida", "id = 48").Should().Be(1);
            Contar("ag_individuo", "id = 33").Should().Be(1);
        }

        [Fact]
        public async Task ReemplazarViolaciones_ConColumnaReglaInteger_ConservaElNombreDeLaRegla()
        {
            // El vault real declara ag_violacion.regla como INTEGER pero persiste el NOMBRE de la regla.
            // Sin CAST en la lectura, el proveedor devuelve 0 y el panel mostraba «regla 0» (Fase 1567).
            InsertarCorrida(48, 3, 2, 33);
            InsertarIndividuo(33, 48);
            InsertarVinculacion(700, 48, 33, 3, 2);
            InsertarGen(1, 33, 31, 148);

            var violaciones = new[]
            {
                new AgViolacion(AgReglaViolacion.R1_SuperposicionDocente, true, "Solape de docente"),
                new AgViolacion(AgReglaViolacion.S1_HuecosHorarios, false, "Hueco horario")
            };

            var rep = await Repo.ReemplazarViolacionesAsync(
                33, violaciones, 1, 1.0, AgEstadoIndividuo.Infactible, 48, CancellationToken.None);
            rep.IsSuccess.Should().BeTrue();

            var leidas = await Repo.GetViolacionesByIndividuoAsync(33, CancellationToken.None);

            leidas.IsSuccess.Should().BeTrue();
            leidas.Value.Should().HaveCount(2);
            leidas.Value.Should().Contain(v => v.Regla == AgReglaViolacion.R1_SuperposicionDocente);
            leidas.Value.Should().Contain(v => v.Regla == AgReglaViolacion.S1_HuecosHorarios);
        }

    }
}

