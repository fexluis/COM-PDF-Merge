using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Sync;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>
    /// Repositorio para la persistencia de datos del AG (aulas, bloques, corridas, individuos, genes y staging).
    /// </summary>
    public class AgScheduleRepository : IAgScheduleRepository
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public AgScheduleRepository(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        private static void NotifyMarkDirty()
        {
            try { SyncContext.Scheduler?.MarkDirty(); } catch { }
        }

        public async Task<Result<IReadOnlyCollection<AgAula>>> GetAulasAsync(string subunidad, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sql = @"
SELECT id as Id, idsede as IdSede, idperiodo as IdPeriodo, idprograma as IdPrograma, codigo as Codigo, 
       nombre as Nombre, capacidad as Capacidad, tipo_aula as TipoAula, es_virtual as EsVirtual, 
       estado as Estado, caracteristicas_json as CaracteristicasJson, semestre as Semestre, 
       grupo as Grupo, modalidadVinculacion as ModalidadVinculacion, fecha_inicio as FechaInicio, 
       fecha_fin as FechaFin, semestre_inicio as SemestreInicio, semestre_fin as SemestreFin, 
       ventana_origen as VentanaOrigen, aula_ref as AulaRef, subunidad as Subunidad,
       jornada as Jornada, horas_dia as HorasDia, dias_ventana as DiasVentana
FROM ag_aula 
WHERE estado IN (1, 4)
  AND (subunidad = @subunidad OR subunidad = '' OR subunidad IS NULL)
  AND COALESCE(subunidad, '') <> 'DEMO'
  AND NOT (idperiodo = 0 AND codigo IN ('A-101', 'A-102', 'A-201', 'LAB-1', 'VIR-01'))
ORDER BY nombre";
                    var items = (await db.QueryAsync<AgAula>(new CommandDefinition(sql, new { subunidad }, cancellationToken: ct))).ToList();
                    return Result.Success<IReadOnlyCollection<AgAula>>(items.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgAula>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgScheduleRepository] Error al obtener aulas", ex);
                return Result.Failure<IReadOnlyCollection<AgAula>>($"Error al leer aulas: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<AgAula>>> GetAulasPorPeriodoAsync(
            int idPeriodo, int? idSede, int? idPrograma, string semestre, string grupo, string subunidad, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sql = @"
SELECT id as Id, idsede as IdSede, idperiodo as IdPeriodo, idprograma as IdPrograma, codigo as Codigo, 
       nombre as Nombre, capacidad as Capacidad, tipo_aula as TipoAula, es_virtual as EsVirtual, 
       estado as Estado, caracteristicas_json as CaracteristicasJson, semestre as Semestre, 
       grupo as Grupo, modalidadVinculacion as ModalidadVinculacion, fecha_inicio as FechaInicio, 
       fecha_fin as FechaFin, semestre_inicio as SemestreInicio, semestre_fin as SemestreFin, 
       ventana_origen as VentanaOrigen, aula_ref as AulaRef, subunidad as Subunidad,
       jornada as Jornada, horas_dia as HorasDia, dias_ventana as DiasVentana
FROM ag_aula 
WHERE estado IN (1, 4)
  AND idperiodo = @idPeriodo
  AND (@idSede IS NULL OR idsede = @idSede)
  AND (@idPrograma IS NULL OR idprograma = @idPrograma)
  AND (@semestre IS NULL OR @semestre = '' OR semestre = @semestre)
  AND (@grupo IS NULL OR @grupo = '' OR grupo = @grupo)
  AND (subunidad = @subunidad OR subunidad = '' OR subunidad IS NULL)
  AND COALESCE(subunidad, '') <> 'DEMO'
  AND NOT (idperiodo = 0 AND codigo IN ('A-101', 'A-102', 'A-201', 'LAB-1', 'VIR-01'))
ORDER BY id DESC";

                    var items = (await db.QueryAsync<AgAula>(new CommandDefinition(
                        sql,
                        new { idPeriodo, idSede, idPrograma, semestre, grupo, subunidad },
                        cancellationToken: ct))).ToList();

                    return Result.Success<IReadOnlyCollection<AgAula>>(items.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgAula>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgScheduleRepository] Error al obtener aulas por período", ex);
                return Result.Failure<IReadOnlyCollection<AgAula>>($"Error al leer aulas por período: {ex.Message}");
            }
        }

        public async Task<Result<AgAula>> GetAulaByIdAsync(long idAula, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sql = "SELECT id as Id, idsede as IdSede, idperiodo as IdPeriodo, idprograma as IdPrograma, codigo as Codigo, nombre as Nombre, capacidad as Capacidad, tipo_aula as TipoAula, es_virtual as EsVirtual, estado as Estado, caracteristicas_json as CaracteristicasJson, semestre as Semestre, grupo as Grupo, modalidadVinculacion as ModalidadVinculacion, fecha_inicio as FechaInicio, fecha_fin as FechaFin, semestre_inicio as SemestreInicio, semestre_fin as SemestreFin, ventana_origen as VentanaOrigen, aula_ref as AulaRef, subunidad as Subunidad, jornada as Jornada, horas_dia as HorasDia, dias_ventana as DiasVentana FROM ag_aula WHERE id = @idAula";
                    var item = await db.QuerySingleOrDefaultAsync<AgAula>(new CommandDefinition(sql, new { idAula }, cancellationToken: ct));
                    if (item == null) return Result.Failure<AgAula>("Aula no encontrada.");
                    return Result.Success(item);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<AgAula>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error al obtener aula {idAula}", ex);
                return Result.Failure<AgAula>($"Error: {ex.Message}");
            }
        }

        public async Task<Result<long>> InsertAulaAsync(AgAula aula, CancellationToken ct)
        {
            if (aula == null) return Result.Failure<long>("Aula nula");
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sql = @"
INSERT INTO ag_aula (idsede, idperiodo, idprograma, codigo, nombre, capacidad, tipo_aula, es_virtual, estado, caracteristicas_json, semestre, grupo, modalidadVinculacion, fecha_inicio, fecha_fin, semestre_inicio, semestre_fin, ventana_origen, aula_ref, subunidad, jornada, horas_dia, dias_ventana)
VALUES (@IdSede, @IdPeriodo, @IdPrograma, @Codigo, @Nombre, @Capacidad, @TipoAula, @EsVirtual, @Estado, @CaracteristicasJson, @Semestre, @Grupo, @ModalidadVinculacion, @FechaInicio, @FechaFin, @SemestreInicio, @SemestreFin, @VentanaOrigen, @AulaRef, @Subunidad, @Jornada, @HorasDia, @DiasVentana);
SELECT last_insert_rowid();";

                    long id = await db.ExecuteScalarAsync<long>(new CommandDefinition(sql, aula, cancellationToken: ct));
                    aula.Id = (int)id;
                    NotifyMarkDirty();
                    return Result.Success(id);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<long>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgScheduleRepository] Error insertando aula", ex);
                return Result.Failure<long>($"Error al insertar aula: {ex.Message}");
            }
        }

        public async Task<Result> InsertAulaAsignaturasAsync(long idAula, IEnumerable<int> idAsignaturas, CancellationToken ct)
        {
            if (idAsignaturas == null) return Result.Success();
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    using (var trans = db.BeginTransaction())
                    {
                        // Fase 1750: No borrar indiscriminadamente; preservar patrones y origen FIJO/LEGADO existentes
                        string sqlDel = @"
DELETE FROM ag_aula_asignatura
WHERE id_aula = @idAula
  AND UPPER(COALESCE(origen_patron, '')) = 'AUTOMATICO'
  AND id_asignatura NOT IN @idAsignaturas;";
                        await db.ExecuteAsync(new CommandDefinition(sqlDel, new { idAula, idAsignaturas }, trans, cancellationToken: ct));

                        string sqlIns = @"
INSERT INTO ag_aula_asignatura (id_aula, id_asignatura, exclusivo, prioridad, dias_patron, subunidad, origen_patron)
VALUES (@idAula, @idAsignatura, 0, 0, '', @subunidad, 'AUTOMATICO')
ON CONFLICT(id_aula, id_asignatura) DO NOTHING;";

                        string subunidad = SyncContext.ActiveSubunidad ?? string.Empty;
                        foreach (var idAsig in idAsignaturas)
                        {
                            ct.ThrowIfCancellationRequested();
                            await db.ExecuteAsync(new CommandDefinition(sqlIns, new { idAula, idAsignatura = idAsig, subunidad }, trans, cancellationToken: ct));
                        }

                        trans.Commit();
                        NotifyMarkDirty();
                        return Result.Success();
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error insertando asignaturas para aula {idAula}", ex);
                return Result.Failure($"Error al asociar asignaturas al aula: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<AgBloqueHorario>>> GetBloquesHorariosAsync(string subunidad, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    // Fase 1002 & Fase 1007: incluir bloques globales ('') y específicos de subunidad, priorizando los específicos
                    string sql = @"SELECT id as Id, dia_semana as DiaSemanaNumero, hora_inicio as HoraInicio, hora_fin as HoraFin,
       duracion_horas as DuracionHoras, franja as Franja, activo as Activo, subunidad as Subunidad
FROM ag_bloque_horario
WHERE activo = 1
  AND (@Subunidad = '' OR COALESCE(subunidad, '') IN ('', @Subunidad))
ORDER BY CASE WHEN COALESCE(subunidad, '') = @Subunidad THEN 0 ELSE 1 END, dia_semana, hora_inicio";
                    var items = (await db.QueryAsync<AgBloqueHorario>(
                        new CommandDefinition(sql, new { Subunidad = subunidad ?? string.Empty }, cancellationToken: ct))).ToList();

                    return Result.Success<IReadOnlyCollection<AgBloqueHorario>>(items.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgBloqueHorario>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgScheduleRepository] Error al obtener bloques horarios", ex);
                return Result.Failure<IReadOnlyCollection<AgBloqueHorario>>($"Error al leer bloques: {ex.Message}");
            }
        }

        public async Task<Result<long>> InsertCorridaAsync(AgCorrida corrida, CancellationToken ct)
        {
            if (corrida == null) return Result.Failure<long>("Corrida nula");
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    bool hayColVac = ColumnaExiste(db, null, "ag_corrida", "vacantes_consentidas");
                    string cols = hayColVac ? ", vacantes_consentidas" : "";
                    string vals = hayColVac ? ", @VacantesConsentidas" : "";

                    string sql = $@"
INSERT INTO ag_corrida (idperiodo, idprograma, idaula, fecha_inicio_corrida, fecha_fin_corrida, estado, poblacion_tamano, generaciones_max, generaciones_ejecutadas, tasa_cruce, tasa_mutacion, elitismo, mejor_fitness, hard_violations, soft_violations, vacantes, id_individuo_optimo, parametros_json, subunidad{cols})
VALUES (@IdPeriodo, @IdPrograma, @IdAula, @FechaInicioCorrida, @FechaFinCorrida, @Estado, @PoblacionTamano, @GeneracionesMax, @GeneracionesEjecutadas, @TasaCruce, @TasaMutacion, @Elitismo, @MejorFitness, @HardViolations, @SoftViolations, @Vacantes, @IdIndividuoOptimo, @ParametrosJson, @Subunidad{vals});
SELECT last_insert_rowid();";

                    var param = new
                    {
                        corrida.IdPeriodo,
                        corrida.IdPrograma,
                        corrida.IdAula,
                        FechaInicioCorrida = corrida.FechaInicioCorrida ?? string.Empty,
                        FechaFinCorrida = corrida.FechaFinCorrida ?? string.Empty,
                        corrida.Estado,
                        corrida.PoblacionTamano,
                        corrida.GeneracionesMax,
                        corrida.GeneracionesEjecutadas,
                        corrida.TasaCruce,
                        corrida.TasaMutacion,
                        corrida.Elitismo,
                        corrida.MejorFitness,
                        corrida.HardViolations,
                        corrida.SoftViolations,
                        corrida.Vacantes,
                        corrida.IdIndividuoOptimo,
                        ParametrosJson = corrida.ParametrosJson ?? "{}",
                        Subunidad = corrida.Subunidad ?? string.Empty,
                        VacantesConsentidas = corrida.vacantes_consentidas ?? string.Empty
                    };

                    long id = await db.ExecuteScalarAsync<long>(new CommandDefinition(sql, param, cancellationToken: ct));
                    corrida.Id = id;
                    NotifyMarkDirty();
                    return Result.Success(id);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<long>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgScheduleRepository] Error insertando corrida", ex);
                return Result.Failure<long>($"Error al insertar corrida: {ex.Message}");
            }
        }

        public async Task<Result> UpdateCorridaAsync(AgCorrida corrida, CancellationToken ct)
        {
            if (corrida == null) return Result.Failure("Corrida nula");
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    bool hayColVac = ColumnaExiste(db, null, "ag_corrida", "vacantes_consentidas");
                    string setVac = hayColVac ? ",\n    vacantes_consentidas = @VacantesConsentidas" : "";

                    string sql = $@"
UPDATE ag_corrida SET
    fecha_fin_corrida = @FechaFinCorrida,
    estado = @Estado,
    generaciones_ejecutadas = @GeneracionesEjecutadas,
    mejor_fitness = @MejorFitness,
    hard_violations = @HardViolations,
    soft_violations = @SoftViolations,
    vacantes = @Vacantes,
    id_individuo_optimo = @IdIndividuoOptimo,
    parametros_json = @ParametrosJson{setVac}
WHERE id = @Id;";

                    var param = new
                    {
                        FechaFinCorrida = corrida.FechaFinCorrida ?? string.Empty,
                        corrida.Estado,
                        corrida.GeneracionesEjecutadas,
                        corrida.MejorFitness,
                        corrida.HardViolations,
                        corrida.SoftViolations,
                        corrida.Vacantes,
                        corrida.IdIndividuoOptimo,
                        ParametrosJson = corrida.ParametrosJson ?? "{}",
                        VacantesConsentidas = corrida.vacantes_consentidas ?? string.Empty,
                        corrida.Id
                    };

                    int filasAfectadas = await db.ExecuteAsync(new CommandDefinition(sql, param, cancellationToken: ct));
                    if (filasAfectadas <= 0)
                    {
                        // Fase 1564: un UPDATE sobre una corrida inexistente era un no-op silencioso.
                        // Si la corrida fue purgada o eliminada, el flujo continuaba como si hubiera persistido
                        // (y el Módulo R quedaba luego con staging huérfano sin cabecera).
                        LoggerHelper.Warning(
                            $"[AgScheduleRepository] UpdateCorridaAsync no encontró la corrida #{corrida.Id} (0 filas afectadas).");
                        return Result.Failure(
                            $"La corrida #{corrida.Id} ya no existe en ag_corrida (fue purgada o eliminada).");
                    }

                    NotifyMarkDirty();
                    return Result.Success();
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgScheduleRepository] Error actualizando corrida", ex);
                return Result.Failure($"Error al actualizar corrida: {ex.Message}");
            }
        }

        public async Task<Result<AgCorrida>> GetCorridaByIdAsync(long idCorrida, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    bool hayColVac = ColumnaExiste(db, null, "ag_corrida", "vacantes_consentidas");
                    string colVac = hayColVac ? ", vacantes_consentidas as vacantes_consentidas" : "";
                    string sql = $"SELECT id as Id, idperiodo as IdPeriodo, idprograma as IdPrograma, idaula as IdAula, fecha_inicio_corrida as FechaInicioCorrida, fecha_fin_corrida as FechaFinCorrida, estado as Estado, poblacion_tamano as PoblacionTamano, generaciones_max as GeneracionesMax, generaciones_ejecutadas as GeneracionesEjecutadas, tasa_cruce as TasaCruce, tasa_mutacion as TasaMutacion, elitismo as Elitismo, mejor_fitness as MejorFitness, hard_violations as HardViolations, soft_violations as SoftViolations, vacantes as Vacantes, id_individuo_optimo as IdIndividuoOptimo, parametros_json as ParametrosJson, subunidad as Subunidad{colVac} FROM ag_corrida WHERE id = @idCorrida";
                    var item = await db.QuerySingleOrDefaultAsync<AgCorrida>(new CommandDefinition(sql, new { idCorrida }, cancellationToken: ct));
                    if (item == null) return Result.Failure<AgCorrida>("Corrida no encontrada.");
                    return Result.Success(item);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<AgCorrida>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error al obtener corrida {idCorrida}", ex);
                return Result.Failure<AgCorrida>($"Error: {ex.Message}");
            }
        }

        public async Task<Result<CuadreCorrida>> GetCuadreCorridaAsync(long idCorrida, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sql = @"
SELECT 
    COUNT(*) as Sesiones,
    COALESCE(SUM(s.horas), 0.0) as Horas,
    COALESCE(SUM(CASE WHEN s.es_vacante = 1 THEN 1 ELSE 0 END), 0) as SesionesVacantes
FROM ag_sesiones s
INNER JOIN ag_vinculaciones v ON v.id = s.idvinculacion
WHERE v.id_corrida = @idCorrida;

SELECT 
    v.idprograma as IdPrograma,
    COUNT(*) as Sesiones,
    COALESCE(SUM(s.horas), 0.0) as Horas
FROM ag_sesiones s
INNER JOIN ag_vinculaciones v ON v.id = s.idvinculacion
WHERE v.id_corrida = @idCorrida
GROUP BY v.idprograma;

SELECT COUNT(*) as Huerfanas
FROM ag_sesiones s
LEFT JOIN ag_vinculaciones v ON v.id = s.idvinculacion
WHERE v.id IS NULL;
";
                    using (var multi = await db.QueryMultipleAsync(new CommandDefinition(sql, new { idCorrida }, cancellationToken: ct)))
                    {
                        var resumen = await multi.ReadFirstOrDefaultAsync<dynamic>();
                        int totalSesiones = resumen != null ? Convert.ToInt32(resumen.Sesiones) : 0;
                        double totalHoras = resumen != null ? Convert.ToDouble(resumen.Horas) : 0.0;
                        int totalVacantes = resumen != null ? Convert.ToInt32(resumen.SesionesVacantes) : 0;

                        var progRows = await multi.ReadAsync<dynamic>();
                        var porPrograma = new List<(int IdPrograma, int Sesiones, double Horas)>();
                        foreach (var r in progRows)
                        {
                            int idProg = Convert.ToInt32(r.IdPrograma);
                            int ses = Convert.ToInt32(r.Sesiones);
                            double hrs = Convert.ToDouble(r.Horas);
                            porPrograma.Add((idProg, ses, hrs));
                        }

                        var huerfanasRes = await multi.ReadFirstOrDefaultAsync<dynamic>();
                        int totalHuerfanas = huerfanasRes != null ? Convert.ToInt32(huerfanasRes.Huerfanas) : 0;

                        var cuadre = new CuadreCorrida
                        {
                            Sesiones = totalSesiones,
                            Horas = totalHoras,
                            SesionesVacantes = totalVacantes,
                            SesionesHuerfanas = totalHuerfanas,
                            PorPrograma = porPrograma.AsReadOnly(),
                            RequiereCompleto = true
                        };

                        return Result.Success(cuadre);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<CuadreCorrida>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error al obtener cuadre de corrida #{idCorrida}", ex);
                return Result.Failure<CuadreCorrida>($"Error al obtener cuadre de corrida: {ex.Message}");
            }
        }

        public async Task<Result<AgCorrida>> GetUltimaCorridaAsync(int idPeriodo, int idPrograma, string subunidad, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    bool hayColVac = ColumnaExiste(db, null, "ag_corrida", "vacantes_consentidas");
                    string colVac = hayColVac ? ", vacantes_consentidas as vacantes_consentidas" : "";
                    string sql = "SELECT id as Id, idperiodo as IdPeriodo, idprograma as IdPrograma, idaula as IdAula, " +
                                 "fecha_inicio_corrida as FechaInicioCorrida, fecha_fin_corrida as FechaFinCorrida, estado as Estado, " +
                                 "poblacion_tamano as PoblacionTamano, generaciones_max as GeneracionesMax, " +
                                 "generaciones_ejecutadas as GeneracionesEjecutadas, tasa_cruce as TasaCruce, tasa_mutacion as TasaMutacion, " +
                                 "elitismo as Elitismo, mejor_fitness as MejorFitness, hard_violations as HardViolations, " +
                                 "soft_violations as SoftViolations, vacantes as Vacantes, id_individuo_optimo as IdIndividuoOptimo, " +
                                 $"parametros_json as ParametrosJson, subunidad as Subunidad{colVac} " +
                                 "FROM ag_corrida WHERE idperiodo = @idPeriodo AND idprograma = @idPrograma " +
                                 "AND (@subunidad IS NULL OR subunidad = @subunidad) ORDER BY id DESC LIMIT 1";

                    var item = await db.QuerySingleOrDefaultAsync<AgCorrida>(
                        new CommandDefinition(sql, new { idPeriodo, idPrograma, subunidad }, cancellationToken: ct));

                    if (item == null) return Result.Failure<AgCorrida>("No hay corridas registradas.");
                    return Result.Success(item);
                }
            }
            catch (OperationCanceledException) { return Result.Failure<AgCorrida>("Operación cancelada."); }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error al obtener última corrida {idPeriodo}/{idPrograma}", ex);
                return Result.Failure<AgCorrida>($"Error: {ex.Message}");
            }
        }

        public async Task<Result<long>> SaveIndividuoAsync(AgIndividuo individuo, CancellationToken ct)
        {
            if (individuo == null) return Result.Failure<long>("Individuo nulo");
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    using (var trans = db.BeginTransaction())
                    {
                        string sqlInd = @"
INSERT INTO ag_individuo (id_corrida, generacion, fitness, hard_violations, soft_violations, estado, subunidad)
VALUES (@IdCorrida, @Generacion, @Fitness, @HardViolations, @SoftViolations, @Estado, @Subunidad);
SELECT last_insert_rowid();";

                        var indParam = new
                        {
                            individuo.IdCorrida,
                            individuo.Generacion,
                            individuo.Fitness,
                            individuo.HardViolations,
                            individuo.SoftViolations,
                            individuo.Estado,
                            Subunidad = individuo.Subunidad ?? string.Empty
                        };

                        long idInd = await db.ExecuteScalarAsync<long>(new CommandDefinition(sqlInd, indParam, trans, cancellationToken: ct));
                        individuo.Id = idInd;

                        if (individuo.Genes != null && individuo.Genes.Count > 0)
                        {
                            foreach (var g in individuo.Genes)
                            {
                                ct.ThrowIfCancellationRequested();
                                g.IdIndividuo = idInd;
                                string sqlGen = @"
INSERT INTO ag_gen (id_individuo, idperiodo, idasignatura, idprograma, idsede, semestre, grupo, idtercero, id_aula, id_bloque_horario, dia_semana, horas_semanales, total_horas_asignatura, fecha_inicio_estimada, fecha_fin_estimada, es_fijo, es_vacante, motivo_vacante, dentro_de_ventana, estado, subunidad, indice_dia_patron, patron_semanal, dias_patron, clave_grupo, horas_gen, fechas_gen)
VALUES (@IdIndividuo, @IdPeriodo, @IdAsignatura, @IdPrograma, @IdSede, @Semestre, @Grupo, @IdTercero, @IdAula, @IdBloqueHorario, @DiaSemana, @HorasSemanales, @TotalHorasAsignatura, @FechaInicioEstimada, @FechaFinEstimada, @EsFijo, @EsVacante, @MotivoVacante, @DentroDeVentana, @Estado, @Subunidad, @IndiceDiaPatron, @PatronSemanal, @DiasPatron, @ClaveGrupo, @HorasGen, @FechasGenStr);
SELECT last_insert_rowid();";
                                var genParam = new
                                {
                                    g.IdIndividuo,
                                    g.IdPeriodo,
                                    g.IdAsignatura,
                                    g.IdPrograma,
                                    g.IdSede,
                                    g.Semestre,
                                    Grupo = g.Grupo ?? string.Empty,
                                    g.IdTercero,
                                    g.IdAula,
                                    g.IdBloqueHorario,
                                    g.DiaSemana,
                                    g.HorasSemanales,
                                    g.TotalHorasAsignatura,
                                    FechaInicioEstimada = g.FechaInicioEstimada.HasValue ? g.FechaInicioEstimada.Value.ToString("yyyy-MM-dd") : string.Empty,
                                    FechaFinEstimada = g.FechaFinEstimada.HasValue ? g.FechaFinEstimada.Value.ToString("yyyy-MM-dd") : string.Empty,
                                    g.EsFijo,
                                    g.EsVacante,
                                    MotivoVacante = g.MotivoVacante ?? string.Empty,
                                    g.DentroDeVentana,
                                    g.Estado,
                                    Subunidad = g.Subunidad ?? string.Empty,
                                    g.IndiceDiaPatron,
                                    PatronSemanal = g.PatronSemanal ?? string.Empty,
                                    DiasPatron = g.DiasPatron ?? string.Empty,
                                    ClaveGrupo = g.ClaveGrupo ?? string.Empty,
                                    g.HorasGen,
                                    FechasGenStr = g.FechasGenStr ?? string.Empty
                                };
                                g.Id = await db.ExecuteScalarAsync<long>(new CommandDefinition(sqlGen, genParam, trans, cancellationToken: ct));
                            }
                        }

                        if (individuo.Violaciones != null && individuo.Violaciones.Count > 0 && individuo.Genes != null && individuo.Genes.Count > 0)
                        {
                            var geneIds = individuo.Genes.Select(g => (long)g.Id).ToList();
                            ResolverIndicesGenes(individuo.Violaciones, geneIds);
                        }

                        if (individuo.Violaciones != null && individuo.Violaciones.Count > 0)
                        {
                            foreach (var v in individuo.Violaciones)
                            {
                                ct.ThrowIfCancellationRequested();
                                v.IdIndividuo = idInd;
                                string sqlViol = @"
INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
VALUES (@IdIndividuo, @IdGen, @Regla, @EsDura, @Mensaje, @DetalleJson, @Subunidad);";
                                var violParam = new
                                {
                                    v.IdIndividuo,
                                    v.IdGen,
                                    Regla = (int)v.Regla, // Fase 1914 / Fase 1925: Persistir enum como entero, no como string
                                    v.EsDura,
                                    Mensaje = v.Mensaje ?? string.Empty,
                                    DetalleJson = v.DetalleJson ?? "{}",
                                    Subunidad = v.Subunidad ?? string.Empty
                                };
                                await db.ExecuteAsync(new CommandDefinition(sqlViol, violParam, trans, cancellationToken: ct));
                            }
                        }

                        trans.Commit();
                        NotifyMarkDirty();
                        return Result.Success(idInd);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<long>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgScheduleRepository] Error guardando individuo", ex);
                return Result.Failure<long>($"Error al guardar individuo: {ex.Message}");
            }
        }

        public async Task<Result<AgIndividuo>> GetIndividuoByIdAsync(long idIndividuo, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sqlInd = "SELECT id as Id, id_corrida as IdCorrida, generacion as Generacion, fitness as Fitness, hard_violations as HardViolations, soft_violations as SoftViolations, estado as Estado, subunidad as Subunidad FROM ag_individuo WHERE id = @idIndividuo";
                    var ind = await db.QuerySingleOrDefaultAsync<AgIndividuo>(new CommandDefinition(sqlInd, new { idIndividuo }, cancellationToken: ct));

                    if (ind == null) return Result.Failure<AgIndividuo>("Individuo no encontrado.");

                    // Fase 1566: tolerar fechas vacías en genes legados ('' no es convertible a DateTime? y
                    // abortaría la rehidratación justo después de reparar una corrida huérfana).
                    string sqlGenes = "SELECT id as Id, id_individuo as IdIndividuo, idperiodo as IdPeriodo, idasignatura as IdAsignatura, idprograma as IdPrograma, idsede as IdSede, semestre as Semestre, grupo as Grupo, idtercero as IdTercero, id_aula as IdAula, id_bloque_horario as IdBloqueHorario, dia_semana as DiaSemana, horas_semanales as HorasSemanales, total_horas_asignatura as TotalHorasAsignatura, NULLIF(TRIM(fecha_inicio_estimada), '') as FechaInicioEstimada, NULLIF(TRIM(fecha_fin_estimada), '') as FechaFinEstimada, es_fijo as EsFijo, es_vacante as EsVacante, motivo_vacante as MotivoVacante, dentro_de_ventana as DentroDeVentana, estado as Estado, subunidad as Subunidad, indice_dia_patron as IndiceDiaPatron, patron_semanal as PatronSemanal, dias_patron as DiasPatron, clave_grupo as ClaveGrupo, horas_gen as HorasGen, fechas_gen AS FechasGenStr FROM ag_gen WHERE id_individuo = @idIndividuo";
                    ind.Genes = (await db.QueryAsync<AgGen>(new CommandDefinition(sqlGenes, new { idIndividuo }, cancellationToken: ct))).ToList();

                    return Result.Success(ind);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<AgIndividuo>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error obteniendo individuo {idIndividuo}", ex);
                return Result.Failure<AgIndividuo>($"Error al obtener individuo: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<AgViolacion>>> GetViolacionesByIndividuoAsync(long idIndividuo, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    // Fase 1567: la columna ag_violacion.regla está declarada INTEGER en el vault, pero se
                    // persiste el NOMBRE de la regla como texto. Sin CAST el proveedor devuelve 0 y el panel
                    // mostraba "regla 0". El CAST preserva el nombre y Dapper lo mapea al enum AgReglaViolacion.

                    const string sql = @"
SELECT id AS Id, id_individuo AS IdIndividuo, id_gen AS IdGen, CAST(regla AS TEXT) AS Regla,
       es_dura AS EsDura, mensaje AS Mensaje, detalle_json AS DetalleJson, subunidad AS Subunidad
FROM ag_violacion
WHERE id_individuo = @idIndividuo
ORDER BY es_dura DESC, regla;";

                    var items = (await db.QueryAsync<AgViolacion>(
                        new CommandDefinition(sql, new { idIndividuo }, cancellationToken: ct))).ToList();
                    return Result.Success<IReadOnlyCollection<AgViolacion>>(items.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgViolacion>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error obteniendo violaciones de individuo {idIndividuo}", ex);
                return Result.Failure<IReadOnlyCollection<AgViolacion>>("Error: " + ex.Message);
            }
        }

        public async Task<Result<IReadOnlyCollection<AgIndividuo>>> GetIndividuosByCorridaAsync(long idCorrida, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sql = "SELECT id as Id, id_corrida as IdCorrida, generacion as Generacion, fitness as Fitness, hard_violations as HardViolations, soft_violations as SoftViolations, estado as Estado, subunidad as Subunidad FROM ag_individuo WHERE id_corrida = @idCorrida ORDER BY fitness DESC";
                    var items = (await db.QueryAsync<AgIndividuo>(new CommandDefinition(sql, new { idCorrida }, cancellationToken: ct))).ToList();
                    return Result.Success<IReadOnlyCollection<AgIndividuo>>(items.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgIndividuo>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error obteniendo individuos de corrida {idCorrida}", ex);
                return Result.Failure<IReadOnlyCollection<AgIndividuo>>($"Error: {ex.Message}");
            }
        }

        public async Task<Result> SaveBitacoraAsync(AgBitacora bitacora, CancellationToken ct)
        {
            if (bitacora == null) return Result.Failure("Bitácora nula");
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sql = @"
INSERT INTO ag_bitacora (id_corrida, generacion, mejor_fitness, promedio_fitness, peor_fitness, hard_violations, evaluados, descartados, vacantes, fuera_de_ventana, duracion_ms, fecha_registro, subunidad)
VALUES (@IdCorrida, @Generacion, @MejorFitness, @PromedioFitness, @PeorFitness, @HardViolations, @Evaluados, @Descartados, @Vacantes, @FueraDeVentana, @DuracionMs, @FechaRegistro, @Subunidad);";

                    var bitParam = new
                    {
                        bitacora.IdCorrida,
                        bitacora.Generacion,
                        bitacora.MejorFitness,
                        bitacora.PromedioFitness,
                        bitacora.PeorFitness,
                        bitacora.HardViolations,
                        bitacora.Evaluados,
                        bitacora.Descartados,
                        bitacora.Vacantes,
                        bitacora.FueraDeVentana,
                        bitacora.DuracionMs,
                        FechaRegistro = bitacora.FechaRegistro ?? string.Empty,
                        Subunidad = bitacora.Subunidad ?? string.Empty
                    };

                    await db.ExecuteAsync(new CommandDefinition(sql, bitParam, cancellationToken: ct));
                    NotifyMarkDirty();
                    return Result.Success();
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al guardar bitácora: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<AgBitacora>>> GetBitacorasByCorridaAsync(long idCorrida, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sql = "SELECT id as Id, id_corrida as IdCorrida, generacion as Generacion, mejor_fitness as MejorFitness, promedio_fitness as PromedioFitness, peor_fitness as PeorFitness, hard_violations as HardViolations, evaluados as Evaluados, descartados as Descartados, vacantes as Vacantes, fuera_de_ventana as FueraDeVentana, duracion_ms as DuracionMs, fecha_registro as FechaRegistro, subunidad as Subunidad FROM ag_bitacora WHERE id_corrida = @idCorrida ORDER BY generacion";
                    var items = (await db.QueryAsync<AgBitacora>(new CommandDefinition(sql, new { idCorrida }, cancellationToken: ct))).ToList();
                    return Result.Success<IReadOnlyCollection<AgBitacora>>(items.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgBitacora>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                return Result.Failure<IReadOnlyCollection<AgBitacora>>($"Error: {ex.Message}");
            }
        }

        public async Task<Result> GuardarStagingAsync(long idCorrida, long idIndividuo, IEnumerable<AgVinculacion> vinculaciones, IEnumerable<AgSesion> sesiones, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    using (var trans = db.BeginTransaction())
                    {
                        var mapVinc = new Dictionary<int, int>();

                        foreach (var v in vinculaciones)
                        {
                            ct.ThrowIfCancellationRequested();
                            int localId = v.Id;
                            string sqlV = @"
INSERT INTO ag_vinculaciones (id_corrida, id_individuo, idtercero, idprograma, idsede, idasignatura, idperiodo, semestre, grupo, dia_horas_semanales, modalidadVinculacion, total_horas, fecha_inicio, fecha_fin, estado, isTerminalAereo, idTerminales, gastoTerminalAerea, es_vacante, motivo_vacante, candidatos_pool, subunidad)
VALUES (@IdCorrida, @IdIndividuo, @IdTercero, @IdPrograma, @IdSede, @IdAsignatura, @IdPeriodo, @Semestre, @Grupo, @DiaHorasSemanales, @ModalidadVinculacion, @TotalHoras, @FechaInicio, @FechaFin, @Estado, @IsTerminalAereo, @IdTerminales, @GastoTerminalAerea, @EsVacante, @MotivoVacante, @CandidatosPool, @Subunidad);
SELECT last_insert_rowid();";

                            var vincParam = new
                            {
                                v.IdCorrida,
                                v.IdIndividuo,
                                v.IdTercero,
                                v.IdPrograma,
                                v.IdSede,
                                v.IdAsignatura,
                                v.IdPeriodo,
                                v.Semestre,
                                Grupo = v.Grupo ?? string.Empty,
                                DiaHorasSemanales = v.DiaHorasSemanales ?? string.Empty,
                                ModalidadVinculacion = v.ModalidadVinculacion ?? string.Empty,
                                v.TotalHoras,
                                FechaInicio = v.FechaInicio ?? string.Empty,
                                FechaFin = v.FechaFin ?? string.Empty,
                                v.Estado,
                                v.IsTerminalAereo,
                                v.IdTerminales,
                                v.GastoTerminalAerea,
                                v.EsVacante,
                                MotivoVacante = v.MotivoVacante ?? string.Empty,
                                v.CandidatosPool,
                                Subunidad = v.Subunidad ?? string.Empty
                            };

                            int dbId = (int)await db.ExecuteScalarAsync<long>(new CommandDefinition(sqlV, vincParam, trans, cancellationToken: ct));
                            v.Id = dbId;
                            mapVinc[localId] = dbId;

                            // Fase 139: Alimentar tabla puente ag_vinculacion_gen para genealogía y trazabilidad N:M
                            if (v.GenesOrigen != null && v.GenesOrigen.Count > 0)
                            {
                                foreach (var gId in v.GenesOrigen)
                                {
                                    ct.ThrowIfCancellationRequested();
                                    string sqlVincGen = @"
INSERT OR IGNORE INTO ag_vinculacion_gen (idagvinculacion, idgen, subunidad)
VALUES (@IdAgVinc, @IdGen, @Sub);";
                                    await db.ExecuteAsync(new CommandDefinition(
                                        sqlVincGen,
                                        new { IdAgVinc = dbId, IdGen = gId, Sub = v.Subunidad ?? string.Empty },
                                        trans,
                                        cancellationToken: ct));
                                }
                            }
                        }

                        // Fase 992: fechasesion es NOT NULL y obligatoria; no enmascarar con DateTime.MinValue
                        var sesionSinFecha = sesiones.FirstOrDefault(x =>
                            !x.FechaSesion.HasValue || x.FechaSesion.Value.Year < 1900);
                        if (sesionSinFecha != null)
                        {
                            // Fase 998: restaurar ids en memoria antes de rollback para evitar ids fantasma
                            foreach (var v in vinculaciones)
                            {
                                var par = mapVinc.FirstOrDefault(kv => kv.Value == v.Id);
                                if (par.Key != 0) v.Id = par.Key;
                            }

                            trans.Rollback();
                            return Result.Failure(
                                $"Sesión sin fecha válida (idgen={sesionSinFecha.IdGen}, vinculación={sesionSinFecha.IdVinculacion}); materialización abortada para no escribir fechas 0001-01-01.");
                        }

                        foreach (var s in sesiones)
                        {
                            ct.ThrowIfCancellationRequested();
                            int mappedVincId = mapVinc.TryGetValue(s.IdVinculacion, out int r) ? r : s.IdVinculacion;
                            string sqlS = @"
INSERT INTO ag_sesiones (idvinculacion, idtercero, id_aula, fechasesion, horariosesion, horas, viatica, subunidad, idindividuo, idgen, es_vacante, motivo_vacante, candidatos_pool)
VALUES (@IdVinculacion, @IdTercero, @IdAula, @FechaSesion, @HorarioSesion, @Horas, @Viatica, @Subunidad, @IdIndividuo, @IdGen, @EsVacante, @MotivoVacante, @CandidatosPool);";

                            var param = new
                            {
                                IdVinculacion = mappedVincId,
                                IdTercero = s.IdTercero,
                                IdAula = s.IdAula,
                                FechaSesion = s.FechaSesion.Value.ToString("yyyy-MM-dd"),
                                HorarioSesion = s.HorarioSesion ?? string.Empty,
                                Horas = s.Horas,
                                Viatica = s.Viatica ? 1 : 0,
                                Subunidad = s.Subunidad ?? string.Empty,
                                IdIndividuo = s.IdIndividuo > 0 ? s.IdIndividuo : idIndividuo,
                                IdGen = s.IdGen,
                                s.EsVacante,
                                MotivoVacante = s.MotivoVacante ?? string.Empty,
                                s.CandidatosPool
                            };

                            await db.ExecuteAsync(new CommandDefinition(sqlS, param, trans, cancellationToken: ct));
                        }

                        trans.Commit();
                        NotifyMarkDirty();
                        return Result.Success();
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgScheduleRepository] Error guardando staging", ex);
                return Result.Failure($"Error al guardar staging: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<AgVinculacion>>> GetVinculacionesStagingAsync(long idCorrida, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sql = "SELECT id as Id, id_corrida as IdCorrida, id_individuo as IdIndividuo, idtercero as IdTercero, idprograma as IdPrograma, idsede as IdSede, idasignatura as IdAsignatura, idperiodo as IdPeriodo, semestre as Semestre, grupo as Grupo, dia_horas_semanales as DiaHorasSemanales, modalidadVinculacion as ModalidadVinculacion, total_horas as TotalHoras, fecha_inicio as FechaInicio, fecha_fin as FechaFin, estado as Estado, isTerminalAereo as IsTerminalAereo, idTerminales as IdTerminales, gastoTerminalAerea as GastoTerminalAerea, es_vacante as EsVacante, motivo_vacante as MotivoVacante, candidatos_pool as CandidatosPool, subunidad as Subunidad FROM ag_vinculaciones WHERE id_corrida = @idCorrida";
                    var items = (await db.QueryAsync<AgVinculacion>(new CommandDefinition(sql, new { idCorrida }, cancellationToken: ct))).ToList();
                    return Result.Success<IReadOnlyCollection<AgVinculacion>>(items.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgVinculacion>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                return Result.Failure<IReadOnlyCollection<AgVinculacion>>($"Error al leer vinculaciones staging: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<AgSesion>>> GetSesionesStagingAsync(long idCorrida, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    string sql = @"
SELECT 
    s.id as Id,
    s.idvinculacion as IdVinculacion,
    s.idtercero as IdTercero,
    s.id_aula as IdAula,
    s.fechasesion as FechaSesionRaw,
    s.horariosesion as HorarioSesion,
    s.horas as Horas,
    s.viatica as ViaticaInt,
    s.subunidad as Subunidad,
    s.idindividuo as IdIndividuo,
    s.idgen as IdGen,
    s.es_vacante as EsVacante,
    s.motivo_vacante as MotivoVacante,
    s.candidatos_pool as CandidatosPool
FROM ag_sesiones s
INNER JOIN ag_vinculaciones v ON s.idvinculacion = v.id
WHERE v.id_corrida = @idCorrida";

                    var rawList = (await db.QueryAsync(new CommandDefinition(sql, new { idCorrida }, cancellationToken: ct))).ToList();
                    var result = new List<AgSesion>();

                    foreach (var r in rawList)
                    {
                        DateTime? f = null;
                        string rawF = r.FechaSesionRaw?.ToString();
                        if (DateTime.TryParse(rawF, out DateTime dtParsed)) f = dtParsed;

                        result.Add(new AgSesion
                        {
                            Id = (int)r.Id,
                            IdVinculacion = (int)r.IdVinculacion,
                            IdTercero = (int)r.IdTercero,
                            IdAula = (int)r.IdAula,
                            FechaSesion = f,
                            HorarioSesion = (string)r.HorarioSesion,
                            Horas = (double)r.Horas,
                            Viatica = ((int)r.ViaticaInt) == 1,
                            Subunidad = (string)r.Subunidad,
                            IdIndividuo = r.IdIndividuo != null ? (long)r.IdIndividuo : 0L,
                            IdGen = r.IdGen != null ? (long)r.IdGen : 0L,
                            EsVacante = r.EsVacante != null ? Convert.ToInt32(r.EsVacante) : 0,
                            MotivoVacante = r.MotivoVacante?.ToString() ?? string.Empty,
                            CandidatosPool = r.CandidatosPool != null ? Convert.ToInt32(r.CandidatosPool) : 0
                        });
                    }

                    return Result.Success<IReadOnlyCollection<AgSesion>>(result.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgSesion>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                return Result.Failure<IReadOnlyCollection<AgSesion>>($"Error al leer sesiones staging: {ex.Message}");
            }
        }

        public async Task<Result> LimpiarStagingAsync(long idCorrida, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    using (var trans = db.BeginTransaction())
                    {
                        string sqlDelGen = @"
DELETE FROM ag_vinculacion_gen 
WHERE idagvinculacion IN (SELECT id FROM ag_vinculaciones WHERE id_corrida = @idCorrida);";
                        await db.ExecuteAsync(new CommandDefinition(sqlDelGen, new { idCorrida }, trans, cancellationToken: ct));

                        string sqlDelSes = @"
DELETE FROM ag_sesiones 
WHERE idvinculacion IN (SELECT id FROM ag_vinculaciones WHERE id_corrida = @idCorrida);";
                        await db.ExecuteAsync(new CommandDefinition(sqlDelSes, new { idCorrida }, trans, cancellationToken: ct));

                        string sqlDelVinc = "DELETE FROM ag_vinculaciones WHERE id_corrida = @idCorrida;";
                        await db.ExecuteAsync(new CommandDefinition(sqlDelVinc, new { idCorrida }, trans, cancellationToken: ct));

                        string sqlDelSalv = "DELETE FROM ag_salvedades_conflictos WHERE id_corrida = @idCorrida;";
                        await db.ExecuteAsync(new CommandDefinition(sqlDelSalv, new { idCorrida }, trans, cancellationToken: ct));

                        trans.Commit();
                        return Result.Success();
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al limpiar staging: {ex.Message}");
            }
        }

        public async Task<Result> LimpiarStagingAsync(int idPeriodo, int idPrograma, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    using (var trans = db.BeginTransaction())
                    {
                        // Fase 510: Logging de auditoría de purga
                        int v0 = await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM ag_vinculaciones;", transaction: trans);
                        int s0 = await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM ag_sesiones;", transaction: trans);
                        bool hayGen = TablaExiste(db, trans, "ag_gen");
                        bool hayIndividuo = TablaExiste(db, trans, "ag_individuo");
                        int g0 = hayGen ? await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM ag_gen;", transaction: trans) : 0;
                        int c0 = await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM ag_corrida;", transaction: trans);

                        // Fase 1563: la purga se ancla a los identificadores de corrida, no al par
                        // (idperiodo, idprograma) denormalizado en cada tabla.
                        // Motivo: AgFabricaIndividuos estampa ag_gen/ag_vinculaciones con el programa de la
                        // ASIGNATURA, que puede diferir del programa elegido en el formulario (caso real:
                        // programas APT-53 id=1 y APT-52 id=2 con el mismo nombre). Purgar la cabecera por un
                        // par y el staging por otro dejaba corridas huérfanas: sin ag_corrida ni ag_individuo
                        // (por tanto sin violaciones y sin promoción) pero con 180 sesiones visibles en el Módulo R.
                        var idsCabecera = (await db.QueryAsync<long>(new CommandDefinition(
                            "SELECT id FROM ag_corrida WHERE idperiodo = @idPeriodo AND idprograma = @idPrograma;",
                            new { idPeriodo, idPrograma }, trans, cancellationToken: ct))).ToList();

                        var idsStaging = (await db.QueryAsync<long>(new CommandDefinition(
                            "SELECT DISTINCT id_corrida FROM ag_vinculaciones WHERE idperiodo = @idPeriodo AND idprograma = @idPrograma;",
                            new { idPeriodo, idPrograma }, trans, cancellationToken: ct))).ToList();

                        var idsDel = idsCabecera
                            .Concat(idsStaging)
                            .Where(id => id > 0)
                            .Distinct()
                            .ToList();

                        if (idsDel.Count > 0)
                        {
                            var paramIds = new { Ids = idsDel };

                            // Fase 1563: individuos referenciados por el staging purgado. Se capturan antes de borrar
                            // ag_vinculaciones para poder limpiar también sus genes/violaciones huérfanos.
                            var idsIndStaging = (await db.QueryAsync<long>(new CommandDefinition(
                                "SELECT DISTINCT id_individuo FROM ag_vinculaciones WHERE id_corrida IN @Ids AND id_individuo > 0;",
                                paramIds, trans, cancellationToken: ct))).ToList();


                            if (TablaExiste(db, trans, "ag_vinculacion_gen"))
                            {
                                await db.ExecuteAsync(new CommandDefinition(
                                    "DELETE FROM ag_vinculacion_gen WHERE idagvinculacion IN (SELECT id FROM ag_vinculaciones WHERE id_corrida IN @Ids);",
                                    paramIds, trans, cancellationToken: ct));
                            }

                            await db.ExecuteAsync(new CommandDefinition(
                                "DELETE FROM ag_sesiones WHERE idvinculacion IN (SELECT id FROM ag_vinculaciones WHERE id_corrida IN @Ids);",
                                paramIds, trans, cancellationToken: ct));

                            if (TablaExiste(db, trans, "ag_salvedades_conflictos"))
                            {
                                await db.ExecuteAsync(new CommandDefinition(
                                    "DELETE FROM ag_salvedades_conflictos WHERE id_corrida IN @Ids;",
                                    paramIds, trans, cancellationToken: ct));
                            }

                            await db.ExecuteAsync(new CommandDefinition(
                                "DELETE FROM ag_vinculaciones WHERE id_corrida IN @Ids;",
                                paramIds, trans, cancellationToken: ct));

                            // Fase 504: metadatos genéticos antes de individuos y corridas
                            if (hayIndividuo && TablaExiste(db, trans, "ag_violacion"))
                            {
                                await db.ExecuteAsync(new CommandDefinition(
                                    "DELETE FROM ag_violacion WHERE id_individuo IN (SELECT id FROM ag_individuo WHERE id_corrida IN @Ids);",
                                    paramIds, trans, cancellationToken: ct));
                            }

                            if (TablaExiste(db, trans, "ag_bitacora"))
                            {
                                await db.ExecuteAsync(new CommandDefinition(
                                    "DELETE FROM ag_bitacora WHERE id_corrida IN @Ids;",
                                    paramIds, trans, cancellationToken: ct));
                            }

                            if (TablaExiste(db, trans, "ag_promocion"))
                            {
                                await db.ExecuteAsync(new CommandDefinition(
                                    "DELETE FROM ag_promocion WHERE id_corrida IN @Ids;",
                                    paramIds, trans, cancellationToken: ct));
                            }

                            if (hayGen && hayIndividuo)
                            {
                                await db.ExecuteAsync(new CommandDefinition(
                                    "DELETE FROM ag_gen WHERE id_individuo IN (SELECT id FROM ag_individuo WHERE id_corrida IN @Ids);",
                                    paramIds, trans, cancellationToken: ct));
                            }
                            // Fase 1563: limpieza de genes/violaciones de individuos referenciados por el staging
                            // purgado aunque su fila en ag_individuo ya no exista (huérfanos de purgas anteriores).
                            if (idsIndStaging.Count > 0)
                            {
                                var paramInd = new { IdsInd = idsIndStaging };

                                if (hayGen)
                                {
                                    await db.ExecuteAsync(new CommandDefinition(
                                        "DELETE FROM ag_gen WHERE id_individuo IN @IdsInd;",
                                        paramInd, trans, cancellationToken: ct));
                                }

                                if (TablaExiste(db, trans, "ag_violacion"))
                                {
                                    await db.ExecuteAsync(new CommandDefinition(
                                        "DELETE FROM ag_violacion WHERE id_individuo IN @IdsInd;",
                                        paramInd, trans, cancellationToken: ct));
                                }
                            }



                            if (hayIndividuo)
                            {
                                await db.ExecuteAsync(new CommandDefinition(
                                    "DELETE FROM ag_individuo WHERE id_corrida IN @Ids;",
                                    paramIds, trans, cancellationToken: ct));
                            }

                            await db.ExecuteAsync(new CommandDefinition(
                                "DELETE FROM ag_corrida WHERE id IN @Ids;",
                                paramIds, trans, cancellationToken: ct));
                        }

                        int v1 = await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM ag_vinculaciones;", transaction: trans);
                        int s1 = await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM ag_sesiones;", transaction: trans);
                        int g1 = hayGen ? await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM ag_gen;", transaction: trans) : 0;
                        int c1 = await db.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM ag_corrida;", transaction: trans);

                        LoggerHelper.Info($"[Purga] vinc {v0}->{v1} | ses {s0}->{s1} | gen {g0}->{g1} | corridas {c0}->{c1} | idPeriodo={idPeriodo} idPrograma={idPrograma} corridasPeriodo={string.Join(",", idsDel)}");

                        // Fase 1563: invariante de integridad. Si tras la purga quedan hijos de las corridas
                        // objetivo, la interfaz volvería a mostrar staging sin cabecera: se registra como error.
                        if (idsDel.Count > 0)
                        {
                            var paramVerif = new { Ids = idsDel };
                            int vincResiduales = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                                "SELECT COUNT(*) FROM ag_vinculaciones WHERE id_corrida IN @Ids;",
                                paramVerif, trans, cancellationToken: ct));
                            int sesResiduales = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                                "SELECT COUNT(*) FROM ag_sesiones WHERE idvinculacion IN (SELECT id FROM ag_vinculaciones WHERE id_corrida IN @Ids);",
                                paramVerif, trans, cancellationToken: ct));

                            if (vincResiduales > 0 || sesResiduales > 0)
                            {
                                LoggerHelper.Error(
                                    $"[Purga] Huérfanos residuales tras la purga: {vincResiduales} vinculación(es) y " +
                                    $"{sesResiduales} sesión(es) de las corridas [{string.Join(",", idsDel)}] " +
                                    $"(idPeriodo={idPeriodo} idPrograma={idPrograma}). Revise el staging en el Módulo R.");
                            }
                        }

                        trans.Commit();
                        return Result.Success();
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al limpiar staging por período/programa: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<AgGen>>> GetGenesByIndividuoAsync(long idIndividuo, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    // Fase 1566: lectura tolerante para corridas huérfanas. Los genes sobreviven en ag_gen aunque
                    // la fila de ag_individuo haya sido eliminada (purga antigua por período/programa).
                    string sqlGenes = "SELECT id as Id, id_individuo as IdIndividuo, idperiodo as IdPeriodo, idasignatura as IdAsignatura, idprograma as IdPrograma, idsede as IdSede, semestre as Semestre, grupo as Grupo, idtercero as IdTercero, id_aula as IdAula, id_bloque_horario as IdBloqueHorario, dia_semana as DiaSemana, horas_semanales as HorasSemanales, total_horas_asignatura as TotalHorasAsignatura, NULLIF(TRIM(fecha_inicio_estimada), '') as FechaInicioEstimada, NULLIF(TRIM(fecha_fin_estimada), '') as FechaFinEstimada, es_fijo as EsFijo, es_vacante as EsVacante, motivo_vacante as MotivoVacante, dentro_de_ventana as DentroDeVentana, estado as Estado, subunidad as Subunidad, indice_dia_patron as IndiceDiaPatron, patron_semanal as PatronSemanal, dias_patron as DiasPatron, clave_grupo as ClaveGrupo, horas_gen as HorasGen, fechas_gen AS FechasGenStr FROM ag_gen WHERE id_individuo = @idIndividuo";
                    var genes = (await db.QueryAsync<AgGen>(
                        new CommandDefinition(sqlGenes, new { idIndividuo }, cancellationToken: ct))).ToList();

                    return Result.Success<IReadOnlyCollection<AgGen>>(genes.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgGen>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error obteniendo genes del individuo {idIndividuo}", ex);
                return Result.Failure<IReadOnlyCollection<AgGen>>($"Error al obtener genes: {ex.Message}");
            }
        }

        public async Task<Result<AgCorridaHuerfana>> DetectarCorridaHuerfanaAsync(long idCorrida, CancellationToken ct)
        {
            if (idCorrida <= 0)
                return Result.Failure<AgCorridaHuerfana>("IdCorrida inválido.");

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    if (!TablaExiste(db, null, "ag_corrida") || !TablaExiste(db, null, "ag_vinculaciones"))
                        return Result.Failure<AgCorridaHuerfana>(
                            "El vault no tiene las tablas del AG (ag_corrida / ag_vinculaciones).");

                    // Fase 1566: el staging es la fuente de verdad para período, programa y subunidad cuando
                    // la cabecera desapareció (caso real corrida #48: staging programa APT-52, cabecera APT-53).
                    var datos = await db.QuerySingleAsync<AgCorridaHuerfana>(new CommandDefinition(@"
SELECT COUNT(*) AS TotalVinculaciones,
       (SELECT COUNT(*) FROM ag_sesiones s
         WHERE s.idvinculacion IN (SELECT id FROM ag_vinculaciones WHERE id_corrida = @idCorrida)) AS TotalSesiones,
       COALESCE(MAX(idperiodo), 0) AS IdPeriodo,
       COALESCE(MAX(idprograma), 0) AS IdPrograma,
       COALESCE(MAX(NULLIF(TRIM(subunidad), '')), '') AS Subunidad,
       COALESCE(MAX(id_individuo), 0) AS IdIndividuo
FROM ag_vinculaciones
WHERE id_corrida = @idCorrida;", new { idCorrida }, cancellationToken: ct));

                    datos.IdCorrida = idCorrida;
                    datos.ExisteCabecera = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                        "SELECT COUNT(*) FROM ag_corrida WHERE id = @idCorrida;",
                        new { idCorrida }, cancellationToken: ct)) > 0;

                    if (datos.IdIndividuo > 0 && TablaExiste(db, null, "ag_individuo"))
                    {
                        datos.ExisteIndividuo = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                            "SELECT COUNT(*) FROM ag_individuo WHERE id = @idIndividuo;",
                            new { idIndividuo = datos.IdIndividuo }, cancellationToken: ct)) > 0;
                    }

                    LoggerHelper.Info($"[AG-Reparacion] Diagnóstico de la corrida #{idCorrida}: {datos}");
                    return Result.Success(datos);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<AgCorridaHuerfana>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error diagnosticando la corrida huérfana #{idCorrida}", ex);
                return Result.Failure<AgCorridaHuerfana>($"Error al diagnosticar la corrida: {ex.Message}");
            }
        }


        public async Task<Result> ReinsertarCabeceraHuerfanaAsync(
            AgCorrida corrida, long idCorrida, AgIndividuo individuo, CancellationToken ct)
        {
            if (corrida == null) return Result.Failure("Corrida nula.");
            if (individuo == null) return Result.Failure("Individuo nulo.");
            if (idCorrida <= 0) return Result.Failure("IdCorrida inválido.");
            if (individuo.Id <= 0) return Result.Failure("IdIndividuo inválido.");

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    using (var trans = db.BeginTransaction())
                    {
                        int existeCorrida = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                            "SELECT COUNT(*) FROM ag_corrida WHERE id = @idCorrida;",
                            new { idCorrida }, trans, cancellationToken: ct));
                        if (existeCorrida > 0)
                            return Result.Failure(
                                $"La corrida #{idCorrida} ya tiene cabecera en ag_corrida; no requiere reparación.");

                        int existeIndividuo = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                            "SELECT COUNT(*) FROM ag_individuo WHERE id = @idIndividuo;",
                            new { idIndividuo = individuo.Id }, trans, cancellationToken: ct));
                        if (existeIndividuo > 0)
                            return Result.Failure(
                                $"El individuo #{individuo.Id} ya existe en ag_individuo; no requiere reparación.");

                        // Fase 1566: inserción con id explícito para respetar las referencias ya existentes en
                        // el staging (ag_vinculaciones.id_individuo, ag_sesiones.idindividuo, ag_gen.id_individuo).
                        const string sqlIndividuo = @"
INSERT INTO ag_individuo (id, id_corrida, generacion, fitness, hard_violations, soft_violations, estado, subunidad)
VALUES (@Id, @IdCorrida, @Generacion, @Fitness, @HardViolations, @SoftViolations, @Estado, @Subunidad);";
                        await db.ExecuteAsync(new CommandDefinition(sqlIndividuo, new
                        {
                            Id = individuo.Id,
                            IdCorrida = idCorrida,
                            individuo.Generacion,
                            individuo.Fitness,
                            individuo.HardViolations,
                            individuo.SoftViolations,
                            Estado = (int)individuo.Estado,
                            Subunidad = individuo.Subunidad ?? corrida.Subunidad ?? string.Empty
                        }, trans, cancellationToken: ct));

                        const string sqlCorrida = @"
INSERT INTO ag_corrida (id, idperiodo, idprograma, idaula, fecha_inicio_corrida, fecha_fin_corrida, estado, poblacion_tamano, generaciones_max, generaciones_ejecutadas, tasa_cruce, tasa_mutacion, elitismo, mejor_fitness, hard_violations, soft_violations, vacantes, id_individuo_optimo, parametros_json, subunidad)
VALUES (@Id, @IdPeriodo, @IdPrograma, @IdAula, @FechaInicio, @FechaFin, @Estado, @Poblacion, @Generaciones, @GeneracionesEjecutadas, @TasaCruce, @TasaMutacion, @Elitismo, @Fitness, @Hard, @Soft, @Vacantes, @IdIndividuo, @ParametrosJson, @Subunidad);";
                        await db.ExecuteAsync(new CommandDefinition(sqlCorrida, new
                        {
                            Id = idCorrida,
                            corrida.IdPeriodo,
                            corrida.IdPrograma,
                            corrida.IdAula,
                            FechaInicio = corrida.FechaInicioCorrida ?? string.Empty,
                            FechaFin = corrida.FechaFinCorrida ?? string.Empty,
                            Estado = (int)corrida.Estado,
                            Poblacion = corrida.PoblacionTamano,
                            Generaciones = corrida.GeneracionesMax,
                            corrida.GeneracionesEjecutadas,
                            corrida.TasaCruce,
                            corrida.TasaMutacion,
                            corrida.Elitismo,
                            Fitness = corrida.MejorFitness,
                            Hard = corrida.HardViolations,
                            Soft = corrida.SoftViolations,
                            corrida.Vacantes,
                            IdIndividuo = individuo.Id,
                            ParametrosJson = corrida.ParametrosJson ?? "{}",
                            Subunidad = corrida.Subunidad ?? string.Empty
                        }, trans, cancellationToken: ct));

                        trans.Commit();
                        NotifyMarkDirty();

                        LoggerHelper.Info(
                            $"[AG-Reparacion] Cabecera reconstruida: corrida #{idCorrida} " +
                            $"(período {corrida.IdPeriodo}, programa {corrida.IdPrograma}, subunidad '{corrida.Subunidad}') " +
                            $"e individuo #{individuo.Id}. Staging intacto.");

                        return Result.Success();
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgScheduleRepository] Error reinsertando la cabecera de la corrida #{idCorrida}", ex);
                return Result.Failure($"Error al reconstruir la cabecera: {ex.Message}");
            }
        }

        /// <summary>
        /// Indica si una tabla existe en el vault. Permite purgar vaults legados que aún no tienen
        /// las tablas opcionales del AG (Fase 1563/1890/1892).
        /// </summary>
        private static bool TablaExiste(IDbConnection db, IDbTransaction trans, string tabla)
        {
            return SqliteHelper.TablaExiste(db, tabla, trans) == ExistenciaSql.Existe;
        }

        /// <summary>
        /// Comprueba de forma segura si una columna existe confirmadamente en una tabla SQLite (Fase 1885/1888/1889/1893).
        /// Solo retorna true si la columna fue fehacientemente comprobada (Existe). Ante Indeterminada o NoExiste
        /// retorna false para no abortar Insert/Update con 'no such column'. El consentimiento de vacantes
        /// se conserva íntegro en parametros_json como mecanismo de respaldo.
        /// </summary>
        private static bool ColumnaExiste(IDbConnection db, IDbTransaction trans, string tabla, string columna)
        {
            return SqliteHelper.ColumnaExiste(db, tabla, columna, trans) == ExistenciaSql.Existe;
        }


        public async Task<Result> ReemplazarViolacionesAsync(long idIndividuo, IReadOnlyCollection<AgViolacion> violaciones, int hardViolations, double softViolations, AgEstadoIndividuo estado, long idCorrida, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();

                    using (var trans = db.BeginTransaction())
                    {
                        await db.ExecuteAsync(new CommandDefinition(
                            "DELETE FROM ag_violacion WHERE id_individuo = @IdIndividuo;",
                            new { IdIndividuo = idIndividuo }, trans, cancellationToken: ct));

                        if (violaciones != null && violaciones.Count > 0)
                        {
                            // Fase 1815: Resolver indicesGenes hacia idGenes/id_gen compartiendo la lógica con SaveIndividuoAsync
                            var geneIds = (await db.QueryAsync<long>(new CommandDefinition(
                                "SELECT id FROM ag_gen WHERE id_individuo = @IdIndividuo ORDER BY id ASC;",
                                new { IdIndividuo = idIndividuo }, trans, cancellationToken: ct))).ToList();

                            if (geneIds.Count > 0)
                            {
                                ResolverIndicesGenes(violaciones, geneIds);
                            }

                            const string sqlViol = @"
INSERT INTO ag_violacion (id_individuo, id_gen, regla, es_dura, mensaje, detalle_json, subunidad)
VALUES (@IdIndividuo, @IdGen, @Regla, @EsDura, @Mensaje, @DetalleJson, @Subunidad);";

                            foreach (var v in violaciones)
                            {
                                ct.ThrowIfCancellationRequested();
                                var violParam = new
                                {
                                    IdIndividuo = idIndividuo,
                                    v.IdGen,
                                    Regla = (int)v.Regla, // Fase 1914 / Fase 1925: Persistir enum como entero, no como string
                                    v.EsDura,
                                    Mensaje = v.Mensaje ?? string.Empty,
                                    DetalleJson = v.DetalleJson ?? "{}",
                                    Subunidad = v.Subunidad ?? string.Empty
                                };
                                await db.ExecuteAsync(new CommandDefinition(sqlViol, violParam, trans, cancellationToken: ct));
                            }
                        }

                        await db.ExecuteAsync(new CommandDefinition(
                            "UPDATE ag_individuo SET hard_violations = @Hard, soft_violations = @Soft, estado = @Estado WHERE id = @IdIndividuo;",
                            new { Hard = hardViolations, Soft = softViolations, Estado = (int)estado, IdIndividuo = idIndividuo },
                            trans, cancellationToken: ct));

                        if (idCorrida > 0)
                        {
                            await db.ExecuteAsync(new CommandDefinition(
                                "UPDATE ag_corrida SET hard_violations = @Hard, soft_violations = @Soft WHERE id = @IdCorrida;",
                                new { Hard = hardViolations, Soft = softViolations, IdCorrida = idCorrida },
                                trans, cancellationToken: ct));
                        }

                        trans.Commit();
                        NotifyMarkDirty();
                        LoggerHelper.Info($"[AG-Recalculo] individuo={idIndividuo} corrida={idCorrida} hard={hardViolations} soft={softViolations} violaciones={(violaciones == null ? 0 : violaciones.Count)}");
                        return Result.Success();
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al reemplazar violaciones: {ex.Message}");
            }
        }

        /// <summary>
        /// Resuelve índices posicionales de genes ('indicesGenes') contenidos en el DetalleJson
        /// hacia sus identificadores persistidos ('idGenes' e 'IdGen') para garantizar enlaces referenciales íntegros.
        /// Compartido entre SaveIndividuoAsync y ReemplazarViolacionesAsync dentro de sus transacciones.
        /// </summary>
        private static void ResolverIndicesGenes(IEnumerable<AgViolacion> violaciones, IReadOnlyList<long> geneIds)
        {
            if (violaciones == null || geneIds == null || geneIds.Count == 0) return;
            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            foreach (var violacion in violaciones)
            {
                if (string.IsNullOrWhiteSpace(violacion.DetalleJson)) continue;
                Dictionary<string, object> detalle = null;
                try
                {
                    detalle = serializer.DeserializeObject(violacion.DetalleJson) as Dictionary<string, object>;
                }
                catch
                {
                    detalle = null;
                }
                if (detalle == null || !detalle.ContainsKey("indicesGenes")) continue;
                var indices = detalle["indicesGenes"] as object[]
                    ?? (detalle["indicesGenes"] as System.Collections.IEnumerable)?.Cast<object>().ToArray();
                if (indices == null || indices.Length == 0) continue;
                var ids = indices.Select(x => Convert.ToInt32(x)).Select(indice =>
                {
                    if (indice >= 0 && indice < geneIds.Count && geneIds[indice] > 0)
                        return geneIds[indice];
                    return 0L;
                }).Where(id => id > 0).Distinct().ToArray();
                if (ids.Length > 0)
                {
                    detalle["idGenes"] = ids;
                    violacion.IdGen = ids[0];
                    violacion.DetalleJson = serializer.Serialize(detalle);
                }
            }
        }
    }
}
