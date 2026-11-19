using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Servicio de promoción no destructiva y atómica desde el staging del AG hacia las tablas de producción.
    /// Ejecuta remapeo de IDs, migración de salvedades y sincronización del pool docente.
    /// </summary>
    public class AgPromotorVinculaciones : IAgPromotorVinculaciones
    {
        private readonly Func<IDbConnection> _connectionFactory;
        private readonly IAgScheduleRepository _scheduleRepository;
        private readonly IAgSalvedadRepository _salvedadRepository;
        private readonly ITerceroAsignaturaSyncService _syncService;

        public AgPromotorVinculaciones(
            Func<IDbConnection> connectionFactory,
            IAgScheduleRepository scheduleRepository,
            IAgSalvedadRepository salvedadRepository,
            ITerceroAsignaturaSyncService syncService)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
            _scheduleRepository = scheduleRepository ?? throw new ArgumentNullException(nameof(scheduleRepository));
            _salvedadRepository = salvedadRepository ?? throw new ArgumentNullException(nameof(salvedadRepository));
            _syncService = syncService ?? throw new ArgumentNullException(nameof(syncService));
        }

        public Task<Result<AgPromocion>> PromoverAProduccionAsync(long idCorrida, long idIndividuo, string usuario, CancellationToken ct)
        {
            return PromoverAProduccionAsync(idCorrida, idIndividuo, usuario, null, null, ct);
        }

        public Task<Result<AgPromocion>> PromoverAProduccionAsync(long idCorrida, long idIndividuo, string usuario, string justificacionVacante, CancellationToken ct)
        {
            return PromoverAProduccionAsync(idCorrida, idIndividuo, usuario, justificacionVacante, null, ct);
        }

        public async Task<Result<AgPromocion>> PromoverAProduccionAsync(long idCorrida, long idIndividuo, string usuario, string justificacionVacante, AgContexto contexto, CancellationToken ct)
        {
            try
            {
                // Gate 1: Validación de estado de corrida y aula (Fase 85)
                var corridaRes = await _scheduleRepository.GetCorridaByIdAsync(idCorrida, ct);
                if (corridaRes.IsFailure)
                {
                    return Result.Failure<AgPromocion>($"No se encontró la corrida {idCorrida}: {corridaRes.Error}");
                }

                if (corridaRes.Value.Estado != AgEstadoCorrida.Aprobada)
                {
                    return Result.Failure<AgPromocion>($"La corrida {idCorrida} no está APROBADA (estado: {corridaRes.Value.Estado}).");
                }

                // Fase 1520 (G2): las corridas sin aula asociada (id_aula = 0) no deben bloquear la promoción
                if (corridaRes.Value.IdAula > 0)
                {
                    var aulaRes = await _scheduleRepository.GetAulaByIdAsync(corridaRes.Value.IdAula, ct);
                    if (aulaRes.IsFailure || (aulaRes.Value.Estado != AgEstadoAula.Disponible && aulaRes.Value.Estado != AgEstadoAula.Aprobado))
                    {
                        return Result.Failure<AgPromocion>("El Aula/Salón no está en estado APROBADO/DISPONIBLE.");
                    }
                }
                else
                {
                    LoggerHelper.Warning($"[AgPromotorVinculaciones] Corrida {idCorrida} sin aula asociada (id_aula=0): se omite el gate de aula.");
                }

                var indResult = await _scheduleRepository.GetIndividuoByIdAsync(idIndividuo, ct);
                if (indResult.IsFailure)
                {
                    return Result.Failure<AgPromocion>($"No se encontró el individuo {idIndividuo}: {indResult.Error}");
                }

                var individuo = indResult.Value;
                // Compuerta estricta de seguridad: Cero violaciones duras y cero vacantes no asignadas
                if (individuo.HardViolations > 0)
                {
                    return Result.Failure<AgPromocion>($"No se puede promover una solución con {individuo.HardViolations} restricciones duras activas.");
                }

                int nVac = individuo.Genes?.Count(g => g.EsVacante == 1 || g.IdTercero <= 0) ?? 0;
                if (nVac > 0 && string.IsNullOrWhiteSpace(justificacionVacante))
                {
                    return Result.Failure<AgPromocion>(
                        $"La solución tiene {nVac} asignatura(s) en estado VACANTE (con vacantes pendientes). " +
                        "Resuélvalas en el Panel de Vacantes o proporcione una justificación para promover parcialmente.");
                }

                var stagingVincsResult = await _scheduleRepository.GetVinculacionesStagingAsync(idCorrida, ct);
                if (stagingVincsResult.IsFailure || stagingVincsResult.Value.Count == 0)
                {
                    return Result.Failure<AgPromocion>("No se encontraron vinculaciones en staging para esta corrida.");
                }

                var stagingSesResult = await _scheduleRepository.GetSesionesStagingAsync(idCorrida, ct);
                if (stagingSesResult.IsFailure)
                {
                    return Result.Failure<AgPromocion>($"Error al leer sesiones de staging: {stagingSesResult.Error}");
                }

                // Corrección [P1]: Verificar siempre identidad e integridad del snapshot sellado y del plan fechado,
                // independientemente de si el llamador suministró o no un contexto en memoria.
                var ctxRes = await RecuperarContextoSelladoAsync(corridaRes.Value, individuo, ct);
                if (ctxRes.IsFailure)
                {
                    return Result.Failure<AgPromocion>(
                        $"Fallo en la verificación del plan/contexto sellado de la corrida #{idCorrida}: {ctxRes.Error}");
                }
                contexto = ctxRes.Value;

                // Corrección [P1]: Exigir fechas autorizadas en todos los genes con carga horaria
                foreach (var g in individuo.Genes)
                {
                    if (g.HorasGen > 0.001 && (g.FechasGen == null || g.FechasGen.Count == 0))
                    {
                        return Result.Failure<AgPromocion>(
                            $"El gen #{g.Id} (Asignatura {g.IdAsignatura}) carece de fechas autorizadas (FechasGen) en el plan sellado. No se permite la promoción sin plan fechado.");
                    }
                }

                // Compuerta centralizada estricta de promoción y cobertura multiconjunto (Fase 2029, 2032 & 2045)
                bool permitirVacantes = !string.IsNullOrWhiteSpace(justificacionVacante);
                var (autorizadasAsigs, autorizadasClaves) = ParseAutorizacionVacantes(justificacionVacante);
                var gateRes = AgCompuertaPromocion.ValidarPromocion(
                    corridaRes.Value, individuo, stagingSesResult, contexto, permitirVacantes, autorizadasAsigs, autorizadasClaves);
                if (gateRes.IsFailure)
                {
                    return Result.Failure<AgPromocion>(gateRes.Error);
                }

                if (permitirVacantes)
                {
                    LoggerHelper.Warning($"[AgPromotorVinculaciones] Promoción parcial con vacantes autorizada bajo justificación: '{justificacionVacante}'.");
                }

                var stagingVincs = stagingVincsResult.Value.ToList();
                var stagingSesiones = stagingSesResult.Value.ToList();

                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    var promocion = new AgPromocion
                    {
                        IdCorrida = idCorrida,
                        IdIndividuo = idIndividuo,
                        IdAula = corridaRes.Value.IdAula,
                        Usuario = usuario ?? Environment.UserName,
                        TotalVinculaciones = stagingVincs.Count,
                        TotalSesiones = stagingSesiones.Count,
                        Subunidad = individuo.Subunidad
                    };

                    using (var trans = db.BeginTransaction())
                    {
                        try
                        {
                            var repoAcademic = new AcademicRepository(db, trans, subunidad: individuo.Subunidad, suppressMarkDirty: true);
                            var mapaVincId = new Dictionary<int, int>(); // StagingId -> ProdId
                            var mapaSesionId = new Dictionary<int, int>(); // StagingId -> ProdId

                            // T6: revalidar DENTRO de la transacción que el alcance con aulas aprobadas sigue
                            // autorizado y que la corrida no cambió desde la preparación. Cualquier desvío revierte.
                            var revalidacion = await RevalidarAlcanceAprobadoEnTransaccionAsync(
                                db, trans, corridaRes.Value, contexto, ct);
                            if (revalidacion.IsFailure)
                            {
                                trans.Rollback();
                                return Result.Failure<AgPromocion>(revalidacion.Error);
                            }

                            // 1. Insertar o Actualizar Vinculaciones en Producción (Fase 120: Idempotencia)
                            foreach (var sv in stagingVincs)
                            {
                                var prodVinc = ConversorPDF.XLL.Academic.Application.Mappers.AgVinculacionToVinculacionMapper.MapToProduccion(sv);

                                // Verificar si esta vinculación ya fue promovida previamente para actualizar en vez de duplicar
                                int existingProdId = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                                    "SELECT idvinculacion FROM ag_promocion WHERE idagvinculacion = @IdAg LIMIT 1;",
                                    new { IdAg = sv.Id }, trans, cancellationToken: ct));

                                if (existingProdId > 0)
                                {
                                    prodVinc.Id = existingProdId;
                                    var resUpd = repoAcademic.UpdateVinculacion(prodVinc);
                                    if (resUpd.IsFailure)
                                    {
                                        throw new Exception($"Error actualizando vinculación {existingProdId}: {resUpd.Error}");
                                    }
                                    repoAcademic.DeleteSesionesByVinculacion(existingProdId);
                                    mapaVincId[sv.Id] = existingProdId;
                                }
                                else
                                {
                                    var resVinc = repoAcademic.InsertVinculacion(prodVinc);
                                    if (resVinc.IsFailure)
                                    {
                                        throw new Exception($"Error insertando vinculación: {resVinc.Error}");
                                    }
                                    mapaVincId[sv.Id] = resVinc.Value;
                                }
                            }

                            // 2. Insertar Sesiones en Producción con ID de vinculación remapeado y capturar mapa de IDs (Fase 87)
                            foreach (var ss in stagingSesiones)
                            {
                                if (!mapaVincId.TryGetValue(ss.IdVinculacion, out int newVincId))
                                {
                                    throw new Exception($"Vinculación huérfana en sesión {ss.Id}");
                                }

                                var prodSes = new Sesion
                                {
                                    IdVinculacion = newVincId,
                                    IdTercero = ss.IdTercero,
                                    // Fase 993: fecha obligatoria; fallar antes de contaminar producción
                                    FechaSesion = ss.FechaSesion ?? throw new Exception(
                                        $"Sesión de staging {ss.Id} sin fecha válida; no se puede promover (fechasesion NOT NULL)."),
                                    HorarioSesion = ss.HorarioSesion,
                                    Horas = ss.Horas,
                                    Viatica = ss.Viatica,
                                    EsVacante = ss.EsVacante,
                                    MotivoVacante = ss.MotivoVacante ?? string.Empty
                                };

                                var resSes = repoAcademic.InsertSesion(prodSes);
                                if (resSes.IsFailure)
                                {
                                    throw new Exception($"Error insertando sesión {ss.Id}: {resSes.Error}");
                                }

                                mapaSesionId[ss.Id] = resSes.Value;
                            }

                            // 3. Copiar salvedades con remapeo de sesiones y re-orden min/max (Fase 87)
                            var salvedadesAgResult = await _salvedadRepository.GetByCorridaAsync(idCorrida, ct);
                            if (salvedadesAgResult.IsSuccess && salvedadesAgResult.Value != null)
                            {
                                foreach (var sal in salvedadesAgResult.Value)
                                {
                                    if (!mapaSesionId.TryGetValue((int)sal.IdSesionA, out int a)) continue;
                                    if (!mapaSesionId.TryGetValue((int)sal.IdSesionB, out int b)) continue;

                                    string sqlSalvedad = @"
INSERT OR IGNORE INTO salvedades_conflictos (idsesion_a, idsesion_b, idtercero, idperiodo, justificacion, subunidad)
VALUES (@a, @b, @t, @p, @j, @s);";

                                    await db.ExecuteAsync(new CommandDefinition(
                                        sqlSalvedad,
                                        new
                                        {
                                            a = Math.Min(a, b),
                                            b = Math.Max(a, b),
                                            t = sal.IdTercero,
                                            p = sal.IdPeriodo,
                                            j = sal.Justificacion,
                                            s = individuo.Subunidad
                                        },
                                        trans,
                                        cancellationToken: ct));
                                }
                            }

                            // 4. Registrar Promoción en ag_promocion con trazabilidad id<->id (Fase 86)
                            string sqlPromocion = @"
INSERT INTO ag_promocion (id_corrida, id_individuo, idaula, idagvinculacion, idvinculacion, idtercero, idasignatura, fecha_inicio, fecha_fin, modo, justificacion_vacante, fecha_promocion, usuario, total_vinculaciones, total_sesiones, subunidad)
VALUES (@IdCorrida, @IdIndividuo, @IdAula, @IdAgVinculacion, @IdVinculacion, @IdTercero, @IdAsignatura, @FechaInicio, @FechaFin, @Modo, @JustificacionVacante, @FechaPromocion, @Usuario, @TotalVinculaciones, @TotalSesiones, @Subunidad);
SELECT last_insert_rowid();";

                            foreach (var sv in stagingVincs)
                            {
                                int prodVincId = mapaVincId[sv.Id];
                                var promItem = new AgPromocion
                                {
                                    IdCorrida = idCorrida,
                                    IdIndividuo = idIndividuo,
                                    IdAula = corridaRes.Value.IdAula,
                                    IdAgVinculacion = sv.Id,
                                    IdVinculacion = prodVincId,
                                    IdTercero = sv.IdTercero,
                                    IdAsignatura = sv.IdAsignatura,
                                    FechaInicio = sv.FechaInicio,
                                    FechaFin = sv.FechaFin,
                                    Modo = "COPIA_REMAPEO",
                                    JustificacionVacante = sv.EsVacante == 1 ? (justificacionVacante ?? string.Empty) : string.Empty,
                                    FechaPromocion = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                    Usuario = usuario ?? Environment.UserName,
                                    TotalVinculaciones = stagingVincs.Count,
                                    TotalSesiones = stagingSesiones.Count,
                                    Subunidad = individuo.Subunidad
                                };

                                long idProm = await db.ExecuteScalarAsync<long>(new CommandDefinition(sqlPromocion, promItem, trans, cancellationToken: ct));
                                if (promocion.Id == 0) promocion.Id = idProm;
                            }

                            trans.Commit();
                            individuo.Estado = AgEstadoIndividuo.Promocionado;
                            LoggerHelper.Info($"[AgPromotorVinculaciones] Promoción exitosa. IdPromocion={promocion.Id}");
                        }
                        catch (Exception)
                        {
                            trans.Rollback();
                            throw;
                        }
                    }

                    ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();

                    // 4. Sincronizar Pool Docente Post-Promoción (con origen AG_APROBADO)
                    if (stagingVincs.Count > 0)
                    {
                        try
                        {
                            var poolRepo = new TerceroAsignaturaRepository(_connectionFactory);
                            foreach (var sv in stagingVincs)
                            {
                                if (sv.IdTercero > 0 && sv.IdAsignatura > 0)
                                {
                                    int conteo = await poolRepo.ContarAsignacionesAsync(sv.IdTercero, sv.IdPrograma, sv.IdAsignatura, sv.Semestre, ct);
                                    var tasItem = new TerceroAsignatura
                                    {
                                        Id = TerceroAsignaturaKey.Calcular(sv.IdTercero, sv.IdPrograma, sv.IdAsignatura, sv.Semestre),
                                        IdTercero = sv.IdTercero,
                                        IdPrograma = sv.IdPrograma,
                                        IdAsignatura = sv.IdAsignatura,
                                        Semestre = (sv.Semestre ?? string.Empty).Trim().ToUpperInvariant(),
                                        Activo = 1,
                                        Prioridad = 1,
                                        VecesAsignada = conteo,
                                        FechaUltimaAsignacion = DateTime.Now.ToString("yyyy-MM-dd"),
                                        Origen = "AG_APROBADO",
                                        Observaciones = $"Promovido desde Corrida #{idCorrida}",
                                        Subunidad = individuo.Subunidad ?? string.Empty
                                    };
                                    await poolRepo.UpsertAsync(tasItem, ct);
                                }
                            }

                            var primerVinc = stagingVincs[0];
                            await _syncService.SincronizarPorSemestreAsync(primerVinc.IdPrograma, primerVinc.Semestre, primerVinc.Subunidad, ct);
                        }
                        catch (Exception exSync)
                        {
                            LoggerHelper.Warning($"[AgPromotorVinculaciones] Error no fatal resincronizando pool: {exSync.Message}");
                        }
                    }

                    return Result.Success(promocion);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<AgPromocion>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgPromotorVinculaciones] Error en promoción a producción", ex);
                return Result.Failure<AgPromocion>($"Error al promover a producción: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene un valor de cadena desde un diccionario preservando fielmente null frente a cadena vacía.
        /// </summary>
        private static string ObtenerString(Dictionary<string, object> dict, string key, string fallback = null)
        {
            if (dict == null || !dict.TryGetValue(key, out var val))
            {
                return fallback;
            }
            return val?.ToString();
        }

        /// <summary>
        /// Revalida dentro de la transacción de promoción que el alcance con aulas aprobadas conserva los
        /// parámetros sellados, el permiso vigente y la ratificación persistida. Un fallo revierte sin escribir.
        /// No sustituye la compuerta de cobertura: sólo impide reemplazar producción sin autorización vigente.
        /// </summary>
        private static async Task<Result> RevalidarAlcanceAprobadoEnTransaccionAsync(
            IDbConnection db, IDbTransaction trans, AgCorrida corrida, AgContexto contexto,
            CancellationToken ct)
        {
            var manifiesto = contexto?.ManifiestoCobertura;
            if (manifiesto == null) return Result.Success(); // sin manifiesto se conserva el flujo anterior

            var aulasAprobadas = new HashSet<int>(contexto.Aulas
                .Where(a => a.Estado == AgEstadoAula.Aprobado && manifiesto.Relaciones.Any(r => r.Clave.IdAula == a.Id))
                .Select(a => a.Id));
            if (aulasAprobadas.Count == 0) return Result.Success();

            // 1. La corrida no debe haber cambiado sus parámetros sellados desde la preparación.
            string parametrosPersistidos = await db.ExecuteScalarAsync<string>(new CommandDefinition(
                "SELECT parametros_json FROM ag_corrida WHERE id = @Id LIMIT 1;",
                new { Id = corrida.Id }, trans, cancellationToken: ct));
            if (string.IsNullOrWhiteSpace(parametrosPersistidos) || parametrosPersistidos != corrida.ParametrosJson)
                return Result.Failure("PARAMETROS_CORRIDA_CAMBIADOS: la corrida no coincide con la preparación autorizada.");

            // 2. Permiso vigente del actor real, nunca el texto del llamador.
            string actor = App.DEFAULT_USER;
            long habilitado = await db.ExecuteScalarAsync<long>(new CommandDefinition(@"
SELECT COUNT(1) FROM usuarios u JOIN permisos p ON p.perfil = u.perfil
WHERE u.nombreusuario = @Actor COLLATE NOCASE AND IFNULL(u.activo,1) = 1
  AND p.opcionespermitidas LIKE '%AgHorariosAcademicos%';",
                new { Actor = actor }, trans, cancellationToken: ct));
            if (habilitado <= 0)
                return Result.Failure("PERMISO_REVOCADO: el actor ya no puede autorizar la replanificación de aulas aprobadas.");

            // 3. Cada aula aprobada del alcance debe tener ratificación persistida para la revisión actual.
            foreach (int idAula in aulasAprobadas.OrderBy(x => x))
            {
                if (!contexto.PuedeGenerarAula(idAula))
                    return Result.Failure($"REPLANIFICACION_NO_AUTORIZADA: aula {idAula} sin decisión vigente en el alcance.");
                long ratificada = await db.ExecuteScalarAsync<long>(new CommandDefinition(@"
SELECT COUNT(1) FROM ag_candidato_plan_aprobado
WHERE id_aula = @IdAula AND huella_manifiesto = @Manifiesto AND huella_origen = @Origen
  AND inicio = @Inicio AND fin = @Fin;",
                    new
                    {
                        IdAula = idAula,
                        Manifiesto = manifiesto.Huella,
                        Origen = contexto.PlanAprobadoSnapshot?.HuellaOrigen,
                        Inicio = manifiesto.Inicio,
                        Fin = manifiesto.Fin
                    }, trans, cancellationToken: ct));
                if (ratificada <= 0)
                    return Result.Failure($"RATIFICACION_NO_PERSISTIDA: aula {idAula} sin consentimiento auditado para la revisión actual.");
            }

            return Result.Success();
        }

        /// <summary>
        /// Valida la correspondencia única de cada gen (asignatura, aula, bloque y multiconjunto de fechas autorizadas)
        /// contra el plan fechado sellado en ParametrosJson["planFechado"]. Rechaza discrepancias y completa fechas
        /// únicamente cuando falten en la entidad cargada.
        /// </summary>
        private static Result ValidarYCompletarFechasGenDesdeParametrosJson(AgIndividuo individuo, string parametrosJson)
        {
            if (individuo == null || individuo.Genes == null)
            {
                return Result.Failure("El individuo óptimo o sus genes son nulos.");
            }
            if (string.IsNullOrWhiteSpace(parametrosJson))
            {
                return Result.Failure("Los parámetros sellados (ParametrosJson) están vacíos o ausentes.");
            }

            try
            {
                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                serializer.MaxJsonLength = int.MaxValue;
                var dict = serializer.DeserializeObject(parametrosJson) as Dictionary<string, object>;
                if (dict == null || !dict.ContainsKey("planFechado") || dict["planFechado"] == null)
                {
                    return Result.Failure("No se encontró el plan fechado sellado ('planFechado') en ParametrosJson.");
                }

                var planFechado = dict["planFechado"] as object[]
                    ?? (dict["planFechado"] as System.Collections.IEnumerable)?.Cast<object>().ToArray();
                if (planFechado == null || planFechado.Length == 0)
                {
                    return Result.Failure("El plan fechado sellado en ParametrosJson no contiene registros.");
                }

                var listaPlan = new List<Dictionary<string, object>>();
                foreach (var item in planFechado)
                {
                    if (item is Dictionary<string, object> d)
                    {
                        listaPlan.Add(d);
                    }
                }

                var usadosPlan = new HashSet<int>();
                foreach (var gen in individuo.Genes)
                {
                    if (gen.HorasGen <= 0.001)
                    {
                        if (gen.FechasGen != null && gen.FechasGen.Count > 0)
                        {
                            return Result.Failure(
                                $"El gen #{gen.Id} (Asignatura {gen.IdAsignatura}) tiene 0 horas asignadas pero presenta {gen.FechasGen.Count} fecha(s) cargada(s).");
                        }
                        gen.FechasGen = new List<DateTime>();
                        continue;
                    }

                    int matchIdx = -1;
                    if (gen.Id > 0)
                    {
                        for (int i = 0; i < listaPlan.Count; i++)
                        {
                            if (!usadosPlan.Contains(i) && listaPlan[i].ContainsKey("idGen") &&
                                Convert.ToInt64(listaPlan[i]["idGen"]) == gen.Id)
                            {
                                matchIdx = i;
                                break;
                            }
                        }
                    }

                    if (matchIdx < 0)
                    {
                        for (int i = 0; i < listaPlan.Count; i++)
                        {
                            if (!usadosPlan.Contains(i) &&
                                listaPlan[i].ContainsKey("idAsignatura") && Convert.ToInt32(listaPlan[i]["idAsignatura"]) == gen.IdAsignatura &&
                                listaPlan[i].ContainsKey("idAula") && Convert.ToInt32(listaPlan[i]["idAula"]) == gen.IdAula &&
                                listaPlan[i].ContainsKey("idBloqueHorario") && Convert.ToInt32(listaPlan[i]["idBloqueHorario"]) == gen.IdBloqueHorario)
                            {
                                matchIdx = i;
                                break;
                            }
                        }
                    }

                    if (matchIdx < 0)
                    {
                        return Result.Failure(
                            $"El gen #{gen.Id} (Asignatura {gen.IdAsignatura}, Aula {gen.IdAula}, Bloque {gen.IdBloqueHorario}) no tiene correspondencia única en el plan fechado sellado.");
                    }

                    usadosPlan.Add(matchIdx);
                    var planEntry = listaPlan[matchIdx];

                    int planAsig = planEntry.ContainsKey("idAsignatura") ? Convert.ToInt32(planEntry["idAsignatura"]) : 0;
                    int planAula = planEntry.ContainsKey("idAula") ? Convert.ToInt32(planEntry["idAula"]) : 0;
                    int planBloque = planEntry.ContainsKey("idBloqueHorario") ? Convert.ToInt32(planEntry["idBloqueHorario"]) : 0;

                    if (planAsig != gen.IdAsignatura || planAula != gen.IdAula || planBloque != gen.IdBloqueHorario)
                    {
                        return Result.Failure(
                            $"Discrepancia en correspondencia del gen #{gen.Id}: cargado (Asignatura={gen.IdAsignatura}, Aula={gen.IdAula}, Bloque={gen.IdBloqueHorario}) " +
                            $"vs plan sellado (Asignatura={planAsig}, Aula={planAula}, Bloque={planBloque}).");
                    }

                    var fechasRaw = planEntry.ContainsKey("fechas")
                        ? (planEntry["fechas"] as object[] ?? (planEntry["fechas"] as System.Collections.IEnumerable)?.Cast<object>().ToArray())
                        : null;
                    var planFechas = new List<DateTime>();
                    if (fechasRaw != null)
                    {
                        foreach (var f in fechasRaw)
                        {
                            if (f != null && DateTime.TryParse(f.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                            {
                                planFechas.Add(dt.Date);
                            }
                        }
                    }
                    planFechas = planFechas.OrderBy(d => d).ToList();

                    if (planFechas.Count == 0)
                    {
                        return Result.Failure(
                            $"El gen #{gen.Id} (Asignatura {gen.IdAsignatura}) tiene {gen.HorasGen}h pero carece de fechas autorizadas en el plan sellado.");
                    }

                    if (gen.FechasGen != null && gen.FechasGen.Count > 0)
                    {
                        var cargadas = gen.FechasGen.Select(d => d.Date).OrderBy(d => d).ToList();
                        if (cargadas.Count != planFechas.Count)
                        {
                            return Result.Failure(
                                $"Discrepancia en cantidad de fechas para el gen #{gen.Id} (Asignatura {gen.IdAsignatura}): " +
                                $"cargadas {cargadas.Count} fechas frente a {planFechas.Count} fechas en el plan sellado.");
                        }

                        for (int k = 0; k < planFechas.Count; k++)
                        {
                            if (cargadas[k] != planFechas[k])
                            {
                                return Result.Failure(
                                    $"Discrepancia en fecha #{k + 1} del gen #{gen.Id} (Asignatura {gen.IdAsignatura}): " +
                                    $"cargada ({cargadas[k]:yyyy-MM-dd}) difiere de la sellada ({planFechas[k]:yyyy-MM-dd}).");
                            }
                        }
                    }
                    else
                    {
                        gen.FechasGen = planFechas;
                    }
                }

                if (usadosPlan.Count < listaPlan.Count)
                {
                    int sobrantesConFechas = 0;
                    for (int i = 0; i < listaPlan.Count; i++)
                    {
                        if (!usadosPlan.Contains(i))
                        {
                            var fList = listaPlan[i].ContainsKey("fechas") ? listaPlan[i]["fechas"] as System.Collections.IEnumerable : null;
                            if (fList != null && fList.Cast<object>().Any())
                            {
                                sobrantesConFechas++;
                            }
                        }
                    }
                    if (sobrantesConFechas > 0)
                    {
                        return Result.Failure(
                            $"El plan fechado sellado contiene {sobrantesConFechas} entrada(s) con fechas que no corresponden a ningún gen activo del individuo.");
                    }
                }

                return Result.Success();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgPromotorVinculaciones] Error validando FechasGen contra planFechado: {ex.Message}", ex);
                return Result.Failure($"Error validando fechas contra el plan sellado: {ex.Message}");
            }
        }

        /// <summary>
        /// Recupera y valida el plan/contexto sellado a partir del snapshot persistido en ParametrosJson de la corrida.
        /// Reconstruye la ventana, bloques, aulas, festivos y relaciones del snapshot sellado; verifica versión, identidad y huella;
        /// y consulta los catálogos relacionales canónicos con columnas reales verificadas.
        /// </summary>
        /// <param name="corrida">Entidad de la corrida leída desde ag_corrida.</param>
        /// <param name="individuo">Individuo óptimo.</param>
        /// <param name="ct">Token de cancelación.</param>
        /// <returns>Instancia de <see cref="AgContexto"/> sellada con índices de ejecución construidos.</returns>
        private async Task<Result<AgContexto>> RecuperarContextoSelladoAsync(
            AgCorrida corrida, AgIndividuo individuo, CancellationToken ct)
        {
            if (corrida == null)
            {
                return Result.Failure<AgContexto>("Corrida nula al intentar recuperar el contexto sellado.");
            }

            if (string.IsNullOrWhiteSpace(corrida.ParametrosJson))
            {
                return Result.Failure<AgContexto>($"La corrida #{corrida.Id} no tiene ParametrosJson persistido.");
            }

            try
            {
                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                serializer.MaxJsonLength = int.MaxValue;

                Dictionary<string, object> dict = null;
                try
                {
                    dict = serializer.DeserializeObject(corrida.ParametrosJson) as Dictionary<string, object>;
                }
                catch (Exception ex)
                {
                    return Result.Failure<AgContexto>($"Error al deserializar ParametrosJson de corrida #{corrida.Id}: {ex.Message}");
                }

                if (dict == null)
                {
                    return Result.Failure<AgContexto>("Estructura inválida o vacía en ParametrosJson.");
                }

                // 1. Verificar versión y huella
                if (!dict.ContainsKey("version") || string.IsNullOrWhiteSpace(dict["version"]?.ToString()))
                {
                    return Result.Failure<AgContexto>("El contexto sellado de la corrida carece de versión de esquema.");
                }

                string huellaEsperada = dict.ContainsKey("huella") ? dict["huella"]?.ToString() : null;
                if (string.IsNullOrWhiteSpace(huellaEsperada))
                {
                    return Result.Failure<AgContexto>("El contexto sellado de la corrida carece de huella criptográfica de integridad.");
                }

                // 2. Verificar snapshot y ventana
                var snapshot = dict.ContainsKey("snapshot") ? dict["snapshot"] as Dictionary<string, object> : null;
                var ventanaDict = dict.ContainsKey("ventana") ? dict["ventana"] as Dictionary<string, object> : null;
                if (snapshot == null || ventanaDict == null)
                {
                    return Result.Failure<AgContexto>("ParametrosJson no contiene el snapshot o la ventana semestral sellada.");
                }

                // 3. Verificar identidad del snapshot vs corrida
                int snapPeriodo = snapshot.ContainsKey("idPeriodo") ? Convert.ToInt32(snapshot["idPeriodo"]) : 0;
                int snapPrograma = snapshot.ContainsKey("idPrograma") ? Convert.ToInt32(snapshot["idPrograma"]) : 0;
                if (snapPeriodo != corrida.IdPeriodo || snapPrograma != corrida.IdPrograma)
                {
                    return Result.Failure<AgContexto>(
                        $"Inconsistencia de identidad en snapshot sellado: Periodo ({snapPeriodo} vs corrida {corrida.IdPeriodo}), " +
                        $"Programa ({snapPrograma} vs corrida {corrida.IdPrograma}).");
                }

                var ctx = new AgContexto
                {
                    IdPeriodo = corrida.IdPeriodo,
                    IdPrograma = corrida.IdPrograma,
                    Subunidad = corrida.Subunidad ?? string.Empty,
                    IdCorrida = corrida.Id,
                    IdAulaCorrida = corrida.IdAula
                };

                // 4. Reconstruir VentanaSemestre desde ventanaDict sellada
                if (!DateTime.TryParse(ventanaDict["fechaInicio"]?.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fIni) ||
                    !DateTime.TryParse(ventanaDict["fechaFin"]?.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fFin) ||
                    fFin.Date < fIni.Date)
                {
                    return Result.Failure<AgContexto>("Rango de fechas de ventana semestral inválido en ParametrosJson.");
                }

                string periodoNombre = ObtenerString(ventanaDict, "periodoNombre");
                string observaciones = ObtenerString(ventanaDict, "observaciones");

                var origenStr = ObtenerString(ventanaDict, "origen");
                Enum.TryParse(origenStr, true, out AgOrigenVentana origenVentana);
                var ventanaRes = VentanaSemestre.Crear(corrida.IdPeriodo, periodoNombre, fIni, fFin, origenVentana, observaciones);
                if (ventanaRes.IsFailure)
                {
                    return Result.Failure<AgContexto>($"Fallo al construir ventana semestral sellada: {ventanaRes.Error}");
                }
                ctx.Ventana = ventanaRes.Value;

                // 5. Reconstruir Bloques Horarios desde snapshot
                var bloquesRaw = snapshot.ContainsKey("bloques")
                    ? (snapshot["bloques"] as object[] ?? (snapshot["bloques"] as System.Collections.IEnumerable)?.Cast<object>().ToArray())
                    : null;
                if (bloquesRaw == null || bloquesRaw.Length == 0)
                {
                    return Result.Failure<AgContexto>("El snapshot sellado no contiene la lista de bloques horarios.");
                }

                foreach (var bObj in bloquesRaw)
                {
                    if (bObj is Dictionary<string, object> bDict)
                    {
                        var bloque = new AgBloqueHorario
                        {
                            Id = Convert.ToInt32(bDict["id"]),
                            DiaSemanaNumero = bDict.ContainsKey("dia") && bDict["dia"] != null ? Convert.ToInt32(bDict["dia"]) : 1,
                            Franja = ObtenerString(bDict, "franja"),
                            HoraInicio = ObtenerString(bDict, "horaInicio"),
                            HoraFin = ObtenerString(bDict, "horaFin"),
                            DuracionHoras = bDict.ContainsKey("duracionHoras") && bDict["duracionHoras"] != null ? Convert.ToDouble(bDict["duracionHoras"]) : 4.0,
                            Activo = bDict.ContainsKey("activo") && bDict["activo"] != null ? Convert.ToInt32(bDict["activo"]) : 1,
                            Subunidad = ObtenerString(bDict, "subunidad")
                        };
                        ctx.Bloques.Add(bloque);
                    }
                }

                // 6. Reconstruir Aulas desde snapshot
                var aulasRaw = snapshot.ContainsKey("aulas")
                    ? (snapshot["aulas"] as object[] ?? (snapshot["aulas"] as System.Collections.IEnumerable)?.Cast<object>().ToArray())
                    : null;
                if (aulasRaw == null || aulasRaw.Length == 0)
                {
                    return Result.Failure<AgContexto>("El snapshot sellado no contiene aulas autorizadas.");
                }

                foreach (var aObj in aulasRaw)
                {
                    if (aObj is Dictionary<string, object> au)
                    {
                        var aula = new AgAula
                        {
                            Id = Convert.ToInt32(au["id"]),
                            IdSede = au.ContainsKey("idSede") && au["idSede"] != null ? Convert.ToInt32(au["idSede"]) : 0,
                            IdPeriodo = au.ContainsKey("idPeriodo") && au["idPeriodo"] != null ? Convert.ToInt32(au["idPeriodo"]) : 0,
                            IdPrograma = au.ContainsKey("idPrograma") && au["idPrograma"] != null ? Convert.ToInt32(au["idPrograma"]) : 0,
                            Capacidad = au.ContainsKey("capacidad") && au["capacidad"] != null ? Convert.ToInt32(au["capacidad"]) : 0,
                            HorasDia = au.ContainsKey("horasDia") && au["horasDia"] != null ? Convert.ToInt32(au["horasDia"]) : 0,
                            EsVirtual = au.ContainsKey("esVirtual") && au["esVirtual"] != null ? Convert.ToInt32(au["esVirtual"]) : 0,
                            Estado = au.ContainsKey("estado") && au["estado"] != null ? (AgEstadoAula)Convert.ToInt32(au["estado"]) : AgEstadoAula.Disponible,
                            Codigo = ObtenerString(au, "codigo"),
                            Nombre = ObtenerString(au, "nombre"),
                            TipoAula = ObtenerString(au, "tipoAula", "ESTANDAR"),
                            CaracteristicasJson = ObtenerString(au, "caracteristicasJson", "{}"),
                            Semestre = ObtenerString(au, "semestre"),
                            Grupo = ObtenerString(au, "grupo", "A"),
                            ModalidadVinculacion = ObtenerString(au, "modalidadVinculacion", ModalidadTutoriaCatalog.Presencial),
                            FechaInicio = ObtenerString(au, "fechaInicio"),
                            FechaFin = ObtenerString(au, "fechaFin"),
                            SemestreInicio = ObtenerString(au, "semestreInicio"),
                            SemestreFin = ObtenerString(au, "semestreFin"),
                            VentanaOrigen = ObtenerString(au, "ventanaOrigen"),
                            AulaRef = ObtenerString(au, "aulaRef"),
                            Subunidad = ObtenerString(au, "subunidad"),
                            Jornada = ObtenerString(au, "jornada"),
                            DiasVentana = ObtenerString(au, "diasVentana")
                        };
                        ctx.Aulas.Add(aula);
                    }
                }

                // 7. Reconstruir Festivos desde snapshot
                if (snapshot.ContainsKey("festivos"))
                {
                    var festivosRaw = snapshot["festivos"] as object[]
                        ?? (snapshot["festivos"] as System.Collections.IEnumerable)?.Cast<object>().ToArray();
                    if (festivosRaw != null)
                    {
                        foreach (var fObj in festivosRaw)
                        {
                            if (fObj is Dictionary<string, object> fDict)
                            {
                                if (DateTime.TryParse(fDict["inicio"]?.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var fi) &&
                                    DateTime.TryParse(fDict["fin"]?.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var ff))
                                {
                                    ctx.Festivos.Add((fi.Date, ff.Date));
                                }
                            }
                        }
                    }
                }

                // 8. Reconstruir RelacionesAulaAsignatura desde snapshot
                if (snapshot.ContainsKey("relacionesAulaAsignatura"))
                {
                    var relsRaw = snapshot["relacionesAulaAsignatura"] as object[]
                        ?? (snapshot["relacionesAulaAsignatura"] as System.Collections.IEnumerable)?.Cast<object>().ToArray();
                    if (relsRaw != null)
                    {
                        foreach (var rObj in relsRaw)
                        {
                            if (rObj is Dictionary<string, object> r)
                            {
                                var aa = new AgAulaAsignatura
                                {
                                    IdAula = Convert.ToInt32(r["idAula"]),
                                    IdAsignatura = Convert.ToInt32(r["idAsignatura"]),
                                    HorasObjetivo = r.ContainsKey("horasObjetivo") && r["horasObjetivo"] != null ? Convert.ToDouble(r["horasObjetivo"]) : 0,
                                    OrdenCurricular = r.ContainsKey("ordenCurricular") && r["ordenCurricular"] != null ? Convert.ToInt32(r["ordenCurricular"]) : 0,
                                    DiasPatron = ObtenerString(r, "diasPatron"),
                                    OrigenPatron = ObtenerString(r, "origenPatron", "LEGADO"),
                                    Prioridad = r.ContainsKey("prioridad") && r["prioridad"] != null ? Convert.ToInt32(r["prioridad"]) : 1,
                                    Exclusivo = r.ContainsKey("exclusivo") && r["exclusivo"] != null ? Convert.ToInt32(r["exclusivo"]) : 0,
                                    Requerida = r.ContainsKey("requerida") && r["requerida"] != null ? Convert.ToInt32(r["requerida"]) : 1,
                                    Subunidad = ObtenerString(r, "subunidad")
                                };
                                ctx.AulaAsignaturas.Add(aa);
                            }
                        }
                    }
                }

                // 9. Reconstruir Programas desde snapshot (o fallback DB si es corrida legada sin snapshot de programas)
                if (snapshot.ContainsKey("programas") && snapshot["programas"] != null)
                {
                    var progsRaw = snapshot["programas"] as object[]
                        ?? (snapshot["programas"] as System.Collections.IEnumerable)?.Cast<object>().ToArray();
                    if (progsRaw != null)
                    {
                        foreach (var pObj in progsRaw)
                        {
                            if (pObj is Dictionary<string, object> p)
                            {
                                ctx.Programas.Add(new Programa
                                {
                                    Id = Convert.ToInt32(p["id"]),
                                    NombrePrograma = ObtenerString(p, "nombrePrograma"),
                                    Abreviatura = ObtenerString(p, "abreviatura"),
                                    Pensum = p.ContainsKey("pensum") && p["pensum"] != null ? Convert.ToInt32(p["pensum"]) : 0,
                                    Actual = p.ContainsKey("actual") && p["actual"] != null ? Convert.ToInt32(p["actual"]) : 1,
                                    Activo = p.ContainsKey("activo") && p["activo"] != null ? Convert.ToInt32(p["activo"]) : 1,
                                    Modalidad = ObtenerString(p, "modalidad")
                                });
                            }
                        }
                    }
                }

                // 10. Reconstruir Asignaturas desde snapshot (o fallback DB si es corrida legada)
                if (snapshot.ContainsKey("asignaturas") && snapshot["asignaturas"] != null)
                {
                    var asigsRaw = snapshot["asignaturas"] as object[]
                        ?? (snapshot["asignaturas"] as System.Collections.IEnumerable)?.Cast<object>().ToArray();
                    if (asigsRaw != null)
                    {
                        foreach (var aObj in asigsRaw)
                        {
                            if (aObj is Dictionary<string, object> a)
                            {
                                ctx.Asignaturas.Add(new Asignatura
                                {
                                    Id = Convert.ToInt32(a["id"]),
                                    IdPrograma = a.ContainsKey("idPrograma") && a["idPrograma"] != null ? Convert.ToInt32(a["idPrograma"]) : 0,
                                    Semestre = ObtenerString(a, "semestre"),
                                    NumNivel = a.ContainsKey("numNivel") && a["numNivel"] != null ? Convert.ToInt32(a["numNivel"]) : 0,
                                    CodUnidad = ObtenerString(a, "codUnidad"),
                                    CodMateria = ObtenerString(a, "codMateria"),
                                    UniTeorica = a.ContainsKey("uniTeorica") && a["uniTeorica"] != null ? Convert.ToInt32(a["uniTeorica"]) : 0,
                                    Metodologia = ObtenerString(a, "metodologia"),
                                    Abreviatura = ObtenerString(a, "abreviatura"),
                                    Nucleo = ObtenerString(a, "nucleo"),
                                    Nombre = ObtenerString(a, "nombre"),
                                    Creditos = a.ContainsKey("creditos") && a["creditos"] != null ? Convert.ToInt32(a["creditos"]) : 0,
                                    IntHoraria = a.ContainsKey("intHoraria") && a["intHoraria"] != null ? Convert.ToInt32(a["intHoraria"]) : 0,
                                    TipoAsignatura = ObtenerString(a, "tipoAsignatura"),
                                    TipoCredito = ObtenerString(a, "tipoCredito"),
                                    Prerequisito = ObtenerString(a, "prerequisito"),
                                    HorasSemanales = a.ContainsKey("horasSemanales") && a["horasSemanales"] != null ? Convert.ToDouble(a["horasSemanales"]) : 4.0,
                                    IdSede = a.ContainsKey("idSede") && a["idSede"] != null ? Convert.ToInt32(a["idSede"]) : 0,
                                    Modalidad = ObtenerString(a, "modalidad", "Presencial"),
                                    Subunidad = ObtenerString(a, "subunidad")
                                });
                            }
                        }
                    }
                }

                // Cargar catálogos relacionales canónicos para soporte complementario (docentes, sedes o fallbacks)
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    if (ctx.Asignaturas.Count == 0)
                    {
                        string sqlAsignaturas = @"
SELECT id AS Id, nombre AS Nombre, creditos AS Creditos,
       int_horaria AS IntHoraria, semestre AS Semestre,
       idprograma AS IdPrograma, subunidad AS Subunidad
FROM asignaturas;";
                        ctx.Asignaturas = (await db.QueryAsync<Asignatura>(
                            new CommandDefinition(sqlAsignaturas, cancellationToken: ct))).ToList();
                    }

                    if (ctx.Programas.Count == 0)
                    {
                        string sqlProgramas = @"
SELECT id AS Id, nombreprograma AS NombrePrograma, abreviatura AS Abreviatura,
       modalidad AS Modalidad, subunidad AS Subunidad
FROM programas;";
                        ctx.Programas = (await db.QueryAsync<Programa>(
                            new CommandDefinition(sqlProgramas, cancellationToken: ct))).ToList();
                    }

                    string sqlTerceros = @"
SELECT id AS Id, nombre AS Nombre, cedula AS Cedula,
       residencia AS Residencia, correo AS Correo, telefono AS Telefono,
       activo AS Activo, subunidad AS Subunidad
FROM terceros;";
                    ctx.Docentes = (await db.QueryAsync<Tercero>(
                        new CommandDefinition(sqlTerceros, cancellationToken: ct))).ToList();

                    string sqlSedes = @"
SELECT id AS Id, iddivipola AS IdDivipola, departamento AS Departamento,
       ciudad AS Ciudad, subsede AS Subsede, categoria AS Categoria,
       transporte AS Transporte, activo AS Activo, subunidad AS Subunidad
FROM sedes;";
                    ctx.Sedes = (await db.QueryAsync<Sede>(
                        new CommandDefinition(sqlSedes, cancellationToken: ct))).ToList();

                    var vacante = ctx.Docentes.FirstOrDefault(d => d.Cedula == "1111111111" || string.Equals(d.Nombre?.Trim(), "VACANTE", StringComparison.OrdinalIgnoreCase));
                    if (vacante != null)
                    {
                        ctx.IdTerceroVacante = vacante.Id;
                    }
                }

                ctx.ConstruirIndices();

                // 11. Extraer planFechado y validar huella criptográfica de integridad
                var aulaCorrida = ctx.Aulas.FirstOrDefault(a => a.Id == corrida.IdAula) ?? ctx.Aulas.FirstOrDefault();
                if (aulaCorrida == null)
                {
                    return Result.Failure<AgContexto>("No se encontró el aula autorizada para la corrida en el snapshot sellado.");
                }

                var planFechadoRaw = dict.ContainsKey("planFechado")
                    ? (dict["planFechado"] as object[] ?? (dict["planFechado"] as System.Collections.IEnumerable)?.Cast<object>().ToArray())
                    : null;
                var listaPlan = new List<Dictionary<string, object>>();
                if (planFechadoRaw != null)
                {
                    foreach (var item in planFechadoRaw)
                    {
                        if (item is Dictionary<string, object> d)
                        {
                            listaPlan.Add(d);
                        }
                    }
                }

                string huellaConPlan = AgHuellaPlanVentana.Calcular(ctx, aulaCorrida, planFechado: listaPlan);
                bool huellaValida = listaPlan.Count > 0 &&
                                    string.Equals(huellaConPlan, huellaEsperada, StringComparison.Ordinal);

                if (!huellaValida)
                {
                    if (listaPlan.Count == 0)
                    {
                        LoggerHelper.Warning($"[AgPromotorVinculaciones] La corrida #{corrida.Id} carece de plan fechado sellado en su snapshot.");
                        return Result.Failure<AgContexto>(
                            $"La corrida #{corrida.Id} no tiene plan fechado sellado o es previa a esta versión. " +
                            "Debe pasar primero por revalidación explícita antes de promoverse.");
                    }

                    LoggerHelper.Warning($"[AgPromotorVinculaciones] Huella calculada ({huellaConPlan}) difiere de huella sellada ({huellaEsperada}) para corrida #{corrida.Id}.");
                    return Result.Failure<AgContexto>(
                        $"La huella de integridad del contexto recuperado ({huellaConPlan}) no coincide con la huella sellada ({huellaEsperada}) de la corrida #{corrida.Id}. " +
                        "Los catálogos o parámetros fueron alterados con posterioridad a la ejecución.");
                }

                // 12. Validar correspondencia única y multiconjunto de fechas contra el plan sellado, completando únicamente cuando falten
                var validacionFechasRes = ValidarYCompletarFechasGenDesdeParametrosJson(individuo, corrida.ParametrosJson);
                if (validacionFechasRes.IsFailure)
                {
                    LoggerHelper.Warning($"[AgPromotorVinculaciones] Fallo en la validación de fechas contra el plan sellado para corrida #{corrida.Id}: {validacionFechasRes.Error}");
                    return Result.Failure<AgContexto>(validacionFechasRes.Error);
                }

                return Result.Success(ctx);
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<AgContexto>("Operación cancelada al recuperar el contexto sellado.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgPromotorVinculaciones] Error recuperando contexto sellado para corrida #{corrida.Id}", ex);
                return Result.Failure<AgContexto>($"Error al recuperar el contexto sellado: {ex.Message}");
            }
        }

        /// <summary>
        /// Parsea los IDs de asignaturas y claves granulares asig:sede autorizadas desde la cadena de justificación de vacantes (Fase 2032, 2038 y 2045).
        /// Exige token explícito (fail-closed): "asignaturas: 1, 36:2", "asig=10:1,20" o "materias: 10".
        /// Si no contiene un token explícito de asignaturas/materias, no autoriza ninguna para prevenir autorizaciones espurias.
        /// </summary>
        internal static (HashSet<int> Asigs, HashSet<string> Claves) ParseAutorizacionVacantes(string justificacion)
        {
            var asigs = new HashSet<int>();
            var claves = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(justificacion))
            {
                return (asigs, claves);
            }

            var match = System.Text.RegularExpressions.Regex.Match(
                justificacion,
                @"(?:asig(?:naturas)?|materias?)\s*[=:]\s*([0-9][0-9,:\s]*)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (!match.Success)
            {
                return (asigs, claves); // sin token explícito => sin autorización (fail-closed)
            }

            string rawList = match.Groups[1].Value;
            var tokens = rawList.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var rawToken in tokens)
            {
                string token = rawToken.Trim();
                if (token.Contains(':'))
                {
                    var partes = token.Split(':');
                    if (partes.Length == 2 &&
                        int.TryParse(partes[0].Trim(), out int asigId) && asigId > 0 &&
                        int.TryParse(partes[1].Trim(), out int sedeId))
                    {
                        if (sedeId > 0)
                        {
                            claves.Add($"{asigId}:{sedeId}");
                        }
                        else
                        {
                            asigs.Add(asigId);
                        }
                    }
                }
                else if (int.TryParse(token, out int id) && id > 0)
                {
                    asigs.Add(id);
                }
            }

            return (asigs, claves);
        }
    }
}
