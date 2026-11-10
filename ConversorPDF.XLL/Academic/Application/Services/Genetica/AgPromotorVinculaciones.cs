using System;
using System.Collections.Generic;
using System.Data;
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
            return PromoverAProduccionAsync(idCorrida, idIndividuo, usuario, null, ct);
        }

        public async Task<Result<AgPromocion>> PromoverAProduccionAsync(long idCorrida, long idIndividuo, string usuario, string justificacionVacante, CancellationToken ct)
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
    }
}
