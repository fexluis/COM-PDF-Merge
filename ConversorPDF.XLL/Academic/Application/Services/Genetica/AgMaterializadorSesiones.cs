using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Materializa una solución de un individuo en las tablas de staging 'ag_vinculaciones' y 'ag_sesiones'.
    /// Agrupa genes de la misma asignatura/docente/aula/ventana en una única vinculación y expande sus sesiones.
    /// </summary>
    public class AgMaterializadorSesiones : IAgMaterializadorSesiones
    {
        private readonly IAgScheduleRepository _scheduleRepository;
        private readonly IAgExpansorSesiones _expansorSesiones;

        public AgMaterializadorSesiones(
            IAgScheduleRepository scheduleRepository,
            IAgExpansorSesiones expansorSesiones)
        {
            _scheduleRepository = scheduleRepository ?? throw new ArgumentNullException(nameof(scheduleRepository));
            _expansorSesiones = expansorSesiones ?? throw new ArgumentNullException(nameof(expansorSesiones));
        }

        public async Task<Result> MaterializarEnStagingAsync(long idCorrida, AgIndividuo individuo, AgContexto contexto, CancellationToken ct)
        {
            if (individuo == null) return Result.Failure("El individuo no puede ser nulo.");
            if (contexto == null) return Result.Failure("El contexto no puede ser nulo.");
            if (individuo.Genes == null || individuo.Genes.Count == 0)
                return Result.Failure("Cromosoma vacío: no hay tuplas que materializar.");

            try
            {
                // Fase 502: la purga por período/programa ya ocurrió antes de crear la corrida; limpiar únicamente por idCorrida
                await _scheduleRepository.LimpiarStagingAsync(idCorrida, ct);

                var vinculacionesStaging = new List<AgVinculacion>();
                var sesionesStaging = new List<AgSesion>();
                var incidencias = new List<string>();

                // Fase 131 / Fase 398: Agrupar genes determinísticamente por ClaveGrupo (IdAsignatura|IdAula|Grupo) o tupla
                var grupos = individuo.Genes
                    .GroupBy(g => !string.IsNullOrWhiteSpace(g.ClaveGrupo)
                        ? g.ClaveGrupo
                        : $"{g.IdAsignatura}|{g.IdAula}|{g.Grupo}")
                    .ToList();

                int idSeq = 1;
                foreach (var grp in grupos)
                {
                    var genes = grp.OrderBy(g => g.IndiceDiaPatron).ToList();
                    int idAulaGrupo = genes.Count > 0 ? genes[0].IdAula : 0;
                    if (LoggerHelper.ModoVerboso)
                    {
                        LoggerHelper.Debug($"[AG-RANURAS-TMP] STAGING_ANTES_SELLAR corrida={idCorrida} individuo={individuo.Id} " + DescribirGrupo(genes, genes[0].TotalHorasAsignatura));
                    }

                    // Fase 583 / Fase 589 / Fase 1304: coherencia defensiva SIEMPRE (I1 + desdoble de franjas de sábado)
                    if (genes.Count > 1)
                    {
                        var tmp = new AgIndividuo();
                        foreach (var g in genes) tmp.Genes.Add(g);   // referencias: la reparación impacta al individuo
                        AgCoherenciaGrupo.Sellar(tmp, contexto, individuo.Genes);
                    }

                    // Fase 133: Mapeo centralizado vía AgGenToAgVinculacionMapper (Estado = Borrador)
                    if (LoggerHelper.ModoVerboso)
                    {
                        LoggerHelper.Debug($"[AG-RANURAS-TMP] STAGING_DESPUES_SELLAR corrida={idCorrida} individuo={individuo.Id} " + DescribirGrupo(genes, genes[0].TotalHorasAsignatura));
                    }
                    var agVinc = AgGenToAgVinculacionMapper.MapGrupo(genes, idCorrida, individuo.Id, contexto);
                    if (agVinc == null) continue;

                    // Fase 455: Asignar y validar consistencia del Grupo en la vinculación
                    agVinc.Grupo = genes[0].Grupo ?? "A";
                    if (genes.Any(g => !string.Equals(g.Grupo, agVinc.Grupo, StringComparison.OrdinalIgnoreCase)))
                    {
                        return Result.Failure($"Grupo inconsistente en {genes[0].ClaveGrupo}: {string.Join(",", genes.Select(g => g.Grupo))}.");
                    }

                    // Fase 664: Coherencia defensiva de programa entre aula y vinculación
                    if (contexto.AulasPorId != null && contexto.AulasPorId.TryGetValue(idAulaGrupo, out var aCohorte))
                    {
                        if (aCohorte.IdPrograma > 0 && agVinc.IdPrograma > 0 && aCohorte.IdPrograma != agVinc.IdPrograma)
                        {
                            return Result.Failure($"Incoherencia de programa en aula {idAulaGrupo} ({aCohorte.Codigo}): aula pertenece a programa {aCohorte.IdPrograma} pero vinculación de asignatura {agVinc.IdAsignatura} pertenece a programa {agVinc.IdPrograma}.");
                        }
                    }

                    // Fase 458 / Fase 1108: Coherencia idperiodo ↔ ventana; preservar la sub-ventana de oleada acotada a la ventana
                    if (contexto.Ventana != null && contexto.Ventana.IdPeriodo > 0)
                    {
                        if (agVinc.IdPeriodo > 0 && agVinc.IdPeriodo != contexto.Ventana.IdPeriodo)
                        {
                            return Result.Failure($"Período incoherente: vinculación={agVinc.IdPeriodo} ventana={contexto.Ventana.IdPeriodo} (aula {idAulaGrupo}).");
                        }
                        agVinc.IdPeriodo = contexto.Ventana.IdPeriodo;
                        // Fase 1108 / Fase 1115: preservar la sub-ventana de oleada; acotar a la ventana con parseo invariante
                        bool okIni = DateTime.TryParseExact(agVinc.FechaInicio, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var fiVin);
                        if (string.IsNullOrEmpty(agVinc.FechaInicio) || (okIni && fiVin < contexto.Ventana.FechaInicio))
                        {
                            agVinc.FechaInicio = contexto.Ventana.FechaInicio.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                        }
                        else if (!okIni)
                        {
                            LoggerHelper.Warning($"[Materializador] FechaInicio no parseable '{agVinc.FechaInicio}' (grupo {genes[0].ClaveGrupo}).");
                            agVinc.FechaInicio = contexto.Ventana.FechaInicio.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                        }

                        bool okFin = DateTime.TryParseExact(agVinc.FechaFin, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var ffVin);
                        if (string.IsNullOrEmpty(agVinc.FechaFin) || (okFin && ffVin > contexto.Ventana.FechaFin))
                        {
                            agVinc.FechaFin = contexto.Ventana.FechaFin.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                        }
                        else if (!okFin)
                        {
                            LoggerHelper.Warning($"[Materializador] FechaFin no parseable '{agVinc.FechaFin}' (grupo {genes[0].ClaveGrupo}).");
                            agVinc.FechaFin = contexto.Ventana.FechaFin.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                        }
                    }

                    agVinc.Id = idSeq;

                    // FASE 897: Guarda de bloque horario inexistente o <= 0 en genes con carga horaria (I6_BloqueInexistente)
                    var genBloqueInvalido = genes.FirstOrDefault(g => g.HorasGen > 0.0 && (g.IdBloqueHorario <= 0 || contexto.BloquesPorId == null || !contexto.BloquesPorId.ContainsKey(g.IdBloqueHorario)));
                    if (genBloqueInvalido != null)
                    {
                        double esperadoI6 = genes.Count > 0 ? genes.Max(g => g.TotalHorasAsignatura) : 0.0;
                        var detalleI6 = DescribirGrupo(genes, esperadoI6);
                        LoggerHelper.Warning($"[Materializador] grupo en cuarentena (I6_BloqueInexistente): {detalleI6} | Gen {genBloqueInvalido.Id} con bloque {genBloqueInvalido.IdBloqueHorario} inexistente.");
                        incidencias.Add($"I6_BloqueInexistente {detalleI6}");
                        continue;
                    }

                    string DiaHorasSemanalesDe(IEnumerable<AgGen> gens)
                    {
                        // Fase 990 & Fase 1295: deduplicar por (Día, HoraInicio, HoraFin) en vez de IdBloqueHorario
                        // para admitir múltiples franjas en el mismo día (ej. Sábado AM y PM) sin duplicar ni colapsar.
                        var partes = gens
                            .Where(g => g.DiaSemana >= 1 && g.HorasGen > 0.001)   // Fase 1314: genes sin carga no aportan sesiones
                            .Select(g =>
                            {
                                double horas = g.HorasSemanales > 0 ? g.HorasSemanales : 4.0;
                                string horario = string.Empty;
                                string hIni = string.Empty;
                                string hFin = string.Empty;
                                if (g.IdBloqueHorario > 0 && contexto.BloquesPorId != null
                                    && contexto.BloquesPorId.TryGetValue(g.IdBloqueHorario, out var bDia))
                                {
                                    horas = bDia.DuracionHoras;
                                    horario = bDia.ObtenerHorarioTexto();
                                    hIni = bDia.HoraInicio ?? string.Empty;
                                    hFin = bDia.HoraFin ?? string.Empty;
                                }

                                return new
                                {
                                    g.DiaSemana,
                                    HoraInicio = hIni,
                                    HoraFin = hFin,
                                    Horas = horas,
                                    Horario = horario
                                };
                            })
                            .GroupBy(x => new { x.DiaSemana, x.HoraInicio, x.HoraFin })
                            .OrderBy(x => x.Key.DiaSemana)
                            .ThenBy(x => x.Key.HoraInicio)
                            .Select(grDia =>
                            {
                                var primero = grDia.First();
                                double horas = grDia.Max(x => x.Horas);
                                string horario = primero.Horario;

                                // Fase 995: forzar cultura invariante en horas para evitar que ',' colisione con el separador de lista
                                return $"{DiaSemanaNormalizador.ObtenerNombreDia(grDia.Key.DiaSemana)}|{horas.ToString("0.##", CultureInfo.InvariantCulture)}|{horario}";
                            })
                            .ToList();

                        return string.Join(", ", partes);
                    }

                    agVinc.DiaHorasSemanales = DiaHorasSemanalesDe(genes);

                    // Fase 294 / Fase 409: Asegurar IdSede válido y fail-fast preventivo (resuelto vía idAulaGrupo)
                    if (agVinc.IdSede <= 0)
                    {
                        if (contexto.AulasPorId != null && contexto.AulasPorId.TryGetValue(idAulaGrupo, out var aRef) && aRef.IdSede > 0)
                        {
                            agVinc.IdSede = aRef.IdSede;
                        }
                        else if (contexto.IdSede > 0)
                        {
                            agVinc.IdSede = contexto.IdSede;
                        }
                    }

                    if (agVinc.IdSede <= 0)
                    {
                        return Result.Failure($"No se pudo resolver la sede (IdSede <= 0) para el aula {idAulaGrupo} al materializar en staging.");
                    }

                    // Fase 417 & 573: Validación estricta de Invariante I1 (Total Horas Grupo vs TotalHorasAsignatura)
                    double horasGrupo = genes.Sum(g => g.HorasGen);
                    double esperado = genes.Max(g => g.TotalHorasAsignatura);
                    if (Math.Abs(horasGrupo - esperado) > 0.001)
                    {
                        // Fase 581 & 584: Diagnóstico y cuarentena por grupo al violar I1
                        var detalle = DescribirGrupo(genes, esperado);
                        LoggerHelper.Warning($"[Materializador] grupo en cuarentena (I1): {detalle}");
                        incidencias.Add($"I1 {detalle}");
                        continue;
                    }

                    // Expansión cronológica de sesiones asociada al idSeq de la vinculación agrupada
                    int sesionesAntes = sesionesStaging.Count;
                    bool expansionFallida = false;
                    foreach (var g in genes)
                    {
                        if (g.HorasGen <= 0.0) continue;   // Fase 576: gen sin carga no aporta sesiones
                        var expRes = _expansorSesiones.Expandir(g, idSeq, contexto);
                        if (expRes.IsFailure)
                        {
                            LoggerHelper.Debug($"[AG-RANURAS-TMP] EXPANDIR_ERROR corrida={idCorrida} gen={g.Id} aula={g.IdAula} asig={g.IdAsignatura} motivo={expRes.Error}");
                            // Fase 584: cuarentena por grupo al fallar expansión
                            var detalleExp = DescribirGrupo(genes, esperado);
                            LoggerHelper.Warning($"[Materializador] grupo en cuarentena (expansión): {detalleExp} | {expRes.Error}");
                            incidencias.Add($"EXP {genes[0].ClaveGrupo}: {expRes.Error}");
                            sesionesStaging.RemoveRange(sesionesAntes, sesionesStaging.Count - sesionesAntes);
                            expansionFallida = true;
                            break;
                        }
                        sesionesStaging.AddRange(expRes.Value);
                    }
                    if (expansionFallida) continue;

                    // Fase 509 & Fase 520: Sellar docente/vacante/pool de todas las sesiones del grupo con la vinculación
                    for (int sIdx = sesionesAntes; sIdx < sesionesStaging.Count; sIdx++)
                    {
                        sesionesStaging[sIdx].IdTercero = agVinc.IdTercero;
                        sesionesStaging[sIdx].EsVacante = agVinc.EsVacante;
                        sesionesStaging[sIdx].MotivoVacante = agVinc.EsVacante == 1 ? agVinc.MotivoVacante : string.Empty;
                        sesionesStaging[sIdx].CandidatosPool = agVinc.CandidatosPool;
                    }

                    // Fase 417 & F651: Validación dinámica de Invariante I2 (Sum(Horas) == esperado, soporta 4h y 8h)
                    int sesionesGrupo = sesionesStaging.Count - sesionesAntes;
                    double sumaHorasGrupo = sesionesStaging.Skip(sesionesAntes).Take(sesionesGrupo).Sum(s => s.Horas);
                    if (Math.Abs(sumaHorasGrupo - esperado) > 0.001)
                    {
                        // Fase 581 & 584: Diagnóstico y cuarentena por grupo al violar I2
                        var detalleI2 = DescribirGrupo(genes, esperado);
                        LoggerHelper.Warning($"[Materializador] grupo en cuarentena (I2): {detalleI2}");
                        incidencias.Add($"I2 {detalleI2}");
                        sesionesStaging.RemoveRange(sesionesAntes, sesionesStaging.Count - sesionesAntes);
                        continue;
                    }

                    // Invariante I7 (Fase 1296): No permitir sesiones duplicadas en (FechaSesion, HorarioSesion) en el mismo grupo
                    var sesGrupoLista = sesionesStaging.Skip(sesionesAntes).Take(sesionesGrupo).ToList();
                    var duplicada = sesGrupoLista
                        .Where(s => s.FechaSesion.HasValue && !string.IsNullOrEmpty(s.HorarioSesion))
                        .GroupBy(s => new { Fecha = s.FechaSesion.Value.Date, s.HorarioSesion })
                        .FirstOrDefault(gr => gr.Count() > 1);
                    if (duplicada != null)
                    {
                        var detalleI7 = DescribirGrupo(genes, esperado);
                        LoggerHelper.Warning($"[Materializador] grupo en cuarentena (I7_ColisionSesion): {detalleI7} | Sesión duplicada {duplicada.Key.Fecha:yyyy-MM-dd} {duplicada.Key.HorarioSesion}");
                        incidencias.Add($"I7 {detalleI7}");
                        sesionesStaging.RemoveRange(sesionesAntes, sesionesStaging.Count - sesionesAntes);
                        continue;
                    }

                    vinculacionesStaging.Add(agVinc);
                    idSeq++;
                }

                // Fase 584: Guarda antes de GuardarStagingAsync
                if (vinculacionesStaging.Count == 0)
                {
                    return Result.Failure($"Staging vacío: {incidencias.Count} grupo(s) en cuarentena. Primera: {incidencias.FirstOrDefault()}");
                }
                if (incidencias.Count > 0)
                {
                    LoggerHelper.Warning($"[AgMaterializadorSesiones] {incidencias.Count} grupo(s) en cuarentena; staging parcial: {vinculacionesStaging.Count} vinculaciones, {sesionesStaging.Count} sesiones.");
                }

                var saveResult = await _scheduleRepository.GuardarStagingAsync(idCorrida, individuo.Id, vinculacionesStaging, sesionesStaging, ct);
                if (saveResult.IsFailure)
                {
                    return Result.Failure($"Error al guardar en staging: {saveResult.Error}");
                }

                LoggerHelper.Info($"[AgMaterializadorSesiones] Solución materializada en staging: {vinculacionesStaging.Count} vinculaciones, {sesionesStaging.Count} sesiones.");
                if (LoggerHelper.ModoVerboso)
                {
                    LoggerHelper.Debug($"[AG-RANURAS-TMP] STAGING_GUARDADO corrida={idCorrida} individuo={individuo.Id} sesiones={sesionesStaging.Count} ultima={sesionesStaging.Select(s => (DateTime?)s.FechaSesion).Max():yyyy-MM-dd} limiteEfectivo={contexto.Ventana?.FechaFin:yyyy-MM-dd}");
                }
                return Result.Success();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgMaterializadorSesiones] Error materializando solución en staging", ex);
                return Result.Failure($"Error materializando solución: {ex.Message}");
            }
        }

        // Fase 581: Helper de diagnóstico para describir composición de grupo al violar invariantes o entrar en cuarentena
        private static string DescribirGrupo(IReadOnlyList<AgGen> genes, double esperado) =>
            $"clave='{genes[0].ClaveGrupo}' genes={genes.Count} esperado={esperado:0.#} Σ={genes.Sum(g => g.HorasGen):0.#} [" +
            string.Join("; ", genes.Select(g =>
                $"id={g.Id} asig={g.IdAsignatura} aula={g.IdAula} grp={g.Grupo} sem={g.Semestre} prog={g.IdPrograma} " +
                $"h={g.HorasGen:0.#} tot={g.TotalHorasAsignatura:0.#} est={g.Estado} fechas={g.FechasGen?.Count ?? 0}")) + "]";
    }
}
