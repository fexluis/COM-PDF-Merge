using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Repara corridas huérfanas del AG: staging presente (<c>ag_vinculaciones</c>/<c>ag_sesiones</c>) pero sin
    /// cabecera en <c>ag_corrida</c> ni individuo en <c>ag_individuo</c> (Fase 1566).
    /// Reutiliza los genes de <c>ag_gen</c>, evalúa el individuo con el mismo evaluador del AG y persiste el
    /// detalle de violaciones en <c>ag_violacion</c>, habilitando de nuevo la auditoría y la promoción.
    /// </summary>
    public sealed class AgReparacionCorridaService : IAgReparacionCorridaService
    {
        private readonly IAgScheduleRepository _scheduleRepository;
        private readonly IAgEvaluadorAptitud _evaluador;

        public AgReparacionCorridaService(IAgScheduleRepository scheduleRepository, IAgEvaluadorAptitud evaluador)
        {
            _scheduleRepository = scheduleRepository ?? throw new ArgumentNullException(nameof(scheduleRepository));
            _evaluador = evaluador ?? throw new ArgumentNullException(nameof(evaluador));
        }

        public async Task<Result<AgReparacionCorridaResultado>> RepararAsync(
            long idCorrida, AgContexto contexto, CancellationToken ct)
        {
            if (contexto == null)
                return Result.Failure<AgReparacionCorridaResultado>(
                    "Se requiere el contexto del AG para reparar la corrida.");
            if (idCorrida <= 0)
                return Result.Failure<AgReparacionCorridaResultado>("IdCorrida inválido.");

            var huerfanaRes = await _scheduleRepository.DetectarCorridaHuerfanaAsync(idCorrida, ct);
            if (huerfanaRes.IsFailure)
                return Result.Failure<AgReparacionCorridaResultado>(huerfanaRes.Error);

            var huerfana = huerfanaRes.Value;
            if (huerfana.ExisteCabecera)
                return Result.Failure<AgReparacionCorridaResultado>(
                    $"La corrida #{idCorrida} conserva cabecera en ag_corrida; no es una corrida huérfana.");

            if (!huerfana.PuedeRepararse)
                return Result.Failure<AgReparacionCorridaResultado>(
                    $"La corrida #{idCorrida} no puede repararse: {huerfana.TotalVinculaciones} vinculación(es) en staging, " +
                    $"individuo #{huerfana.IdIndividuo}, período {huerfana.IdPeriodo}, programa {huerfana.IdPrograma}. " +
                    "Se requiere staging materializado y un individuo identificable en ag_vinculaciones.");

            // Fase 1577: el contexto debe pertenecer a la corrida. Con programas homónimos (mismo nombre, distinto id)
            // el contexto podía quedar en el programa equivocado y se persistían violaciones calculadas con otro pensum.
            if (contexto.IdPeriodo != huerfana.IdPeriodo || contexto.IdPrograma != huerfana.IdPrograma)
                return Result.Failure<AgReparacionCorridaResultado>(
                    $"El contexto del AG no corresponde a la corrida #{idCorrida}: " +
                    $"contexto={contexto.IdPeriodo}/{contexto.IdPrograma}, corrida={huerfana.IdPeriodo}/{huerfana.IdPrograma}. " +
                    "Seleccione el período y el programa reales de la corrida (puede haber programas homónimos con distinto id) " +
                    "y vuelva a intentarlo.");

            var genesRes = await _scheduleRepository.GetGenesByIndividuoAsync(huerfana.IdIndividuo, ct);
            if (genesRes.IsFailure)
                return Result.Failure<AgReparacionCorridaResultado>(genesRes.Error);
            if (genesRes.Value.Count == 0)
                return Result.Failure<AgReparacionCorridaResultado>(
                    $"El individuo #{huerfana.IdIndividuo} no conserva genes en ag_gen; no hay base para reconstruir la corrida. " +
                    "Vuelva a ejecutar el AG.");

            var individuo = new AgIndividuo
            {
                Id = huerfana.IdIndividuo,
                IdCorrida = idCorrida,
                Estado = AgEstadoIndividuo.NoEvaluado,
                Subunidad = huerfana.Subunidad ?? string.Empty,
                Genes = genesRes.Value.ToList()
            };

            // Guarda de cobertura: nunca se escribe con un contexto parcial (persistiría violaciones incompletas).
            if (contexto.AulasPorId == null || contexto.AulasPorId.Count == 0)
                contexto.ConstruirIndices();

            var aulasFaltantes = individuo.Genes
                .Where(g => g.IdAula > 0)
                .Select(g => g.IdAula)
                .Distinct()
                .Where(id => contexto.AulasPorId == null || !contexto.AulasPorId.ContainsKey(id))
                .ToList();

            var asignaturasFaltantes = individuo.Genes
                .Where(g => g.IdAsignatura > 0)
                .Select(g => g.IdAsignatura)
                .Distinct()
                .Where(id => contexto.AsignaturasPorId == null || !contexto.AsignaturasPorId.ContainsKey(id))
                .ToList();

            if (aulasFaltantes.Count > 0 || asignaturasFaltantes.Count > 0)
                return Result.Failure<AgReparacionCorridaResultado>(
                    "El contexto reconstruido no cubre la corrida: " +
                    $"{aulasFaltantes.Count} aula(s) y {asignaturasFaltantes.Count} asignatura(s) ausentes " +
                    $"(aulas: {string.Join(", ", aulasFaltantes.Take(5))}; asignaturas: {string.Join(", ", asignaturasFaltantes.Take(5))}). " +
                    "Seleccione la sede y las cohortes de la corrida antes de repararla.");

            // Evaluación con el mismo evaluador del AG: fija Fitness, HARD/SOFT y el detalle por regla.
            long idCorridaPrevio = contexto.IdCorrida;
            contexto.IdCorrida = idCorrida;
            try
            {
                _evaluador.EvaluarAptitud(individuo, contexto);
            }
            finally
            {
                contexto.IdCorrida = idCorridaPrevio;
            }

            individuo.Estado = individuo.HardViolations == 0
                ? AgEstadoIndividuo.Optimo
                : AgEstadoIndividuo.Infactible;

            var ahora = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var corrida = new AgCorrida
            {
                Id = idCorrida,
                IdPeriodo = huerfana.IdPeriodo,
                IdPrograma = huerfana.IdPrograma,
                IdAula = 0,                     // la compuerta de aula se omite con id_aula = 0 (Fase 1520 G2)
                FechaInicioCorrida = ahora,
                FechaFinCorrida = ahora,
                Estado = AgEstadoCorrida.Completada,
                PoblacionTamano = 0,
                GeneracionesMax = 0,
                GeneracionesEjecutadas = 0,
                MejorFitness = individuo.Fitness,
                HardViolations = individuo.HardViolations,
                SoftViolations = individuo.SoftViolations,
                Vacantes = individuo.Vacantes,
                IdIndividuoOptimo = individuo.Id,
                ParametrosJson = "{}",
                Subunidad = huerfana.Subunidad ?? string.Empty
            };

            var insertRes = await _scheduleRepository.ReinsertarCabeceraHuerfanaAsync(
                corrida, idCorrida, individuo, ct);
            if (insertRes.IsFailure)
                return Result.Failure<AgReparacionCorridaResultado>(insertRes.Error);

            var violaciones = individuo.Violaciones ?? new List<AgViolacion>();
            var persistRes = await _scheduleRepository.ReemplazarViolacionesAsync(
                individuo.Id, violaciones, individuo.HardViolations, individuo.SoftViolations,
                individuo.Estado, idCorrida, ct);
            if (persistRes.IsFailure)
                return Result.Failure<AgReparacionCorridaResultado>(
                    "La cabecera se reconstruyó pero no se pudieron persistir las violaciones: " + persistRes.Error);

            var porRegla = violaciones
                .GroupBy(v => v.Regla)
                .OrderByDescending(g => g.Count())
                .ToDictionary(g => g.Key, g => g.Count());

            string resumen = porRegla.Count == 0
                ? "sin violaciones"
                : string.Join(" · ", porRegla.Select(kv => kv.Key + "=" + kv.Value));

            LoggerHelper.Info(
                $"[AG-Reparacion] Corrida #{idCorrida} reparada: individuo #{individuo.Id}, " +
                $"período {huerfana.IdPeriodo}, programa {huerfana.IdPrograma}, genes {individuo.Genes.Count}, " +
                $"HARD={individuo.HardViolations}, SOFT={individuo.SoftViolations:0.##}, " +
                $"vinculaciones={huerfana.TotalVinculaciones}, sesiones={huerfana.TotalSesiones} · {resumen}");

            return Result.Success(new AgReparacionCorridaResultado
            {
                IdCorrida = idCorrida,
                IdIndividuo = individuo.Id,
                IdPeriodo = huerfana.IdPeriodo,
                IdPrograma = huerfana.IdPrograma,
                Subunidad = huerfana.Subunidad ?? string.Empty,
                Hard = individuo.HardViolations,
                Soft = individuo.SoftViolations,
                Fitness = individuo.Fitness,
                GenesEvaluados = individuo.Genes.Count,
                Violaciones = violaciones.Count,
                Vinculaciones = huerfana.TotalVinculaciones,
                Sesiones = huerfana.TotalSesiones,
                PorRegla = porRegla,
                ResumenPorRegla = resumen
            });
        }
    }
}

