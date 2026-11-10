using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Reevalúa el individuo óptimo de una corrida con el mismo validador del AG y reemplaza el
    /// detalle de violaciones persistido (Fase 1527).
    /// </summary>
    public sealed class AgRecalculoViolacionesService : IAgRecalculoViolacionesService
    {
        private readonly IAgScheduleRepository _scheduleRepository;
        private readonly IAgValidadorReglas _validador;

        public AgRecalculoViolacionesService(IAgScheduleRepository scheduleRepository, IAgValidadorReglas validador)
        {
            _scheduleRepository = scheduleRepository ?? throw new System.ArgumentNullException(nameof(scheduleRepository));
            _validador = validador ?? throw new System.ArgumentNullException(nameof(validador));
        }

        public async Task<Result<AgRecalculoResultado>> RecalcularAsync(long idCorrida, AgContexto contexto, CancellationToken ct)
        {
            if (contexto == null)
                return Result.Failure<AgRecalculoResultado>("Se requiere el contexto del AG para recalcular las violaciones.");
            if (idCorrida <= 0)
                return Result.Failure<AgRecalculoResultado>("IdCorrida inválido.");

            var corridaRes = await _scheduleRepository.GetCorridaByIdAsync(idCorrida, ct);
            if (corridaRes.IsFailure)
                return Result.Failure<AgRecalculoResultado>("No se pudo leer la corrida: " + corridaRes.Error);

            var corrida = corridaRes.Value;
            if (!corrida.IdIndividuoOptimo.HasValue || corrida.IdIndividuoOptimo.Value <= 0)
                return Result.Failure<AgRecalculoResultado>("La corrida no tiene individuo óptimo persistido; ejecute el AG antes de recalcular.");

            var indRes = await _scheduleRepository.GetIndividuoByIdAsync(corrida.IdIndividuoOptimo.Value, ct);
            if (indRes.IsFailure)
                return Result.Failure<AgRecalculoResultado>("No se pudo leer el individuo óptimo: " + indRes.Error);

            var individuo = indRes.Value;
            if (individuo.Genes == null || individuo.Genes.Count == 0)
                return Result.Failure<AgRecalculoResultado>("El individuo óptimo no tiene genes persistidos; no es posible recalcular.");

            contexto.IdCorrida = idCorrida;
            if (contexto.AulasPorId == null || contexto.AulasPorId.Count == 0)
                contexto.ConstruirIndices();

            _validador.Evaluar(individuo, contexto);

            var violaciones = individuo.Violaciones ?? new List<AgViolacion>();
            var porRegla = violaciones
                .GroupBy(v => v.Regla)
                .OrderByDescending(g => g.Count())
                .ToDictionary(g => g.Key, g => g.Count());

            var resumen = porRegla.Count == 0
                ? "sin violaciones"
                : string.Join(" · ", porRegla.Select(kv => kv.Key + "=" + kv.Value));

            var persist = await _scheduleRepository.ReemplazarViolacionesAsync(
                individuo.Id, violaciones, individuo.HardViolations, individuo.SoftViolations,
                individuo.Estado, idCorrida, ct);
            if (persist.IsFailure)
                return Result.Failure<AgRecalculoResultado>("No se pudo persistir el recálculo: " + persist.Error);

            return Result.Success(new AgRecalculoResultado
            {
                IdCorrida = idCorrida,
                IdIndividuo = individuo.Id,
                Hard = individuo.HardViolations,
                Soft = individuo.SoftViolations,
                GenesEvaluados = individuo.Genes.Count,
                AsignaturasContexto = contexto.Asignaturas?.Count ?? 0,
                Violaciones = violaciones,
                PorRegla = porRegla,
                ResumenPorRegla = resumen
            });
        }
    }
}