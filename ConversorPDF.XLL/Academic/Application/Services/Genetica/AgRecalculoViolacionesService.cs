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

            // Fase 1532: el contexto debe corresponder a la corrida (evita sobrescribir con datos ajenos)
            if (contexto.IdPeriodo != corrida.IdPeriodo || contexto.IdPrograma != corrida.IdPrograma)
                return Result.Failure<AgRecalculoResultado>(
                    $"El contexto del AG no corresponde a la corrida #{idCorrida}: " +
                    $"contexto={contexto.IdPeriodo}/{contexto.IdPrograma}, corrida={corrida.IdPeriodo}/{corrida.IdPrograma}. " +
                    "Seleccione el período y programa de la corrida y vuelva a intentarlo.");

            // Fase 1538: solo comparar cuando ambos lados aportan subunidad (evita bloquear sesiones sin sync)
            bool contextoConSubunidad = !string.IsNullOrWhiteSpace(contexto.Subunidad);
            bool corridaConSubunidad = !string.IsNullOrWhiteSpace(corrida.Subunidad);
            if (contextoConSubunidad && corridaConSubunidad &&
                !string.Equals(contexto.Subunidad.Trim(), corrida.Subunidad.Trim(), System.StringComparison.OrdinalIgnoreCase))
                return Result.Failure<AgRecalculoResultado>(
                    $"La territorial del contexto ('{contexto.Subunidad}') no coincide con la de la corrida ('{corrida.Subunidad}').");

            var indRes = await _scheduleRepository.GetIndividuoByIdAsync(corrida.IdIndividuoOptimo.Value, ct);
            if (indRes.IsFailure)
                return Result.Failure<AgRecalculoResultado>("No se pudo leer el individuo óptimo: " + indRes.Error);

            var individuo = indRes.Value;
            if (individuo.Genes == null || individuo.Genes.Count == 0)
                return Result.Failure<AgRecalculoResultado>("El individuo óptimo no tiene genes persistidos; no es posible recalcular.");

            // Fase 1533 & Fase 1540: Validar que las colecciones requeridas del contexto estén instanciadas antes de ConstruirIndices
            if (contexto.Aulas == null || contexto.Bloques == null || contexto.Asignaturas == null ||
                contexto.Docentes == null || contexto.Sedes == null || contexto.PoolDocentes == null ||
                contexto.Festivos == null || contexto.AulaAsignaturas == null)
                return Result.Failure<AgRecalculoResultado>(
                    "El contexto del AG está incompleto (aulas/bloques/asignaturas/docentes/sedes/pool/festivos). Reconstruya el contexto y reintente.");

            // Fase 1534: Preservar IdCorrida previo del contexto para no mutar el estado compartido del llamador
            long idCorridaPrevio = contexto.IdCorrida;
            contexto.IdCorrida = idCorrida;
            try
            {
                if (contexto.AulasPorId == null || contexto.AulasPorId.Count == 0)
                    contexto.ConstruirIndices();

                // Fase 1549, 1551 & 1553: abortar si el contexto no cubre las aulas o asignaturas de la corrida
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

                // Fase 1549, 1551 & 1552: solo aulas y asignaturas condicionan la evaluación
                // (AgValidadorReglas usa AulasPorId y AsignaturasPorId). Los docentes no participan
                // en el cálculo de violaciones: exigirlos rechazaría recálculos legítimos.
                if (aulasFaltantes.Count > 0 || asignaturasFaltantes.Count > 0)
                    return Result.Failure<AgRecalculoResultado>(
                        "El contexto reconstruido no cubre la corrida: " +
                        $"{aulasFaltantes.Count} aula(s), {asignaturasFaltantes.Count} asignatura(s) ausentes " +
                        $"(aulas: {string.Join(", ", aulasFaltantes.Take(5))}; " +
                        $"asignaturas: {string.Join(", ", asignaturasFaltantes.Take(5))}). " +
                        "Seleccione la sede y las cohortes de la corrida antes de recalcular; " +
                        "de lo contrario se persistirían violaciones incompletas.");

                _validador.Evaluar(individuo, contexto);
            }
            finally
            {
                contexto.IdCorrida = idCorridaPrevio;
            }

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