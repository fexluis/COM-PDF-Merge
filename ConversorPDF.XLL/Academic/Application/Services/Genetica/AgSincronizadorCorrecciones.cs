using System;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Servicio de sincronización de correcciones y ajustes manuales aplicados a soluciones del AG.
    /// </summary>
    public class AgSincronizadorCorrecciones
    {
        private readonly IAgScheduleRepository _scheduleRepo;
        private readonly ITerceroAsignaturaSyncService _poolSyncService;

        public AgSincronizadorCorrecciones(
            IAgScheduleRepository scheduleRepo,
            ITerceroAsignaturaSyncService poolSyncService)
        {
            _scheduleRepo = scheduleRepo ?? throw new ArgumentNullException(nameof(scheduleRepo));
            _poolSyncService = poolSyncService ?? throw new ArgumentNullException(nameof(poolSyncService));
        }

        public async Task<Result> SincronizarGenModificadoAsync(AgGen gen, string subunidad, CancellationToken ct)
        {
            if (gen == null) return Result.Failure("Gen nulo");

            if (gen.IdTercero > 0 && gen.IdAsignatura > 0)
            {
                var syncRes = await _poolSyncService.SincronizarDocenteAsignaturaAsync(
                    gen.IdTercero,
                    gen.IdPrograma,
                    gen.IdAsignatura,
                    gen.Semestre,
                    subunidad,
                    ct);

                if (syncRes.IsFailure)
                {
                    return Result.Failure($"Error sincronizando pool: {syncRes.Error}");
                }
            }

            return Result.Success();
        }

        public async Task<Result> SincronizarIndividuoAsync(AgCorrida corrida, AgIndividuo individuo, CancellationToken ct)
        {
            if (individuo == null) return Result.Failure("Individuo nulo");

            string subunidad = corrida?.Subunidad ?? individuo.Subunidad ?? string.Empty;

            if (individuo.Genes != null)
            {
                foreach (var gen in individuo.Genes)
                {
                    var res = await SincronizarGenModificadoAsync(gen, subunidad, ct);
                    if (res.IsFailure) return res;
                }
            }

            return Result.Success();
        }
    }
}
