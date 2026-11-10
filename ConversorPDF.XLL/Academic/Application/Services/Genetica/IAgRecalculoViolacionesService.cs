using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Recalcula las violaciones de una corrida ya materializada con el mismo validador y contexto
    /// que el AG, y persiste el desglose en ag_violacion (Fase 1527).
    /// </summary>
    public interface IAgRecalculoViolacionesService
    {
        Task<Result<AgRecalculoResultado>> RecalcularAsync(long idCorrida, AgContexto contexto, CancellationToken ct);
    }
}