using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>Lee origen productivo para reconciliación, sin escribir ni acreditar aprobación implícita.</summary>
    public interface IAgPlanAprobadoRepository
    {
        /// <summary>Obtiene todas las vinculaciones/sesiones del alcance, sin corte temporal ni paginación parcial.
        /// Fallos técnicos y cancelación son Failure; ausencia y ambigüedad se conservan en el snapshot.</summary>
        Task<Result<AgPlanAprobadoSnapshot>> ObtenerAsync(AgManifiestoCobertura manifiesto, CancellationToken ct);
    }
}
