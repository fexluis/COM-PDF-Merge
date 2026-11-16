using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Promueve una solución validada de staging a las tablas productivas ('vinculaciones', 'sesiones', 'salvedades_conflictos').
    /// Aplica compuerta de seguridad estricta: W1=0 y HardViolations=0, validando la cobertura contra el plan/contexto sellado.
    /// </summary>
    public interface IAgPromotorVinculaciones
    {
        Task<Result<AgPromocion>> PromoverAProduccionAsync(long idCorrida, long idIndividuo, string usuario, CancellationToken ct);
        Task<Result<AgPromocion>> PromoverAProduccionAsync(long idCorrida, long idIndividuo, string usuario, string justificacionVacante, CancellationToken ct);
        Task<Result<AgPromocion>> PromoverAProduccionAsync(long idCorrida, long idIndividuo, string usuario, string justificacionVacante, AgContexto contexto, CancellationToken ct);
    }
}
