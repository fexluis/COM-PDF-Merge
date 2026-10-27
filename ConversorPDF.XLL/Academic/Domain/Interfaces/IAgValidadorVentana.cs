using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Validador pre-vuelo (pre-flight) de la ventana del semestre y suficiencia de recursos antes de iniciar el AG.
    /// </summary>
    public interface IAgValidadorVentana
    {
        Task<Result<InformeFactibilidad>> ValidarFactibilidadAsync(AgContexto contexto, CancellationToken ct);
        Result<InformeFactibilidad> ValidarFactibilidad(AgContexto contexto);
    }
}
