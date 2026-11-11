using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Reconstruye la cabecera (<c>ag_corrida</c>) y el individuo (<c>ag_individuo</c>) de una corrida que
    /// conserva staging pero perdió sus metadatos genéticos, reutilizando los genes de <c>ag_gen</c> y el
    /// contexto del AG (Fase 1566). No vuelve a ejecutar el algoritmo genético.
    /// </summary>
    public interface IAgReparacionCorridaService
    {
        Task<Result<AgReparacionCorridaResultado>> RepararAsync(long idCorrida, AgContexto contexto, CancellationToken ct);
    }
}
