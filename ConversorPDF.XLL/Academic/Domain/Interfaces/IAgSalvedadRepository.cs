using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Repositorio para la gestión de salvedades de traslados/conflictos en la capa de staging del AG.
    /// </summary>
    public interface IAgSalvedadRepository
    {
        Task<Result<IReadOnlyCollection<AgSalvedad>>> GetByCorridaAsync(long idCorrida, CancellationToken ct);
        Task<Result> SaveAsync(AgSalvedad salvedad, CancellationToken ct);
        Task<Result> DeleteAsync(long idCorrida, int idSesionA, int idSesionB, CancellationToken ct);
    }
}
