using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Materializa la solución óptima de un individuo en registros concretos dentro de las tablas de staging 'ag_vinculaciones' y 'ag_sesiones'.
    /// </summary>
    public interface IAgMaterializadorSesiones
    {
        Task<Result> MaterializarEnStagingAsync(long idCorrida, AgIndividuo individuo, AgContexto contexto, CancellationToken ct);
    }
}
