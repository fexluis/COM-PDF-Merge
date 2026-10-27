using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    /// <summary>
    /// Servicio de aplicación para la gestión de la lista de vinculaciones.
    /// </summary>
    public interface IVinculacionListService
    {
        /// <summary>
        /// Obtiene la lista de vinculaciones filtrada.
        /// </summary>
        Task<Result<IReadOnlyCollection<SesionListItem>>> GetListaSesionesAsync(
            VinculacionFilterDto filtros, 
            CancellationToken ct);

        /// <summary>
        /// Obtiene los datos para los combos de filtrado.
        /// </summary>
        Task<Result<(IReadOnlyCollection<LookupItem> Sedes, IReadOnlyCollection<LookupItem> Terceros, IReadOnlyCollection<LookupItem> Periodos, IReadOnlyCollection<LookupItem> Programas)>> GetCatalogosFiltroAsync(CancellationToken ct);
    }
}
