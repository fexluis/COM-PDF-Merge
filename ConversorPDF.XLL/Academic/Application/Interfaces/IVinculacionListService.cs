using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
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

        /// <summary>
        /// Obtiene las corridas del staging AG listas para el selector (incluye el comodín
        /// "-- Todas las corridas --" y las etiquetas ya formateadas). Fases 1490/1492.
        /// </summary>
        Task<Result<IReadOnlyCollection<CorridaStagingItem>>> GetCorridasStagingAsync(
            int? idPeriodo,
            int? idPrograma,
            string subunidad,
            CancellationToken ct);

        /// <summary>
        /// Resumen de diagnóstico del staging de una corrida con los filtros activos de la UI (Fases 1490/1493).
        /// </summary>
        Task<Result<AgStagingResumen>> GetResumenStagingAsync(
            long idCorrida,
            VinculacionFilterDto filtros,
            CancellationToken ct);

        /// <summary>
        /// Elimina el staging de la corrida indicada (solo MODO AG) y devuelve los conteos borrados (Fase 1517).
        /// </summary>
        Task<Result<(int Sesiones, int Vinculaciones)>> EliminarCorridaStagingAsync(
            long idCorrida,
            CancellationToken ct);
    }
}
