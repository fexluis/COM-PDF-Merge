using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    public interface IAgFiltroSeleccionRepository
    {
        Result<List<int>> Obtener(string subunidad, string tipo);
        Task<Result<List<int>>> ObtenerAsync(string subunidad, string tipo, CancellationToken ct = default);
        
        Result Reemplazar(string subunidad, string tipo, IEnumerable<int> refIds, AgDimensionesSeleccion dims = null);
        Task<Result> ReemplazarAsync(string subunidad, string tipo, IEnumerable<int> refIds, AgDimensionesSeleccion dims = null, CancellationToken ct = default);

        Result Limpiar(string subunidad, string tipo);
        Task<Result> LimpiarAsync(string subunidad, string tipo, CancellationToken ct = default);

        Task<Result<AgFiltrosSeleccion>> ObtenerFiltrosAsync(string subunidad, CancellationToken ct = default);
        Task<Result> GuardarFiltrosAsync(AgFiltrosSeleccion filtros, string subunidad, CancellationToken ct = default);
    }
}
