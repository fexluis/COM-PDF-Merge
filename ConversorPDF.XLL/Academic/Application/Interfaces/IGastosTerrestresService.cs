using System.Collections.Generic;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.DTOs;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    public interface IGastosTerrestresService
    {
        Task<Result<IEnumerable<GastoTerrestreGridDto>>> GetAllAsync();
        Task<Result> UpdateAsync(GastoTerrestreGridDto dto);
        Task<Result> DeleteAsync(int id);
        Task<Result<string>> CrearTrayectoAsync(string territorial, string trayecto, decimal? valor);
    }
}
