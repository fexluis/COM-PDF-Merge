using System.Collections.Generic;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    public interface ISalvedadConflictoRepository
    {
        Task<IReadOnlyCollection<SalvedadConflicto>> GetByPeriodoAsync(int idPeriodo);
        Task<Result> SaveAsync(SalvedadConflicto salvedad);
        Task<Result> DeleteAsync(int idSesionA, int idSesionB);
    }
}
