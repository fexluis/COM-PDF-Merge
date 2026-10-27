using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    public interface ISedeLookupRepository
    {
        Task<Result<IReadOnlyCollection<LookupItem>>> GetSedesLookupAsync(CancellationToken ct);
    }
}
