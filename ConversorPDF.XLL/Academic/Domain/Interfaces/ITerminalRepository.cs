using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    public interface ITerminalRepository
    {
        Result<IReadOnlyCollection<Terminal>> GetAll();
        Result Add(Terminal terminal);
        Result Update(Terminal terminal);
        Result Delete(int id);
    }
}
