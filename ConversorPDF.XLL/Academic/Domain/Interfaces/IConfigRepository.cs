using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    public interface IConfigRepository
    {
        Result<Config> GetConfig();
        Result UpdateConfig(Config config);
    }
}
