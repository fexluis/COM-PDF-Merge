using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.DTOs;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    /// <summary>
    /// Contrato de servicio para operaciones de consulta, creación, actualización y eliminación de períodos académicos.
    /// </summary>
    public interface IPeriodoCrudService
    {
        Result<IEnumerable<PeriodoGridDto>> ObtenerTodos();
        Result<PeriodoGridDto> ObtenerPorId(int idPeriodo);
        Result<int> Guardar(PeriodoGridDto dto);
        Result Eliminar(int idPeriodo);
    }
}
