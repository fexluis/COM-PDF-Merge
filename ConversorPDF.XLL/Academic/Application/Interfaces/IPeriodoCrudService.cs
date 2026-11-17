using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;
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

        /// <summary>
        /// Extiende la ventana temporal del período de manera controlada y auditada (Fase 1854).
        /// Solo permite extensión (nuevoFin &gt;= finActual y nuevoInicio &lt;= inicioActual) con tope de 2 años.
        /// </summary>
        Result<Periodo> ExtenderVentanaFin(int idPeriodo, DateTime nuevoFin, string actor, string origen, string motivo = null, DateTime? nuevoInicio = null);
    }
}
