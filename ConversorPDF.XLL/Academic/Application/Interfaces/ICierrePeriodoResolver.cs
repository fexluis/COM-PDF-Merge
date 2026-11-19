using System;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    /// <summary>
    /// Contrato para la evaluación de fechas límite de períodos académicos sin dependencias de UI.
    /// </summary>
    public interface ICierrePeriodoResolver
    {
        /// <summary>
        /// Evalúa si la fecha final de sesiones supera el límite de cierre del período.
        /// </summary>
        /// <param name="periodo">Instancia del período seleccionado o null.</param>
        /// <param name="fechaFinal">Fecha final configurada para las sesiones académicas.</param>
        /// <param name="hoy">Fecha base de referencia temporal (para determinismo en tests).</param>
        /// <returns>Resultado con el veredicto, fecha límite y mensaje correspondiente.</returns>
        ResultadoCierrePeriodo Evaluar(Periodo periodo, DateTime fechaFinal, DateTime hoy);
    }
}