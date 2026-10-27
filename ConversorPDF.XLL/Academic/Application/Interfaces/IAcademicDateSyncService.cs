using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    /// <summary>
    /// Interfaz para el servicio de sincronización de fechas académicas.
    /// </summary>
    public interface IAcademicDateSyncService
    {
        /// <summary>
        /// Calcula el ajuste de fechas basado en las sesiones generadas.
        /// </summary>
        AcademicDateAdjustment CalculateAdjustment(IEnumerable<Sesion> sesiones, DateTime currentStart, DateTime currentEnd);
    }
}
