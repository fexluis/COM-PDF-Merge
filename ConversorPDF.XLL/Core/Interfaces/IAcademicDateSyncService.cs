using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Core.Models.Academic;

namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Servicio para sincronizar y validar rangos de fechas académicos.
    /// </summary>
    public interface IAcademicDateSyncService
    {
        /// <summary>
        /// Calcula el ajuste necesario para los controles de fecha basados en las sesiones generadas.
        /// </summary>
        AcademicDateAdjustment CalculateAdjustment(
            IEnumerable<Sesion> sesiones, 
            DateTime currentStart, 
            DateTime currentEnd);
    }
}
