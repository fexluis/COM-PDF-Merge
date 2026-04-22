using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Models.Academic;

namespace ConversorPDF.XLL.Core.Services.Academic
{
    /// <summary>
    /// Implementación del servicio de sincronización de fechas.
    /// </summary>
    public class AcademicDateSyncService : IAcademicDateSyncService
    {
        /// <inheritdoc />
        public AcademicDateAdjustment CalculateAdjustment(
            IEnumerable<Sesion> sesiones, 
            DateTime currentStart, 
            DateTime currentEnd)
        {
            var result = new AcademicDateAdjustment
            {
                CalculatedStart = currentStart,
                CalculatedEnd = currentEnd,
                NeedsAdjustment = false
            };

            var listaSesiones = sesiones?.ToList();
            if (listaSesiones == null || !listaSesiones.Any())
            {
                return result;
            }

            // 1. Obtener los extremos reales de la generación
            DateTime realStart = listaSesiones.Min(s => s.fechasesionInicial).Date;
            DateTime realEnd = listaSesiones.Max(s => s.fechasesionFinal).Date;

            // 2. Comparar Inicio
            if (realStart != currentStart.Date)
            {
                result.NeedsAdjustment = true;
                result.CalculatedStart = realStart;
                result.Messages.Add($"- Fecha Inicio ajustada de {currentStart:dd/MM/yyyy} a {realStart:dd/MM/yyyy}.");
            }

            // 3. Comparar Fin
            if (realEnd != currentEnd.Date)
            {
                result.NeedsAdjustment = true;
                result.CalculatedEnd = realEnd;
                result.Messages.Add($"- Fecha Final ajustada de {currentEnd:dd/MM/yyyy} a {realEnd:dd/MM/yyyy}.");

                if (realEnd > currentEnd.Date)
                {
                    result.ExceededOriginalLimit = true;
                    result.Messages.Add("\n⚠️ AVISO: La fecha final generada ha SUPERADO la fecha límite previamente establecida para poder cumplir con las horas totales.");
                }
            }

            return result;
        }
    }
}
