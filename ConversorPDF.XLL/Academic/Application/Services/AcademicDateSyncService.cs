using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Application.Interfaces;

namespace ConversorPDF.XLL.Academic.Application.Services
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
            var builder = new AcademicDateAdjustmentBuilder
            {
                CalculatedStart = currentStart,
                CalculatedEnd = currentEnd,
                NeedsAdjustment = false
            };

            var listaSesiones = sesiones?.ToList();
            if (listaSesiones == null || !listaSesiones.Any())
            {
                return builder.Build();
            }

            // 1. Obtener los extremos reales de la generación
            var fechasValidas = listaSesiones
                .Where(s => s.FechaSesion > new DateTime(2000, 1, 1))
                .Select(s => s.FechaSesion.Date)
                .ToList();

            if (!fechasValidas.Any())
                return builder.Build();

            DateTime realStart = fechasValidas.Min();
            DateTime realEnd = fechasValidas.Max();

            // 2. Comparar Inicio
            if (realStart != currentStart.Date)
            {
                builder.NeedsAdjustment = true;
                builder.CalculatedStart = realStart;
                builder.AddMessage($"- Fecha Inicio ajustada de {currentStart:dd/MM/yyyy} a {realStart:dd/MM/yyyy}.");
            }

            // 3. Comparar Fin
            if (realEnd != currentEnd.Date)
            {
                builder.NeedsAdjustment = true;
                builder.CalculatedEnd = realEnd;
                builder.AddMessage($"- Fecha Final ajustada de {currentEnd:dd/MM/yyyy} a {realEnd:dd/MM/yyyy}.");

                if (realEnd > currentEnd.Date)
                {
                    builder.ExceededOriginalLimit = true;
                    builder.AddMessage("\n⚠️ AVISO: La fecha final generada ha SUPERADO la fecha límite previamente establecida para poder cumplir con las horas totales.");
                }
            }

            return builder.Build();
        }
    }
}
