using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Planificador de recurrencias temporales que proyecta un bloque semanal en fechas discretas dentro de la ventana semestral.
    /// </summary>
    public class AgPlanificadorTemporal
    {
        public List<DateTime> GenerarFechasValidas(
            DayOfWeek diaSemana,
            DateTime fechaInicio,
            DateTime fechaFin,
            double duracionBloqueHoras,
            double totalHorasObjetivo,
            Func<DateTime, bool> esFestivoPredicate)
        {
            var fechas = new List<DateTime>();
            double horasAcumuladas = 0.0;
            var cursor = fechaInicio.Date;

            // Avanzar al primer día correspondiente
            while (cursor.DayOfWeek != diaSemana && cursor <= fechaFin.Date)
            {
                cursor = cursor.AddDays(1);
            }

            int safetyGuard = 0;
            while (cursor <= fechaFin.Date && horasAcumuladas < totalHorasObjetivo && safetyGuard < 100)
            {
                safetyGuard++;

                if (esFestivoPredicate == null || !esFestivoPredicate(cursor))
                {
                    fechas.Add(cursor);
                    horasAcumuladas += duracionBloqueHoras;
                }

                cursor = cursor.AddDays(7);
            }

            return fechas;
        }

        /// <summary>
        /// Cuenta las ocurrencias exactas no festivas de un día de la semana ISO (1=Lunes..7=Domingo) dentro de una ventana [ini..fin] (Fase 586 / Fase 590).
        /// </summary>
        public static int ContarOcurrencias(int diaIso, DateTime ini, DateTime fin, Func<DateTime, bool> esFestivo)
        {
            if (diaIso < 1 || diaIso > 7) return 0;   // Fase 590: guarda de rango
            var dow = diaIso == 7 ? DayOfWeek.Sunday : (DayOfWeek)diaIso;
            return ContarOcurrencias(dow, ini, fin, esFestivo);
        }

        /// <summary>
        /// Cuenta las ocurrencias exactas no festivas de un día de la semana dentro de una ventana [ini..fin] (Fase 380).
        /// </summary>
        /// <param name="diaSemana">Día de la semana a contar.</param>
        /// <param name="ini">Fecha inicial de la ventana inclusiva.</param>
        /// <param name="fin">Fecha final de la ventana inclusiva.</param>
        /// <param name="esFestivo">Predicado que determina si una fecha es festiva.</param>
        /// <returns>Número de días hábiles no festivos en que se manifiesta la franja.</returns>
        public static int ContarOcurrencias(DayOfWeek diaSemana, DateTime ini, DateTime fin, Func<DateTime, bool> esFestivo)
        {
            int n = 0;
            var c = ini.Date;
            while (c.DayOfWeek != diaSemana && c <= fin.Date)
            {
                c = c.AddDays(1);
            }

            while (c <= fin.Date)
            {
                if (esFestivo == null || !esFestivo(c))
                {
                    n++;
                }
                c = c.AddDays(7);
            }

            return n;
        }
    }
}
