using System;
using System.Collections.Generic;
using System.Globalization;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Utilitario canónico para numeración y agrupamiento de semanas ISO 8601 (Fase 950 / Fase 958).
    /// Compatible con .NET Framework 4.8 (sin System.Globalization.ISOWeek).
    /// </summary>
    public static class AgSemanaIso
    {
        public static IReadOnlyList<int> DesdeVentana(string diasVentana)
        {
            return AgDiasVentana.Dias(diasVentana);
        }

        /// <summary>
        /// Retorna el día de la semana en formato ISO (Lunes=1 .. Domingo=7) (Fase 1029).
        /// </summary>
        public static int DiaIso(DateTime fecha)
        {
            return fecha.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)fecha.DayOfWeek;
        }

        // Fase 1033: se elimina el alias De(DateTime); colisionaba en el registro
        // automático de Excel-DNA 1.8.0 con AgClaveCohorte.De(...) (registro por nombre).

        public static int Clave(DateTime fecha)
        {
            // Algoritmo estándar ISO 8601 compatible con .NET Framework 4.8
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(fecha);
            DateTime f = (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday) ? fecha.AddDays(3) : fecha;
            int year = f.Year;
            int week = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(f, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return year * 100 + week;
        }
    }
}
