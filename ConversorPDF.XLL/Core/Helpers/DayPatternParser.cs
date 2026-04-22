using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Utilidad para el parseo y traducción de patrones de días seleccionados (Español -> Inglés).
    /// </summary>
    public static class DayPatternParser
    {
        public static Dictionary<DayOfWeek, double> Parse(Dictionary<string, int> input)
        {
            var result = new Dictionary<DayOfWeek, double>();
            if (input == null || !input.Any()) return result;

            foreach (var kvp in input)
            {
                string diaNombreStr = TranslateDay(kvp.Key.Trim());
                if (Enum.TryParse<DayOfWeek>(diaNombreStr, true, out var day))
                {
                    result[day] = kvp.Value;
                }
            }
            return result;
        }

        public static string TranslateDay(string day)
        {
            if (string.IsNullOrWhiteSpace(day)) return string.Empty;

            switch (day.ToUpper().Trim())
            {
                case "LUNES": return "Monday";
                case "MARTES": return "Tuesday";
                case "MIERCOLES":
                case "MIÉRCOLES": return "Wednesday";
                case "JUEVES": return "Thursday";
                case "VIERNES": return "Friday";
                case "SABADO":
                case "SÁBADO": return "Saturday";
                case "DOMINGO": return "Sunday";
                default: return day;
            }
        }
    }
}
