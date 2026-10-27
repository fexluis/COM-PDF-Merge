using System;
using System.Linq;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Helper estático para analizar cadenas de semestres (ej. "VIII-", "I")
    /// y convertirlas a su representación entera para facilitar la ordenación.
    /// Extraído para mantener el principio DRY entre las vistas y las exportaciones.
    /// </summary>
    public static class SemestreParser
    {
        public static int Parsear(string semestreStr)
        {
            if (string.IsNullOrWhiteSpace(semestreStr)) return 999;
            string s = semestreStr.Trim().ToUpper();

            if (s.StartsWith("VIII-") || s == "VIII") return 8;
            if (s.StartsWith("VII-") || s == "VII") return 7;
            if (s.StartsWith("VI-") || s == "VI") return 6;
            if (s.StartsWith("V-") || s == "V") return 5;
            if (s.StartsWith("IX-") || s == "IX") return 9;
            if (s.StartsWith("X-") || s == "X") return 10;
            if (s.StartsWith("IV-") || s == "IV") return 4;
            if (s.StartsWith("III-") || s == "III") return 3;
            if (s.StartsWith("II-") || s == "II") return 2;
            if (s.StartsWith("I-") || s == "I") return 1;

            if (int.TryParse(new string(s.TakeWhile(char.IsDigit).ToArray()), out int num))
                return num;

            return 999;
        }
    }
}
