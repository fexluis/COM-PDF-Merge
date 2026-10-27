using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Catálogo y utilitario de opciones canónicas de ventana de días (β) para la toma de decisiones.
    /// </summary>
    public static class AgDiasVentana
    {
        public static readonly IReadOnlyDictionary<string, string> Opciones =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "LUN-JUE", "1,2,3,4" },
                { "MAR-VIE", "2,3,4,5" },
                { "JUE-SAB", "4,5,6" },
                { "VIE-SAB", "5,6" },
                { "VIE-DOM", "5,6,7" }
            };

        private static int DiaIso(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            if (int.TryParse(s.Trim(), out int d) && d >= 1 && d <= 7) return d;
            string t = s.Trim().ToUpperInvariant();
            if (t.StartsWith("LUN")) return 1;
            if (t.StartsWith("MAR")) return 2;
            if (t.StartsWith("MIE") || t.StartsWith("MIÉ")) return 3;
            if (t.StartsWith("JUE")) return 4;
            if (t.StartsWith("VIE")) return 5;
            if (t.StartsWith("SAB") || t.StartsWith("SÁB")) return 6;
            if (t.StartsWith("DOM")) return 7;
            return 0;
        }

        public static string AIso(string ventana)
        {
            if (string.IsNullOrWhiteSpace(ventana)) return string.Empty;
            var v = ventana.Trim();
            if (Opciones.TryGetValue(v, out var iso)) return iso;

            // Fase 1021 / Fase 1136: tolerar ISO numérico ('1,2,3,4') o abreviaturas ('LUN,MIE') ya persistidas
            var dias = v.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(DiaIso)
                        .Where(d => d > 0).Distinct().OrderBy(d => d).ToList();
            return dias.Count > 0 ? string.Join(",", dias) : string.Empty;
        }

        public static IReadOnlyList<int> Dias(string ventana)
        {
            string iso = AIso(ventana);
            if (string.IsNullOrWhiteSpace(iso)) return new List<int>();

            return iso.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                      .Select(s => int.TryParse(s, out int d) ? d : 0)
                      .Where(d => d >= 1 && d <= 7)
                      .Distinct()
                      .ToList();
        }
    }
}
