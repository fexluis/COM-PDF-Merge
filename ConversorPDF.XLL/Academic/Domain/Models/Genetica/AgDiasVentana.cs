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
        /// <summary>Opciones canónicas de ventanas de días y sus representaciones ISO.</summary>
        public static readonly IReadOnlyDictionary<string, string> Opciones =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "LUN-JUE", "1,2,3,4" },
                { "LUN-VIE", "1,2,3,4,5" },
                { "MAR-VIE", "2,3,4,5" },
                { "JUE-SAB", "4,5,6" },
                { "VIE-SAB", "5,6" },
                { "VIE-DOM", "5,6,7" },
                { "JUE-DOM", "4,5,6,7" },
                { "MIE-DOM", "3,4,5,6,7" }
            };

        private static int DiaIso(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                return 0;
            }
            string t = s.Trim().ToUpperInvariant();
            if (t == "6M" || t == "6T")
            {
                return 0;
            }
            if (int.TryParse(t, out int d) && d >= 1 && d <= 7)
            {
                return d;
            }
            if (t == "LUN" || t == "LUNES")
            {
                return 1;
            }
            if (t == "MAR" || t == "MARTES")
            {
                return 2;
            }
            if (t == "MIE" || t == "MIÉ" || t == "MIERCOLES" || t == "MIÉRCOLES")
            {
                return 3;
            }
            if (t == "JUE" || t == "JUEVES")
            {
                return 4;
            }
            if (t == "VIE" || t == "VIERNES")
            {
                return 5;
            }
            if (t == "SAB" || t == "SÁB" || t == "SABADO" || t == "SÁBADO")
            {
                return 6;
            }
            if (t == "DOM" || t == "DOMINGO")
            {
                return 7;
            }
            return 0;
        }

        /// <summary>Convierte la ventana de días a su representación ISO delimitada por comas.</summary>
        public static string AIso(string ventana)
        {
            if (string.IsNullOrWhiteSpace(ventana))
            {
                return string.Empty;
            }
            var v = ventana.Trim();
            if (Opciones.TryGetValue(v, out var iso))
            {
                return iso;
            }

            // Rango con guión AAA-BBB (no envolvente: inicio <= fin)
            if (v.Contains("-"))
            {
                var partes = v.Split('-');
                if (partes.Length == 2)
                {
                    int ini = DiaIso(partes[0]);
                    int fin = DiaIso(partes[1]);
                    if (ini > 0 && fin > 0)
                    {
                        if (ini <= fin)
                        {
                            var rango = Enumerable.Range(ini, fin - ini + 1);
                            return string.Join(",", rango);
                        }
                        return string.Empty; // Invertido / circular sin envolvente
                    }
                    return string.Empty; // Extremos inválidos
                }
            }

            // Fase 1021 / Fase 1136: tolerar ISO numérico ('1,2,3,4') o abreviaturas ('LUN,MIE') ya persistidas
            var dias = v.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(DiaIso)
                        .Where(d => d > 0).Distinct().OrderBy(d => d).ToList();
            return dias.Count > 0 ? string.Join(",", dias) : string.Empty;
        }

        /// <summary>Obtiene la lista de enteros ISO de los días que componen la ventana.</summary>
        public static IReadOnlyList<int> Dias(string ventana)
        {
            string iso = AIso(ventana);
            if (string.IsNullOrWhiteSpace(iso))
            {
                return new List<int>();
            }

            return iso.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                      .Select(s => int.TryParse(s, out int d) ? d : 0)
                      .Where(d => d >= 1 && d <= 7)
                      .Distinct()
                      .ToList();
        }
    }
}
