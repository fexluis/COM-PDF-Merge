using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Escalera determinista de ampliación de ventana semanal (Fase 1228 / Fase 1711).
    /// VIE-SAB → JUE-SAB → "4,5,6,7"; LUN-JUE/MAR-VIE → LUN-VIE; VIE-DOM → JUE-DOM → MIE-DOM.
    /// Máximo 2 pasos; idempotente.
    /// Internal para prevenir registro espurio de UDFs en Excel-DNA 1.8.0.
    /// </summary>
    internal static class AgEscaleraVentanas
    {
        public const int MaxPasos = 2;

        private static readonly Dictionary<string, string[]> Escalones =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                { "LUN-JUE",    new[] { "LUN-VIE" } },
                { "MAR-VIE",    new[] { "LUN-VIE" } },
                { "LUN-VIE",    new string[0] },
                { "VIE-SAB",    new[] { "JUE-SAB", "4,5,6,7" } },
                { "JUE-SAB",    new[] { "4,5,6,7" } },
                { "4,5,6,7",    new string[0] },
                { "VIE-DOM",    new[] { "JUE-DOM", "MIE-DOM" } },
                { "JUE-DOM",    new[] { "MIE-DOM" } },
                { "MIE-DOM",    new string[0] },
                { "3,4,5,6,7",  new string[0] },
                // Alias ISO de legado
                { "1,2,3,4",    new[] { "LUN-VIE" } },
                { "2,3,4,5",    new[] { "LUN-VIE" } },
                { "1,2,3,4,5",  new string[0] },
                { "5,6",        new[] { "4,5,6" } },
                { "4,5,6",      new[] { "4,5,6,7" } },
                { "5,6,7",      new[] { "JUE-DOM" } }
            };

        /// <summary>Normaliza a la etiqueta canónica del catálogo o al ISO 'd,d,...'.</summary>
        public static string Normalizar(string diasVentana)
        {
            if (string.IsNullOrWhiteSpace(diasVentana))
            {
                return string.Empty;
            }
            string texto = diasVentana.Trim();
            foreach (var opcion in AgDiasVentana.Opciones)
            {
                if (string.Equals(opcion.Key, texto, StringComparison.OrdinalIgnoreCase))
                {
                    return opcion.Key;
                }
            }
            return AgDiasVentana.AIso(texto);
        }

        /// <summary>Pasos siguientes sugeridos (máximo MaxPasos), sin repetir el punto de partida.</summary>
        public static IReadOnlyList<string> Escalera(string diasVentana)
        {
            var pasos = new List<string>();
            string actual = Normalizar(diasVentana);
            for (int i = 0; i < MaxPasos && !string.IsNullOrEmpty(actual); i++)
            {
                if (!Escalones.TryGetValue(actual, out var siguientes) || siguientes.Length == 0)
                {
                    break;
                }
                string paso = siguientes[0];
                if (string.Equals(paso, Normalizar(diasVentana), StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
                if (pasos.Contains(paso))
                {
                    break;
                }
                pasos.Add(paso);
                actual = paso;
            }
            return pasos;
        }

        /// <summary>Primer escalón disponible ('' si la ventana ya está en el tope).</summary>
        public static string Siguiente(string diasVentana)
        {
            var pasos = Escalera(diasVentana);
            return pasos.Count > 0 ? pasos[0] : string.Empty;
        }

        /// <summary>Indica si la ventana admite escalamiento adicional.</summary>
        public static bool PuedeEscalar(string diasVentana) => Escalera(diasVentana).Count > 0;

        /// <summary>Rótulo formateado de ranuras de la ventana respetando franja 6M/6T si corresponde.</summary>
        internal static string Rotular(string diasVentana, double unidad)
        {
            return string.Join(",", AgRanuras
                .PorVentana(diasVentana, unidad: unidad)
                .Select(r => r.Dia == 6 && r.Horas < 8
                    ? (string.Equals(r.Franja, AgFranjaJornada.MANANA, StringComparison.OrdinalIgnoreCase) ? "6M" : "6T")
                    : r.Dia.ToString()));
        }
    }
}
