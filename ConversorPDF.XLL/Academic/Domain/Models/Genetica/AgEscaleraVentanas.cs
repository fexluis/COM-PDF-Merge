using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Escalera determinista de ampliación de ventana semanal (Fase 1228).
    /// VIE-SAB → JUE-SAB → "4,5,6,7"; máximo 2 pasos; idempotente.
    /// Internal para prevenir registro espurio de UDFs en Excel-DNA 1.8.0.
    /// </summary>
    internal static class AgEscaleraVentanas
    {
        public const int MaxPasos = 2;

        private static readonly Dictionary<string, string[]> Escalones =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                { "VIE-SAB",  new[] { "JUE-SAB", "4,5,6,7" } },
                { "JUE-SAB",  new[] { "4,5,6,7" } },
                { "MAR-VIE",  new[] { "1,2,3,4,5" } },   // Fase 1262: sin escalones laterales
                { "LUN-JUE",  new[] { "1,2,3,4,5" } },
                { "VIE-DOM",  new string[0] },
                { "4,5,6,7",  new string[0] }
            };

        /// <summary>Normaliza a la etiqueta canónica del catálogo o al ISO 'd,d,...'.</summary>
        public static string Normalizar(string diasVentana)
        {
            if (string.IsNullOrWhiteSpace(diasVentana)) return string.Empty;
            string texto = diasVentana.Trim();
            foreach (var opcion in AgDiasVentana.Opciones)
            {
                if (string.Equals(opcion.Key, texto, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(opcion.Value, texto, StringComparison.OrdinalIgnoreCase))
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
                if (!Escalones.TryGetValue(actual, out var siguientes) || siguientes.Length == 0) break;
                string paso = siguientes[0];
                if (string.Equals(paso, Normalizar(diasVentana), StringComparison.OrdinalIgnoreCase)) break;
                if (pasos.Contains(paso)) break;
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

        public static bool PuedeEscalar(string diasVentana) => Escalera(diasVentana).Count > 0;

        internal static string Rotular(string diasVentana, double unidad)
        {
            throw new NotImplementedException();
        }
    }
}
