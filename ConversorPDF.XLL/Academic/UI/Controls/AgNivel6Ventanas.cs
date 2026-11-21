using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.UI.Controls
{
    /// <summary>
    /// Helper de interfaz para el Nivel 6 (Modalidad / Dimensiones) del árbol de contexto académico.
    /// Centraliza las 9 opciones de ventana de días (8 canónicas de AgDiasVentana más el alias ISO terminal
    /// explícito "4,5,6,7"), su normalización ISO para renderizado y su resolución por coordenadas de clic (hit-test).
    /// </summary>
    internal static class AgNivel6Ventanas
    {
        private static readonly string[] OpcionesOrdenadas = new string[]
        {
            "LUN-JUE",
            "LUN-VIE",
            "MAR-VIE",
            "JUE-SAB",
            "VIE-SAB",
            "VIE-DOM",
            "JUE-DOM",
            "MIE-DOM",
            "4,5,6,7" // Fase 2243: Alias ISO terminal explícito seleccionable independientemente de JUE-DOM
        };

        private static readonly IReadOnlyList<string> ClavesSoloLectura = Array.AsReadOnly(OpcionesOrdenadas);

        private static readonly Dictionary<string, string> NombresVisuales = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "LUN-JUE", "Lun-Jue" },
            { "LUN-VIE", "Lun-Vie" },
            { "MAR-VIE", "Mar-Vie" },
            { "JUE-SAB", "Jue-Sab" },
            { "VIE-SAB", "Vie-Sab" },
            { "VIE-DOM", "Vie-Dom" },
            { "JUE-DOM", "Jue-Dom" },
            { "MIE-DOM", "Mie-Dom" },
            { "4,5,6,7", "4,5,6,7" }
        };

        private static readonly Dictionary<string, string> NombresVisualesCortos = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "LUN-JUE", "L-J" },
            { "LUN-VIE", "L-V" },
            { "MAR-VIE", "M-V" },
            { "JUE-SAB", "J-S" },
            { "VIE-SAB", "V-S" },
            { "VIE-DOM", "V-D" },
            { "JUE-DOM", "J-D" },
            { "MIE-DOM", "Mi-D" },
            { "4,5,6,7", "4-6,7" }
        };

        public static IReadOnlyList<string> Claves => ClavesSoloLectura;

        public static int Cantidad => OpcionesOrdenadas.Length;

        /// <summary>
        /// Normaliza una cadena de días o alias eliminando espacios intermedios y separadores redundantes.
        /// Asegura que variantes con espacios como "4, 5, 6, 7" se unifiquen a "4,5,6,7" antes de comparar.
        /// </summary>
        private static string NormalizarToken(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                return string.Empty;
            }

            return string.Join(",", s.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim())
                .Where(p => p.Length > 0));
        }

        /// <summary>
        /// Obtiene el índice (0..8) de la ventana correspondiente comparando primero coincidencia exacta
        /// y luego por ISO canónico (excluyendo el alias explícito '4,5,6,7' para no colapsar 'JUE-DOM').
        /// Retorna -1 si la ventana es nula, vacía o no coincide con ninguna opción canónica.
        /// </summary>
        public static int Indice(string diasVentana)
        {
            string t = NormalizarToken(diasVentana);
            if (t.Length == 0)
            {
                return -1;
            }

            // 1) Coincidencia textual exacta de clave (incluye "4,5,6,7" y "JUE-DOM")
            for (int i = 0; i < OpcionesOrdenadas.Length; i++)
            {
                if (string.Equals(t, OpcionesOrdenadas[i], StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            // 2) Coincidencia por ISO canónico, excluyendo la opción explícita "4,5,6,7" para no colapsar JUE-DOM
            string iso = AgDiasVentana.AIso(t);
            if (string.IsNullOrWhiteSpace(iso))
            {
                return -1;
            }

            for (int i = 0; i < OpcionesOrdenadas.Length; i++)
            {
                if (string.Equals(OpcionesOrdenadas[i], "4,5,6,7", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (AgDiasVentana.Opciones.TryGetValue(OpcionesOrdenadas[i], out string isoCanonica)
                    && string.Equals(iso, isoCanonica, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        /// <summary>
        /// Determina si dos representaciones de ventana de días corresponden al mismo conjunto
        /// de días. Protege la distinción entre '4,5,6,7' y 'JUE-DOM' impidiendo su colapso mutuo por ISO.
        /// </summary>
        public static bool Coincide(string diasVentana, string claveCanonica)
        {
            if (string.IsNullOrWhiteSpace(diasVentana) || string.IsNullOrWhiteSpace(claveCanonica))
            {
                return false;
            }

            string t1 = NormalizarToken(diasVentana);
            string t2 = NormalizarToken(claveCanonica);

            // 1) Coincidencia textual exacta
            if (string.Equals(t1, t2, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // Si cualquiera de las dos es el alias explícito terminal "4,5,6,7", no colapsar por ISO
            if (string.Equals(t1, "4,5,6,7", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(t2, "4,5,6,7", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // 2) Coincidencia por ISO canónico
            string iso1 = AgDiasVentana.AIso(t1);
            string iso2 = AgDiasVentana.AIso(t2);
            if (string.IsNullOrWhiteSpace(iso1) || string.IsNullOrWhiteSpace(iso2))
            {
                return false;
            }

            return string.Equals(iso1, iso2, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Obtiene la clave canónica (ej. "LUN-VIE") correspondiente al valor dado (ej. "1,2,3,4,5").
        /// Si no coincide con ninguna opción conocida, retorna el valor original en mayúsculas recortado.
        /// </summary>
        public static string ClaveCanonica(string diasVentana)
        {
            int idx = Indice(diasVentana);
            if (idx >= 0 && idx < OpcionesOrdenadas.Length)
            {
                return OpcionesOrdenadas[idx];
            }

            return (diasVentana ?? string.Empty).Trim().ToUpperInvariant();
        }

        /// <summary>
        /// Obtiene el texto para renderizar el botón de opción según el paso horizontal disponible.
        /// Si el paso es inferior a 58 píxeles se usa la versión compacta ("L-V", "J-D", etc.).
        /// </summary>
        public static string FormatearRadio(string clave, int paso)
        {
            if (string.IsNullOrWhiteSpace(clave))
            {
                return string.Empty;
            }

            if (paso < 58)
            {
                if (NombresVisualesCortos.TryGetValue(clave, out string corto))
                {
                    return corto;
                }
            }

            if (NombresVisuales.TryGetValue(clave, out string regular))
            {
                return regular;
            }

            return clave;
        }

        /// <summary>
        /// Determina la clave de ventana seleccionada según la posición X del clic dentro del panel.
        /// Protege el hueco muerto posterior al último botón para no seleccionar accidentalmente
        /// la última opción al hacer clic en el espacio vacío anterior al separador.
        /// </summary>
        public static string ResolverPorX(int dx, int[] posiciones, int fin, int step)
        {
            if (posiciones == null || posiciones.Length != OpcionesOrdenadas.Length)
            {
                return null;
            }

            for (int i = 0; i < posiciones.Length; i++)
            {
                int inicio = posiciones[i];
                int finBoton;
                if (i + 1 < posiciones.Length)
                {
                    finBoton = posiciones[i + 1];
                }
                else
                {
                    // Limitar el último botón al paso efectivo o a fin
                    finBoton = (step > 0) ? Math.Min(inicio + step, fin) : fin;
                }

                if (dx >= inicio && dx < finBoton)
                {
                    return OpcionesOrdenadas[i];
                }
            }

            return null;
        }
    }
}
