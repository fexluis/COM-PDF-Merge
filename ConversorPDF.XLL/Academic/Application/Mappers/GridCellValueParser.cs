using System;
using System.Globalization;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    /// <summary>
    /// Parser defensivo y desacoplado para la interpretación de celdas de texto
    /// provenientes del DataGridView de sesiones académicas.
    /// Garantiza inmunidad ante discrepancias culturales (comas/puntos) y valida
    /// formatos canónicos de fechas y números sin dependencias de UI.
    /// (Fase 235 / F1.1).
    /// </summary>
    public static class GridCellValueParser
    {
        private static readonly string[] FormatosFecha = new[]
        {
            "dd/MM/yyyy",
            "d/M/yyyy",
            "dd-MM-yyyy",
            "yyyy-MM-dd",
            "d/M/yy",
            "dd/MM/yy",
            "yyyy/MM/dd",
            "yyyyMMdd"
        };

        /// <summary>
        /// Intenta parsear un valor a una cantidad positiva de horas (double) entre 0 y 24.
        /// Soporta tipos numéricos nativos (<see cref="double"/>) o cadenas de texto.
        /// Reemplaza comas decimales por puntos y utiliza <see cref="NumberStyles.Float"/>
        /// con <see cref="CultureInfo.InvariantCulture"/> para evitar que estilos como
        /// <see cref="NumberStyles.Any"/> interpreten la coma como separador de miles en cultura es-CO.
        /// (Fase 235 / F1.1, Fase 2 / F2.5).
        /// </summary>
        /// <param name="value">Objeto, número o texto a parsear (ej. "4.5", "4,5", 4, 4.0d).</param>
        /// <param name="horas">Salida con el número de horas parseado.</param>
        /// <returns><c>true</c> si el parseo fue exitoso y el valor está en el rango (0, 24]; de lo contrario, <c>false</c>.</returns>
        public static bool TryParseHoras(object value, out double horas)
        {
            horas = 0;
            if (value == null) return false;

            if (value is double d)
            {
                if (d > 0 && d <= 24 && !double.IsNaN(d) && !double.IsInfinity(d))
                {
                    horas = d;
                    return true;
                }
                return false;
            }

            string txt = (value as string ?? value.ToString())?.Trim();
            if (string.IsNullOrWhiteSpace(txt)) return false;

            // Normalizar separador decimal local para cultura es-CO/es-ES
            txt = txt.Replace(',', '.');

            return double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out horas)
                   && horas > 0 && horas <= 24;
        }

        /// <summary>
        /// Intenta parsear un valor a <see cref="DateTime"/> usando una lista cerrada de formatos canónicos.
        /// Soporta objetos <see cref="DateTime"/> nativos, descarta componentes horarios ("HH:mm:ss", espacios/tabs/NBSP)
        /// y normaliza años de 2 dígitos a partir del siglo XXI (Fases 693, 696, 698).
        /// </summary>
        /// <param name="value">Objeto o texto con la fecha a evaluar (ej. "28/09/2026", "10/10/26", DateTime).</param>
        /// <param name="fecha">Salida con la fecha resultante.</param>
        /// <returns><c>true</c> si la fecha cumple con alguno de los formatos canónicos; de lo contrario, <c>false</c>.</returns>
        public static bool TryParseFecha(object value, out DateTime fecha)
        {
            fecha = default(DateTime);
            if (value is DateTime dt)
            {
                fecha = dt.Date;
                return true;
            }
            if (value == null) return false;

            string txt = (value as string ?? value.ToString())?.Trim();
            if (string.IsNullOrWhiteSpace(txt)) return false;

            // Fase 698: Descartar componente horario recortando por espacios, tabulaciones o NBSP
            var partes = txt.Split(new[] { ' ', '\t', '\r', '\n', '\u00A0' }, StringSplitOptions.RemoveEmptyEntries);
            if (partes.Length > 0) txt = partes[0];

            bool ok = DateTime.TryParseExact(
                txt,
                FormatosFecha,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out fecha);

            if (!ok) return false;

            // Fase 696: Normalizar años de 2 dígitos resueltos a 19xx por GregorianCalendar.TwoDigitYearMax
            if (fecha.Year < 2000) fecha = fecha.AddYears(100);

            return true;
        }
    }
}
