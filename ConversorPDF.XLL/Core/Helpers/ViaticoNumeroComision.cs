using System;
using System.Text.RegularExpressions;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Utilidad para interpretar y comparar números de comisión de viáticos bajo la convención '#YY'.
    /// Permite evaluar rangos de comisiones por tupla (año, consecutivo).
    /// </summary>
    internal static class ViaticoNumeroComision
    {
        internal const int LongitudMinimaConsecutivoAnio = 3;
        private const int AnioVigenciaPorDefecto = 2026;

        /// <summary>
        /// Determina si los límites dados representan un filtro activo.
        /// </summary>
        /// <param name="inicio">Límite inicial.</param>
        /// <param name="fin">Límite final.</param>
        /// <returns>True si ambos límites son mayores a cero; false en caso contrario.</returns>
        internal static bool HayFiltro(int inicio, int fin)
        {
            return inicio > 0 && fin > 0;
        }

        /// <summary>
        /// Intenta interpretar una cadena de comisión en sus componentes de año y consecutivo.
        /// Si la cadena tiene al menos 3 dígitos y termina en un sufijo de año de 2 dígitos (YY), extrae año (2000 + YY) y consecutivo.
        /// Si tiene menos de 3 dígitos, se interpreta como consecutivo del año de vigencia actual (App.APP_VIGENCIA).
        /// </summary>
        /// <param name="valor">Texto o número a interpretar.</param>
        /// <param name="anio">Año resultante de cuatro dígitos.</param>
        /// <param name="consecutivo">Número consecutivo resultante.</param>
        /// <returns>True si pudo interpretarse válidamente; false si es nulo, vacío o no numérico.</returns>
        internal static bool TryParse(string valor, out int anio, out int consecutivo)
        {
            anio = 0;
            consecutivo = 0;

            if (string.IsNullOrWhiteSpace(valor))
            {
                return false;
            }

            string limpio = valor.Trim();
            if (!Regex.IsMatch(limpio, @"^\d+$"))
            {
                return false;
            }

            int anioVigenciaActual = AnioVigenciaPorDefecto;
            if (int.TryParse(App.APP_VIGENCIA, out int parsedVigencia) && parsedVigencia > 2000)
            {
                anioVigenciaActual = parsedVigencia;
            }

            if (limpio.Length < LongitudMinimaConsecutivoAnio)
            {
                if (!int.TryParse(limpio, out consecutivo) || consecutivo <= 0)
                {
                    return false;
                }
                anio = anioVigenciaActual;
                return true;
            }

            string sufijo = limpio.Substring(limpio.Length - 2);
            string prefijo = limpio.Substring(0, limpio.Length - 2);

            if (int.TryParse(sufijo, out int yy) && int.TryParse(prefijo, out int c) && c > 0)
            {
                anio = 2000 + yy;
                consecutivo = c;
                return true;
            }

            if (int.TryParse(limpio, out int consPuro) && consPuro > 0)
            {
                anio = anioVigenciaActual;
                consecutivo = consPuro;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Intenta interpretar un límite de rango (inicio o fin).
        /// Si el límite parece un consecutivo simple (ej. 1, 45, 145), se interpreta como consecutivo del año de vigencia.
        /// Si tiene longitud y formato de consecutivo con año (termina en el sufijo vigente o posee >= 5 dígitos), se parsea como '#YY'.
        /// </summary>
        /// <param name="limite">Valor entero del límite.</param>
        /// <param name="anio">Año asignado.</param>
        /// <param name="consecutivo">Consecutivo asignado.</param>
        /// <returns>True si el límite es válido.</returns>
        internal static bool TryParseLimite(int limite, out int anio, out int consecutivo)
        {
            anio = 0;
            consecutivo = 0;
            if (limite <= 0)
            {
                return false;
            }

            string sLimite = limite.ToString();
            int anioVigenciaActual = AnioVigenciaPorDefecto;
            if (int.TryParse(App.APP_VIGENCIA, out int parsedVigencia) && parsedVigencia > 2000)
            {
                anioVigenciaActual = parsedVigencia;
            }
            string vigenciaSuffix = (anioVigenciaActual % 100).ToString("D2");

            bool pareceConsecutivoAnio = sLimite.Length >= 3 &&
                                         (sLimite.EndsWith(vigenciaSuffix) || sLimite.Length >= 5);

            if (!pareceConsecutivoAnio)
            {
                anio = anioVigenciaActual;
                consecutivo = limite;
                return true;
            }

            return TryParse(sLimite, out anio, out consecutivo);
        }

        /// <summary>
        /// Evalúa si un valor de comisión se encuentra dentro del rango inclusivo [inicio, fin].
        /// Normaliza rangos invertidos (inicio > fin). Si no hay filtro activo o los límites no son interpretables, retorna true.
        /// </summary>
        /// <param name="valor">Número o texto de la comisión a verificar.</param>
        /// <param name="inicio">Límite inicial del rango.</param>
        /// <param name="fin">Límite final del rango.</param>
        /// <returns>True si está dentro del rango o si no hay filtro; false en caso contrario.</returns>
        internal static bool EnRango(string valor, int inicio, int fin)
        {
            if (!HayFiltro(inicio, fin))
            {
                return true;
            }

            if (!TryParse(valor, out int anioVal, out int consVal))
            {
                return false;
            }

            if (!TryParseLimite(inicio, out int aIni, out int cIni) ||
                !TryParseLimite(fin, out int aFin, out int cFin))
            {
                return true;
            }

            if (aIni > aFin || (aIni == aFin && cIni > cFin))
            {
                int tempA = aIni;
                int tempC = cIni;
                aIni = aFin;
                cIni = cFin;
                aFin = tempA;
                cFin = tempC;
            }

            bool mayorOIgualMin = (anioVal > aIni) || (anioVal == aIni && consVal >= cIni);
            bool menorOIgualMax = (anioVal < aFin) || (anioVal == aFin && consVal <= cFin);

            return mayorOIgualMin && menorOIgualMax;
        }
    }
}
