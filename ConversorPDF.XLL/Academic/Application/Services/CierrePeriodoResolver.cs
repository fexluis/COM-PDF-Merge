using System;
using System.Globalization;
using System.Text.RegularExpressions;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio de aplicación para resolver y evaluar límites de cierre de períodos académicos.
    /// Resuelve la precedencia canónica (ventana definida del período) vs. fallback histórico (15/12).
    /// </summary>
    public class CierrePeriodoResolver : ICierrePeriodoResolver
    {
        /// <summary>
        /// Día del cierre de semestre para la regla heredada (hallazgo H8 en DocsPdf\AG-2.txt:30).
        /// </summary>
        public const int DiaCierreSemestre = 15;

        /// <summary>
        /// Mes del cierre de semestre para la regla heredada (hallazgo H8 en DocsPdf\AG-2.txt:30).
        /// </summary>
        public const int MesCierreSemestre = 12;

        private static readonly Regex RegexPeriodoAnio = new Regex(@"^(\d{4})-\d{2}$", RegexOptions.Compiled);

        private static readonly string[] FormatosFecha = new[]
        {
            "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "yyyyMMdd"
        };

        /// <summary>
        /// Evalúa si la fecha final de sesiones supera el límite de cierre del período académico.
        /// </summary>
        /// <param name="periodo">Instancia del período seleccionado o null.</param>
        /// <param name="fechaFinal">Fecha final configurada para las sesiones académicas.</param>
        /// <param name="hoy">Fecha base de referencia temporal (para determinismo en pruebas).</param>
        /// <returns>Resultado con el veredicto, fecha límite y mensaje correspondiente.</returns>
        public ResultadoCierrePeriodo Evaluar(Periodo periodo, DateTime fechaFinal, DateTime hoy)
        {
            DateTime fechaFinalSolo = fechaFinal.Date;

            // Prioridad 1: Ventana canónica del período si está configurada y es válida
            if (periodo != null && periodo.TryObtenerVentana(out _, out DateTime finPeriodo))
            {
                bool excede = fechaFinalSolo > finPeriodo.Date;
                string nombre = !string.IsNullOrWhiteSpace(periodo.PeriodoNombre) ? periodo.PeriodoNombre : periodo.Nombre;

                return new ResultadoCierrePeriodo
                {
                    ExcedeLimite = excede,
                    Limite = finPeriodo.Date,
                    LimiteEsVentanaPeriodo = true,
                    Mensaje = excede
                        ? string.Format("La fecha final ({0:dd/MM/yyyy}) supera el cierre del período {1} ({2:dd/MM/yyyy}). ¿Desea continuar de todos modos?", fechaFinalSolo, nombre, finPeriodo.Date)
                        : string.Empty
                };
            }

            // Prioridad 2: Fallback histórico al 15 de diciembre del año resuelto
            int anioFallback = ResolverAnioFallback(periodo, hoy);
            DateTime limiteFallback = new DateTime(anioFallback, MesCierreSemestre, DiaCierreSemestre);
            bool excedeFallback = fechaFinalSolo > limiteFallback;

            return new ResultadoCierrePeriodo
            {
                ExcedeLimite = excedeFallback,
                Limite = limiteFallback,
                LimiteEsVentanaPeriodo = false,
                Mensaje = excedeFallback
                    ? string.Format("La fecha final supera el 15 de diciembre del año {0}. ¿Desea continuar de todos modos?", anioFallback)
                    : string.Empty
            };
        }

        private static int ResolverAnioFallback(Periodo periodo, DateTime hoy)
        {
            if (periodo != null)
            {
                // Intento a: extrae año de PeriodoNombre con formato YYYY-MM
                if (!string.IsNullOrWhiteSpace(periodo.PeriodoNombre))
                {
                    var match = RegexPeriodoAnio.Match(periodo.PeriodoNombre.Trim());
                    if (match.Success && int.TryParse(match.Groups[1].Value, out int anioNombre)
                        && anioNombre >= 1 && anioNombre <= 9999)
                    {
                        return anioNombre;
                    }
                }

                // Intento b: extrae año de PeriodoFin parseable
                if (!string.IsNullOrWhiteSpace(periodo.PeriodoFin))
                {
                    if (DateTime.TryParseExact(periodo.PeriodoFin.Trim(), FormatosFecha, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fechaFin)
                        || DateTime.TryParse(periodo.PeriodoFin.Trim(), out fechaFin))
                    {
                        if (fechaFin.Year >= 1 && fechaFin.Year <= 9999)
                        {
                            return fechaFin.Year;
                        }
                    }
                }
            }

            // Intento c: año del parámetro hoy
            return hoy.Year;
        }
    }
}