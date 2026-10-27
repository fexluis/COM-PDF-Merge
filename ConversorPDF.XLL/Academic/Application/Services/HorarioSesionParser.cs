using System;
using System.Text.RegularExpressions;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio de aplicación para parsear cadenas de horario en estructuras FranjaHoraria.
    /// </summary>
    public static class HorarioSesionParser
    {
        private static readonly Regex RangoRegex = new Regex(@"^(.*?)\s+A\s+(.*?)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Intenta extraer la franja horaria desde una cadena de texto como "6:00 PM A 10:00 PM".
        /// </summary>
        public static Result<FranjaHoraria> TryParse(string horarioText)
        {
            if (string.IsNullOrWhiteSpace(horarioText))
                return Result.Failure<FranjaHoraria>("El texto del horario está vacío.");

            // Limpieza básica
            string cleanedText = horarioText.Trim().ToUpperInvariant();
            
            // Reemplazar variaciones de "12:00 M", "12 M", "12:30 M" por "PM"
            // Se asume que "M" en este contexto significa Mediodía/PM.
            cleanedText = Regex.Replace(cleanedText, @"\b(12:?[0-5]?[0-9]?)\s*M\b", "$1 PM", RegexOptions.IgnoreCase);

            // Buscar el patrón "INICIO A FIN"
            var match = RangoRegex.Match(cleanedText);
            if (!match.Success)
            {
                return Result.Failure<FranjaHoraria>($"Formato no reconocido: '{horarioText}'.");
            }

            string startStr = match.Groups[1].Value.Trim();
            string endStr = match.Groups[2].Value.Trim();

            // Formatos permitidos
            string[] formats = new[] { "h:mm tt", "hh:mm tt", "H:mm", "HH:mm", "h tt", "hh tt" };

            if (!DateTime.TryParseExact(startStr, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime startDt))
                return Result.Failure<FranjaHoraria>($"No se pudo parsear la hora de inicio: '{startStr}'.");

            if (!DateTime.TryParseExact(endStr, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime endDt))
                return Result.Failure<FranjaHoraria>($"No se pudo parsear la hora de fin: '{endStr}'.");

            try
            {
                var franja = new FranjaHoraria(startDt.TimeOfDay, endDt.TimeOfDay);
                return Result.Success(franja);
            }
            catch (ArgumentException ex)
            {
                return Result.Failure<FranjaHoraria>(ex.Message);
            }
        }
    }
}
