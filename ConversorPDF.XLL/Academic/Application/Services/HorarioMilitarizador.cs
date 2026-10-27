using System;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Helper puro e idempotente para normalizar cadenas de horario (12h AM/PM o militar) al formato canónico "HH:mm A HH:mm".
    /// </summary>
    public static class HorarioMilitarizador
    {
        /// <summary>
        /// Normaliza un rango horario a formato militar "HH:mm A HH:mm" (Fase 660: renombrado para evitar colisión UDF).
        /// Si no puede parsearse o está vacío, retorna la cadena original limpia o vacía.
        /// </summary>
        public static string NormalizarHorario(string horario)
        {
            if (string.IsNullOrWhiteSpace(horario)) return string.Empty;
            var r = HorarioSesionParser.TryParse(horario);
            if (r.IsFailure) return horario.Trim();

            var f = r.Value;
            // En TimeSpan 'hh' representa horas de 00 a 23 (formato 24h militar); 'HH' no existe en TimeSpan y lanza FormatException
            return $"{f.Inicio:hh\\:mm} A {f.Fin:hh\\:mm}";
        }
    }
}
