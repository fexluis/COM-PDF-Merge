using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Formateador canónico de la columna K (días y horas de sesión por semana) (Fase 941 / Fase 991 / Fase 996).
    /// Genera la representación canónica de 3 partes "DÍA|horas|horario" compatible con RebuildSelectedDays (ej. "JUEVES|4|18:00-22:00").
    /// </summary>
    public static class DiasHorasSemanaFormatter
    {
        public static string DesdeSesiones(IEnumerable<Sesion> sesiones)
        {
            if (sesiones == null) return string.Empty;

            var pares = new List<(int orden, string texto)>();
            var visitados = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var s in sesiones)
            {
                if (s == null) continue;

                string dia = s.FechaSesion != DateTime.MinValue && s.FechaSesion != default(DateTime)
                    ? ObtenerNombreDia(s.FechaSesion.DayOfWeek)
                    : null;

                if (string.IsNullOrWhiteSpace(dia) && !string.IsNullOrWhiteSpace(s.HorarioSesion))
                {
                    var partes = s.HorarioSesion.Split('|');
                    dia = NormalizarDia(partes[0]);
                }

                if (string.IsNullOrWhiteSpace(dia)) continue;

                string horario = (s.HorarioSesion ?? string.Empty).Trim();
                double horas = s.Horas > 0 ? s.Horas : 4.0;

                // Fase 991 / Fase 995: contrato canónico de 3 partes (DIA|horas|HORARIO) con formato invariante
                string clave = $"{dia}|{horas.ToString("0.##", CultureInfo.InvariantCulture)}|{horario}";

                if (visitados.Add(clave))
                {
                    pares.Add((ObtenerOrdenDia(dia), clave));
                }
            }

            if (pares.Count == 0) return string.Empty;

            // Fase 997: separador ", " acorde al contrato de RebuildSelectedDays / DayHoursDropdown
            return string.Join(", ", pares.OrderBy(p => p.orden).Select(p => p.texto));
        }

        /// <summary>
        /// Formatea la columna K para elementos SesionListItem agrupados en el exportador Excel (Fase 1136).
        /// Emite pares canónicos "DÍA|horas|horario" ordenados cronológicamente por día de la semana.
        /// </summary>
        /// <param name="items">Colección de sesiones representadas como SesionListItem.</param>
        /// <returns>Cadena formateada con días, horas y horario separados por comas.</returns>
        public static string DesdeSesionListItems(IEnumerable<SesionListItem> items)
        {
            if (items == null) return string.Empty;

            var pares = new List<(int orden, string texto)>();
            var visitados = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var s in items)
            {
                if (s == null) continue;

                string dia = !string.IsNullOrWhiteSpace(s.DiaSemana)
                    ? NormalizarDia(s.DiaSemana)
                    : (s.FechaSesion.HasValue ? ObtenerNombreDia(s.FechaSesion.Value.DayOfWeek) : string.Empty);

                if (string.IsNullOrWhiteSpace(dia)) continue;

                double horas = s.Horas > 0 ? s.Horas : 4.0;
                string horario = (s.HorarioSesion ?? string.Empty).Trim();
                string clave = $"{dia}|{horas.ToString("0.##", CultureInfo.InvariantCulture)}|{horario}";

                if (visitados.Add(clave))
                {
                    pares.Add((ObtenerOrdenDia(dia), clave));
                }
            }

            if (pares.Count == 0) return string.Empty;

            return string.Join(", ", pares.OrderBy(p => p.orden).Select(p => p.texto));
        }

        // Fase 996: delegar en DiaSemanaNormalizador para emitir nombres canónicos completos ("LUNES", ...)
        private static string NormalizarDia(string dia)
        {
            int orden = ObtenerOrdenDia(dia);
            return orden < 99 ? DiaSemanaNormalizador.ObtenerNombreDia(orden) : string.Empty;
        }

        private static string ObtenerNombreDia(DayOfWeek dow)
            => DiaSemanaNormalizador.ObtenerNombreDia(DiaSemanaNormalizador.DayOfWeekANumero(dow));

        private static int ObtenerOrdenDia(string dia)
        {
            if (string.IsNullOrWhiteSpace(dia)) return 99;
            string d = dia.Trim().ToUpperInvariant();
            if (d.StartsWith("LUN")) return 1;
            if (d.StartsWith("MAR")) return 2;
            if (d.StartsWith("MIE") || d.StartsWith("MIÉ")) return 3;
            if (d.StartsWith("JUE")) return 4;
            if (d.StartsWith("VIE")) return 5;
            if (d.StartsWith("SAB") || d.StartsWith("SÁB")) return 6;
            if (d.StartsWith("DOM")) return 7;
            return 99;
        }
    }
}
