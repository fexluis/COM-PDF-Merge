using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio de normalización y conversión de días de la semana.
    /// Garantiza la eliminación sistemática de tildes (p. ej. MIÉRCOLES -> MIERCOLES, SÁBADO -> SABADO),
    /// remoción de marcas diacríticas combinantes Unicode (FormD), colapso de espacios/NBSP,
    /// consistencia en el esquema de base de datos γ y traducción bidireccional con <see cref="DayOfWeek"/>.
    /// </summary>
    public static class DiaSemanaNormalizador
    {
        private static readonly Dictionary<string, string> DiasSinTilde = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "LUNES", "LUNES" },
            { "MARTES", "MARTES" },
            { "MIERCOLES", "MIERCOLES" },
            { "JUEVES", "JUEVES" },
            { "VIERNES", "VIERNES" },
            { "SABADO", "SABADO" },
            { "DOMINGO", "DOMINGO" }
        };

        private static readonly Dictionary<DayOfWeek, string> DayOfWeekToNombre = new Dictionary<DayOfWeek, string>
        {
            { DayOfWeek.Monday, "LUNES" },
            { DayOfWeek.Tuesday, "MARTES" },
            { DayOfWeek.Wednesday, "MIERCOLES" },
            { DayOfWeek.Thursday, "JUEVES" },
            { DayOfWeek.Friday, "VIERNES" },
            { DayOfWeek.Saturday, "SABADO" },
            { DayOfWeek.Sunday, "DOMINGO" }
        };

        private static readonly Dictionary<string, DayOfWeek> NombreToDayOfWeek = new Dictionary<string, DayOfWeek>(StringComparer.OrdinalIgnoreCase)
        {
            { "LUNES", DayOfWeek.Monday },
            { "MARTES", DayOfWeek.Tuesday },
            { "MIERCOLES", DayOfWeek.Wednesday },
            { "JUEVES", DayOfWeek.Thursday },
            { "VIERNES", DayOfWeek.Friday },
            { "SABADO", DayOfWeek.Saturday },
            { "DOMINGO", DayOfWeek.Sunday }
        };

        private static readonly Dictionary<string, int> NombreToNumero = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            { "LUNES", 1 },
            { "MARTES", 2 },
            { "MIERCOLES", 3 },
            { "JUEVES", 4 },
            { "VIERNES", 5 },
            { "SABADO", 6 },
            { "DOMINGO", 7 }
        };

        private static readonly Dictionary<int, string> NumeroToNombre = new Dictionary<int, string>
        {
            { 1, "LUNES" },
            { 2, "MARTES" },
            { 3, "MIERCOLES" },
            { 4, "JUEVES" },
            { 5, "VIERNES" },
            { 6, "SABADO" },
            { 7, "DOMINGO" }
        };

        /// <summary>
        /// Normaliza un nombre de día eliminando acentos, caracteres diacríticos combinantes (FormD),
        /// caracteres no separables (NBSP \u00A0) y colapsando espacios múltiples.
        /// Retorna la representación canónica en mayúsculas (ej. MIERCOLES, SABADO).
        /// </summary>
        /// <param name="dia">Nombre del día con o sin tilde.</param>
        /// <returns>Nombre en mayúsculas sin tildes o cadena vacía si es nulo o vacío.</returns>
        public static string NormalizarDiaSemana(string dia)
        {
            if (string.IsNullOrWhiteSpace(dia))
            {
                return string.Empty;
            }

            string limpio = dia.Replace('\u00A0', ' ').Replace('\t', ' ').Trim();
            var sb = new StringBuilder(limpio.Length);
            bool espacioPrevio = false;
            foreach (char c in limpio)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (!espacioPrevio)
                    {
                        sb.Append(' ');
                    }
                    espacioPrevio = true;
                    continue;
                }
                espacioPrevio = false;
                sb.Append(c);
            }

            string sinTildes = new string(sb.ToString()
                .Normalize(NormalizationForm.FormD)
                .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                .ToArray())
                .ToUpperInvariant();

            return DiasSinTilde.TryGetValue(sinTildes, out string canonico) ? canonico : sinTildes;
        }

        /// <summary>
        /// Convierte una instancia de <see cref="DayOfWeek"/> a su nombre canónico en español sin tilde.
        /// </summary>
        /// <param name="dayOfWeek">Valor del enum DayOfWeek.</param>
        /// <returns>Nombre canónico en español en mayúsculas (LUNES..DOMINGO).</returns>
        public static string DesdeDayOfWeek(DayOfWeek dayOfWeek)
        {
            if (DayOfWeekToNombre.TryGetValue(dayOfWeek, out string nombre))
            {
                return nombre;
            }

            return dayOfWeek.ToString().ToUpperInvariant();
        }

        /// <summary>
        /// Intenta convertir un nombre de día en español (con o sin tilde, con cualquier formato de espacios) a <see cref="DayOfWeek"/>.
        /// </summary>
        /// <param name="dia">Nombre del día.</param>
        /// <param name="dayOfWeek">Resultado del enum correspondiente si es válido.</param>
        /// <returns>True si el día fue reconocido; false en caso contrario.</returns>
        public static bool TryParse(string dia, out DayOfWeek dayOfWeek)
        {
            dayOfWeek = DayOfWeek.Monday;
            string normalizado = NormalizarDiaSemana(dia);
            if (string.IsNullOrEmpty(normalizado))
            {
                return false;
            }

            return NombreToDayOfWeek.TryGetValue(normalizado, out dayOfWeek);
        }

        /// <summary>
        /// Convierte un número ordinal de presentación (1 = Lunes, ..., 7 = Domingo) a <see cref="DayOfWeek"/> canónico de .NET/SQLite (0 = Domingo, 1 = Lunes, ..., 6 = Sábado).
        /// Retorna null si el número está fuera de rango [1..7].
        /// </summary>
        /// <param name="numeroDia">Número entre 1 (Lunes) y 7 (Domingo).</param>
        /// <returns>Instancia de DayOfWeek correspondiente o null si es inválido.</returns>
        public static DayOfWeek? NumeroDiaADayOfWeek(int numeroDia)
        {
            if (numeroDia < 1 || numeroDia > 7)
            {
                return null;
            }

            return (DayOfWeek)(numeroDia == 7 ? 0 : numeroDia);
        }

        /// <summary>
        /// Convierte una instancia de <see cref="DayOfWeek"/> a su número ordinal semanal de presentación (1 = Lunes, ..., 7 = Domingo).
        /// </summary>
        /// <param name="day">Día de la semana.</param>
        /// <returns>Entero entre 1 (Lunes) y 7 (Domingo).</returns>
        public static int DayOfWeekANumero(DayOfWeek day)
        {
            return day == DayOfWeek.Sunday ? 7 : (int)day;
        }

        /// <summary>
        /// Convierte un nombre de día a su número ordinal semanal de presentación (1 = Lunes, ..., 7 = Domingo).
        /// Retorna 0 si no es válido.
        /// NOTA: Para modelos internos del motor genético y persistencia en 'ag_bloque_horario.dia_semana', debe usarse <see cref="DayOfWeek"/>.
        /// </summary>
        /// <param name="dia">Nombre del día en español con o sin tildes.</param>
        /// <returns>Entero entre 1 y 7, o 0 si no es reconocido.</returns>
        public static int ObtenerNumeroDia(string dia)
        {
            string normalizado = NormalizarDiaSemana(dia);
            if (string.IsNullOrEmpty(normalizado))
            {
                return 0;
            }

            if (NombreToNumero.TryGetValue(normalizado, out int num))
            {
                return num;
            }

            return 0;
        }

        /// <summary>
        /// Convierte un número ordinal semanal (1 = Lunes, ..., 7 = Domingo) a su nombre canónico en español sin tilde.
        /// </summary>
        /// <param name="numeroDia">Número entre 1 y 7.</param>
        /// <returns>Nombre del día en mayúsculas sin tilde o cadena vacía si el número está fuera de rango.</returns>
        public static string ObtenerNombreDia(int numeroDia)
        {
            if (NumeroToNombre.TryGetValue(numeroDia, out string nombre))
            {
                return nombre;
            }

            return string.Empty;
        }

        /// <summary>
        /// Compara si dos representaciones de día son equivalentes ignorando tildes, mayúsculas, NBSP y espacios.
        /// </summary>
        /// <param name="diaA">Primer nombre de día.</param>
        /// <param name="diaB">Segundo nombre de día.</param>
        /// <returns>True si representan el mismo día.</returns>
        public static bool SonIguales(string diaA, string diaB)
        {
            return string.Equals(NormalizarDiaSemana(diaA), NormalizarDiaSemana(diaB), StringComparison.OrdinalIgnoreCase);
        }
    }
}
