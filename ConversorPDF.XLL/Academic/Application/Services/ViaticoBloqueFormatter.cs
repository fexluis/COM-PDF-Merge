using System;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Servicio de presentación y formateo de información de bloques de viáticos para vistas de árbol.
    /// </summary>
    public static class ViaticoBloqueFormatter
    {
        /// <summary>
        /// Obtiene el rango efectivo (inicio, fin) de la comisión de un bloque,
        /// aplicando el día adicional de regreso por pernocta en el último día si aplica.
        /// </summary>
        /// <param name="bloque">Bloque de viático evaluado.</param>
        /// <returns>Tupla con fecha de inicio y fin de la comisión, o null si las fechas no están definidas.</returns>
        public static (DateTime Inicio, DateTime Fin)? ObtenerRangoComision(ViaticoBloque bloque)
        {
            if (bloque?.FechaMinima == null || bloque.FechaFinComision == null) return null;
            return (bloque.FechaMinima.Value.Date, bloque.FechaFinComision.Value.Date);
        }

        /// <summary>
        /// Obtiene el texto de cierre del bloque de viático: Día de la semana y HorarioSesion
        /// de la sesión con la hora de inicio más tardía en la FechaMaxima del bloque.
        /// </summary>
        public static string ObtenerTextoCierre(ViaticoBloque bloque)
        {
            if (bloque?.FechaMaxima == null || bloque.Sesiones == null) return "S/F";
            var delDia = bloque.Sesiones.Where(s => s.FechaSesion.HasValue
                                && s.FechaSesion.Value.Date == bloque.FechaMaxima.Value.Date).ToList();
            if (delDia.Count == 0) return "S/F";

            var cierre = delDia.OrderByDescending(ObtenerHoraInicio)
                               .ThenByDescending(s => s.HorarioSesion, StringComparer.OrdinalIgnoreCase)
                               .First();

            string dia = (cierre.DiaSemana ?? "S/D").ToUpperInvariant();
            return $"{dia} {cierre.HorarioSesion}".Trim();
        }

        private static TimeSpan ObtenerHoraInicio(SesionListItem s)
        {
            if (string.IsNullOrWhiteSpace(s.HorarioSesion)) return TimeSpan.MinValue;
            var r = HorarioSesionParser.TryParse(s.HorarioSesion);
            return r.IsSuccess ? r.Value.Inicio : TimeSpan.MinValue;
        }

        /// <summary>
        /// Formatea la cadena de texto para el nodo de Nivel 2 (Bloque de viático) en el TreeView de viáticos.
        /// Formato: 📅 [FechaInicio] - [FechaFin] | [Duración] d | [Aéreo/Terrestre] | [Ruta] | $[ValorTransporte] | [DIA HORA] | [Pernocta Ultimo Dia]
        /// </summary>
        /// <param name="b">Bloque de viático.</param>
        /// <param name="fechaIni">Fecha de inicio formateada (dd/MM/yyyy).</param>
        /// <param name="fechaFin">Fecha de fin efectiva formateada (dd/MM/yyyy).</param>
        /// <param name="tipoTransporte">Modalidad de transporte ("Aéreo" o "Terrestre").</param>
        /// <param name="ruta">Ruta formateada del trayecto.</param>
        /// <param name="valorTransporte">Valor de transporte calculado.</param>
        /// <returns>Cadena formateada para la fila del bloque, o cadena vacía si el bloque es nulo.</returns>
        public static string FormatearNodoNivel2(ViaticoBloque b, string fechaIni, string fechaFin,
            string tipoTransporte, string ruta, decimal valorTransporte)
        {
            if (b == null) return string.Empty;
            string marca = b.TienePernoctaUltimoDia ? " | Pernocta Ultimo Dia" : string.Empty;
            return $"📅 {fechaIni} - {fechaFin} | {b.DuracionTotal} d | {tipoTransporte} | {ruta} | ${valorTransporte:N0} | {ObtenerTextoCierre(b)}{marca}";
        }
    }
}
