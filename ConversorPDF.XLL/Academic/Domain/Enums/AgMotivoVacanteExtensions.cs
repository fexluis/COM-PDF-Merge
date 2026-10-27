using System;

namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Métodos de extensión y utilidades de conversión determinista para <see cref="AgMotivoVacante"/> (Fase 329).
    /// </summary>
    public static class AgMotivoVacanteExtensions
    {
        /// <summary>
        /// Obtiene el código textual canónico en mayúsculas del motivo de vacante.
        /// </summary>
        /// <param name="m">Instancia del enum <see cref="AgMotivoVacante"/>.</param>
        /// <returns>Cadena en mayúsculas o cadena vacía si es Ninguno.</returns>
        public static string Codigo(this AgMotivoVacante m)
        {
            return m == AgMotivoVacante.Ninguno ? string.Empty : m.ToString();
        }

        /// <summary>
        /// Parsea de forma segura e insensible a mayúsculas un código textual al enum <see cref="AgMotivoVacante"/>.
        /// </summary>
        /// <param name="codigo">Código textual a convertir.</param>
        /// <returns>Valor del enum <see cref="AgMotivoVacante"/> o Ninguno si no se reconoce.</returns>
        public static AgMotivoVacante Desde(string codigo)
        {
            if (string.IsNullOrWhiteSpace(codigo)) return AgMotivoVacante.Ninguno;
            return Enum.TryParse<AgMotivoVacante>(codigo, true, out var m) ? m : AgMotivoVacante.Ninguno;
        }
    }
}
