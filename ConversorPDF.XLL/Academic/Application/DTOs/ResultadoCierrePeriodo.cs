using System;

namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// Representa el resultado de la evaluación de la fecha de cierre de un período académico.
    /// </summary>
    public sealed class ResultadoCierrePeriodo
    {
        /// <summary>
        /// Indica si la fecha final analizada excede el límite de cierre determinado.
        /// </summary>
        public bool ExcedeLimite { get; set; }

        /// <summary>
        /// Fecha límite contra la cual se evaluó la fecha final.
        /// </summary>
        public DateTime Limite { get; set; }

        /// <summary>
        /// Indica si el límite proviene de la ventana canónica del período (periodo_fin)
        /// o del mecanismo de fallback histórico (15 de diciembre).
        /// </summary>
        public bool LimiteEsVentanaPeriodo { get; set; }

        /// <summary>
        /// Mensaje informativo o de advertencia para el usuario en caso de exceder el límite.
        /// </summary>
        public string Mensaje { get; set; }
    }
}