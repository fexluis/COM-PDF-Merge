using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Core.Models.Academic
{
    /// <summary>
    /// Representa el resultado del análisis de ajuste de fechas para una vinculación.
    /// </summary>
    public sealed class AcademicDateAdjustment
    {
        /// <summary>
        /// Indica si las fechas generadas difieren de las originales y requieren actualización en la UI.
        /// </summary>
        public bool NeedsAdjustment { get; set; }

        /// <summary>
        /// La fecha de inicio calculada a partir de la primera sesión.
        /// </summary>
        public DateTime CalculatedStart { get; set; }

        /// <summary>
        /// La fecha final calculada a partir de la última sesión.
        /// </summary>
        public DateTime CalculatedEnd { get; set; }

        /// <summary>
        /// Mensajes detallados sobre los cambios realizados (ej: "Fecha inicio ajustada...").
        /// </summary>
        public List<string> Messages { get; set; } = new List<string>();

        /// <summary>
        /// Indica si la fecha final excedió el límite original.
        /// </summary>
        public bool ExceededOriginalLimit { get; set; }
    }
}
