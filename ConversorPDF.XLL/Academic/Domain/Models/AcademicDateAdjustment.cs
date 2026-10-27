using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Representa el resultado inmutable del análisis de ajuste de fechas para una vinculación.
    /// </summary>
    public sealed class AcademicDateAdjustment
    {
        /// <summary>
        /// Indica si las fechas generadas difieren de las originales y requieren actualización en la UI.
        /// </summary>
        public bool NeedsAdjustment { get; }

        /// <summary>
        /// La fecha de inicio calculada a partir de la primera sesión.
        /// </summary>
        public DateTime CalculatedStart { get; }

        /// <summary>
        /// La fecha final calculada a partir de la última sesión.
        /// </summary>
        public DateTime CalculatedEnd { get; }

        /// <summary>
        /// Mensajes detallados sobre los cambios realizados (ej: "Fecha inicio ajustada...").
        /// </summary>
        public IReadOnlyList<string> Messages { get; }

        /// <summary>
        /// Indica si la fecha final excedió el límite original.
        /// </summary>
        public bool ExceededOriginalLimit { get; }

        /// <summary>
        /// Constructor interno: solo puede instanciarse desde <see cref="AcademicDateAdjustmentBuilder"/>.
        /// </summary>
        internal AcademicDateAdjustment(
            bool needsAdjustment,
            DateTime calculatedStart,
            DateTime calculatedEnd,
            bool exceededOriginalLimit,
            IReadOnlyList<string> messages)
        {
            NeedsAdjustment = needsAdjustment;
            CalculatedStart = calculatedStart;
            CalculatedEnd = calculatedEnd;
            ExceededOriginalLimit = exceededOriginalLimit;
            Messages = messages ?? new List<string>().AsReadOnly();
        }
    }
}
