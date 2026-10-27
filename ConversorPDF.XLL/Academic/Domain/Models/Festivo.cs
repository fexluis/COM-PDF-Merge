using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Representa un día festivo dentro de un periodo académico.
    /// </summary>
    public class Festivo
    {
        /// <summary> Identificador único del festivo. </summary>
        public int IdFestivo { get; set; }

        /// <summary> Identificador del periodo académico al que pertenece el festivo. </summary>
        public int IdPeriodo { get; set; }

        /// <summary> Fecha de inicio del día festivo. </summary>
        public DateTime FechaFestivoInicio { get; set; }

        /// <summary> Fecha de finalización del día festivo. </summary>
        public DateTime FechaFestivoFinal { get; set; }

        /// <summary> Detalle o descripción del festivo (ej. "Día de la Independencia"). </summary>
        public string DetalleFestivo { get; set; }

        /// <summary> Tipo de periodo, ej. "FESTIVO" o "PRESENCIAL". </summary>
        public string TipoFestivo { get; set; } = "FESTIVO";
    }
}
