using System;

namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// DTO que representa una fila del DataGridView, sin dependencia de WinForms.
    /// Utilizado para el mapeo y testeo unitario de la lógica de extracción.
    /// </summary>
    public class GridRowDto
    {
        public bool IsChecked { get; set; }
        public bool IsNewRow { get; set; }
        public object Tag { get; set; }
        public bool Viatica { get; set; }
        public double Horas { get; set; }
        /// <summary>
        /// Fecha de la sesión (columna fechasesion en DB).
        /// </summary>
        public DateTime FechaSesion { get; set; }
        /// <summary>
        /// Horario descriptivo de la sesión, ej: "MARTES|4" (columna horariosesion en DB).
        /// </summary>
        public string HorarioSesion { get; set; }
    }
}

