using System;

namespace ConversorPDF.XLL.Core.Helpers
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
        public double Duracion { get; set; }
        public DateTime FechaInicial { get; set; }
        public DateTime FechaFinal { get; set; }
    }
}
