using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Representa un ítem de la lista de vinculaciones para visualización en el grid.
    /// Esta entidad es de solo lectura desde la perspectiva de la consulta de lista.
    /// </summary>
    public class VinculacionListItem
    {
        public int IdVinculacion { get; set; }
        public long IdCorrida { get; set; }
        public string PeriodoNombre { get; set; }
        public int IdPeriodo { get; set; }
        public string ProgramaTipo { get; set; }
        public string ProgramaAbreviatura { get; set; }
        public string TerceroNombre { get; set; }
        public string Semestre { get; set; }
        public string AsignaturaNombre { get; set; }
        public string SedeCiudad { get; set; }
        public DateTime? SesionFechaInicial { get; set; }
        public DateTime? SesionFechaFinal { get; set; }
        public double SesionHoras { get; set; }
        public bool SesionViatica { get; set; }
        public string Modalidad { get; set; }

        /// <summary>
        /// Texto formateado para mostrar el rango de fechas en la UI.
        /// </summary>
        public string RangoFechas => (SesionFechaInicial.HasValue && SesionFechaFinal.HasValue)
            ? $"{SesionFechaInicial.Value:dd/MM/yyyy} - {SesionFechaFinal.Value:dd/MM/yyyy}"
            : "No definida";
            
        /// <summary>
        /// Texto amigable para el tipo de programa.
        /// </summary>
        public string ProgramaDetalle => $"{ProgramaAbreviatura} ({ProgramaTipo})";
    }
}
