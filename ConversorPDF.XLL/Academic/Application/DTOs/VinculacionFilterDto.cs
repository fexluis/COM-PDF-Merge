using System;

namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// DTO para filtrar la lista de vinculaciones.
    /// </summary>
    public class VinculacionFilterDto
    {
        public int? IdPeriodo { get; set; }
        public int? IdSede { get; set; }
        public int? IdTercero { get; set; }
        public int? IdPrograma { get; set; }
        public bool SoloViaticas { get; set; }
        public DateTime? FechaDesde { get; set; }
        public DateTime? FechaHasta { get; set; }
        public string ModalidadPrograma { get; set; }
        
        /// <summary>
        /// Filtro por subunidad (Territorial).
        /// </summary>
        public string Subunidad { get; set; }
        
        /// <summary>
        /// Filtro por semestre.
        /// </summary>
        public string Semestre { get; set; }

        /// <summary>
        /// Filtro por corrida de optimización (Modo AG / Staging).
        /// </summary>
        public long? IdCorrida { get; set; }
    }
}
