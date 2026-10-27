using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Representa una salvedad o justificación manual de conflicto en la capa de staging del AG.
    /// Mapea a la tabla 'ag_salvedades_conflictos'.
    /// </summary>
    public class AgSalvedad
    {
        public int Id { get; set; }
        public long IdCorrida { get; set; }
        public int IdSesionA { get; set; }
        public int IdSesionB { get; set; }
        public int IdTercero { get; set; }
        public int IdPeriodo { get; set; }
        public string Justificacion { get; set; }
        public string Subunidad { get; set; }

        public AgSalvedad()
        {
            Justificacion = string.Empty;
            Subunidad = string.Empty;
        }
    }
}
