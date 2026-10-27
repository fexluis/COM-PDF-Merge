using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Registro de auditoría de la promoción de una solución factible de staging a producción.
    /// Mapea a la tabla 'ag_promocion'.
    /// </summary>
    public class AgPromocion
    {
        public long Id { get; set; }
        public long IdCorrida { get; set; }
        public long IdIndividuo { get; set; }
        public string FechaPromocion { get; set; }
        public string Usuario { get; set; }
        public int TotalVinculaciones { get; set; }
        public int TotalSesiones { get; set; }
        public int IdAula { get; set; }
        public long IdAgVinculacion { get; set; }
        public long IdVinculacion { get; set; }
        public int IdTercero { get; set; }
        public int IdAsignatura { get; set; }
        public string FechaInicio { get; set; }
        public string FechaFin { get; set; }
        public string Modo { get; set; }
        public string JustificacionVacante { get; set; }
        public string Subunidad { get; set; }

        public AgPromocion()
        {
            FechaPromocion = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Usuario = Environment.UserName;
            Modo = "COPIA_REMAPEO";
            FechaInicio = string.Empty;
            FechaFin = string.Empty;
            JustificacionVacante = string.Empty;
            Subunidad = string.Empty;
        }
    }
}
