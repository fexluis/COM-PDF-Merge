using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Entidad de sesión académica en la capa de staging del AG.
    /// Mapea a la tabla 'ag_sesiones'.
    /// </summary>
    public class AgSesion
    {
        public int Id { get; set; }
        public int IdVinculacion { get; set; }
        public int IdTercero { get; set; }
        public int IdAula { get; set; }
        public DateTime? FechaSesion { get; set; }
        public string HorarioSesion { get; set; }
        public double Horas { get; set; }
        public bool Viatica { get; set; }
        public string Subunidad { get; set; }
        public long IdIndividuo { get; set; }
        public long IdGen { get; set; }
        public int EsVacante { get; set; }
        public string MotivoVacante { get; set; } = string.Empty;
        public int CandidatosPool { get; set; }

        public AgSesion()
        {
            HorarioSesion = string.Empty;
            Subunidad = string.Empty;
            EsVacante = 0;
            MotivoVacante = string.Empty;
            CandidatosPool = 0;
        }

        private static readonly string[] NombresDia =
        { "DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO" };

        public string DiaSemanaTexto
        {
            get
            {
                if (!FechaSesion.HasValue) { return "S/D"; }
                int idx = (int)FechaSesion.Value.DayOfWeek;
                return (idx >= 0 && idx < NombresDia.Length) ? NombresDia[idx] : "S/D";
            }
        }

        public override string ToString()
        {
            string f = FechaSesion.HasValue ? FechaSesion.Value.ToString("yyyy-MM-dd") : "S/F";
            return $"AgSesion[Vinc:{IdVinculacion}, Fecha:{f}, Horas:{Horas}, Horario:{HorarioSesion}, Viatica:{Viatica}]";
        }
    }
}
