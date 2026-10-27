using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class Sesion
    {
        public int Id { get; set; }
        public int IdVinculacion { get; set; }
        public int IdTercero { get; set; }
        /// <summary>
        /// Fecha de la sesión en formato yyyy-MM-dd (columna fechasesion).
        /// </summary>
        public DateTime FechaSesion { get; set; }
        /// <summary>
        /// Horario descriptivo de la sesión, ej: "MARTES|4" (columna horariosesion).
        /// </summary>
        public string HorarioSesion { get; set; }
        public double Horas { get; set; }
        public bool Viatica { get; set; }
        public int EsVacante { get; set; }
        public string MotivoVacante { get; set; } = string.Empty;
    }
}
