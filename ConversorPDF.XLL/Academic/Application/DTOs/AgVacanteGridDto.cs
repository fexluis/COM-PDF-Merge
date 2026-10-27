using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// DTO para la visualización de vacantes académicas en la grilla del Panel de Vacantes (Módulo R, Fases 318 y 322).
    /// </summary>
    public class AgVacanteGridDto
    {
        public int IdSesion { get; set; }
        public int IdVinculacion { get; set; }
        public int IdPrograma { get; set; }
        public string Programa { get; set; }
        public int IdAsignatura { get; set; }
        public string Asignatura { get; set; }
        public string Semestre { get; set; }
        public AgMotivoVacante MotivoVacante { get; set; }
        public string Motivo => MotivoVacante.Codigo();
        public int CandidatosPool { get; set; }
        public string Fecha { get; set; }
        public string Dia { get; set; }
        public string Horario { get; set; }
        public string Aula { get; set; }
        public string AulaRef { get; set; }
    }
}
