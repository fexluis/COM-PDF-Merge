namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// Parámetros de filtrado para consultas sobre el pool de docentes-asignaturas.
    /// </summary>
    public class FiltroPoolDto
    {
        public int? IdTercero { get; set; }
        public int? IdPrograma { get; set; }
        public int? IdAsignatura { get; set; }
        public string Semestre { get; set; } = string.Empty;
        public int? Activo { get; set; }
        public string Subunidad { get; set; } = string.Empty;
    }
}
