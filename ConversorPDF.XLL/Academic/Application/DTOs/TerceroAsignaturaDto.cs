namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// Data Transfer Object para la visualización y edición en interfaz de una habilitación docente-asignatura.
    /// </summary>
    public class TerceroAsignaturaDto
    {
        public long Id { get; set; }
        public int IdTercero { get; set; }
        public string NombreTercero { get; set; } = string.Empty;
        public int IdPrograma { get; set; }
        public string NombrePrograma { get; set; } = string.Empty;
        public int IdAsignatura { get; set; }
        public string NombreAsignatura { get; set; } = string.Empty;
        public string Semestre { get; set; } = string.Empty;
        public int Activo { get; set; } = 1;
        public int Prioridad { get; set; }
        public int VecesAsignada { get; set; }
        public string FechaUltimaAsignacion { get; set; } = string.Empty;
        public string Origen { get; set; } = "MANUAL";
        public string Observaciones { get; set; } = string.Empty;
        public string Subunidad { get; set; } = string.Empty;
    }
}
