namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Información descriptiva de una vacante estructural detectada sobre la demanda real de aula-asignatura (Fase 2052).
    /// </summary>
    internal sealed class VacanteEstructuralInfo
    {
        public int IdAsignatura { get; set; }
        public int IdSede { get; set; }
        public string Semestre { get; set; }
        public string Grupo { get; set; }
        public string Motivo { get; set; }
        public int IdAula { get; set; }
    }
}
