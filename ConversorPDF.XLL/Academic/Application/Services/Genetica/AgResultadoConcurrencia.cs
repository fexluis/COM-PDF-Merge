namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Representa el resultado del cálculo de concurrencia de asignaturas por cohorte y carriles (Fase 2304).
    /// </summary>
    internal sealed class AgResultadoConcurrencia
    {
        public bool Evaluado { get; set; }
        public int MaxConcurrentes { get; set; }
        public int CarrilesDisponibles { get; set; }
        public bool Factible { get; set; }
    }
}
