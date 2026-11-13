namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Resultado de la evaluación de demanda y saturación por carriles para una ventana (Fase 1737).
    /// </summary>
    internal sealed class AgResultadoEvaluacionVentana
    {
        public int DemandaTotal { get; set; }
        public int CapacidadTotal { get; set; }
        public bool HaySaturacion { get; set; }
        public bool SaturacionCarriles { get; set; }
        public string Motivo { get; set; } = string.Empty;
    }
}

