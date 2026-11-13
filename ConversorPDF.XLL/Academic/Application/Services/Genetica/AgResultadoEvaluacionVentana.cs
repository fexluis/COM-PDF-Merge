using System;
using System.Collections.Generic;

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
        public int SemanasRequeridas { get; set; }
        public DateTime? FechaFinMinima { get; set; }
        public int RanurasOciosas { get; set; }
        public int DemandaPendiente { get; set; }
        public int DiaCuello { get; set; }
        public string FranjaCuello { get; set; } = string.Empty;
        public List<int> MateriasFueraVentana { get; set; } = new List<int>();
        public List<int> IdAsignaturasSaturadas { get; set; } = new List<int>();
    }
}

