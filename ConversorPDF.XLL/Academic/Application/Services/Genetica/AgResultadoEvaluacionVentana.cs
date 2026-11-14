using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

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
        public DateTime FechaInicioVentana { get; set; }
        public DateTime FechaLimite { get; set; }
        public DateTime? FechaFinProyectada { get; set; }
        // Compatibilidad con consumidores antiguos. La proyección no acredita un mínimo global.
        public DateTime? FechaFinMinima { get { return FechaFinProyectada; } set { FechaFinProyectada = value; } }
        public bool CabeEnPlazo { get; set; }
        public bool BusquedaAgotada { get; set; }
        public bool HorizonteAgotado { get; set; }
        public string Error { get; set; } = string.Empty;
        public AgPlanOleadasResultado Plan { get; set; }
        public IReadOnlyList<AgSesionPlanificada> Sesiones { get; set; } = new List<AgSesionPlanificada>();
        public int RanurasOciosas { get; set; }
        public int DemandaPendiente { get; set; }
        public double DeficitHoras { get; set; }
        public int DeficitSesiones => DemandaPendiente;
        public int DiaCuello { get; set; }
        public string FranjaCuello { get; set; } = string.Empty;
        public List<int> MateriasFueraVentana { get; set; } = new List<int>();
        public List<int> IdAsignaturasSaturadas { get; set; } = new List<int>();
        public IReadOnlyDictionary<int, int> SesionesPendientesPorAsignatura { get; set; } = new Dictionary<int, int>();
        public IReadOnlyDictionary<int, double> HorasPendientesPorAsignatura { get; set; } = new Dictionary<int, double>();
    }
}

