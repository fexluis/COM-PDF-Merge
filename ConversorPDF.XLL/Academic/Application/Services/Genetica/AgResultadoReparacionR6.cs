using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Resultado y telemetría del reparador de secuencia curricular R6 (Fase 2369, Fase 2373, Fase 2386).
    /// </summary>
    internal sealed class AgResultadoReparacionR6
    {
        public AgIndividuo Individuo { get; set; }
        public int Intentos { get; set; }
        public int HardPre { get; set; }
        public int HardPost { get; set; }
        public bool HuboMejora { get; set; }
        public List<string> Detalle { get; set; } = new List<string>();
        public string Error { get; set; } = string.Empty;

        /// <summary>Motivo diagnóstico de infactibilidad o descarte del carril (Fase 2373).</summary>
        public string Motivo { get; set; } = string.Empty;

        /// <summary>Motivo de fallo o descarte en re-ranuración hacia aulas alternativas (Fase 2386).</summary>
        public string MotivoReranuracion { get; set; } = string.Empty;

        /// <summary>Semanas requeridas calculadas en el carril (Fase 2373).</summary>
        public int SemanasRequeridas { get; set; }

        /// <summary>Semanas disponibles en el carril (Fase 2373).</summary>
        public int SemanasDisponibles { get; set; }
    }
}
