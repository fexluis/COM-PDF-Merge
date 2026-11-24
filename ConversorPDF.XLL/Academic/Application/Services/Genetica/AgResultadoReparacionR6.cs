using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Resultado y telemetría del reparador de secuencia curricular R6 (Fase 2369).
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
    }
}
