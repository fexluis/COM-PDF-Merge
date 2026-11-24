using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Resultado y telemetría emitidos por el motor de reparación de sobrecargas R5 (Fase 2325 / Fase 2329).
    /// </summary>
    internal sealed class AgResultadoReparacionR5
    {
        public AgIndividuo Individuo { get; set; }
        public int Intentos { get; set; }
        public int Evaluaciones { get; set; }
        public int HardPre { get; set; }
        public int HardPost { get; set; }
        public double DesperdicioPre { get; set; }
        public double DesperdicioPost { get; set; }
        public bool HuboMejora { get; set; }
        public int RippleIntentos { get; set; }
        public int RippleSinCola { get; set; }
        public List<string> Detalle { get; set; } = new List<string>();
        public string Error { get; set; } = string.Empty;
    }
}
