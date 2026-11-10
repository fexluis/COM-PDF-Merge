using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Resultado del recálculo de violaciones de una corrida (Fase 1527).
    /// Devuelve el desglose por regla para diagnóstico.
    /// </summary>
    public sealed class AgRecalculoResultado
    {
        public long IdCorrida { get; set; }
        public long IdIndividuo { get; set; }
        public int Hard { get; set; }
        public double Soft { get; set; }
        public int GenesEvaluados { get; set; }
        public int AsignaturasContexto { get; set; }
        public IReadOnlyList<AgViolacion> Violaciones { get; set; }
        public IReadOnlyDictionary<AgReglaViolacion, int> PorRegla { get; set; }
        public string ResumenPorRegla { get; set; }
    }
}