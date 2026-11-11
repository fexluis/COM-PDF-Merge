using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Resultado de la reparación de una corrida huérfana del AG (Fase 1566):
    /// cabecera e individuo reconstruidos desde el staging y violaciones repobladas en ag_violacion.
    /// </summary>
    public sealed class AgReparacionCorridaResultado
    {
        public long IdCorrida { get; set; }
        public long IdIndividuo { get; set; }
        public int IdPeriodo { get; set; }
        public int IdPrograma { get; set; }
        public string Subunidad { get; set; }

        public int Hard { get; set; }
        public double Soft { get; set; }
        public double Fitness { get; set; }

        public int GenesEvaluados { get; set; }
        public int Violaciones { get; set; }
        public int Vinculaciones { get; set; }
        public int Sesiones { get; set; }

        public IReadOnlyDictionary<AgReglaViolacion, int> PorRegla { get; set; }
        public string ResumenPorRegla { get; set; }

        public AgReparacionCorridaResultado()
        {
            Subunidad = string.Empty;
            ResumenPorRegla = string.Empty;
        }
    }
}
