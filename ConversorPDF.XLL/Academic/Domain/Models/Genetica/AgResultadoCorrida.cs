using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Resultado global emitido tras la ejecución del motor del Algoritmo Genético.
    /// </summary>
    public class AgResultadoCorrida
    {
        public long IdCorrida { get; set; }
        public AgEstadoCorrida Estado { get; set; }
        public int GeneracionesEjecutadas { get; set; }
        public long TiempoTotalMs { get; set; }
        public AgIndividuo MejorIndividuo { get; set; }
        public List<AgBitacora> HistorialConvergencia { get; set; }
        public string Mensaje { get; set; }

        private bool? _esFactible;
        public bool EsFactible
        {
            get => _esFactible ?? (MejorIndividuo != null && MejorIndividuo.EsFactible);
            set => _esFactible = value;
        }

        public int Vacantes => MejorIndividuo?.Vacantes ?? 0;

        public AgResultadoCorrida()
        {
            Estado = AgEstadoCorrida.Iniciada;
            HistorialConvergencia = new List<AgBitacora>();
            Mensaje = string.Empty;
        }
    }
}
