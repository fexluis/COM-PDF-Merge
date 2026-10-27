using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Modelo inmutable de asignación de carril y oleada para una asignatura dentro de una cohorte/aula (Fase 1057).
    /// Acota la sub-ventana de fechas [FechaInicio, FechaFin] y el patrón semanal de días.
    /// </summary>
    internal sealed class AgPlanOleada
    {
        public int IdAsignatura { get; }
        public int Carril { get; }
        public int Oleada { get; }
        public IReadOnlyList<int> DiasPatron { get; }
        public IReadOnlyList<AgRanura> RanurasPatron { get; }
        public DateTime FechaInicio { get; }
        public DateTime FechaFin { get; }
        public int SemanasOleada { get; }
        public int SesionesRequeridas { get; }
        public int CapacidadRanuras { get; }

        public AgPlanOleada(
            int idAsignatura,
            int carril,
            int oleada,
            IReadOnlyList<int> diasPatron,
            DateTime fechaInicio,
            DateTime fechaFin,
            int semanasOleada,
            int sesionesRequeridas,
            int capacidadRanuras,
            IReadOnlyList<AgRanura> ranurasPatron = null)
        {
            IdAsignatura = idAsignatura;
            Carril = carril;
            Oleada = oleada;
            DiasPatron = diasPatron ?? new int[0];
            FechaInicio = fechaInicio;
            FechaFin = fechaFin;
            SemanasOleada = semanasOleada;
            SesionesRequeridas = sesionesRequeridas;
            CapacidadRanuras = capacidadRanuras;
            RanurasPatron = ranurasPatron ?? new List<AgRanura>();
        }

        public override string ToString() =>
            $"Asig {IdAsignatura} | Carril {Carril} [{string.Join(",", DiasPatron)}] | Oleada {Oleada} ({FechaInicio:yyyy-MM-dd}..{FechaFin:yyyy-MM-dd}) | {SesionesRequeridas}/{CapacidadRanuras} ses";
    }
}
