using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Resultado de la planificación y secuenciación determinista de una cadena de carril (Fase 2371, Fase 2406, Fase 2411, Fase 2420).
    /// </summary>
    internal sealed class AgResultadoPlanCarril
    {
        public bool Factible { get; set; }
        public string Motivo { get; set; } = string.Empty; // SIN_CARRIL | SIN_FECHAS | CAPACIDAD | SIN_ORDEN | FIJO
        /// <summary>Sesiones totales requeridas por los sujetos del carril (Fase 2421).</summary>
        public int SesionesRequeridas { get; set; }
        /// <summary>Semanas requeridas equivalentes deducidas a partir de SesionesPorSemana (Fase 2421, Fase 2423).</summary>
        public int SemanasRequeridas
        {
            get
            {
                if (SesionesRequeridas <= 0) return 0;
                int porSemana = System.Math.Max(1, SesionesPorSemana);
                return (int)System.Math.Ceiling(SesionesRequeridas / (double)porSemana);
            }
        }
        /// <summary>Fechas/sesiones totales disponibles en el carril (Fase 2406, Fase 2411).</summary>
        public int SesionesDisponibles { get; set; }
        /// <summary>Sesiones requeridas/disponibles por semana del carril (Fase 2420).</summary>
        public int SesionesPorSemana { get; set; } = 2;
        /// <summary>Semanas equivalentes deducidas a partir de SesionesPorSemana del carril (Fase 2406, Fase 2411, Fase 2420).</summary>
        public int SemanasEquivalentes
        {
            get
            {
                if (SesionesDisponibles <= 0) return 0;
                int porSemana = System.Math.Max(1, SesionesPorSemana);
                return (int)System.Math.Ceiling(SesionesDisponibles / (double)porSemana);
            }
        }
        /// <summary>Semanas requeridas equivalentes deducidas a partir de SesionesPorSemana del carril (Fase 2420, Fase 2421).</summary>
        public int SemanasRequeridasEquivalentes => SemanasRequeridas;
        /// <summary>Propiedad de sólo lectura equivalente a SemanasEquivalentes para compatibilidad (Fase 2411).</summary>
        public int SemanasDisponibles => SemanasEquivalentes;
        public AgIndividuo Individuo { get; set; }
    }
}
