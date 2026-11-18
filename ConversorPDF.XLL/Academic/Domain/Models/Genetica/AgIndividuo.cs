using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Representa un cromosoma o solución candidata en el espacio de búsqueda del AG.
    /// Contiene un conjunto de genes que asignan docente, aula y bloques a cada curso/asignatura.
    /// Mapea a la tabla 'ag_individuo'.
    /// </summary>
    public class AgIndividuo
    {
        public long Id { get; set; }
        public long IdCorrida { get; set; }
        public int Generacion { get; set; }
        public double Fitness { get; set; }
        public int HardViolations { get; set; }
        public double SoftViolations { get; set; }
        public AgEstadoIndividuo Estado { get; set; }
        public string Subunidad { get; set; }
        public long HuellaEstructural { get; set; }

        public List<AgGen> Genes { get; set; }
        public List<AgViolacion> Violaciones { get; set; }

        public bool EsFactible => HardViolations == 0;
        public int Vacantes => Genes?.FindAll(g => g.EsVacante == 1 || g.IdTercero <= 0).Count ?? 0;

        public AgIndividuo()
        {
            Estado = AgEstadoIndividuo.NoEvaluado;
            Genes = new List<AgGen>();
            Violaciones = new List<AgViolacion>();
            Subunidad = string.Empty;
        }

        public AgIndividuo Clonar()
        {
            var clon = new AgIndividuo
            {
                Id = 0,
                IdCorrida = this.IdCorrida,
                Generacion = this.Generacion,
                Fitness = this.Fitness,
                HardViolations = this.HardViolations,
                SoftViolations = this.SoftViolations,
                Estado = this.Estado,
                Subunidad = this.Subunidad,
                HuellaEstructural = this.HuellaEstructural,
                Genes = new List<AgGen>(this.Genes.Count),
                Violaciones = new List<AgViolacion>(this.Violaciones)
            };

            foreach (var g in this.Genes)
            {
                clon.Genes.Add(g.Clonar());
            }

            return clon;
        }

        /// <summary>
        /// Calcula una huella estructural de 64 bits rápida (FNV-1a) sobre el genotipo del individuo
        /// para determinar identidad de asignaciones en caché de aptitud.
        /// </summary>
        public long CalcularHuellaEstructural()
        {
            if (Genes == null || Genes.Count == 0)
            {
                HuellaEstructural = 0L;
                return 0L;
            }

            unchecked
            {
                ulong hash = 14695981039346656037UL;
                const ulong prime = 1099511628211UL;

                hash = (hash ^ (ulong)Genes.Count) * prime;
                for (int i = 0; i < Genes.Count; i++)
                {
                    var g = Genes[i];
                    if (g != null)
                    {
                        hash = (hash ^ (ulong)g.CalcularHuellaEstructural()) * prime;
                    }
                }

                HuellaEstructural = (long)hash;
                return (long)hash;
            }
        }

        public override string ToString()
        {
            return $"Individuo Gen:{Generacion} | Fit:{Fitness:F4} | Hard:{HardViolations} | Soft:{SoftViolations:F2} | Genes:{Genes?.Count ?? 0}";
        }
    }
}
