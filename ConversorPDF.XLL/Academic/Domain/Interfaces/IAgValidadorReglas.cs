using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Validador de reglas duras (R1-R7, W1) y evaluación de penalizaciones blandas (S1-S5) sobre un individuo.
    /// </summary>
    public interface IAgValidadorReglas
    {
        /// <summary>
        /// Evalúa todas las reglas sobre el individuo y llena su lista de violaciones y contadores.
        /// </summary>
        void Evaluar(AgIndividuo individuo, AgContexto contexto);
    }
}
