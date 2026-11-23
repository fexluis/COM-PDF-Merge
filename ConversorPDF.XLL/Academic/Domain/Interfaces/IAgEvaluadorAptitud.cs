using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Evaluador de la función objetivo y asignación de aptitud escalar (Fitness) para cada individuo.
    /// </summary>
    public interface IAgEvaluadorAptitud
    {
        void EvaluarAptitud(AgIndividuo individuo, AgContexto contexto);
        void EvaluarAptitud(AgIndividuo individuo, AgContexto contexto, bool forzar, bool incluirDetalleJson = false);
    }
}
