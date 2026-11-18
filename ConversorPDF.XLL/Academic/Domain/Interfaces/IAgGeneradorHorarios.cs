using System;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Motor principal de ejecución y orquestación del Algoritmo Genético para optimización de horarios académicos.
    /// </summary>
    public interface IAgGeneradorHorarios
    {
        Task<Result<AgResultadoCorrida>> EjecutarOptimizacionAsync(
            AgContexto contexto,
            int poblacionTamano,
            int generacionesMax,
            double tasaCruce,
            double tasaMutacion,
            int elitismo,
            IProgress<(int Generacion, double MejorFitness, int HardViolations, string Mensaje)> progreso,
            CancellationToken ct,
            double umbralFitness = 0.0,
            int estancamientoMax = AgParametrosPorDefecto.EstancamientoMax);
    }
}
