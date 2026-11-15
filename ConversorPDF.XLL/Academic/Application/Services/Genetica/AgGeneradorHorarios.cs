using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Motor genético principal para optimización de horarios académicos.
    /// Ejecuta el ciclo evolutivo generacional con elitismo, cruce, mutación y telemetría en tiempo real.
    /// </summary>
    public class AgGeneradorHorarios : IAgGeneradorHorarios
    {
        private readonly IAgValidadorVentana _validadorVentana;
        private readonly IAgEvaluadorAptitud _evaluadorAptitud;
        private readonly AgFabricaIndividuos _fabricaIndividuos;
        private readonly AgOperadorSeleccion _operadorSeleccion;
        private readonly AgOperadorCruce _operadorCruce;
        private readonly AgOperadorMutacion _operadorMutacion;

        public AgGeneradorHorarios(
            IAgValidadorVentana validadorVentana,
            IAgEvaluadorAptitud evaluadorAptitud)
        {
            _validadorVentana = validadorVentana ?? throw new ArgumentNullException(nameof(validadorVentana));
            _evaluadorAptitud = evaluadorAptitud ?? throw new ArgumentNullException(nameof(evaluadorAptitud));

            _fabricaIndividuos = new AgFabricaIndividuos();
            _operadorSeleccion = new AgOperadorSeleccion();
            var asignador = new AgAsignadorTercero();
            _operadorCruce = new AgOperadorCruce(new Random(), asignador);
            _operadorMutacion = new AgOperadorMutacion();
        }

        public async Task<Result<AgResultadoCorrida>> EjecutarOptimizacionAsync(
            AgContexto contexto,
            int poblacionTamano,
            int generacionesMax,
            double tasaCruce,
            double tasaMutacion,
            int elitismo,
            IProgress<(int Generacion, double MejorFitness, int HardViolations, string Mensaje)> progreso,
            CancellationToken ct,
            double umbralFitness = 0.0,
            int estancamientoMax = 20)
        {
            if (contexto == null) return Result.Failure<AgResultadoCorrida>("El contexto del AG no puede ser nulo.");

            try
            {
                var stopwatch = Stopwatch.StartNew();
                var resultado = new AgResultadoCorrida();

                var aprobadas = ValidarAprobadas(contexto);
                if (aprobadas.IsFailure)
                    return Result.Failure<AgResultadoCorrida>(aprobadas.Error);

                // 1. Diagnóstico de Factibilidad Pre-vuelo
                var factibilidadResult = await _validadorVentana.ValidarFactibilidadAsync(contexto, ct);
            if (factibilidadResult.IsFailure)
            {
                return Result.Failure<AgResultadoCorrida>($"Fallo en validación de factibilidad: {factibilidadResult.Error}");
            }

            var informe = factibilidadResult.Value;
            if (!informe.EsFactible)
            {
                string errores = string.Join("\n• ", informe.BloqueosCriticos);
                // F1601: Se reemplaza «inviabilidad matemática» por «bloqueos de factibilidad detectados»
                return Result.Failure<AgResultadoCorrida>($"Corrida abortada por bloqueos de factibilidad detectados:\n• {errores}");
            }

            // 2. Construir Índices de Contexto
            contexto.ConstruirIndices();

            // 3. Inicializar Población
            var poblacion = new List<AgIndividuo>(poblacionTamano);
            for (int i = 0; i < poblacionTamano; i++)
            {
                ct.ThrowIfCancellationRequested();
                var ind = _fabricaIndividuos.CrearAleatorio(contexto, resultado.IdCorrida, 0);
                _evaluadorAptitud.EvaluarAptitud(ind, contexto);
                poblacion.Add(ind);
            }

            AgIndividuo mejorGlobal = poblacion.OrderByDescending(ind => ind.Fitness).First().Clonar();
            double mejorHist = double.NegativeInfinity;
            int estancadas = 0;

            // 4. Bucle Evolutivo
            for (int gen = 1; gen <= generacionesMax; gen++)
            {
                ct.ThrowIfCancellationRequested();
                var swGen = Stopwatch.StartNew();

                // Ordenar población por fitness descendente
                poblacion = poblacion.OrderByDescending(ind => ind.Fitness).ToList();

                var mejorGen = poblacion[0];
                if (mejorGen.Fitness > mejorGlobal.Fitness)
                {
                    mejorGlobal = mejorGen.Clonar();
                }

                if (mejorGlobal.Fitness > mejorHist + 1e-9)
                {
                    mejorHist = mejorGlobal.Fitness;
                    estancadas = 0;
                }
                else
                {
                    estancadas++;
                }

                // Bitácora de convergencia
                double promedioFit = poblacion.Average(ind => ind.Fitness);
                var bitacora = new AgBitacora
                {
                    Generacion = gen,
                    MejorFitness = mejorGen.Fitness,
                    PromedioFitness = promedioFit,
                    HardViolations = mejorGen.HardViolations,
                    DuracionMs = swGen.ElapsedMilliseconds,
                    Subunidad = contexto.Subunidad
                };
                resultado.HistorialConvergencia.Add(bitacora);

                // Reportar progreso UI
                progreso?.Report((gen, mejorGen.Fitness, mejorGen.HardViolations, $"Generación {gen}/{generacionesMax} - Mejor Fitness: {mejorGen.Fitness:F4} (Hard: {mejorGen.HardViolations})"));

                // Condición de parada por umbral
                bool sinViolaciones = mejorGlobal.HardViolations == 0;
                bool alcanzoUmbral = umbralFitness > 0 && mejorGlobal.Fitness >= umbralFitness;
                if (sinViolaciones && alcanzoUmbral)
                {
                    LoggerHelper.Info($"[AgGeneradorHorarios] Óptimo alcanzado en gen {gen}: Fitness={mejorGlobal.Fitness:F4}");
                    resultado.GeneracionesEjecutadas = gen;
                    break;
                }

                // Condición de parada por estancamiento
                if (estancamientoMax > 0 && estancadas >= estancamientoMax)
                {
                    LoggerHelper.Info($"[AgGeneradorHorarios] Estancamiento ({estancadas} generaciones) alcanzado en gen {gen}. Fin.");
                    resultado.GeneracionesEjecutadas = gen;
                    break;
                }

                if (gen == generacionesMax)
                {
                    resultado.GeneracionesEjecutadas = gen;
                    break;
                }

                // Crear nueva generación
                var nuevaPoblacion = new List<AgIndividuo>(poblacionTamano);

                // Elitismo
                int eliteCount = Math.Min(elitismo, poblacion.Count);
                for (int e = 0; e < eliteCount; e++)
                {
                    nuevaPoblacion.Add(poblacion[e].Clonar());
                }

                // Reproducción (Selección + Cruce + Mutación)
                while (nuevaPoblacion.Count < poblacionTamano)
                {
                    var padre1 = _operadorSeleccion.SeleccionarPorTorneo(poblacion);
                    var padre2 = _operadorSeleccion.SeleccionarPorTorneo(poblacion);

                    var (hijo1, hijo2) = _operadorCruce.Cruzar(padre1, padre2, tasaCruce, gen, contexto);

                    _operadorMutacion.Mutar(hijo1, tasaMutacion, contexto);
                    _operadorMutacion.Mutar(hijo2, tasaMutacion, contexto);

                    _evaluadorAptitud.EvaluarAptitud(hijo1, contexto);
                    _evaluadorAptitud.EvaluarAptitud(hijo2, contexto);

                    nuevaPoblacion.Add(hijo1);
                    if (nuevaPoblacion.Count < poblacionTamano)
                    {
                        nuevaPoblacion.Add(hijo2);
                    }
                }

                poblacion = nuevaPoblacion;
            }

            stopwatch.Stop();

            // Fase 582: sellado canónico + re-evaluación del campeón antes de publicarse
            AgCoherenciaGrupo.Sellar(mejorGlobal, contexto);
            _evaluadorAptitud.EvaluarAptitud(mejorGlobal, contexto);

            mejorGlobal.Estado = mejorGlobal.HardViolations == 0
                ? AgEstadoIndividuo.Optimo
                : AgEstadoIndividuo.Infactible;
            resultado.TiempoTotalMs = stopwatch.ElapsedMilliseconds;
            resultado.MejorIndividuo = mejorGlobal;
            resultado.Estado = AgEstadoCorrida.Completada;

            // Fase 214: Diagnóstico estricto de soluciones degeneradas con 0 genes
            if (mejorGlobal.Genes == null || mejorGlobal.Genes.Count == 0)
            {
                resultado.EsFactible = false;
                resultado.Estado = AgEstadoCorrida.Error;
                resultado.Mensaje = "Solución VACÍA: no se generaron tuplas (revise contexto/aulas/asignaturas).";
            }
            else
            {
                resultado.EsFactible = mejorGlobal.EsFactible;
                resultado.Mensaje = mejorGlobal.EsFactible
                    ? $"Optimización completada sin violaciones duras en {resultado.GeneracionesEjecutadas} generaciones ({stopwatch.ElapsedMilliseconds} ms)."
                    : $"Optimización finalizada con {mejorGlobal.HardViolations} restricciones duras no resueltas en {resultado.GeneracionesEjecutadas} generaciones.";
            }

            return Result.Success(resultado);
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<AgResultadoCorrida>("Optimización cancelada por el usuario.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgGeneradorHorarios] Error no controlado durante la optimización", ex);
                return Result.Failure<AgResultadoCorrida>($"Error durante la optimización: {ex.Message}");
            }
        }

        internal static Result ValidarAprobadas(AgContexto contexto)
        {
            if (contexto == null)
                return Result.Failure("El contexto del AG no puede ser nulo.");

            var aulasAprobadas = (contexto.Aulas ?? Enumerable.Empty<AgAula>())
                .Where(a => a.Estado == AgEstadoAula.Aprobado)
                .ToList();

            if (aulasAprobadas.Count == 0)
                return Result.Success();

            if (contexto.Ventana == null)
                return Result.Failure("I1_VENTANA_NO_DEFINIDA");

            foreach (var aula in aulasAprobadas)
            {
                var programa = contexto.Programas != null ? contexto.Programas.FirstOrDefault(p => p.Id == aula.IdPrograma) : null;
                double unidad = AgUnidadesSemanales.HorasSesionAula(aula,
                    AgNivelAcademicoCatalogo.Desde(programa?.Modalidad));
                var evaluacion = AgEvaluadorDemandaVentana.Evaluar(
                    aula.DiasVentana, aula.Jornada, unidad, contexto, aula,
                    contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin);
                if (!string.IsNullOrWhiteSpace(evaluacion.Error))
                    return Result.Failure(evaluacion.Error);
                if (!evaluacion.CabeEnPlazo)
                    return Result.Failure(
                        $"I5_AULA_APROBADA_CON_DEFICIT: {aula.AulaRef}; " +
                        "requiere revisión explícita de aprobación para replanificar. " +
                        "El escalador no modifica aulas aprobadas.");
            }
            return Result.Success();
        }
    }
}
