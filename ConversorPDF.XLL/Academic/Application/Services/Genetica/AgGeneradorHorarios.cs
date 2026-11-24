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

        /// <summary>
        /// Controla si se ejecuta la reparación post-evolución del campeón mediante forward shifting por carril (Fase 2130 / Fase 2143).
        /// Activo por defecto; puede desactivarse internamente para pruebas de replay/determinismo.
        /// </summary>
        internal static bool RealineacionOleadasActiva = true;

        /// <summary>
        /// Controla si se ejecuta la reparación post-evolución de sobrecargas R5 (Fase 2328).
        /// </summary>
        internal static bool ReparacionR5Activa = true;
        internal static AgModoReparacionR5 ReparacionR5Modo = AgModoReparacionR5.Completo;

        /// <summary>
        /// Peso agregado de las violaciones duras según la tabla de pesos (Fase 2363).
        /// Permite comparar reparaciones que intercambian reglas (R6=850 &gt; R5=800).
        /// </summary>
        internal static double PesoDuras(AgIndividuo ind, AgPesosFitness pesos)
        {
            if (ind?.Violaciones == null) return 0.0;
            double total = 0.0;
            foreach (var v in ind.Violaciones)
            {
                if (v.EsDura) total += pesos != null ? pesos.PorRegla(v.Regla) : 800.0;
            }
            return total;
        }

        /// <summary>Conteo de violaciones duras de una regla (Fase 2363).</summary>
        internal static int ConteoDuras(AgIndividuo ind, AgReglaViolacion regla)
        {
            if (ind?.Violaciones == null) return 0;
            int n = 0;
            foreach (var v in ind.Violaciones)
            {
                if (v.EsDura && v.Regla == regla) n++;
            }
            return n;
        }

        /// <summary>
        /// Criterio de aceptación de una reparación: menor peso, o mismo peso con menos R6 (Fase 2363).
        /// </summary>
        internal static bool MejoraReparacion(double pesoPre, int r6Pre, double pesoPost, int r6Post)
        {
            if (pesoPost < pesoPre - 0.001) return true;
            return Math.Abs(pesoPost - pesoPre) < 0.001 && r6Post < r6Pre;
        }

        private static string ConstruirEstadoFechas(AgIndividuo ind)
        {
            if (ind?.Genes == null) return string.Empty;
            var sb = new System.Text.StringBuilder(ind.Genes.Count * 20);
            for (int i = 0; i < ind.Genes.Count; i++)
            {
                var g = ind.Genes[i];
                sb.Append(i).Append(':')
                  .Append(g.FechaInicioEstimada.HasValue ? g.FechaInicioEstimada.Value.ToString("yyyyMMdd") : "-").Append(':')
                  .Append(g.FechaFinEstimada.HasValue ? g.FechaFinEstimada.Value.ToString("yyyyMMdd") : "-").Append(';');
            }
            return sb.ToString();
        }

        public AgGeneradorHorarios(
            IAgValidadorVentana validadorVentana,
            IAgEvaluadorAptitud evaluadorAptitud)
            : this(validadorVentana, evaluadorAptitud, (Random)null)
        {
        }

        public AgGeneradorHorarios(
            IAgValidadorVentana validadorVentana,
            IAgEvaluadorAptitud evaluadorAptitud,
            int seed)
            : this(validadorVentana, evaluadorAptitud, new Random(seed))
        {
        }

        public AgGeneradorHorarios(
            IAgValidadorVentana validadorVentana,
            IAgEvaluadorAptitud evaluadorAptitud,
            Random random)
        {
            _validadorVentana = validadorVentana ?? throw new ArgumentNullException(nameof(validadorVentana));
            _evaluadorAptitud = evaluadorAptitud ?? throw new ArgumentNullException(nameof(evaluadorAptitud));

            _fabricaIndividuos = random != null ? new AgFabricaIndividuos(random) : new AgFabricaIndividuos();
            _operadorSeleccion = random != null ? new AgOperadorSeleccion(random) : new AgOperadorSeleccion();
            var asignador = random != null ? new AgAsignadorTercero(random) : new AgAsignadorTercero();
            _operadorCruce = new AgOperadorCruce(random ?? new Random(), asignador);
            _operadorMutacion = random != null ? new AgOperadorMutacion(random) : new AgOperadorMutacion();
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
            int estancamientoMax = AgParametrosPorDefecto.EstancamientoMax)
        {
            if (contexto == null) return Result.Failure<AgResultadoCorrida>("El contexto del AG no puede ser nulo.");

            try
            {
                var stopwatch = Stopwatch.StartNew();
                var resultado = new AgResultadoCorrida
                {
                    IdCorrida = contexto.IdCorrida
                };

                var aprobadas = ValidarAprobadas(contexto);
                if (aprobadas.IsFailure)
                    return Result.Failure<AgResultadoCorrida>(aprobadas.Error);

                // 1. Diagnóstico de Factibilidad Pre-vuelo
                var swFact = Stopwatch.StartNew();
                var factibilidadResult = await _validadorVentana.ValidarFactibilidadAsync(contexto, ct);
                swFact.Stop();
                resultado.TiempoFactibilidadMs = swFact.ElapsedMilliseconds;

                if (factibilidadResult.IsFailure)
                {
                    return Result.Failure<AgResultadoCorrida>($"Fallo en validación de factibilidad: {factibilidadResult.Error}");
                }

                var informe = factibilidadResult.Value;
                if (!informe.EsFactible)
                {
                    string errores = string.Join("\n• ", informe.BloqueosCriticos);
                    return Result.Failure<AgResultadoCorrida>($"Corrida abortada por bloqueos de factibilidad detectados:\n• {errores}");
                }

                // 2. Construir Índices de Contexto
                var swIndices = Stopwatch.StartNew();
                contexto.ConstruirIndices();
                if (contexto.MaxVacantesForzadas < 0)
                {
                    string fuente = informe?.DeficitDocentePorAsignatura != null && informe.DeficitDocentePorAsignatura.Count > 0
                        ? "informe"
                        : "fallback";
                    contexto.MaxVacantesForzadas = AgValidadorVentana.CalcularDeficitEstructural(contexto, informe);
                    if (LoggerHelper.ModoVerboso)
                    {
                        LoggerHelper.Debug($"[AG-TOPE] MaxVacantesForzadas={contexto.MaxVacantesForzadas} fuente={fuente} asignaturasConsentidas={contexto.AsignaturasVacantesConsentidas?.Count ?? 0} claves={contexto.ClavesVacantesConsentidas?.Count ?? 0}");
                    }
                }
                else if (LoggerHelper.ModoVerboso)
                {
                    LoggerHelper.Debug($"[AG-TOPE] MaxVacantesForzadas={contexto.MaxVacantesForzadas} fuente=predefinido asignaturasConsentidas={contexto.AsignaturasVacantesConsentidas?.Count ?? 0} claves={contexto.ClavesVacantesConsentidas?.Count ?? 0}");
                }
                swIndices.Stop();
                resultado.TiempoIndicesMs = swIndices.ElapsedMilliseconds;

                // 3. Inicializar Población
                var swPob = Stopwatch.StartNew();
                var cacheAptitud = new Dictionary<long, AgEntradaCacheAptitud>();
                var poblacion = new List<AgIndividuo>(poblacionTamano);
                for (int i = 0; i < poblacionTamano; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    var ind = _fabricaIndividuos.CrearAleatorio(contexto, resultado.IdCorrida, 0);
                    _evaluadorAptitud.EvaluarAptitud(ind, contexto);
                    poblacion.Add(ind);
                    cacheAptitud[ind.CalcularHuellaEstructural()] = new AgEntradaCacheAptitud(ind);
                }
                swPob.Stop();
                resultado.TiempoPoblacionInicialMs = swPob.ElapsedMilliseconds;

                var swPrimerCiclo = Stopwatch.StartNew();
                AgIndividuo mejorGlobal = (poblacion.Count > 0 ? poblacion.OrderByDescending(ind => ind.Fitness).FirstOrDefault() : null)?.Clonar() ?? new AgIndividuo();
                double mejorHist = double.NegativeInfinity;
                int estancadas = 0;
                int mejorGenNumero = 0;
                var parada = AgMotivoParada.NoDefinida;
                long duracionUltimaTransicionMs = resultado.TiempoPoblacionInicialMs;

                // 4. Bucle Evolutivo
                for (int gen = 1; gen <= generacionesMax; gen++)
                {
                    ct.ThrowIfCancellationRequested();

                    // Ordenar población por fitness descendente
                    poblacion = poblacion.OrderByDescending(ind => ind.Fitness).ToList();

                    var mejorGen = poblacion[0];
                    var peorGen = poblacion[poblacion.Count - 1];
                    if (mejorGen.Fitness > mejorGlobal.Fitness)
                    {
                        mejorGlobal = mejorGen.Clonar();
                    }

                    if (mejorGlobal.Fitness > mejorHist + 1e-9)
                    {
                        mejorHist = mejorGlobal.Fitness;
                        estancadas = 0;
                        mejorGenNumero = gen;
                    }
                    else
                    {
                        estancadas++;
                    }

                    if (gen == 1 && swPrimerCiclo.IsRunning)
                    {
                        swPrimerCiclo.Stop();
                        resultado.TiempoPrimerCicloMs = swPrimerCiclo.ElapsedMilliseconds;
                    }

                    // Bitácora de convergencia
                    double promedioFit = poblacion.Average(ind => ind.Fitness);
                    int vacantes = mejorGen.Vacantes;
                    int fueraDeVentana = mejorGen.Genes != null ? mejorGen.Genes.Count(g => g.DentroDeVentana == 0) : 0;
                    long duracionGenMs = gen == 1 ? (resultado.TiempoPoblacionInicialMs + resultado.TiempoPrimerCicloMs) : duracionUltimaTransicionMs;
                    int evaluadosGen = gen == 1 ? poblacionTamano : (resultado.PerfilesGeneracion.Count > 0 ? resultado.PerfilesGeneracion[resultado.PerfilesGeneracion.Count - 1].EvaluacionesAptitud : poblacionTamano);

                    var bitacora = new AgBitacora
                    {
                        Generacion = gen,
                        MejorFitness = mejorGen.Fitness,
                        PromedioFitness = promedioFit,
                        PeorFitness = peorGen.Fitness,
                        HardViolations = mejorGen.HardViolations,
                        Evaluados = evaluadosGen,
                        Vacantes = vacantes,
                        FueraDeVentana = fueraDeVentana,
                        DuracionMs = duracionGenMs,
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
                        parada = AgMotivoParada.Umbral;
                        resultado.GeneracionesEjecutadas = gen;
                        break;
                    }

                    // Condición de parada por estancamiento
                    if (estancamientoMax > 0 && estancadas >= estancamientoMax)
                    {
                        LoggerHelper.Info($"[AgGeneradorHorarios] Estancamiento ({estancadas} generaciones) alcanzado en gen {gen}. Fin.");
                        parada = AgMotivoParada.Estancamiento;
                        resultado.GeneracionesEjecutadas = gen;
                        break;
                    }

                    if (gen == generacionesMax)
                    {
                        parada = AgMotivoParada.MaxGeneraciones;
                        resultado.GeneracionesEjecutadas = gen;
                        break;
                    }

                    // Crear nueva generación y medir transición reproductiva
                    var swTransicion = Stopwatch.StartNew();
                    var perfil = new AgResultadoCorrida.AgPerfilGeneracion
                    {
                        Generacion = gen,
                        ElitesClonados = Math.Min(elitismo, poblacion.Count),
                        Vacantes = vacantes,
                        FueraDeVentana = fueraDeVentana
                    };

                    long acumSel = 0, acumCruce = 0, acumMut = 0, acumEval = 0;
                    int aciertosGen = 0;
                    var swOp = new Stopwatch();

                    var nuevaPoblacion = new List<AgIndividuo>(poblacionTamano);
                    for (int e = 0; e < perfil.ElitesClonados; e++)
                    {
                        nuevaPoblacion.Add(poblacion[e].Clonar());
                    }

                    while (nuevaPoblacion.Count < poblacionTamano)
                    {
                        swOp.Restart();
                        var padre1 = _operadorSeleccion.SeleccionarPorTorneo(poblacion);
                        var padre2 = _operadorSeleccion.SeleccionarPorTorneo(poblacion);
                        swOp.Stop();
                        acumSel += swOp.ElapsedMilliseconds;

                        swOp.Restart();
                        var (hijo1, hijo2) = _operadorCruce.Cruzar(padre1, padre2, tasaCruce, gen, contexto);
                        swOp.Stop();
                        acumCruce += swOp.ElapsedMilliseconds;

                        swOp.Restart();
                        _operadorMutacion.Mutar(hijo1, tasaMutacion, contexto);
                        _operadorMutacion.Mutar(hijo2, tasaMutacion, contexto);
                        swOp.Stop();
                        acumMut += swOp.ElapsedMilliseconds;

                        swOp.Restart();
                        EvaluarConCache(hijo1, contexto, cacheAptitud, ref aciertosGen);
                        EvaluarConCache(hijo2, contexto, cacheAptitud, ref aciertosGen);
                        swOp.Stop();
                        acumEval += swOp.ElapsedMilliseconds;

                        perfil.HijosCreados += 2;
                        perfil.EvaluacionesAptitud += 2;

                        nuevaPoblacion.Add(hijo1);
                        perfil.HijosIncorporados++;

                        if (nuevaPoblacion.Count < poblacionTamano)
                        {
                            nuevaPoblacion.Add(hijo2);
                            perfil.HijosIncorporados++;
                        }
                        else
                        {
                            perfil.EvaluacionesExtraPorCupoImpar++;
                        }
                    }

                    swTransicion.Stop();
                    duracionUltimaTransicionMs = swTransicion.ElapsedMilliseconds;
                    perfil.DuracionTotalMs = duracionUltimaTransicionMs;
                    perfil.TiempoSeleccionMs = acumSel;
                    perfil.TiempoCruceNetoMs = acumCruce;
                    perfil.TiempoMutacionNetoMs = acumMut;
                    perfil.TiempoEvaluacionMs = acumEval;
                    resultado.PerfilesGeneracion.Add(perfil);

                    LoggerHelper.Info($"[AG-PERF] Gen {gen}->{gen + 1}: Duracion={perfil.DuracionTotalMs}ms | Sel={perfil.TiempoSeleccionMs}ms, CruceNeto={perfil.TiempoCruceNetoMs}ms, MutNeto={perfil.TiempoMutacionNetoMs}ms, Eval={perfil.TiempoEvaluacionMs}ms | Evals={perfil.EvaluacionesAptitud} (Hijos={perfil.HijosCreados}, Incorp={perfil.HijosIncorporados}, Elites={perfil.ElitesClonados}, Impar={perfil.EvaluacionesExtraPorCupoImpar}, CacheHits={aciertosGen}) | MejorFit={mejorGen.Fitness:F4}, PeorFit={peorGen.Fitness:F4}, Vac={vacantes}");

                    poblacion = nuevaPoblacion;
                }

                stopwatch.Stop();

                // Fase 582 / Fase 2130: sellado canónico + re-evaluación y reparación de campeón (R1 Forward Shifting)
                var swCamp = Stopwatch.StartNew();
                AgCoherenciaGrupo.Sellar(mejorGlobal, contexto);

                void EvaluarCampeon()
                {
                    if (_evaluadorAptitud is AgEvaluadorAptitud evalAptInstance)
                    {
                        evalAptInstance.EvaluarAptitud(mejorGlobal, contexto, forzar: true, incluirDetalleJson: true);
                    }
                    else
                    {
                        _evaluadorAptitud.EvaluarAptitud(mejorGlobal, contexto);
                    }
                }

                EvaluarCampeon();
                int hardPre = mejorGlobal.HardViolations;
                int vacPre = mejorGlobal.Vacantes;

                if (RealineacionOleadasActiva)
                {
                    // Fase 2363: multipasada con aceptación por peso (R6=850 > R5=800) y hash de estado
                    var pesos = contexto.Pesos ?? AgPesosFitness.PorDefecto();
                    var aceptado = mejorGlobal.Clonar();
                    double pesoOk = PesoDuras(aceptado, pesos);
                    int r6Ok = ConteoDuras(aceptado, AgReglaViolacion.R6_SecuenciaAsignaturas);
                    int carrilesRep = 0, carrilesEval = 0, asigsDesp = 0, fallos = 0;
                    string errorRep = string.Empty;
                    var detalleRep = new List<string>();
                    var visitados = new HashSet<string>(StringComparer.Ordinal);

                    for (int pasada = 0; pasada < 3; pasada++)
                    {
                        var rep = AgPlanificadorOleadas.RealinearOleadas(mejorGlobal, contexto);
                        carrilesEval = rep.CarrilesEvaluados;
                        if (rep.CarrilesReparados > carrilesRep) carrilesRep = rep.CarrilesReparados;
                        asigsDesp += rep.AsignaturasDesplazadas;
                        fallos += rep.Fallos;
                        if (!string.IsNullOrEmpty(rep.Error)) errorRep = rep.Error;
                        if (rep.Detalle != null && rep.Detalle.Count > 0) detalleRep.AddRange(rep.Detalle);

                        if (!rep.HuboCambios) break;

                        AgCoherenciaGrupo.Sellar(mejorGlobal, contexto);
                        EvaluarCampeon();

                        double peso = PesoDuras(mejorGlobal, pesos);
                        int r6 = ConteoDuras(mejorGlobal, AgReglaViolacion.R6_SecuenciaAsignaturas);
                        if (!MejoraReparacion(pesoOk, r6Ok, peso, r6))
                        {
                            mejorGlobal = aceptado;
                            detalleRep.Add($"rollback: peso {pesoOk:F0}->{peso:F0} r6 {r6Ok}->{r6}");
                            break;
                        }

                        string estado = ConstruirEstadoFechas(mejorGlobal);
                        if (!visitados.Add(estado))
                        {
                            mejorGlobal = aceptado;
                            detalleRep.Add("rollback: ciclo detectado");
                            break;
                        }

                        aceptado = mejorGlobal.Clonar();
                        pesoOk = peso;
                        r6Ok = r6;
                    }

                    LoggerHelper.Info(
                        $"[AG-REPARACION] Hard {hardPre}->{mejorGlobal.HardViolations} peso={pesoOk:F0} r6={r6Ok} " +
                        $"carriles={carrilesRep}/{carrilesEval} asigs={asigsDesp} fallos={fallos} " +
                        $"vacantes {vacPre}->{mejorGlobal.Vacantes} " +
                        $"{(errorRep.Length > 0 ? "error=" + errorRep : string.Empty)}");

                    if (detalleRep.Count > 0 && LoggerHelper.ModoVerboso)
                    {
                        foreach (var d in detalleRep)
                        {
                            LoggerHelper.Debug($"[AG-REPARACION] {d}");
                        }
                    }
                }

                // Fase 2328: Hook de reparación automática de sobrecargas R5
                if (ReparacionR5Activa && ReparacionR5Modo != AgModoReparacionR5.Off && mejorGlobal.HardViolations > 0)
                {
                    var repR5 = AgReparadorR5.Ejecutar(mejorGlobal, contexto, _evaluadorAptitud, ReparacionR5Modo);
                    if (repR5 != null)
                    {
                        resultado.ReparacionR5Intentos = repR5.Intentos;
                        resultado.ReparacionR5HardPre = repR5.HardPre;
                        resultado.ReparacionR5HardPost = repR5.HardPost;
                        resultado.ReparacionR5DesperdicioPre = repR5.DesperdicioPre;
                        resultado.ReparacionR5DesperdicioPost = repR5.DesperdicioPost;
                        resultado.ReparacionR5RippleIntentos = repR5.RippleIntentos;
                        resultado.ReparacionR5RippleSinCola = repR5.RippleSinCola;

                        if (repR5.HuboMejora && repR5.Individuo != null)
                        {
                            mejorGlobal = repR5.Individuo;
                        }
                    }
                }

                swCamp.Stop();
                resultado.TiempoCampeonFinalMs = swCamp.ElapsedMilliseconds;

                // Fase 1979/1983/1984/1995: Telemetría causal estructurada. Las restricciones HARD se desglosan por regla
                // calculando descuadre con signo frente al contador (entradas - contador); las vacantes por motivo con
                // el mismo predicado estricto que AgIndividuo.Vacantes (EsVacante == 1 || IdTercero <= 0).
                var duras = (mejorGlobal.Violaciones ?? new List<AgViolacion>())
                    .Where(v => v.EsDura)
                    .GroupBy(v => v.Regla)
                    .Select(g => new { Regla = g.Key, Total = g.Count() })
                    .ToList();
                var porRegla = duras
                    .Select(g => $"{g.Regla.ToString()}x{g.Total}")
                    .ToList();
                int detalleDuro = duras.Sum(g => g.Total);
                int descuadre = detalleDuro - mejorGlobal.HardViolations;

                var porMotivo = (mejorGlobal.Genes ?? new List<AgGen>())
                    .Where(g => g.EsVacante == 1 || g.IdTercero <= 0)
                    .GroupBy(g => string.IsNullOrEmpty(g.MotivoVacante) ? "SIN_MOTIVO" : g.MotivoVacante)
                    .Select(g => $"{g.Key}x{g.Count()}")
                    .ToList();

                // Fase 2018 / Fase 2022: Métrica de vacantes forzadas excluyendo consentidas
                var consentidas = contexto.AsignaturasVacantesConsentidas ?? new HashSet<int>();
                int vacantesForzadas = (mejorGlobal.Genes ?? new List<AgGen>())
                    .Count(g => g.EsVacante == 1 && g.IdTercero <= 0 && !consentidas.Contains(g.IdAsignatura));

                string maxForz = contexto.MaxVacantesForzadas >= 0
                    ? contexto.MaxVacantesForzadas.ToString(System.Globalization.CultureInfo.InvariantCulture) : "n/a";

                resultado.MotivoParada = parada;
                resultado.Estancadas = estancadas;
                LoggerHelper.Info(
                    $"[AG-CAUSAS] Corrida {resultado.IdCorrida}: Parada={parada} " +
                    $"Gen={resultado.GeneracionesEjecutadas}/{generacionesMax} Estancadas={estancadas}/{estancamientoMax} " +
                    $"HARD={mejorGlobal.HardViolations} Detalle={detalleDuro} Descuadre={descuadre:+#;-#;0} " +
                    $"[{string.Join(",", porRegla)}] Vacantes={mejorGlobal.Vacantes} Forzadas={vacantesForzadas} MaxForzadas={maxForz} [{string.Join(",", porMotivo)}]");

                // Fase 1994: Desglose forense del campeón infactible (causalidad de conflictos y vacantes)
                if (mejorGlobal.HardViolations > 0)
                {
                    try
                    {
                        foreach (var linea in AgDiagnosticoInfactibilidad.Describir(mejorGlobal, contexto))
                        {
                            LoggerHelper.Info(linea);
                        }
                    }
                    catch (Exception exDiag)
                    {
                        LoggerHelper.WarningLimitado(
                            $"DiagInfactibleFallido|{resultado.IdCorrida}",
                            $"[AG-CAUSAS-DETALLE] no generado: {exDiag.Message}");
                    }
                }

                // Fase 2086: Diagnóstico forense detallado en modo verboso ([AG-HARD], [AG-VAC], [AG-SNAPSHOT], [AG-PARADA])
                if (LoggerHelper.ModoVerboso)
                {
                    LoggerHelper.Debug($"[AG-HARD] campeon HARD={mejorGlobal.HardViolations} detalle={detalleDuro} descuadre={descuadre:+#;-#;0} reglas=[{string.Join(", ", duras.Select(d => $"{d.Regla} x{d.Total}"))}]");

                    int vIdx = 1;
                    var listaDuras = (mejorGlobal.Violaciones ?? new List<AgViolacion>()).Where(v => v.EsDura).Take(50).ToList();
                    foreach (var v in listaDuras)
                    {
                        AgGen genRef = null;
                        if (v.IdGen.HasValue && mejorGlobal.Genes != null)
                        {
                            genRef = mejorGlobal.Genes.FirstOrDefault(g => g.Id == v.IdGen.Value);
                        }

                        int idAsig = genRef?.IdAsignatura ?? 0;
                        int idSede = genRef?.IdSede ?? 0;
                        int idAula = genRef?.IdAula ?? 0;
                        int idTercero = genRef?.IdTercero ?? 0;
                        string consentidaStr = genRef != null
                            ? (contexto.EstaConsentida(idAsig, idSede) ? "True" : "False")
                            : "—";

                        LoggerHelper.Debug($"[AG-HARD] v{vIdx++} regla={v.Regla} gen={v.IdGen ?? 0} asig={idAsig} sede={idSede} aula={idAula} tercero={idTercero} consentida={consentidaStr} msg=\"{v.Mensaje}\"");
                    }

                    // [AG-VAC] (Fase 2086, 2092)
                    var genesVac = (mejorGlobal.Genes ?? new List<AgGen>())
                        .Where(g => g.EsVacante == 1 || g.IdTercero <= 0)
                        .ToList();

                    int cConsentidas = genesVac.Count(g => contexto.EstaConsentida(g.IdAsignatura, g.IdSede));
                    var estructuralesSesiones = new Dictionary<int, int>();
                    var cohortesLiberadasSet = new HashSet<string>(StringComparer.Ordinal);
                    int cLiberadasSesiones = 0;
                    int cPoolVacio = 0;

                    foreach (var g in genesVac)
                    {
                        if (contexto.EstaConsentida(g.IdAsignatura, g.IdSede)) continue;

                        bool poolVacio = !contexto.PoolPorAsignatura.TryGetValue(g.IdAsignatura, out var pl) || pl == null || pl.Count == 0 || !pl.Any(p => p.Activo == 1);
                        if (poolVacio) cPoolVacio++;

                        var origen = AgClasificadorVacante.Clasificar(g.MotivoVacante);
                        if (origen == AgOrigenVacante.EstructuralDocente)
                        {
                            estructuralesSesiones[g.IdAsignatura] = estructuralesSesiones.TryGetValue(g.IdAsignatura, out var n) ? n + 1 : 1;
                        }
                        else if (origen == AgOrigenVacante.LiberadaAg)
                        {
                            cLiberadasSesiones++;
                            cohortesLiberadasSet.Add(AgClasificadorVacante.ClaveCohorte(g));
                        }
                    }

                    int cLiberadasCohortes = cohortesLiberadasSet.Count;
                    string estSesionesStr = string.Join(",", estructuralesSesiones.Select(kv => $"{kv.Key}x{kv.Value}"));
                    LoggerHelper.Debug(
                        $"[AG-VAC] consentidasEnGenes={cConsentidas} liberadasCohortes={cLiberadasCohortes} " +
                        $"liberadasSesiones={cLiberadasSesiones} estructuralesAsignaturas={estructuralesSesiones.Count} " +
                        $"estructuralesSesiones=[{estSesionesStr}] poolVacio={cPoolVacio}");

                    var liberadasGenes = genesVac
                        .Where(g => !contexto.EstaConsentida(g.IdAsignatura, g.IdSede) &&
                                    AgClasificadorVacante.Clasificar(g.MotivoVacante) == AgOrigenVacante.LiberadaAg)
                        .Take(50)
                        .ToList();

                    foreach (var lg in liberadasGenes)
                    {
                        LoggerHelper.Debug($"[AG-VAC] liberada asig={lg.IdAsignatura} sede={lg.IdSede} aula={lg.IdAula} cohorte={lg.Semestre}/{lg.Grupo} gen={lg.Id} motivo={lg.MotivoVacante} consentida={contexto.EstaConsentida(lg.IdAsignatura, lg.IdSede)}");
                    }

                    // [AG-SNAPSHOT] para genes que originaron violaciones o vacantes liberadas
                    var genesSnapshot = new List<AgGen>();
                    foreach (var v in listaDuras)
                    {
                        if (v.IdGen.HasValue)
                        {
                            var g = mejorGlobal.Genes?.FirstOrDefault(x => x.Id == v.IdGen.Value);
                            if (g != null && !genesSnapshot.Contains(g)) genesSnapshot.Add(g);
                        }
                        else if (!string.IsNullOrWhiteSpace(v.DetalleJson))
                        {
                            try
                            {
                                var match = System.Text.RegularExpressions.Regex.Match(v.DetalleJson, @"\""indicesGenes\""\s*:\s*\[([^\]]+)\]");
                                if (match.Success)
                                {
                                    var idxStrs = match.Groups[1].Value.Split(',');
                                    foreach (var s in idxStrs)
                                    {
                                        if (int.TryParse(s.Trim(), out int idx) && idx >= 0 && idx < (mejorGlobal.Genes?.Count ?? 0))
                                        {
                                            var g = mejorGlobal.Genes[idx];
                                            if (!genesSnapshot.Contains(g)) genesSnapshot.Add(g);
                                        }
                                    }
                                }
                            }
                            catch (Exception exSnap)
                            {
                                LoggerHelper.Debug($"[AG-SNAPSHOT] Error extrayendo genes: {exSnap.Message}");
                            }
                        }
                    }

                    if (duras.Any(d => d.Regla == AgReglaViolacion.R15_VacanteExcedida || d.Regla == AgReglaViolacion.R16_VacanteEstructuralNoConsentida))
                    {
                        foreach (var lg in liberadasGenes)
                        {
                            if (!genesSnapshot.Contains(lg)) genesSnapshot.Add(lg);
                        }
                    }

                    foreach (var sg in genesSnapshot.Take(50))
                    {
                        int idTerceroReal = 0;
                        if (contexto.PoolPorAsignatura != null && contexto.PoolPorAsignatura.TryGetValue(sg.IdAsignatura, out var pl) && pl != null)
                        {
                            var primeroActivo = pl.FirstOrDefault(p => p.Activo == 1 && (contexto.IdTerceroVacante <= 0 || p.IdTercero != contexto.IdTerceroVacante));
                            if (primeroActivo != null) idTerceroReal = primeroActivo.IdTercero;
                        }

                        string franjaStr = contexto.BloquesPorId != null && contexto.BloquesPorId.TryGetValue(sg.IdBloqueHorario, out var blq) ? (blq.Franja ?? blq.HorarioTexto ?? blq.Id.ToString()) : "N/D";
                        LoggerHelper.Debug($"[AG-SNAPSHOT] gen={sg.Id} asig={sg.IdAsignatura} prog={sg.IdPrograma} sede={sg.IdSede} sem={sg.Semestre} grupo={sg.Grupo} aula={sg.IdAula} bloque={sg.IdBloqueHorario} dia={sg.DiaSemana} patrón={sg.PatronSemanal} franja={franjaStr} tercero={sg.IdTercero} esVacante={sg.EsVacante} motivo={sg.MotivoVacante} terceroRealPool={idTerceroReal}");
                    }

                    // [AG-PARADA]
                    LoggerHelper.Debug($"[AG-PARADA] motivo={parada} gen={resultado.GeneracionesEjecutadas}/{generacionesMax} estancadas={estancadas}/{estancamientoMax} mejorGen={mejorGenNumero}");
                    if (resultado.HistorialConvergencia != null && resultado.HistorialConvergencia.Count > 0)
                    {
                        var ultimas = resultado.HistorialConvergencia.Skip(Math.Max(0, resultado.HistorialConvergencia.Count - 10)).ToList();
                        string histStr = string.Join(" | ", ultimas.Select(b => $"{b.Generacion}:f={b.MejorFitness:F2},h={b.HardViolations},v={b.Vacantes}"));
                        LoggerHelper.Debug($"[AG-PARADA] hist=[{histStr}]");
                    }
                }

                mejorGlobal.Estado = mejorGlobal.HardViolations == 0
                    ? AgEstadoIndividuo.Optimo
                    : AgEstadoIndividuo.Infactible;
                resultado.TiempoTotalMs = stopwatch.ElapsedMilliseconds;
                resultado.MejorIndividuo = mejorGlobal;
                resultado.Estado = mejorGlobal.HardViolations == 0
                    ? AgEstadoCorrida.Completada
                    : AgEstadoCorrida.SinFactibilidad;

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

            // Fase 1848 & Fase 1851: Fail-closed si el manifiesto declara aulas Aprobado pero no hay snapshot canónico o es divergente.
            if (contexto.ManifiestoCobertura != null)
            {
                bool declaraAprobadas = contexto.ManifiestoCobertura.Relaciones.Any(r => r.EstadoAula == (int)AgEstadoAula.Aprobado);
                var s = contexto.PlanAprobadoSnapshot;
                if (declaraAprobadas && s == null)
                    return Result.Failure("I5_SNAPSHOT_APROBADO_REQUERIDO: manifiesto con aulas Aprobado sin snapshot canónico.");
                if (declaraAprobadas && (s.Manifiesto.Huella != contexto.ManifiestoCobertura.Huella
                    || !s.Manifiesto.Relaciones.Select(r => r.Clave.Valor)
                         .SequenceEqual(contexto.ManifiestoCobertura.Relaciones.Select(r => r.Clave.Valor))))
                    return Result.Failure("I5_SNAPSHOT_DIVERGENTE: el snapshot canónico no corresponde al manifiesto vigente.");
            }

            var aulasAprobadas = (contexto.Aulas ?? Enumerable.Empty<AgAula>())
                .Where(a => a.Estado == AgEstadoAula.Aprobado
                    && (contexto.ManifiestoCobertura == null || contexto.TieneOrigenConservable(a.Id)))
                .ToList();

            if (aulasAprobadas.Count == 0)
                return Result.Success();

            if (contexto.ManifiestoCobertura != null)
            {
                if (contexto.SesionesConservadas == null || contexto.CandidatosCobertura.Count != aulasAprobadas.Count ||
                    aulasAprobadas.Any(a => !contexto.CandidatosCobertura.Any(c => c.IdAula == a.Id &&
                        (c.Desde == c.Hacia || contexto.PuedeGenerarAula(a.Id)))))
                    return Result.Failure("I5_AULA_APROBADA_SIN_RATIFICACION");
                return Result.Success(); // La cobertura final del manifiesto se comprueba después de materializar.
            }

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

        private void EvaluarConCache(
            AgIndividuo individuo,
            AgContexto contexto,
            Dictionary<long, AgEntradaCacheAptitud> cacheAptitud,
            ref int aciertosCache)
        {
            if (individuo == null) return;

            // Clones exactos ya evaluados en su progenitor
            if (individuo.Estado != AgEstadoIndividuo.NoEvaluado)
            {
                return;
            }

            long huella = individuo.CalcularHuellaEstructural();
            if (cacheAptitud != null && cacheAptitud.TryGetValue(huella, out var entrada))
            {
                entrada.AplicarA(individuo);
                aciertosCache++;
                return;
            }

            _evaluadorAptitud.EvaluarAptitud(individuo, contexto);

            if (cacheAptitud != null)
            {
                cacheAptitud[huella] = new AgEntradaCacheAptitud(individuo);
            }
        }

        private struct AgEntradaCacheAptitud
        {
            public readonly double Fitness;
            public readonly int HardViolations;
            public readonly double SoftViolations;
            public readonly AgEstadoIndividuo Estado;
            public readonly List<AgViolacion> Violaciones;

            public AgEntradaCacheAptitud(AgIndividuo ind)
            {
                Fitness = ind.Fitness;
                HardViolations = ind.HardViolations;
                SoftViolations = ind.SoftViolations;
                Estado = ind.Estado;
                Violaciones = ind.Violaciones != null ? new List<AgViolacion>(ind.Violaciones) : new List<AgViolacion>();
            }

            public void AplicarA(AgIndividuo ind)
            {
                ind.Fitness = Fitness;
                ind.HardViolations = HardViolations;
                ind.SoftViolations = SoftViolations;
                ind.Estado = Estado;
                ind.Violaciones = new List<AgViolacion>(Violaciones);
            }
        }
    }
}
