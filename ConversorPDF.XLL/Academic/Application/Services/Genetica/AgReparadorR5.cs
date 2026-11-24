using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Motor heurístico de reparación automática de sobrecargas R5 mediante inventario de ranuras libres,
    /// forward shifting, compresión y re-ranuración determinista (Fases 2325–2327 / 2332–2340 / 2343–2351).
    /// </summary>
    internal sealed class AgReparadorR5
    {
        public const int MaxIntentos = 16;
        public const int MaxEvaluaciones = 20;
        public const int MaxTiempoMs = 3000;

        /// <summary>
        /// Sobrecarga directa compatible con el hook de AgGeneradorHorarios (Fase 2328).
        /// </summary>
        public static AgIndividuo Reparar(
            AgIndividuo campeon,
            AgContexto contexto,
            IAgEvaluadorAptitud evaluadorAptitud,
            AgModoReparacionR5 modo = AgModoReparacionR5.Completo)
        {
            var res = Ejecutar(campeon, contexto, evaluadorAptitud, modo);
            return (res != null && res.HuboMejora && res.Individuo != null)
                ? res.Individuo
                : campeon;
        }

        /// <summary>
        /// Ejecuta el proceso completo de diagnóstico y reparación R5 devolviendo telemetría detallada.
        /// </summary>
        public static AgResultadoReparacionR5 Ejecutar(
            AgIndividuo campeon,
            AgContexto contexto,
            IAgEvaluadorAptitud evaluadorAptitud,
            AgModoReparacionR5 modo = AgModoReparacionR5.Completo)
        {
            var resultado = new AgResultadoReparacionR5
            {
                Individuo = campeon,
                HardPre = campeon?.HardViolations ?? 0,
                HardPost = campeon?.HardViolations ?? 0,
                HuboMejora = false
            };

            if (campeon == null || contexto == null || evaluadorAptitud == null || modo == AgModoReparacionR5.Off)
            {
                return resultado;
            }

            if (campeon.HardViolations <= 0)
            {
                return resultado;
            }

            var sw = Stopwatch.StartNew();
            double despPre = 0.0;
            string primerCohorte = null;

            try
            {
                var sobrecargas = Detectar(campeon, contexto);
                if (sobrecargas.Count == 0)
                {
                    return resultado;
                }

                primerCohorte = sobrecargas[0].Cohorte;
                var invInicial = AgInventarioRanuras.Construir(campeon, contexto, primerCohorte);
                despPre = CalcularDesperdicio(invInicial);
                resultado.DesperdicioPre = despPre;
                resultado.DesperdicioPost = despPre;

                var puntajeMejor = new AgPuntajeReparacion(campeon, despPre, 0);
                AgIndividuo mejorClon = null;
                var actual = campeon; // Fase 2339: Iteración acumulativa
                string cohorteReparada = null;

                int evaluaciones = 0;
                int intentosEvaluados = 0;
                int intentosAceptados = 0;

                while (intentosAceptados < MaxIntentos && evaluaciones < MaxEvaluaciones && sw.ElapsedMilliseconds < MaxTiempoMs)
                {
                    var sobrecargasActual = Detectar(actual, contexto);
                    if (sobrecargasActual.Count == 0)
                    {
                        break;
                    }

                    bool mejoraEnIteracion = false;

                    // Fase 2364: duras preexistentes (R2/R4/R6) que no deben bloquear al candidato
                    var preexistentes = ClavesDuras(actual);

                    foreach (var sob in sobrecargasActual)
                    {
                        if (intentosAceptados >= MaxIntentos || evaluaciones >= MaxEvaluaciones || sw.ElapsedMilliseconds >= MaxTiempoMs)
                        {
                            break;
                        }

                        var inv = AgInventarioRanuras.Construir(actual, contexto, sob.Cohorte);
                        // Fase 2360: Calcular desperdicio pre-reparación sobre la cohorte de la sobrecarga actual
                        var invPre = inv;
                        double despPreCohorte = CalcularDesperdicio(invPre);
                        var puntajeBaseCohorte = new AgPuntajeReparacion(actual, despPreCohorte, 0);

                        // Candidatos C1: Desplazamiento (Forward Shifting) sobre ranuras libres (Fase 2347)
                        var candidatosC1 = GenerarDesplazamientos(actual, contexto, sob, inv);
                        foreach (var cand in candidatosC1)
                        {
                            intentosEvaluados++;
                            if (evaluaciones >= MaxEvaluaciones || sw.ElapsedMilliseconds >= MaxTiempoMs)
                            {
                                break;
                            }

                            // Fase 2348: Sellar primero, luego Validar
                            AgCoherenciaGrupo.Sellar(cand, contexto);
                            if (cand.Vacantes > actual.Vacantes)
                            {
                                continue;
                            }
                            if (!Validar(cand, actual, contexto, preexistentes))
                            {
                                continue;
                            }

                            // Fase 2332: Forzar reevaluación reiniciando Estado
                            cand.Estado = AgEstadoIndividuo.NoEvaluado;
                            evaluadorAptitud.EvaluarAptitud(cand, contexto, forzar: true, incluirDetalleJson: true);
                            evaluaciones++;

                            var invCand = AgInventarioRanuras.Construir(cand, contexto, sob.Cohorte);
                            double despCand = CalcularDesperdicio(invCand);
                            var puntajeCand = new AgPuntajeReparacion(cand, despCand, 1);

                            if (puntajeCand.Hard < puntajeBaseCohorte.Hard || (puntajeCand.Hard == puntajeBaseCohorte.Hard && puntajeCand.CompareTo(puntajeBaseCohorte) < 0))
                            {
                                if (puntajeCand.Hard < actual.HardViolations)
                                {
                                    int hardAntes = actual.HardViolations; // Fase 2351: Capturar antes de actualizar
                                    puntajeMejor = puntajeCand;
                                    mejorClon = cand;
                                    actual = cand;
                                    cohorteReparada = sob.Cohorte;
                                    intentosAceptados++;
                                    mejoraEnIteracion = true;
                                    resultado.Detalle.Add($"[AG-REPARACION-R5] C1 cohorte={sob.Cohorte} hard {hardAntes}->{puntajeCand.Hard}, Desp {despPreCohorte:F3}->{despCand:F3}");
                                    break;
                                }
                            }
                        }

                        if (mejoraEnIteracion) break;
                        if (modo == AgModoReparacionR5.SoloDesplazamiento) continue;

                        // Candidatos C2: Compresión real sobre huecos libres (Fase 2349)
                        var candidatosC2 = GenerarCompresiones(actual, contexto, sob, inv);
                        foreach (var cand in candidatosC2)
                        {
                            intentosEvaluados++;
                            if (evaluaciones >= MaxEvaluaciones || sw.ElapsedMilliseconds >= MaxTiempoMs)
                            {
                                break;
                            }

                            // Fase 2348: Sellar primero, luego Validar
                            AgCoherenciaGrupo.Sellar(cand, contexto);
                            if (cand.Vacantes > actual.Vacantes)
                            {
                                continue;
                            }
                            if (!Validar(cand, actual, contexto, preexistentes))
                            {
                                continue;
                            }

                            cand.Estado = AgEstadoIndividuo.NoEvaluado;
                            evaluadorAptitud.EvaluarAptitud(cand, contexto, forzar: true, incluirDetalleJson: true);
                            evaluaciones++;

                            var invCand = AgInventarioRanuras.Construir(cand, contexto, sob.Cohorte);
                            double despCand = CalcularDesperdicio(invCand);
                            var puntajeCand = new AgPuntajeReparacion(cand, despCand, 1);

                            if (puntajeCand.Hard < puntajeBaseCohorte.Hard || (puntajeCand.Hard == puntajeBaseCohorte.Hard && puntajeCand.CompareTo(puntajeBaseCohorte) < 0))
                            {
                                if (puntajeCand.Hard < actual.HardViolations)
                                {
                                    int hardAntes = actual.HardViolations;
                                    puntajeMejor = puntajeCand;
                                    mejorClon = cand;
                                    actual = cand;
                                    cohorteReparada = sob.Cohorte;
                                    intentosAceptados++;
                                    mejoraEnIteracion = true;
                                    resultado.Detalle.Add($"[AG-REPARACION-R5] C2 cohorte={sob.Cohorte} hard {hardAntes}->{puntajeCand.Hard}, Desp {despPreCohorte:F3}->{despCand:F3}");
                                    break;
                                }
                            }
                        }

                        if (mejoraEnIteracion) break;

                        // Candidatos C3 / C4: Re-ranuración e intercambio
                        var candidatosC3C4 = GenerarReRanuracionEIntercambio(actual, contexto, sob, inv);
                        foreach (var cand in candidatosC3C4)
                        {
                            intentosEvaluados++;
                            if (evaluaciones >= MaxEvaluaciones || sw.ElapsedMilliseconds >= MaxTiempoMs)
                            {
                                break;
                            }

                            // Fase 2348: Sellar primero, luego Validar
                            AgCoherenciaGrupo.Sellar(cand, contexto);
                            if (cand.Vacantes > actual.Vacantes)
                            {
                                continue;
                            }
                            if (!Validar(cand, actual, contexto, preexistentes))
                            {
                                continue;
                            }

                            cand.Estado = AgEstadoIndividuo.NoEvaluado;
                            evaluadorAptitud.EvaluarAptitud(cand, contexto, forzar: true, incluirDetalleJson: true);
                            evaluaciones++;

                            var invCand = AgInventarioRanuras.Construir(cand, contexto, sob.Cohorte);
                            double despCand = CalcularDesperdicio(invCand);
                            var puntajeCand = new AgPuntajeReparacion(cand, despCand, 1);

                            if (puntajeCand.Hard < puntajeBaseCohorte.Hard || (puntajeCand.Hard == puntajeBaseCohorte.Hard && puntajeCand.CompareTo(puntajeBaseCohorte) < 0))
                            {
                                if (puntajeCand.Hard < actual.HardViolations)
                                {
                                    int hardAntes = actual.HardViolations;
                                    puntajeMejor = puntajeCand;
                                    mejorClon = cand;
                                    actual = cand;
                                    cohorteReparada = sob.Cohorte;
                                    intentosAceptados++;
                                    mejoraEnIteracion = true;
                                    resultado.Detalle.Add($"[AG-REPARACION-R5] C3/C4 cohorte={sob.Cohorte} hard {hardAntes}->{puntajeCand.Hard}, Desp {despPreCohorte:F3}->{despCand:F3}");
                                    break;
                                }
                            }
                        }

                        if (mejoraEnIteracion) break;

                        // Fase 2367: Operador A (ripple) sobre la cadena del carril
                        var candidatosRipple = GenerarDesplazamientosRipple(actual, contexto, sob, inv);
                        resultado.RippleIntentos += candidatosRipple.Count;
                        if (candidatosRipple.Count == 0) resultado.RippleSinCola++;

                        foreach (var cand in candidatosRipple)
                        {
                            intentosEvaluados++;
                            if (evaluaciones >= MaxEvaluaciones || sw.ElapsedMilliseconds >= MaxTiempoMs)
                            {
                                break;
                            }

                            AgCoherenciaGrupo.Sellar(cand, contexto);
                            if (cand.Vacantes > actual.Vacantes)
                            {
                                continue;
                            }
                            if (!Validar(cand, actual, contexto, preexistentes))
                            {
                                continue;
                            }

                            cand.Estado = AgEstadoIndividuo.NoEvaluado;
                            evaluadorAptitud.EvaluarAptitud(cand, contexto, forzar: true, incluirDetalleJson: true);
                            evaluaciones++;

                            var invCand = AgInventarioRanuras.Construir(cand, contexto, sob.Cohorte);
                            double despCand = CalcularDesperdicio(invCand);
                            var puntajeCand = new AgPuntajeReparacion(cand, despCand, 1);

                            if (puntajeCand.Hard < actual.HardViolations)
                            {
                                int hardAntes = actual.HardViolations;
                                puntajeMejor = puntajeCand;
                                mejorClon = cand;
                                actual = cand;
                                cohorteReparada = sob.Cohorte;
                                intentosAceptados++;
                                mejoraEnIteracion = true;
                                resultado.Detalle.Add($"[AG-REPARACION-R5] A ripple cohorte={sob.Cohorte} hard {hardAntes}->{puntajeCand.Hard}, Desp {despPreCohorte:F3}->{despCand:F3}");
                                break;
                            }
                        }
                    }

                    if (!mejoraEnIteracion)
                    {
                        break; // No se encontraron más mejoras viables
                    }
                }

                resultado.Intentos = intentosEvaluados;
                resultado.Evaluaciones = evaluaciones;

                if (mejorClon != null && puntajeMejor.Hard < campeon.HardViolations)
                {
                    resultado.HuboMejora = true;
                    resultado.Individuo = mejorClon;
                    resultado.HardPost = puntajeMejor.Hard;

                    // Fase 2356 / 2360: Recalcular desperdicio pre y post sobre la misma cohorte reparada
                    string cohorteFinal = cohorteReparada ?? primerCohorte;
                    var invPreFinal = AgInventarioRanuras.Construir(campeon, contexto, cohorteFinal);
                    var invPostFinal = AgInventarioRanuras.Construir(mejorClon, contexto, cohorteFinal);
                    resultado.DesperdicioPre = CalcularDesperdicio(invPreFinal);
                    resultado.DesperdicioPost = CalcularDesperdicio(invPostFinal);

                    LoggerHelper.Info(
                        $"[AG-REPARACION-R5] Hard {resultado.HardPre}->{resultado.HardPost} " +
                        $"intentos={intentosEvaluados} aceptados={intentosAceptados} evals={evaluaciones} " +
                        $"desperdicio {resultado.DesperdicioPre:F3}->{resultado.DesperdicioPost:F3} " +
                        $"tiempo={sw.ElapsedMilliseconds}ms");
                }
                else
                {
                    resultado.HuboMejora = false;
                    resultado.Individuo = campeon;
                    resultado.HardPost = campeon.HardViolations;
                    resultado.DesperdicioPost = despPre;
                    resultado.Detalle.Add($"rollback: sin mejora ({campeon.HardViolations})");

                    LoggerHelper.Info(
                        $"[AG-REPARACION-R5] Rollback sin mejora: Hard {campeon.HardViolations} " +
                        $"intentos={intentosEvaluados} evals={evaluaciones} tiempo={sw.ElapsedMilliseconds}ms");
                }

                return resultado;
            }
            catch (Exception ex)
            {
                // Fase 2339: Contención de excepciones de reparación
                LoggerHelper.Warning($"[AG-REPARACION-R5] omitida: {ex.Message}");
                resultado.HuboMejora = false;
                resultado.Individuo = campeon;
                resultado.HardPost = campeon.HardViolations;
                resultado.DesperdicioPost = despPre;
                resultado.Error = ex.Message;
                resultado.Detalle.Add($"excepcion: {ex.Message}");
                return resultado;
            }
        }

        /// <summary>
        /// Detecta sobrecargas R5 estructuradas recomputando con AgConcurrenciaCohorte.
        /// </summary>
        internal static List<AgSobrecargaR5> Detectar(AgIndividuo ind, AgContexto ctx)
        {
            var res = new List<AgSobrecargaR5>();
            if (ind?.Genes == null || ctx == null) return res;

            var porCohorte = ind.Genes
                .Where(g => g.EsVacante == 0)
                .GroupBy(g => AgCalendarioConflictos.Cohorte(g));

            foreach (var gCohorte in porCohorte)
            {
                string cohorte = gCohorte.Key;
                int maxCarriles = AgValidadorReglas.MaxCarrilesCohorte(gCohorte, ctx);

                var intervalos = gCohorte
                    .Where(x => x.FechaInicioEstimada.HasValue && x.FechaFinEstimada.HasValue)
                    .Select(x => (IdAsignatura: x.IdAsignatura, Ini: x.FechaInicioEstimada.Value.Date, Fin: x.FechaFinEstimada.Value.Date));

                var sobrecargas = AgConcurrenciaCohorte.DetectarSobrecargas(intervalos, maxCarriles);
                foreach (var s in sobrecargas)
                {
                    res.Add(new AgSobrecargaR5
                    {
                        Cohorte = cohorte,
                        TramoIni = s.TramoIni,
                        TramoFin = s.TramoFin,
                        MaxConcurrentes = s.MaxConcurrentes,
                        MaxCarriles = maxCarriles,
                        GenesCohorte = gCohorte.ToList()
                    });
                }
            }

            return res;
        }

        /// <summary>
        /// Candidatos C1: Forward shifting para liberar el tramo sobrecargado usando ranuras libres (Fase 2347).
        /// </summary>
        internal static List<AgIndividuo> GenerarDesplazamientos(
            AgIndividuo campeon,
            AgContexto ctx,
            AgSobrecargaR5 sob,
            AgInventarioRanuras inv)
        {
            var candidatos = new List<AgIndividuo>();
            if (campeon?.Genes == null || sob == null || inv == null) return candidatos;

            int C = sob.MaxConcurrentes;
            int M = sob.MaxCarriles;
            int liberar = Math.Max(1, C - M);

            // Asignaturas solapadas en el tramo de sobrecarga
            var solapadas = sob.GenesCohorte
                .Where(g => g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue
                         && g.FechaInicioEstimada.Value.Date <= sob.TramoFin
                         && g.FechaFinEstimada.Value.Date >= sob.TramoIni)
                .GroupBy(g => g.IdAsignatura)
                .Select(grp => new
                {
                    IdAsignatura = grp.Key,
                    Ini = grp.Min(g => g.FechaInicioEstimada.Value.Date),
                    Fin = grp.Max(g => g.FechaFinEstimada.Value.Date),
                    Genes = grp.ToList()
                })
                .OrderByDescending(x => x.Ini)
                .ThenByDescending(x => x.Fin)
                .ToList();

            foreach (var candAsig in solapadas)
            {
                var primerGen = candAsig.Genes.FirstOrDefault();
                if (primerGen == null || primerGen.EsFijo == 1) continue;

                var otras = solapadas
                    .Where(x => x.IdAsignatura != candAsig.IdAsignatura)
                    .Select(x => x.Fin)
                    .OrderBy(d => d)
                    .ToList();

                if (otras.Count < liberar) continue;

                DateTime nuevoIniMin = otras[liberar - 1].AddDays(1);

                // Fase 2347: Ranuras libres reales excluyendo las ocupadas por la propia víctima
                var candAsigGenes = campeon.Genes
                    .Where(g => g.EsVacante == 0
                             && g.IdAsignatura == candAsig.IdAsignatura
                             && string.Equals(AgCalendarioConflictos.Cohorte(g), sob.Cohorte, StringComparison.Ordinal))
                    .ToList();

                if (candAsigGenes.Count == 0) continue;

                var fechasVictima = new HashSet<DateTime>(
                    candAsigGenes.Where(g => g.FechasGen != null).SelectMany(g => g.FechasGen).Select(d => d.Date));
                var ocupadasAjenas = new HashSet<DateTime>(
                    inv.Ocupadas.Keys.Where(f => !fechasVictima.Contains(f.Date)));

                // Fase 2354 / 2358: Priorizar ranuras libres reales incluyendo fechas propias de la víctima; fallback a liberables
                var fechasValidasSet = new HashSet<DateTime>(
                    (inv.FechasValidas ?? Enumerable.Empty<DateTime>()).Select(f => f.Date));
                var libres = (inv.FechasValidas ?? Enumerable.Empty<DateTime>())
                    .Where(f => !ocupadasAjenas.Contains(f.Date) || (inv.Liberables != null && inv.Liberables.Contains(f.Date)))
                    .Distinct()
                    .OrderBy(f => f)
                    .ToList();
                var libresConVictima = (inv.LibresReales ?? Enumerable.Empty<DateTime>())
                    .Concat(fechasVictima)
                    .Where(f => fechasValidasSet.Contains(f.Date))
                    .Select(f => f.Date)
                    .Distinct()
                    .OrderBy(f => f)
                    .ToList();
                var opciones = libresConVictima.Count > 0 ? libresConVictima : libres;

                int diaIso = (primerGen.DiaSemana >= 1 && primerGen.DiaSemana <= 7) ? primerGen.DiaSemana : 1;
                var posiblesInicios = opciones
                    .Where(f => ((f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek) == diaIso && f >= nuevoIniMin)
                    .OrderBy(f => f)
                    .Take(8)
                    .ToList();

                foreach (var inicio in posiblesInicios)
                {
                    var clon = campeon.Clonar();
                    bool exito = true;
                    DateTime cursor = inicio; // Fase 2355: Cursor temporal para asignaturas multi-gen
                    var usadas = new HashSet<DateTime>(); // Fase 2359: Evitar solape multi-gen

                    foreach (var g in candAsigGenes.OrderBy(x => x.DiaSemana))
                    {
                        int idx = campeon.Genes.IndexOf(g);
                        if (idx < 0 || idx >= clon.Genes.Count) { exito = false; break; }
                        var gClon = clon.Genes[idx];

                        int gDiaIso = (gClon.DiaSemana >= 1 && gClon.DiaSemana <= 7) ? gClon.DiaSemana : 1;
                        double unidad = (ctx.UnidadHorasPorAulaId != null && ctx.UnidadHorasPorAulaId.TryGetValue(gClon.IdAula, out var u) && u > 0) ? u : 4.0;
                        int targetSesiones = (int)Math.Max(1, Math.Round(gClon.HorasGen / unidad));

                        var nuevasFechas = opciones
                            .Where(f => f >= cursor
                                     && !usadas.Contains(f.Date)
                                     && ((f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek) == gDiaIso)
                            .Take(targetSesiones)
                            .ToList();

                        if (nuevasFechas.Count < targetSesiones)
                        {
                            exito = false;
                            break;
                        }

                        foreach (var f in nuevasFechas)
                        {
                            usadas.Add(f.Date);
                        }

                        gClon.FechaInicioEstimada = nuevasFechas.Min();
                        gClon.FechaFinEstimada = nuevasFechas.Max();
                        gClon.FechasGen = nuevasFechas;

                        cursor = nuevasFechas.Min().AddDays(1);
                    }

                    if (exito)
                    {
                        candidatos.Add(clon);
                    }
                }
            }

            return candidatos;
        }

        /// <summary>
        /// Candidatos C2: Inserción real en huecos libres intermedios (Fase 2349).
        /// </summary>
        internal static List<AgIndividuo> GenerarCompresiones(
            AgIndividuo campeon,
            AgContexto ctx,
            AgSobrecargaR5 sob,
            AgInventarioRanuras inv)
        {
            var candidatos = new List<AgIndividuo>();
            if (campeon?.Genes == null || sob == null || inv == null) return candidatos;
            if (inv.HuecosIntermedios <= 0 || inv.LibresReales == null || inv.LibresReales.Count == 0) return candidatos;

            var solapadas = sob.GenesCohorte
                .Where(g => g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue
                         && g.FechaInicioEstimada.Value.Date <= sob.TramoFin
                         && g.FechaFinEstimada.Value.Date >= sob.TramoIni)
                .GroupBy(g => g.IdAsignatura)
                .Select(grp => new { IdAsignatura = grp.Key, Genes = grp.ToList() })
                .ToList();

            foreach (var asig in solapadas)
            {
                var primerGen = asig.Genes.FirstOrDefault();
                if (primerGen == null || primerGen.EsFijo == 1) continue;

                string origen = ObtenerOrigenPatron(primerGen, ctx);
                if (!string.Equals(origen, "AUTOMATICO", StringComparison.OrdinalIgnoreCase)) continue;

                var genesAMover = campeon.Genes
                    .Where(g => g.EsVacante == 0
                             && g.IdAsignatura == asig.IdAsignatura
                             && string.Equals(AgCalendarioConflictos.Cohorte(g), sob.Cohorte, StringComparison.Ordinal))
                    .ToList();

                if (genesAMover.Count == 0) continue;

                var clon = campeon.Clonar();
                bool exito = true;

                foreach (var g in genesAMover)
                {
                    int idx = campeon.Genes.IndexOf(g);
                    if (idx < 0 || idx >= clon.Genes.Count) { exito = false; break; }
                    var gClon = clon.Genes[idx];

                    if (gClon.FechasGen == null || gClon.FechasGen.Count <= 1) continue;

                    int diaIso = (gClon.DiaSemana >= 1 && gClon.DiaSemana <= 7) ? gClon.DiaSemana : 1;
                    DateTime ini = gClon.FechaInicioEstimada ?? gClon.FechasGen.Min();
                    DateTime fin = gClon.FechaFinEstimada ?? gClon.FechasGen.Max();

                    // Buscar huecos libres intermedios reales en el mismo día Iso
                    var huecosLibres = inv.LibresReales
                        .Where(f => f > ini && f < fin && ((f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek) == diaIso)
                        .OrderBy(f => f)
                        .ToList();

                    if (huecosLibres.Count == 0)
                    {
                        exito = false;
                        break;
                    }

                    // Compactar fechas insertando en huecos libres y acortando por la cola
                    var combinadas = gClon.FechasGen.Concat(huecosLibres)
                        .Distinct()
                        .OrderBy(d => d)
                        .Take(gClon.FechasGen.Count)
                        .ToList();

                    if (combinadas.Count == gClon.FechasGen.Count && combinadas.Max() < fin)
                    {
                        gClon.FechasGen = combinadas;
                        gClon.FechaInicioEstimada = combinadas.Min();
                        gClon.FechaFinEstimada = combinadas.Max();
                    }
                    else
                    {
                        exito = false;
                        break;
                    }
                }

                if (exito)
                {
                    candidatos.Add(clon);
                }
            }

            return candidatos;
        }

        /// <summary>
        /// Operador A (Fase 2367): ripple de la cadena del carril. Libera las fechas de la víctima y de las
        /// asignaturas posteriores del mismo carril y las reproyecta secuencialmente sobre ranuras reales.
        /// </summary>
        internal static List<AgIndividuo> GenerarDesplazamientosRipple(
            AgIndividuo actual,
            AgContexto ctx,
            AgSobrecargaR5 sob,
            AgInventarioRanuras inv)
        {
            var candidatos = new List<AgIndividuo>();
            if (actual?.Genes == null || ctx == null || sob == null || inv?.Carriles == null) return candidatos;

            int liberar = Math.Max(1, sob.MaxConcurrentes - sob.MaxCarriles);

            var solapadas = sob.GenesCohorte
                .Where(g => g.EsVacante == 0 && g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue
                         && g.FechaInicioEstimada.Value.Date <= sob.TramoFin
                         && g.FechaFinEstimada.Value.Date >= sob.TramoIni)
                .GroupBy(g => g.IdAsignatura)
                .Select(x => new
                {
                    IdAsignatura = x.Key,
                    Ini = x.Min(g => g.FechaInicioEstimada.Value.Date),
                    Fin = x.Max(g => g.FechaFinEstimada.Value.Date)
                })
                .OrderByDescending(x => x.Ini)
                .ToList();

            foreach (var victima in solapadas)
            {
                var genesVictima = actual.Genes
                    .Where(g => g.EsVacante == 0 && g.IdAsignatura == victima.IdAsignatura
                             && string.Equals(AgCalendarioConflictos.Cohorte(g), sob.Cohorte, StringComparison.Ordinal))
                    .ToList();
                if (genesVictima.Count == 0 || genesVictima.Any(g => g.EsFijo == 1)) continue;

                var otras = solapadas.Where(x => x.IdAsignatura != victima.IdAsignatura)
                                     .Select(x => x.Fin).OrderBy(d => d).ToList();
                if (otras.Count < liberar) continue;
                DateTime nuevoIniMin = otras[liberar - 1].AddDays(1);

                string claveCarril = AgCarriles.ClaveCarril(genesVictima[0]);
                var carril = inv.Carriles.FirstOrDefault(c => string.Equals(c.Clave, claveCarril, StringComparison.Ordinal));
                if (carril == null || carril.FechasValidas.Count == 0) continue;

                var cadena = actual.Genes
                    .Where(g => g.EsVacante == 0
                             && string.Equals(AgCarriles.ClaveCarril(g), claveCarril, StringComparison.Ordinal)
                             && g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue
                             && g.FechaInicioEstimada.Value >= victima.Ini)
                    .OrderBy(g => g.FechaInicioEstimada).ThenBy(g => g.DiaSemana)
                    .ToList();
                if (cadena.Count == 0 || cadena.Any(g => g.EsFijo == 1)) continue;

                var idxCadena = new HashSet<int>();
                foreach (var g in cadena)
                {
                    int idx = actual.Genes.IndexOf(g);
                    if (idx >= 0) idxCadena.Add(idx);
                }
                if (idxCadena.Count == 0) continue;

                int idAulaVictima = genesVictima[0].IdAula;
                var ocupadasFuera = new HashSet<DateTime>();
                for (int i = 0; i < actual.Genes.Count; i++)
                {
                    if (idxCadena.Contains(i)) continue;
                    var g = actual.Genes[i];
                    if (g.EsVacante != 0) continue;
                    bool mismaAula = g.IdAula == idAulaVictima;
                    bool mismaCohorte = string.Equals(AgCalendarioConflictos.Cohorte(g), sob.Cohorte, StringComparison.Ordinal);
                    if (!mismaAula && !mismaCohorte) continue;

                    foreach (var f in AgInventarioRanuras.FechasDeGen(g)) ocupadasFuera.Add(f.Date);
                }

                var libres = carril.FechasValidas.Where(f => !ocupadasFuera.Contains(f)).ToList();
                if (libres.Count == 0) continue;

                DateTime inicioMin = nuevoIniMin;
                for (int i = 0; i < actual.Genes.Count; i++)
                {
                    if (idxCadena.Contains(i)) continue;
                    var g = actual.Genes[i];
                    if (g.EsVacante != 0) continue;
                    if (!string.Equals(AgCarriles.ClaveCarril(g), claveCarril, StringComparison.Ordinal)) continue;
                    if (!g.FechaInicioEstimada.HasValue || !g.FechaFinEstimada.HasValue) continue;
                    if (g.FechaInicioEstimada.Value >= victima.Ini) continue;
                    var fin = g.FechaFinEstimada.Value.Date.AddDays(1);
                    if (fin > inicioMin) inicioMin = fin;
                }

                int diaVictima = genesVictima[0].DiaSemana >= 1 && genesVictima[0].DiaSemana <= 7
                    ? genesVictima[0].DiaSemana : 1;
                var inicios = libres
                    .Where(f => AgInventarioRanuras.Iso(f) == diaVictima && f >= inicioMin)
                    .OrderBy(f => f)
                    .Take(8)
                    .ToList();

                foreach (var inicio in inicios)
                {
                    var clon = actual.Clonar();
                    var usadas = new HashSet<DateTime>();
                    DateTime cursor = inicio;
                    bool exito = true;

                    foreach (var idx in idxCadena.OrderBy(i => actual.Genes[i].FechaInicioEstimada).ThenBy(i => actual.Genes[i].DiaSemana))
                    {
                        var gClon = clon.Genes[idx];
                        double unidad = (ctx.UnidadHorasPorAulaId != null
                                         && ctx.UnidadHorasPorAulaId.TryGetValue(gClon.IdAula, out var u) && u > 0) ? u : 4.0;
                        int n = (int)Math.Max(1, Math.Round(gClon.HorasGen / unidad));

                        var fechas = libres
                            .Where(f => f >= cursor && AgInventarioRanuras.Iso(f) == gClon.DiaSemana && !usadas.Contains(f))
                            .Take(n)
                            .ToList();
                        if (fechas.Count < n)
                        {
                            exito = false;
                            break;
                        }

                        foreach (var f in fechas) usadas.Add(f);
                        gClon.FechasGen = fechas;
                        gClon.FechaInicioEstimada = fechas.Min();
                        gClon.FechaFinEstimada = fechas.Max();
                        cursor = fechas[fechas.Count - 1].AddDays(1);
                    }

                    if (exito) candidatos.Add(clon);
                }
            }

            return candidatos;
        }

        /// <summary>
        /// Candidatos C3 / C4: Re-ranuración e intercambio de días en patrones AUTOMATICO (Fases 2335–2336).
        /// </summary>
        internal static List<AgIndividuo> GenerarReRanuracionEIntercambio(
            AgIndividuo campeon,
            AgContexto ctx,
            AgSobrecargaR5 sob,
            AgInventarioRanuras inv)
        {
            var candidatos = new List<AgIndividuo>();
            if (campeon?.Genes == null || sob == null || inv == null) return candidatos;

            var solapadas = sob.GenesCohorte
                .Where(g => g.EsVacante == 0
                         && g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue
                         && g.FechaInicioEstimada.Value.Date <= sob.TramoFin
                         && g.FechaFinEstimada.Value.Date >= sob.TramoIni)
                .ToList();

            // C3: Re-ranuración de día individual
            foreach (var g in solapadas)
            {
                if (g.EsFijo == 1) continue;

                string origen = ObtenerOrigenPatron(g, ctx);
                if (!string.Equals(origen, "AUTOMATICO", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var aula = ctx.AulasPorId != null && ctx.AulasPorId.TryGetValue(g.IdAula, out var a) ? a : null;
                var prog = ctx.ProgramasPorId != null && ctx.ProgramasPorId.TryGetValue(g.IdPrograma, out var p) ? p : null;
                double unidad = AgUnidadesSemanales.HorasSesionAula(aula, AgNivelAcademicoCatalogo.Desde(prog?.Modalidad));
                if (unidad <= 0) unidad = 4.0;

                foreach (var nuevoDia in inv.DiasVentana)
                {
                    if (nuevoDia == g.DiaSemana) continue;

                    if (ctx.TryGetBloque(nuevoDia, unidad, out var nuevoBloque))
                    {
                        var clon = campeon.Clonar();
                        int idx = campeon.Genes.IndexOf(g);
                        if (idx < 0 || idx >= clon.Genes.Count) continue;
                        var gClon = clon.Genes[idx];

                        gClon.DiaSemana = nuevoDia;
                        gClon.IdBloqueHorario = nuevoBloque.Id;
                        gClon.DiasPatron = nuevoDia.ToString();
                        gClon.IndiceDiaPatron = 0;
                        gClon.PatronSemanal = string.Empty;

                        DateTime ini = gClon.FechaInicioEstimada ?? (ctx.Ventana != null ? ctx.Ventana.FechaInicio : DateTime.MinValue);
                        int targetSesiones = (gClon.FechasGen != null && gClon.FechasGen.Count > 0) ? gClon.FechasGen.Count : 1;
                        var nuevasFechas = inv.FechasValidas
                            .Where(f => f >= ini && ((f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek) == nuevoDia)
                            .Take(targetSesiones)
                            .ToList();

                        if (nuevasFechas.Count == targetSesiones)
                        {
                            gClon.FechasGen = nuevasFechas;
                            gClon.FechaInicioEstimada = nuevasFechas.Min();
                            gClon.FechaFinEstimada = nuevasFechas.Max();
                            candidatos.Add(clon);
                        }
                    }
                }
            }

            // C4: Intercambio de días entre dos asignaturas automáticas de la cohorte
            for (int i = 0; i < solapadas.Count; i++)
            {
                var g1 = solapadas[i];
                if (g1.EsFijo == 1) continue;
                if (!string.Equals(ObtenerOrigenPatron(g1, ctx), "AUTOMATICO", StringComparison.OrdinalIgnoreCase)) continue;

                for (int j = i + 1; j < solapadas.Count; j++)
                {
                    var g2 = solapadas[j];
                    if (g2.EsFijo == 1 || g1.IdAsignatura == g2.IdAsignatura) continue;
                    if (!string.Equals(ObtenerOrigenPatron(g2, ctx), "AUTOMATICO", StringComparison.OrdinalIgnoreCase)) continue;
                    if (g1.DiaSemana == g2.DiaSemana) continue;

                    var clon = campeon.Clonar();
                    int idx1 = campeon.Genes.IndexOf(g1);
                    int idx2 = campeon.Genes.IndexOf(g2);
                    if (idx1 < 0 || idx1 >= clon.Genes.Count || idx2 < 0 || idx2 >= clon.Genes.Count) continue;

                    var gClon1 = clon.Genes[idx1];
                    var gClon2 = clon.Genes[idx2];

                    int dia1 = gClon1.DiaSemana;
                    int dia2 = gClon2.DiaSemana;
                    int b1 = gClon1.IdBloqueHorario;
                    int b2 = gClon2.IdBloqueHorario;

                    gClon1.DiaSemana = dia2;
                    gClon1.IdBloqueHorario = b2;
                    gClon1.DiasPatron = dia2.ToString();
                    gClon1.IndiceDiaPatron = 0;
                    gClon1.PatronSemanal = string.Empty;

                    gClon2.DiaSemana = dia1;
                    gClon2.IdBloqueHorario = b1;
                    gClon2.DiasPatron = dia1.ToString();
                    gClon2.IndiceDiaPatron = 0;
                    gClon2.PatronSemanal = string.Empty;

                    // Reproyectar fechas
                    DateTime ini1 = gClon1.FechaInicioEstimada ?? (ctx.Ventana != null ? ctx.Ventana.FechaInicio : DateTime.MinValue);
                    DateTime ini2 = gClon2.FechaInicioEstimada ?? (ctx.Ventana != null ? ctx.Ventana.FechaInicio : DateTime.MinValue);

                    int count1 = gClon1.FechasGen?.Count ?? 1;
                    int count2 = gClon2.FechasGen?.Count ?? 1;

                    var f1 = inv.FechasValidas
                        .Where(f => f >= ini1 && ((f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek) == dia2)
                        .Take(count1).ToList();
                    var f2 = inv.FechasValidas
                        .Where(f => f >= ini2 && ((f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek) == dia1)
                        .Take(count2).ToList();

                    if (f1.Count == count1 && f2.Count == count2)
                    {
                        gClon1.FechasGen = f1;
                        gClon1.FechaInicioEstimada = f1.Min();
                        gClon1.FechaFinEstimada = f1.Max();

                        gClon2.FechasGen = f2;
                        gClon2.FechaInicioEstimada = f2.Min();
                        gClon2.FechaFinEstimada = f2.Max();

                        candidatos.Add(clon);
                    }
                }
            }

            return candidatos;
        }

        /// <summary>
        /// Validación V1–V6 sobre un candidato generado (Fases 2333, 2335, 2338, 2343–2345).
        /// </summary>
        internal static bool Validar(AgIndividuo clon, AgIndividuo original, AgContexto ctx, ISet<string> preexistentes = null)
        {
            if (clon?.Genes == null || original?.Genes == null || ctx == null) return false;
            if (ctx.Ventana == null) return false;
            preexistentes = preexistentes ?? ClavesDuras(original);

            // V3: Sin vacantes nuevas
            if (clon.Vacantes > original.Vacantes) return false;

            for (int i = 0; i < clon.Genes.Count; i++)
            {
                var g = clon.Genes[i];
                var gOrig = (i < original.Genes.Count) ? original.Genes[i] : null;

                if (g.EsVacante == 1) continue;

                // V4: No mover genes con EsFijo == 1
                if (gOrig != null && gOrig.EsFijo == 1)
                {
                    if (g.FechaInicioEstimada != gOrig.FechaInicioEstimada || g.FechaFinEstimada != gOrig.FechaFinEstimada)
                    {
                        return false;
                    }
                }

                // Fase 2333: Para LEGADO/FIJO comparar patrón (días/bloque), no fechas
                string origen = ObtenerOrigenPatron(g, ctx);
                bool esLegadoOFijo = string.Equals(origen, "LEGADO", StringComparison.OrdinalIgnoreCase) ||
                                     string.Equals(origen, "FIJO", StringComparison.OrdinalIgnoreCase);

                if (esLegadoOFijo && gOrig != null &&
                    (g.DiaSemana != gOrig.DiaSemana ||
                     g.IdBloqueHorario != gOrig.IdBloqueHorario ||
                     !string.Equals(g.DiasPatron, gOrig.DiasPatron, StringComparison.Ordinal)))
                {
                    return false;
                }

                // V1: Fechas válidas, día de ventana, no festivo
                if (g.FechasGen != null && g.FechasGen.Count > 0)
                {
                    foreach (var f in g.FechasGen)
                    {
                        if (!ctx.Ventana.Contiene(f)) return false;
                        if (ctx.EsFestivo(f)) return false;

                        int diaIso = (f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek;
                        if (diaIso != g.DiaSemana) return false;

                        // Fase 2344: Verificar si docente ocupado en calendario global con FranjaHoraria
                        if (g.IdTercero > 0)
                        {
                            var bloque = ctx.Bloques != null ? ctx.Bloques.FirstOrDefault(b => b.Id == g.IdBloqueHorario) : null;
                            if (bloque != null && ctx.TerceroOcupado(g.IdTercero, f, new FranjaHoraria(bloque.TimeSpanInicio, bloque.TimeSpanFin)))
                            {
                                return false;
                            }
                        }
                    }

                    // V2: Invariante de horas: Count * unidad == HorasGen
                    double unidad = (ctx.UnidadHorasPorAulaId != null && ctx.UnidadHorasPorAulaId.TryGetValue(g.IdAula, out var u) && u > 0) ? u : 4.0;
                    double horasEsperadas = Math.Round(g.FechasGen.Count * unidad, 2);
                    if (Math.Abs(horasEsperadas - Math.Round(g.HorasGen, 2)) > 0.01)
                    {
                        return false;
                    }
                }
            }

            // V5: Conflicto docente y aula
            var genesActivos = clon.Genes.Where(g => g.EsVacante == 0 && g.FechasGen != null && g.FechasGen.Count > 0).ToList();

            // Cruce docente
            var porDocente = genesActivos.Where(g => g.IdTercero > 0).GroupBy(g => g.IdTercero);
            foreach (var docGroup in porDocente)
            {
                var sesiones = new HashSet<(DateTime Fecha, int IdBloque)>();
                foreach (var g in docGroup)
                {
                    foreach (var f in g.FechasGen)
                    {
                        if (!sesiones.Add((f.Date, g.IdBloqueHorario)))
                        {
                            return false; // Colisión de docente
                        }
                    }
                }
            }

            // Cruce de aula (R2) y R4/R6: se validan por claves al final (Fase 2364)

            // Fase 2345: Cruce contra SesionesConservadas comparando solape de FranjaHoraria
            if (ctx.SesionesConservadas != null && ctx.SesionesConservadas.Count > 0)
            {
                foreach (var sc in ctx.SesionesConservadas)
                {
                    var f = sc.Fecha.Date;
                    foreach (var g in genesActivos)
                    {
                        if (g.FechasGen != null && g.FechasGen.Any(df => df.Date == f))
                        {
                            var bloque = ctx.Bloques != null ? ctx.Bloques.FirstOrDefault(b => b.Id == g.IdBloqueHorario) : null;
                            if (bloque != null)
                            {
                                var franjaGen = new FranjaHoraria(bloque.TimeSpanInicio, bloque.TimeSpanFin);
                                var franjaSc = new FranjaHoraria(sc.Inicio, sc.Fin);
                                if (franjaGen.SeSolapaCon(franjaSc))
                                {
                                    if (g.IdAula == sc.Clave.IdAula)
                                    {
                                        return false; // Choque con sesión conservada en aula
                                    }
                                    if (g.IdTercero > 0 && g.IdTercero == sc.IdTercero)
                                    {
                                        return false; // Choque con sesión conservada en docente
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Fase 2364: R2/R4/R6 se evalúan por claves; sólo se rechazan las NUEVAS.
            var clavesClon = ClavesDuras(clon);
            foreach (var clave in clavesClon)
            {
                if (!preexistentes.Contains(clave))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Claves canónicas de conflictos duros R2/R4/R6 de un individuo (Fase 2364).
        /// Permiten distinguir una violación preexistente de una introducida por un candidato.
        /// </summary>
        internal static HashSet<string> ClavesDuras(AgIndividuo ind)
        {
            var set = new HashSet<string>(StringComparer.Ordinal);
            if (ind?.Genes == null) return set;

            var activos = ind.Genes.Where(g => g.EsVacante == 0).ToList();

            // R2: dos asignaturas en la misma aula, fecha y bloque
            foreach (var grp in activos
                .SelectMany(g => (g.FechasGen ?? new List<DateTime>()).Select(f => new { g, F = f.Date }))
                .GroupBy(x => (x.g.IdAula, x.F, x.g.IdBloqueHorario)))
            {
                var asigs = grp.Select(x => x.g.IdAsignatura).Distinct().OrderBy(x => x).ToList();
                for (int i = 0; i < asigs.Count; i++)
                {
                    for (int j = i + 1; j < asigs.Count; j++)
                    {
                        set.Add("R2|" + grp.Key.Item1 + "|" + grp.Key.Item2.ToString("yyyyMMdd") + "|" +
                                grp.Key.Item3 + "|" + asigs[i] + "|" + asigs[j]);
                    }
                }
            }

            // R4: una cohorte con más de un bloque el mismo día
            foreach (var co in activos.GroupBy(g => AgCalendarioConflictos.Cohorte(g)))
            {
                var porFecha = co
                    .SelectMany(g => (g.FechasGen ?? new List<DateTime>()).Select(f => new { g, F = f.Date }))
                    .GroupBy(x => x.F);

                foreach (var cf in porFecha)
                {
                    var asigs = cf.Select(x => x.g.IdAsignatura).Distinct().OrderBy(x => x).ToList();
                    for (int i = 0; i < asigs.Count; i++)
                    {
                        bool propioMultiBloque = cf
                            .Where(x => x.g.IdAsignatura == asigs[i])
                            .Select(x => x.g.IdBloqueHorario)
                            .Distinct()
                            .Count() > 1;
                        if (propioMultiBloque)
                        {
                            set.Add("R4|" + co.Key + "|" + cf.Key.ToString("yyyyMMdd") + "|" + asigs[i] + "|" + asigs[i]);
                        }

                        for (int j = i + 1; j < asigs.Count; j++)
                        {
                            var bloques = cf.Select(x => x.g.IdBloqueHorario).Distinct().ToList();
                            if (bloques.Count > 1)
                            {
                                set.Add("R4|" + co.Key + "|" + cf.Key.ToString("yyyyMMdd") + "|" + asigs[i] + "|" + asigs[j]);
                            }
                        }
                    }
                }
            }

            // R6: solape estricto de secuencia curricular por cohorte y carril
            foreach (var cc in activos.GroupBy(g => (Cohorte: AgCalendarioConflictos.Cohorte(g), Carril: AgCarriles.ClaveCarril(g))))
            {
                var asigs = cc.GroupBy(g => g.IdAsignatura)
                    .Select(a => (Id: a.Key,
                                  Ini: a.Min(x => x.FechaInicioEstimada ?? DateTime.MaxValue),
                                  Fin: a.Max(x => x.FechaFinEstimada ?? DateTime.MinValue)))
                    .ToList();

                for (int i = 0; i < asigs.Count; i++)
                {
                    for (int j = i + 1; j < asigs.Count; j++)
                    {
                        if (asigs[i].Ini < asigs[j].Fin && asigs[j].Ini < asigs[i].Fin)
                        {
                            int a = Math.Min(asigs[i].Id, asigs[j].Id);
                            int b = Math.Max(asigs[i].Id, asigs[j].Id);
                            set.Add("R6|" + cc.Key.Cohorte + "|" + cc.Key.Carril + "|" + a + "|" + b);
                        }
                    }
                }
            }

            return set;
        }

        private static string ObtenerOrigenPatron(AgGen g, AgContexto ctx)
        {
            if (g == null || ctx == null) return "LEGADO";

            if (ctx.AulaAsignaturaPorAulaAsig != null && ctx.AulaAsignaturaPorAulaAsig.TryGetValue((g.IdAula, g.IdAsignatura), out var aa))
            {
                return aa.OrigenPatron;
            }
            if (ctx.AulaAsignaturasPorId != null && ctx.AulaAsignaturasPorId.TryGetValue(g.IdAsignatura, out var aa2))
            {
                return aa2.OrigenPatron;
            }
            return "LEGADO";
        }

        /// <summary>
        /// Desperdicio normalizado sin doble conteo de huecos (Fase 2340).
        /// </summary>
        private static double CalcularDesperdicio(AgInventarioRanuras inv)
        {
            if (inv == null || inv.FechasValidas == null || inv.FechasValidas.Count == 0) return 0.0;
            double n = inv.FechasValidas.Count;
            double libres = inv.LibresReales != null ? inv.LibresReales.Count : 0;
            return (libres / n) + (inv.BalancePorDia / n);
        }

        internal struct AgPuntajeReparacion : IComparable<AgPuntajeReparacion>
        {
            public int Hard;
            public int Vacantes;
            public double Desperdicio;
            public int MakespanDias;
            public double Soft;
            public int Movimientos;

            public AgPuntajeReparacion(AgIndividuo ind, double desperdicio, int movimientos)
            {
                Hard = ind?.HardViolations ?? 999;
                Vacantes = ind?.Vacantes ?? 999;
                Desperdicio = desperdicio;
                MakespanDias = CalcularMakespan(ind);
                Soft = ind?.SoftViolations ?? 0.0;
                Movimientos = movimientos;
            }

            public int CompareTo(AgPuntajeReparacion other)
            {
                // 1. Hard violations
                int c = Hard.CompareTo(other.Hard);
                if (c != 0) return c;

                // 2. Vacantes
                c = Vacantes.CompareTo(other.Vacantes);
                if (c != 0) return c;

                // 3. Desperdicio de ranuras
                c = Desperdicio.CompareTo(other.Desperdicio);
                if (c != 0) return c;

                // 4. Makespan
                c = MakespanDias.CompareTo(other.MakespanDias);
                if (c != 0) return c;

                // 5. Soft violations
                c = Soft.CompareTo(other.Soft);
                if (c != 0) return c;

                // 6. Número de movimientos
                return Movimientos.CompareTo(other.Movimientos);
            }

            private static int CalcularMakespan(AgIndividuo ind)
            {
                if (ind?.Genes == null || ind.Genes.Count == 0) return 0;
                var fechas = ind.Genes
                    .Where(g => g.EsVacante == 0 && g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue)
                    .Select(g => (Ini: g.FechaInicioEstimada.Value.Date, Fin: g.FechaFinEstimada.Value.Date))
                    .ToList();

                if (fechas.Count == 0) return 0;
                var min = fechas.Min(x => x.Ini);
                var max = fechas.Max(x => x.Fin);
                return (int)(max - min).TotalDays;
            }
        }
    }
}
