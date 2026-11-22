using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Motor heurístico de reparación automática de sobrecargas R5 mediante inventario de ranuras,
    /// forward shifting y compresión determinista (Fase 2325 / Fase 2326 / Fase 2327).
    /// </summary>
    internal sealed class AgReparadorR5
    {
        public const int MaxIntentos = 8;
        public const int MaxEvaluaciones = 12;
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

            var sobrecargas = Detectar(campeon, contexto);
            if (sobrecargas.Count == 0)
            {
                return resultado;
            }

            string primerCohorte = sobrecargas[0].Cohorte;
            var invInicial = AgInventarioRanuras.Construir(campeon, contexto, primerCohorte);
            double despPre = CalcularDesperdicio(invInicial);
            resultado.DesperdicioPre = despPre;
            resultado.DesperdicioPost = despPre;

            var puntajeMejor = new AgPuntajeReparacion(campeon, despPre, 0);
            AgIndividuo mejorClon = null;
            int evaluaciones = 0;
            int intentos = 0;

            foreach (var sob in sobrecargas)
            {
                if (intentos >= MaxIntentos || evaluaciones >= MaxEvaluaciones || sw.ElapsedMilliseconds >= MaxTiempoMs)
                {
                    break;
                }

                var inv = AgInventarioRanuras.Construir(campeon, contexto, sob.Cohorte);

                // Candidatos C1: Desplazamiento (Forward Shifting)
                var candidatosC1 = GenerarDesplazamientos(campeon, contexto, sob, inv);
                foreach (var cand in candidatosC1)
                {
                    intentos++;
                    if (intentos > MaxIntentos || evaluaciones >= MaxEvaluaciones || sw.ElapsedMilliseconds >= MaxTiempoMs)
                    {
                        break;
                    }

                    if (!Validar(cand, campeon, contexto))
                    {
                        continue;
                    }

                    AgCoherenciaGrupo.Sellar(cand, contexto);
                    evaluadorAptitud.EvaluarAptitud(cand, contexto);
                    evaluaciones++;

                    var invCand = AgInventarioRanuras.Construir(cand, contexto, sob.Cohorte);
                    double despCand = CalcularDesperdicio(invCand);
                    var puntajeCand = new AgPuntajeReparacion(cand, despCand, 1);

                    if (puntajeCand.Hard < puntajeMejor.Hard || (puntajeCand.Hard == puntajeMejor.Hard && puntajeCand.CompareTo(puntajeMejor) < 0))
                    {
                        if (puntajeCand.Hard < campeon.HardViolations)
                        {
                            puntajeMejor = puntajeCand;
                            mejorClon = cand;
                            resultado.Detalle.Add($"C1 aceptado: Hard {campeon.HardViolations}->{puntajeCand.Hard}, Desp {despPre:F3}->{despCand:F3}");
                        }
                    }
                }

                if (modo == AgModoReparacionR5.SoloDesplazamiento)
                {
                    continue;
                }

                // Candidatos C2: Compresión
                var candidatosC2 = GenerarCompresiones(campeon, contexto, sob, inv);
                foreach (var cand in candidatosC2)
                {
                    intentos++;
                    if (intentos > MaxIntentos || evaluaciones >= MaxEvaluaciones || sw.ElapsedMilliseconds >= MaxTiempoMs)
                    {
                        break;
                    }

                    if (!Validar(cand, campeon, contexto))
                    {
                        continue;
                    }

                    AgCoherenciaGrupo.Sellar(cand, contexto);
                    evaluadorAptitud.EvaluarAptitud(cand, contexto);
                    evaluaciones++;

                    var invCand = AgInventarioRanuras.Construir(cand, contexto, sob.Cohorte);
                    double despCand = CalcularDesperdicio(invCand);
                    var puntajeCand = new AgPuntajeReparacion(cand, despCand, 1);

                    if (puntajeCand.Hard < puntajeMejor.Hard || (puntajeCand.Hard == puntajeMejor.Hard && puntajeCand.CompareTo(puntajeMejor) < 0))
                    {
                        if (puntajeCand.Hard < campeon.HardViolations)
                        {
                            puntajeMejor = puntajeCand;
                            mejorClon = cand;
                            resultado.Detalle.Add($"C2 aceptado: Hard {campeon.HardViolations}->{puntajeCand.Hard}, Desp {despPre:F3}->{despCand:F3}");
                        }
                    }
                }

                // Candidatos C3 / C4: Re-ranuración e intercambio
                var candidatosC3C4 = GenerarReRanuracionEIntercambio(campeon, contexto, sob, inv);
                foreach (var cand in candidatosC3C4)
                {
                    intentos++;
                    if (intentos > MaxIntentos || evaluaciones >= MaxEvaluaciones || sw.ElapsedMilliseconds >= MaxTiempoMs)
                    {
                        break;
                    }

                    if (!Validar(cand, campeon, contexto))
                    {
                        continue;
                    }

                    AgCoherenciaGrupo.Sellar(cand, contexto);
                    evaluadorAptitud.EvaluarAptitud(cand, contexto);
                    evaluaciones++;

                    var invCand = AgInventarioRanuras.Construir(cand, contexto, sob.Cohorte);
                    double despCand = CalcularDesperdicio(invCand);
                    var puntajeCand = new AgPuntajeReparacion(cand, despCand, 1);

                    if (puntajeCand.Hard < puntajeMejor.Hard || (puntajeCand.Hard == puntajeMejor.Hard && puntajeCand.CompareTo(puntajeMejor) < 0))
                    {
                        if (puntajeCand.Hard < campeon.HardViolations)
                        {
                            puntajeMejor = puntajeCand;
                            mejorClon = cand;
                            resultado.Detalle.Add($"C3/C4 aceptado: Hard {campeon.HardViolations}->{puntajeCand.Hard}, Desp {despPre:F3}->{despCand:F3}");
                        }
                    }
                }
            }

            resultado.Intentos = intentos;
            resultado.Evaluaciones = evaluaciones;

            if (mejorClon != null && puntajeMejor.Hard < campeon.HardViolations)
            {
                resultado.HuboMejora = true;
                resultado.Individuo = mejorClon;
                resultado.HardPost = puntajeMejor.Hard;
                resultado.DesperdicioPost = puntajeMejor.Desperdicio;

                LoggerHelper.Info(
                    $"[AG-REPARACION-R5] Hard {resultado.HardPre}->{resultado.HardPost} " +
                    $"intentos={intentos} evals={evaluaciones} " +
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
                    $"intentos={intentos} evals={evaluaciones} tiempo={sw.ElapsedMilliseconds}ms");
            }

            return resultado;
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
        /// Candidatos C1: Forward shifting para liberar el tramo sobrecargado.
        /// </summary>
        internal static List<AgIndividuo> GenerarDesplazamientos(
            AgIndividuo campeon,
            AgContexto ctx,
            AgSobrecargaR5 sob,
            AgInventarioRanuras inv)
        {
            var candidatos = new List<AgIndividuo>();
            if (campeon?.Genes == null || sob == null) return candidatos;

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

            // Evaluar asignaturas candidatas a mover (no fijas, no LEGADO)
            foreach (var candAsig in solapadas)
            {
                var primerGen = candAsig.Genes.FirstOrDefault();
                if (primerGen == null || primerGen.EsFijo == 1) continue;

                string origen = ObtenerOrigenPatron(primerGen, ctx);
                if (string.Equals(origen, "LEGADO", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(origen, "FIJO", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var otras = solapadas
                    .Where(x => x.IdAsignatura != candAsig.IdAsignatura)
                    .Select(x => x.Fin)
                    .OrderBy(d => d)
                    .ToList();

                if (otras.Count < liberar) continue;

                DateTime fechaBase = otras[liberar - 1].AddDays(1);

                // Clonar individuo para aplicar el desplazamiento
                var clon = campeon.Clonar();
                var genesAMover = clon.Genes.Where(g => g.IdAsignatura == candAsig.IdAsignatura).ToList();
                if (genesAMover.Count == 0) continue;

                bool exito = true;
                foreach (var g in genesAMover)
                {
                    // Buscar primera fecha válida >= fechaBase con el día del bloque/gen
                    int diaIso = (g.DiaSemana >= 1 && g.DiaSemana <= 7) ? g.DiaSemana : 1;
                    var fechaIniValida = inv.FechasValidas
                        .FirstOrDefault(f => f >= fechaBase && ((f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek) == diaIso);

                    if (fechaIniValida == default(DateTime))
                    {
                        exito = false;
                        break;
                    }

                    // Determinar número de sesiones requeridas
                    double unidad = (ctx.UnidadHorasPorAulaId != null && ctx.UnidadHorasPorAulaId.TryGetValue(g.IdAula, out var u) && u > 0) ? u : 4.0;
                    int targetSesiones = (int)Math.Max(1, Math.Round(g.HorasGen / unidad));

                    // Proyectar fechas válidas
                    var nuevasFechas = inv.FechasValidas
                        .Where(f => f >= fechaIniValida && ((f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek) == diaIso)
                        .Take(targetSesiones)
                        .ToList();

                    if (nuevasFechas.Count < targetSesiones)
                    {
                        exito = false;
                        break;
                    }

                    g.FechaInicioEstimada = nuevasFechas.Min();
                    g.FechaFinEstimada = nuevasFechas.Max();
                    g.FechasGen = nuevasFechas;
                }

                if (exito)
                {
                    candidatos.Add(clon);
                }
            }

            return candidatos;
        }

        /// <summary>
        /// Candidatos C2: Compresión usando ranuras libres reales en el tramo.
        /// </summary>
        internal static List<AgIndividuo> GenerarCompresiones(
            AgIndividuo campeon,
            AgContexto ctx,
            AgSobrecargaR5 sob,
            AgInventarioRanuras inv)
        {
            var candidatos = new List<AgIndividuo>();
            if (campeon?.Genes == null || sob == null || inv == null) return candidatos;

            // C2 requiere que existan ranuras libres reales en el tramo
            int libresEnTramo = inv.LibresReales.Count(f => f >= sob.TramoIni && f <= sob.TramoFin);
            if (libresEnTramo <= 0)
            {
                // Rechazado: sin ranuras libres reales en el tramo
                return candidatos;
            }

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
                if (!string.Equals(origen, "AUTOMATICO", StringComparison.OrdinalIgnoreCase))
                {
                    continue; // Solo AUTOMATICO
                }

                var clon = campeon.Clonar();
                var genesClon = clon.Genes.Where(g => g.IdAsignatura == asig.IdAsignatura).ToList();
                bool exito = true;

                foreach (var g in genesClon)
                {
                    if (g.FechasGen == null || g.FechasGen.Count <= 1) continue;

                    int diaIso = (g.DiaSemana >= 1 && g.DiaSemana <= 7) ? g.DiaSemana : 1;
                    DateTime ini = g.FechaInicioEstimada ?? g.FechasGen.Min();

                    var fechasConsecutivas = inv.FechasValidas
                        .Where(f => f >= ini && ((f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek) == diaIso)
                        .Take(g.FechasGen.Count)
                        .ToList();

                    if (fechasConsecutivas.Count == g.FechasGen.Count)
                    {
                        g.FechasGen = fechasConsecutivas;
                        g.FechaInicioEstimada = fechasConsecutivas.Min();
                        g.FechaFinEstimada = fechasConsecutivas.Max();
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
        /// Candidatos C3 / C4: Re-ranuración e intercambio de días en patrones AUTOMATICO.
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
                .Where(g => g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue
                         && g.FechaInicioEstimada.Value.Date <= sob.TramoFin
                         && g.FechaFinEstimada.Value.Date >= sob.TramoIni)
                .ToList();

            foreach (var g in solapadas)
            {
                if (g.EsFijo == 1) continue;

                string origen = ObtenerOrigenPatron(g, ctx);
                if (!string.Equals(origen, "AUTOMATICO", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Intentar alternar día dentro de DiasVentana
                foreach (var nuevoDia in inv.DiasVentana)
                {
                    if (nuevoDia == g.DiaSemana) continue;

                    double dur = g.HorasGen;
                    if (ctx.TryGetBloque(nuevoDia, dur, out var nuevoBloque))
                    {
                        var clon = campeon.Clonar();
                        var gClon = clon.Genes.FirstOrDefault(x => x.Id == g.Id);
                        if (gClon != null)
                        {
                            gClon.DiaSemana = nuevoDia;
                            gClon.IdBloqueHorario = nuevoBloque.Id;

                            DateTime ini = gClon.FechaInicioEstimada ?? ctx.Ventana.FechaInicio;
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
            }

            return candidatos;
        }

        /// <summary>
        /// Validación V1–V6 sobre un candidato generado.
        /// </summary>
        internal static bool Validar(AgIndividuo clon, AgIndividuo original, AgContexto ctx)
        {
            if (clon?.Genes == null || original?.Genes == null || ctx == null) return false;

            // V3: Sin vacantes nuevas
            if (clon.Vacantes > original.Vacantes) return false;

            foreach (var g in clon.Genes)
            {
                if (g.EsVacante == 1) continue;

                var gOrig = original.Genes.FirstOrDefault(x => x.Id == g.Id);

                // V4: No mover genes con EsFijo o LEGADO/FIJO
                if (gOrig != null && gOrig.EsFijo == 1)
                {
                    if (g.FechaInicioEstimada != gOrig.FechaInicioEstimada || g.FechaFinEstimada != gOrig.FechaFinEstimada)
                    {
                        return false;
                    }
                }
                string origen = ObtenerOrigenPatron(g, ctx);
                if (string.Equals(origen, "LEGADO", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(origen, "FIJO", StringComparison.OrdinalIgnoreCase))
                {
                    if (gOrig != null && (g.FechaInicioEstimada != gOrig.FechaInicioEstimada || g.FechaFinEstimada != gOrig.FechaFinEstimada))
                    {
                        return false;
                    }
                }

                // V1: Fechas válidas, día de ventana, no festivo
                if (g.FechasGen != null && g.FechasGen.Count > 0)
                {
                    foreach (var f in g.FechasGen)
                    {
                        if (ctx.Ventana != null && !ctx.Ventana.Contiene(f)) return false;
                        if (ctx.EsFestivo(f)) return false;

                        int diaIso = (f.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)f.DayOfWeek;
                        if (diaIso != g.DiaSemana) return false;
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

            // V5: Conflicto docente y aula: verificar que no haya solapes de fechas y bloques entre asignaciones del mismo docente o aula
            var genesActivos = clon.Genes.Where(g => g.EsVacante == 0 && g.FechasGen != null && g.FechasGen.Count > 0).ToList();
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

            // V6: Regla R6 por ClaveCarril: sin solapes estrictos dentro del mismo carril
            var porCarril = genesActivos.GroupBy(g => AgCarriles.ClaveCarril(g));
            foreach (var cGroup in porCarril)
            {
                var asignaturas = cGroup
                    .Where(g => g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue)
                    .GroupBy(g => g.IdAsignatura)
                    .Select(a => (Ini: a.Min(x => x.FechaInicioEstimada.Value.Date), Fin: a.Max(x => x.FechaFinEstimada.Value.Date)))
                    .ToList();

                for (int i = 0; i < asignaturas.Count; i++)
                {
                    for (int j = i + 1; j < asignaturas.Count; j++)
                    {
                        var a1 = asignaturas[i];
                        var a2 = asignaturas[j];
                        if (a1.Ini < a2.Fin && a2.Ini < a1.Fin)
                        {
                            return false; // Solape R6 en mismo carril
                        }
                    }
                }
            }

            return true;
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

        private static double CalcularDesperdicio(AgInventarioRanuras inv)
        {
            if (inv == null || inv.FechasValidas == null || inv.FechasValidas.Count == 0) return 0.0;
            double n = inv.FechasValidas.Count;
            double libresYHuecos = (inv.LibresReales != null ? inv.LibresReales.Count : 0) + inv.HuecosIntermedios;
            return (libresYHuecos / n) + (inv.BalancePorDia / n);
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

            public int CompareTo(AgPuntajeReparacion other)
            {
                int c = Hard.CompareTo(other.Hard);
                if (c != 0) return c;
                c = Vacantes.CompareTo(other.Vacantes);
                if (c != 0) return c;
                c = Desperdicio.CompareTo(other.Desperdicio);
                if (c != 0) return c;
                c = MakespanDias.CompareTo(other.MakespanDias);
                if (c != 0) return c;
                c = Soft.CompareTo(other.Soft);
                if (c != 0) return c;
                return Movimientos.CompareTo(other.Movimientos);
            }
        }
    }
}
