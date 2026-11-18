using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Operador de cruce (Crossover) para recombinar dos soluciones progenitoras (P1, P2) en dos descendientes (H1, H2).
    /// Implementa cruce ordenado (OX), preservación de la invariante del aula y reparación en cascada con diagnóstico de vacantes.
    /// </summary>
    public class AgOperadorCruce
    {
        private readonly Random _random;
        private readonly AgAsignadorTercero _asignador;
        private readonly AgPlanificadorPatronSemanal _planificadorPatron;

        public AgOperadorCruce(Random random = null, AgAsignadorTercero asignador = null, AgPlanificadorPatronSemanal planificadorPatron = null)
        {
            _random = random ?? new Random();
            _asignador = asignador ?? new AgAsignadorTercero(_random);
            _planificadorPatron = planificadorPatron ?? new AgPlanificadorPatronSemanal();
        }

        public AgOperadorCruce(int seed, AgAsignadorTercero asignador = null, AgPlanificadorPatronSemanal planificadorPatron = null)
            : this(new Random(seed), asignador, planificadorPatron)
        {
        }

        public (AgIndividuo Hijo1, AgIndividuo Hijo2) Cruzar(AgIndividuo padre1, AgIndividuo padre2, double tasaCruce, int generacion, AgContexto contexto = null)
        {
            if (padre1 == null) throw new ArgumentNullException(nameof(padre1));
            if (padre2 == null) throw new ArgumentNullException(nameof(padre2));

            var h1 = padre1.Clonar();
            var h2 = padre2.Clonar();
            h1.Generacion = generacion;
            h2.Generacion = generacion;

            if (_random.NextDouble() > tasaCruce || padre1.Genes.Count < 1)
            {
                return (h1, h2);
            }

            // Aplicar OX (Order Crossover)
            var dirtyH1 = new HashSet<long>();
            var dirtyH2 = new HashSet<long>();
            bool mod1 = HeredarOX(h1, padre2, contexto, dirtyH1);
            bool mod2 = HeredarOX(h2, padre1, contexto, dirtyH2);

            // Reparación en cascada de docentes
            if (contexto != null)
            {
                var pares = new[] { (h1, mod1, dirtyH1), (h2, mod2, dirtyH2) };
                foreach (var par in pares)
                {
                    var hijo = par.Item1;
                    var modificado = par.Item2;
                    var directos = par.Item3;
                    if (!modificado)
                    {
                        continue;
                    }

                    var proyeccion = new AgProyeccionIndividuo(hijo, contexto);
                    var dirtySet = ConstruirConjuntoSucio(hijo, directos, contexto);

                    foreach (var gen in hijo.Genes)
                    {
                        if (gen.Estado == AgEstadoGen.Fijado || (contexto.ManifiestoCobertura != null && !contexto.PuedeGenerarAula(gen.IdAula))) continue;

                        if (gen.EsVacante == 1 || gen.IdBloqueHorario <= 0)
                        {
                            gen.Estado = AgEstadoGen.Incompleto;
                            continue;
                        }

                        // Si el gen no está en el conjunto sucio conservador y ya tiene docente válido, omitir búsqueda
                        if (dirtySet != null && !dirtySet.Contains(gen.Id) && gen.Estado == AgEstadoGen.Valido && gen.IdTercero > 0)
                        {
                            continue;
                        }

                        if (proyeccion.DisponibleDocente(gen.IdTercero, gen))
                        {
                            gen.Estado = AgEstadoGen.Valido;
                            continue;
                        }

                        int alterno = _asignador.BuscarAlterno(gen, contexto, hijo, proyeccion);
                        if (alterno > 0)
                        {
                            gen.IdTercero = alterno;
                            gen.Estado = AgEstadoGen.Valido;
                            continue;
                        }

                        gen.IdTercero = 0;
                        gen.EsVacante = 1;
                        gen.MotivoVacante = _asignador.DiagnosticarMotivo(gen, contexto);
                        gen.Estado = AgEstadoGen.Incompleto;
                    }

                    AgCoherenciaGrupo.Sellar(hijo, contexto);   // Fase 521
                    AgMetricas.Recalcular(hijo, contexto);
                }
            }

            return (h1, h2);
        }

        private bool HeredarOX(AgIndividuo hijo, AgIndividuo donante, AgContexto contexto, HashSet<long> dirtyIds = null)
        {
            if (hijo.Genes == null || donante.Genes == null) return false;
            int n = Math.Min(hijo.Genes.Count, donante.Genes.Count);
            if (n < 1) return false;

            AgAula AulaSrcDst(int idAula) =>
                contexto?.AulasPorId != null && contexto.AulasPorId.TryGetValue(idAula, out var au) ? au : null;

            int punto1 = _random.Next(n);
            int punto2 = _random.Next(n);
            if (punto1 > punto2)
            {
                int t = punto1;
                punto1 = punto2;
                punto2 = t;
            }

            bool cambios = false;
            for (int i = punto1; i <= punto2; i++)
            {
                var dst = hijo.Genes[i];
                if (dst.Estado == AgEstadoGen.Fijado || (contexto?.ManifiestoCobertura != null && !contexto.PuedeGenerarAula(dst.IdAula))) continue;
                if (dst.EsFijoTemporal) continue;   // F659: conserva bloque/día/patrón decididos

                var src = donante.Genes[i];
                if (contexto?.ManifiestoCobertura != null && (src.IdAula != dst.IdAula
                    || !contexto.ManifiestoCobertura.Relaciones.Any(r => r.Clave.IdAula == dst.IdAula && r.Clave.IdAsignatura == src.IdAsignatura))) continue;

                // Fase 541 & Fase 664: no heredar asignatura de otro semestre ni de otro programa
                Asignatura asigSrc = null;
                AgAula aDst = null;
                string semAsig = (contexto?.AsignaturasPorId != null &&
                                  contexto.AsignaturasPorId.TryGetValue(src.IdAsignatura, out asigSrc))
                                 ? AgFormatoAsignatura.SemestreRomano(asigSrc.Semestre) : "SIN ASIGNAR";
                string semAula = AgFormatoAsignatura.SemestreRomano(dst.IdAula > 0 && contexto?.AulasPorId != null &&
                                  contexto.AulasPorId.TryGetValue(dst.IdAula, out aDst) ? aDst.Semestre : null);
                if (semAsig != "SIN ASIGNAR" && semAula != "SIN ASIGNAR" && semAsig != semAula) continue;

                int progAsig = asigSrc != null ? asigSrc.IdPrograma : 0;
                int progDst = dst.IdPrograma;
                int progAula = (aDst != null && aDst.IdPrograma > 0) ? aDst.IdPrograma : progDst;
                if (progAsig > 0 && progAula > 0 && progAsig != progAula) continue;

                int origAsig = dst.IdAsignatura;
                int origTercero = dst.IdTercero;
                int origBloque = dst.IdBloqueHorario;

                dst.IdAsignatura = src.IdAsignatura;
                dst.IdTercero = src.IdTercero;
                if (contexto != null && contexto.BloquesPorId != null && contexto.BloquesPorId.TryGetValue(src.IdBloqueHorario, out var bloque))
                {
                    dst.AsignarBloque(bloque);
                }
                else
                {
                    dst.IdBloqueHorario = src.IdBloqueHorario;
                    dst.DiaSemana = src.DiaSemana;
                    dst.HorasSemanales = src.HorasSemanales;
                }
                // Fase 1072: El plan de carril/oleada del descendiente manda: heredar sólo si está vacío
                if (!dst.FechaInicioEstimada.HasValue) dst.FechaInicioEstimada = src.FechaInicioEstimada;
                if (!dst.FechaFinEstimada.HasValue)    dst.FechaFinEstimada = src.FechaFinEstimada;
                dst.TotalHorasAsignatura = src.TotalHorasAsignatura;
                dst.EsVacante = src.EsVacante;
                dst.MotivoVacante = src.MotivoVacante;
                dst.DentroDeVentana = src.DentroDeVentana;
                dst.Estado = src.Estado;

                // Fase 546 & Fase 547: recalcular patrón y horas del gen heredado
                AgAulaAsignatura aaSrc = null;
                if (contexto?.AulaAsignaturaPorAulaAsig != null && dst.IdAula > 0 &&
                    contexto.AulaAsignaturaPorAulaAsig.TryGetValue((dst.IdAula, src.IdAsignatura), out var aaExacta))
                {
                    aaSrc = aaExacta;
                }
                else if (contexto?.AulaAsignaturasPorId != null &&
                         contexto.AulaAsignaturasPorId.TryGetValue(src.IdAsignatura, out var aaDef))
                {
                    aaSrc = aaDef;
                }

                // Fase 987 / Fase 1019: el carril del descendiente manda; pasar aula destino para contención de ventana
                var diasK = !string.IsNullOrWhiteSpace(dst.DiasPatron)
                    ? AgPlanificadorPatronSemanal.DiasKDe(dst.DiasPatron, null, AulaSrcDst(dst.IdAula)).ToList()
                    : (aaSrc != null
                        ? AgPlanificadorPatronSemanal.DiasKDe(aaSrc, AulaSrcDst(dst.IdAula)).ToList()
                        : new List<int> { dst.DiaSemana > 0 ? dst.DiaSemana : 2 });

                // Fase 548: el patrón de días se deriva canónicamente de diasK (AgAulaAsignatura no expone PatronSemanal)
                dst.DiasPatron = string.Join(",", diasK);

                // Fase 985 / D4': el patrón y los bloques se ajustan a la unidad del nivel del aula destino
                var aulaDst = AulaSrcDst(dst.IdAula);
                var progDstObj = aulaDst != null ? contexto?.Programas?.FirstOrDefault(p => p.Id == aulaDst.IdPrograma) : null;
                double unidadCruce = aulaDst != null
                    ? AgUnidadesSemanales.HorasSesionAula(aulaDst, AgNivelAcademicoCatalogo.Desde(progDstObj?.Modalidad))
                    : 4.0;

                // Fase 553, 554, 556, Fase 986, Fase 1086, Fase 1131, Fase 1135 & Fase 1168: re-asignar bloque coherente preservando franjaActual
                int diaObjetivo = diasK[Math.Min(dst.IndiceDiaPatron, diasK.Count - 1)];
                string franjaActual = (contexto?.BloquesPorId != null && dst.IdBloqueHorario > 0
                    && contexto.BloquesPorId.TryGetValue(dst.IdBloqueHorario, out var bAct)) ? bAct.Franja : null;
                var bloquePatron = AgSeleccionBloque.ParaDia(aulaDst, diaObjetivo, unidadCruce, contexto, franjaActual);

                if (bloquePatron != null)
                {
                    dst.AsignarBloque(bloquePatron);
                }
                else
                {
                    // Fase 1086 / Fase 1131 / Fase 1135 / Fase 1168: último recurso en días del patrón respetando franja y unidad
                    AgBloqueHorario bloqueUltimo = null;
                    foreach (var dAlt in diasK)
                    {
                        bloqueUltimo = AgSeleccionBloque.ParaDia(aulaDst, dAlt, unidadCruce, contexto, franjaActual);
                        if (bloqueUltimo != null) break;
                    }

                    if (bloqueUltimo != null)
                    {
                        dst.AsignarBloque(bloqueUltimo);
                    }
                    else
                    {
                        dst.IdBloqueHorario = 0;
                        dst.HorasSemanales = unidadCruce;
                        dst.EsVacante = 1;
                        dst.MotivoVacante = AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo();
                    }
                    dst.DiasPatron = string.Join(",", diasK);   // reafirmar patrón tras fijar bloque
                    dst.Estado = AgEstadoGen.Incompleto;
                    LoggerHelper.WarningLimitado(
                        $"CruceSinBloque|{contexto?.IdCorrida}|{diaObjetivo}",
                        $"[Cruce] Gen {dst.Id}: sin bloque de {unidadCruce}h para el día {diaObjetivo}; bloque={bloqueUltimo?.Id} (Incompleto).",
                        maxVeces: 1);
                }

                // Fase 550 & Fase 985: regenerar el histograma semanal desde la cohorte real usando unidadCruce
                DateTime fIni = dst.FechaInicioEstimada ?? (contexto?.Ventana != null ? contexto.Ventana.FechaInicio : DateTime.UtcNow.Date);
                DateTime fFin = dst.FechaFinEstimada ?? (contexto?.Ventana != null ? contexto.Ventana.FechaFin : DateTime.UtcNow.Date.AddMonths(4));
                var resPatron = _planificadorPatron.Construir(
                    diasK,
                    (int)dst.TotalHorasAsignatura,
                    fIni,
                    fFin,
                    contexto != null ? (Func<DateTime, bool>)contexto.EsFestivo : null,
                    unidadCruce);

                dst.PatronSemanal = resPatron.IsSuccess && resPatron.Value?.SesionesPorSemana != null
                    ? string.Join(",", resPatron.Value.SesionesPorSemana)
                    : string.Empty;

                // Fase 551, 564 & 569: vaciar fechas heredadas y delegar el reparto a Sellar solo si contexto != null
                dst.FechasGen = new List<DateTime>();
                dst.IndiceDiaPatron = Math.Min(Math.Max(0, dst.IndiceDiaPatron), Math.Max(0, diasK.Count - 1));
                if (contexto != null) dst.HorasGen = 0.0;

                // Fase 540: re-derivar cohorte desde el aula del hijo (preserva la invariante del aula)
                if (aulaDst != null)
                {
                    dst.IdSede = aulaDst.IdSede;
                    dst.IdPrograma = aulaDst.IdPrograma;
                    dst.Semestre = string.IsNullOrWhiteSpace(aulaDst.Semestre) ? dst.Semestre : aulaDst.Semestre;
                    dst.Grupo = string.IsNullOrWhiteSpace(aulaDst.Grupo) ? dst.Grupo : aulaDst.Grupo;
                }
                dst.ClaveGrupo = $"{dst.IdAsignatura}|{dst.IdAula}|{dst.Grupo}";

                if (dst.IdAsignatura != origAsig || dst.IdTercero != origTercero || dst.IdBloqueHorario != origBloque)
                {
                    cambios = true;
                    if (dirtyIds != null) dirtyIds.Add(dst.Id);
                }
            }

            if (cambios)
            {
                hijo.Estado = AgEstadoIndividuo.NoEvaluado;
            }
            return cambios;
        }

        private static HashSet<long> ConstruirConjuntoSucio(AgIndividuo hijo, HashSet<long> directos, AgContexto contexto)
        {
            var dirty = new HashSet<long>(directos ?? Enumerable.Empty<long>());
            if (hijo.Genes == null || directos == null || directos.Count == 0) return dirty;

            var dirtyGrupos = new HashSet<string>(StringComparer.Ordinal);
            var dirtyDocentes = new HashSet<int>();
            var dirtyAulas = new HashSet<int>();

            foreach (var g in hijo.Genes)
            {
                if (directos.Contains(g.Id))
                {
                    if (!string.IsNullOrWhiteSpace(g.ClaveGrupo)) dirtyGrupos.Add(g.ClaveGrupo);
                    if (g.IdTercero > 0) dirtyDocentes.Add(g.IdTercero);
                    if (g.IdAula > 0) dirtyAulas.Add(g.IdAula);
                }
            }

            foreach (var g in hijo.Genes)
            {
                if (dirty.Contains(g.Id)) continue;

                if (!string.IsNullOrWhiteSpace(g.ClaveGrupo) && dirtyGrupos.Contains(g.ClaveGrupo))
                {
                    dirty.Add(g.Id);
                    continue;
                }

                if (g.IdTercero > 0 && dirtyDocentes.Contains(g.IdTercero))
                {
                    dirty.Add(g.Id);
                    continue;
                }

                if (g.IdAula > 0 && dirtyAulas.Contains(g.IdAula))
                {
                    dirty.Add(g.Id);
                    continue;
                }
            }

            return dirty;
        }
    }
}
