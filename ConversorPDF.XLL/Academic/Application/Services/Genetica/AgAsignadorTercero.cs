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
    /// Asignador y reparador de docentes desde el pool de habilitaciones para genes individuales.
    /// </summary>
    public class AgAsignadorTercero
    {
        private readonly Random _random;

        public AgAsignadorTercero(Random random = null)
        {
            _random = random ?? new Random();
        }

        public AgAsignadorTercero(int seed)
            : this(new Random(seed))
        {
        }

        internal static bool MismoGrupo(AgGen a, AgGen b) =>
            a != null && b != null
            && a.IdAula == b.IdAula && a.IdAsignatura == b.IdAsignatura
            && a.IdPrograma == b.IdPrograma && a.IdSede == b.IdSede
            && string.Equals(a.Grupo ?? "", b.Grupo ?? "", StringComparison.Ordinal);

        /// <summary>
        /// Determina si un docente está disponible para todo el grupo de genes co-programados de una asignatura/aula,
        /// comprobando habilitación en pool, ausencia de solapamiento fechado contra otros genes del individuo
        /// y disponibilidad de agenda externa en AgContexto.TerceroOcupado.
        /// </summary>
        public static bool DisponibleGrupo(int docente, AgGen gen, AgIndividuo individuo, AgContexto contexto, AgProyeccionIndividuo proyeccion = null)
        {
            if (proyeccion != null)
            {
                return proyeccion.DisponibleDocente(docente, gen);
            }

            long t0 = 0;
            if (AgTelemetriaSubetapas.Activo)
            {
                AgTelemetriaSubetapas.LlamadasDisponibleGrupo++;
                t0 = System.Diagnostics.Stopwatch.GetTimestamp();
            }

            try
            {
                if (gen == null || contexto == null) return false;
                if (docente <= 0 || docente == contexto.IdTerceroVacante) return false;
                var grupo = (individuo != null && individuo.Genes != null)
                    ? individuo.Genes.Where(g => MismoGrupo(g, gen)).ToList()
                    : new List<AgGen>();
                if (grupo.Count == 0) grupo.Add(gen);
                if (grupo.Any(g => g.Estado == AgEstadoGen.Fijado && g.IdTercero != docente))
                    return false;
                if (grupo.Any(g => !contexto.EsDocenteHabilitado(docente, g.IdPrograma, g.IdAsignatura, g.Semestre)))
                    return false;

                var propias = AgCalendarioConflictos.Crear(grupo, contexto);
                if (propias.IsFailure) return false;

                var otrosGenes = (individuo != null && individuo.Genes != null)
                    ? individuo.Genes.Where(g => !MismoGrupo(g, gen) && AgCalendarioConflictos.DocenteReal(g, contexto) && g.IdTercero == docente)
                    : Enumerable.Empty<AgGen>();
                var ajenas = AgCalendarioConflictos.Crear(otrosGenes, contexto);
                if (ajenas.IsFailure) return false;

                foreach (var o in propias.Value)
                {
                    if (ajenas.Value.Any(a => AgCalendarioConflictos.SeCruzan(o, a))) return false;
                    if (contexto.SesionesConservadas != null && contexto.SesionesConservadas.Count > 0)
                    {
                        foreach (var s in contexto.SesionesConservadas)
                        {
                            if (s.IdTercero == docente && s.Fecha.Date == o.Fecha && s.Inicio < o.Fin && o.Inicio < s.Fin)
                                return false;
                        }
                    }
                    if (contexto.TerceroOcupado(docente, o.Fecha, new ConversorPDF.XLL.Academic.Domain.Models.FranjaHoraria(o.Inicio, o.Fin)))
                        return false;
                }
                return true;
            }
            finally
            {
                if (AgTelemetriaSubetapas.Activo)
                {
                    AgTelemetriaSubetapas.TiempoDisponibleGrupoTicks += (System.Diagnostics.Stopwatch.GetTimestamp() - t0);
                }
            }
        }

        public bool AsignarDocente(AgGen gen, AgContexto contexto, HashSet<int> docentesExcluidos = null)
        {
            if (gen == null || contexto == null) return false;

            if (contexto.PoolPorAsignatura.TryGetValue(gen.IdAsignatura, out var pool) && pool.Count > 0)
            {
                var candidatos = pool.Where(p => p.Activo == 1).ToList();
                if (docentesExcluidos != null && docentesExcluidos.Count > 0)
                {
                    candidatos = candidatos.Where(p => !docentesExcluidos.Contains(p.IdTercero)).ToList();
                }

                if (candidatos.Count > 0)
                {
                    // Ordenar por prioridad ascendente (1 es mayor prioridad)
                    var ordenados = candidatos.OrderBy(c => c.Prioridad).ThenBy(c => c.VecesAsignada).ToList();
                    int mejorPrioridad = ordenados[0].Prioridad;
                    var mejores = ordenados.Where(c => c.Prioridad == mejorPrioridad).ToList();

                    var elegido = mejores[_random.Next(mejores.Count)];
                    gen.IdTercero = elegido.IdTercero;
                    gen.EsVacante = 0;
                    gen.MotivoVacante = string.Empty;
                    return true;
                }
            }

            // Fase 1896-B: Reutilizar clasificación canónica DiagnosticarMotivo (habilitación, ventana, disponibilidad)
            gen.IdTercero = 0;
            gen.EsVacante = 1;
            gen.MotivoVacante = DiagnosticarMotivo(gen, contexto);
            return false;
        }

        public bool DocenteDisponible(AgGen gen, AgIndividuo individuo, AgContexto contexto)
        {
            if (gen == null || contexto == null || gen.IdTercero <= 0) return false;
            return DisponibleGrupo(gen.IdTercero, gen, individuo, contexto);
        }

        public int BuscarAlterno(AgGen gen, AgContexto contexto, AgIndividuo individuo)
        {
            return BuscarAlterno(gen, contexto, individuo, null);
        }

        public int BuscarAlterno(AgGen gen, AgContexto contexto, AgIndividuo individuo, AgProyeccionIndividuo proyeccion)
        {
            if (gen == null || contexto == null) return 0;

            if (contexto.PoolPorAsignatura.TryGetValue(gen.IdAsignatura, out var pool) && pool.Count > 0)
            {
                var habilitados = pool.Where(p => p.Activo == 1).OrderBy(p => p.Prioridad).ThenBy(p => p.VecesAsignada).ToList();
                foreach (var cand in habilitados)
                {
                    if (cand.IdTercero == gen.IdTercero) continue;

                    bool disp = proyeccion != null
                        ? proyeccion.DisponibleDocente(cand.IdTercero, gen)
                        : DisponibleGrupo(cand.IdTercero, gen, individuo, contexto);

                    if (disp)
                    {
                        if (individuo != null && individuo.Genes != null)
                        {
                            foreach (var m in individuo.Genes.Where(g => MismoGrupo(g, gen) && g.Estado != AgEstadoGen.Fijado))
                            {
                                m.IdTercero = cand.IdTercero;
                                m.EsVacante = 0;
                                m.MotivoVacante = string.Empty;
                            }
                        }
                        else if (gen.Estado != AgEstadoGen.Fijado)
                        {
                            gen.IdTercero = cand.IdTercero;
                            gen.EsVacante = 0;
                            gen.MotivoVacante = string.Empty;
                        }

                        if (proyeccion != null)
                        {
                            proyeccion.ActualizarDocenteGrupo(gen, cand.IdTercero);
                        }

                        return cand.IdTercero;
                    }
                }
            }

            return 0;
        }

        public int BuscarYAsignarAlternoGrupo(AgGen gen, AgContexto contexto, AgIndividuo individuo)
        {
            return BuscarAlterno(gen, contexto, individuo);
        }

        /// <summary>
        /// Diagnostica el motivo formal de vacante para un gen no asignado o liberado (Fase 1902 / Fase 2013).
        /// Versión estática invocable desde operadores genéticos sin requerir instancia.
        /// </summary>
        public static string DiagnosticarMotivoVacante(AgGen gen, AgContexto contexto)
        {
            if (gen == null || contexto == null) return AgMotivoVacante.SIN_CANDIDATOS.Codigo();

            string motivo;
            int poolCount = 0;
            int activosCount = 0;
            int habilitadosCount = 0;

            if (!contexto.PoolPorAsignatura.TryGetValue(gen.IdAsignatura, out var pool) || pool.Count == 0 || !pool.Any(p => p.Activo == 1))
            {
                motivo = AgMotivoVacante.SIN_CANDIDATOS.Codigo();
            }
            else
            {
                poolCount = pool.Count;
                activosCount = pool.Count(p => p.Activo == 1 && (contexto.IdTerceroVacante <= 0 || p.IdTercero != contexto.IdTerceroVacante));

                // Fase 1902: si el gen carece de programa o semestre (mutación/recombinación), no atribuir erróneamente a NO_CUMPLE_REQUISITOS
                if (gen.IdPrograma <= 0 || string.IsNullOrWhiteSpace(gen.Semestre))
                {
                    motivo = AgMotivoVacante.SIN_DISPONIBILIDAD.Codigo();
                }
                else
                {
                    habilitadosCount = pool.Count(p => p.Activo == 1 && contexto.EsDocenteHabilitado(p.IdTercero, gen.IdPrograma, gen.IdAsignatura, gen.Semestre));

                    // Si existen candidatos en pool para la materia pero ninguno habilitado para este programa/semestre
                    if (habilitadosCount == 0)
                    {
                        motivo = AgMotivoVacante.NO_CUMPLE_REQUISITOS.Codigo();
                    }
                    // Si las fechas del gen caen fuera de la ventana semestral
                    else if (contexto.Ventana != null && gen.FechaInicioEstimada.HasValue && !contexto.Ventana.Contiene(gen.FechaInicioEstimada.Value))
                    {
                        motivo = AgMotivoVacante.FUERA_DE_VENTANA.Codigo();
                    }
                    else
                    {
                        motivo = AgMotivoVacante.SIN_DISPONIBILIDAD.Codigo();
                    }
                }
            }

            if (LoggerHelper.ModoVerboso)
            {
                LoggerHelper.Debug($"[AG-VACACION] Diagnosticar gen={gen.Id} asig={gen.IdAsignatura} pool={poolCount} activos={activosCount} habilitados={habilitadosCount} gen prog={gen.IdPrograma}/sem={gen.Semestre} motivo={motivo}");
            }

            return motivo;
        }

        /// <summary>
        /// Sobrecarga de instancia para preservar compatibilidad con llamadores existentes.
        /// </summary>
        public string DiagnosticarMotivo(AgGen gen, AgContexto contexto)
        {
            return DiagnosticarMotivoVacante(gen, contexto);
        }
    }

    public static class AgTelemetriaSubetapas
    {
        [ThreadStatic]
        private static bool _activo;

        [ThreadStatic]
        public static long LlamadasDisponibleGrupo;

        [ThreadStatic]
        public static long CalendariosCreados;

        [ThreadStatic]
        public static long SesionesProyectadas;

        [ThreadStatic]
        public static long CandidatosEvaluados;

        [ThreadStatic]
        public static long ConteoSellados;

        [ThreadStatic]
        public static long TiempoSelladoTicks;

        [ThreadStatic]
        public static long TiempoDisponibleGrupoTicks;

        public static bool Activo => _activo;

        public static void Iniciar()
        {
            _activo = true;
            Reiniciar();
        }

        public static void Detener()
        {
            _activo = false;
        }

        public static void Reiniciar()
        {
            LlamadasDisponibleGrupo = 0;
            CalendariosCreados = 0;
            SesionesProyectadas = 0;
            CandidatosEvaluados = 0;
            ConteoSellados = 0;
            TiempoSelladoTicks = 0;
            TiempoDisponibleGrupoTicks = 0;
        }

        public static long TiempoSelladoMs => System.Diagnostics.Stopwatch.Frequency > 0
            ? TiempoSelladoTicks * 1000 / System.Diagnostics.Stopwatch.Frequency
            : 0;

        public static long TiempoDisponibleGrupoMs => System.Diagnostics.Stopwatch.Frequency > 0
            ? TiempoDisponibleGrupoTicks * 1000 / System.Diagnostics.Stopwatch.Frequency
            : 0;
    }

    public sealed class AgProyeccionIndividuo
    {
        private readonly Dictionary<long, List<AgCalendarioConflictos.Ocupacion>> _ocupacionesPorGen =
            new Dictionary<long, List<AgCalendarioConflictos.Ocupacion>>();
        private readonly Dictionary<int, List<AgCalendarioConflictos.Ocupacion>> _ocupacionesPorDocente =
            new Dictionary<int, List<AgCalendarioConflictos.Ocupacion>>();
        private readonly Dictionary<string, List<AgGen>> _genesPorGrupo =
            new Dictionary<string, List<AgGen>>(StringComparer.Ordinal);

        public AgIndividuo Individuo { get; }
        public AgContexto Contexto { get; }

        public AgProyeccionIndividuo(AgIndividuo individuo, AgContexto contexto)
        {
            Individuo = individuo ?? throw new ArgumentNullException(nameof(individuo));
            Contexto = contexto ?? throw new ArgumentNullException(nameof(contexto));
            ConstruirIndices();
        }

        public void ConstruirIndices()
        {
            _ocupacionesPorGen.Clear();
            _ocupacionesPorDocente.Clear();
            _genesPorGrupo.Clear();

            if (Individuo.Genes == null) return;

            foreach (var g in Individuo.Genes)
            {
                string clave = string.IsNullOrWhiteSpace(g.ClaveGrupo) ? $"{g.IdAsignatura}|{g.IdAula}|{g.Grupo}" : g.ClaveGrupo;
                if (!_genesPorGrupo.TryGetValue(clave, out var lista))
                {
                    lista = new List<AgGen>();
                    _genesPorGrupo[clave] = lista;
                }
                lista.Add(g);

                if (g.HorasGen > 0 && Contexto.BloquesPorId != null && Contexto.BloquesPorId.TryGetValue(g.IdBloqueHorario, out var bloque) && bloque.Activo == 1 && bloque.EsHorarioValido)
                {
                    var resFechas = AgExpansorSesiones.ProyectarFechas(g, Contexto);
                    if (resFechas.IsSuccess)
                    {
                        var oList = new List<AgCalendarioConflictos.Ocupacion>(resFechas.Value.Count);
                        foreach (var f in resFechas.Value)
                        {
                            var oc = new AgCalendarioConflictos.Ocupacion
                            {
                                Gen = g,
                                Fecha = f.Date,
                                Inicio = bloque.TimeSpanInicio,
                                Fin = bloque.TimeSpanFin,
                                Franja = bloque.Franja ?? string.Empty
                            };
                            oList.Add(oc);

                            if (AgCalendarioConflictos.DocenteReal(g, Contexto) && g.IdTercero > 0)
                            {
                                if (!_ocupacionesPorDocente.TryGetValue(g.IdTercero, out var docList))
                                {
                                    docList = new List<AgCalendarioConflictos.Ocupacion>();
                                    _ocupacionesPorDocente[g.IdTercero] = docList;
                                }
                                docList.Add(oc);
                            }
                        }
                        _ocupacionesPorGen[g.Id] = oList;
                    }
                }
            }
        }

        public IReadOnlyList<AgGen> ObtenerGrupo(AgGen gen)
        {
            if (gen == null) return Array.Empty<AgGen>();
            string clave = string.IsNullOrWhiteSpace(gen.ClaveGrupo) ? $"{gen.IdAsignatura}|{gen.IdAula}|{gen.Grupo}" : gen.ClaveGrupo;
            return _genesPorGrupo.TryGetValue(clave, out var l) ? l : (IReadOnlyList<AgGen>)new[] { gen };
        }

        public bool DisponibleDocente(int docente, AgGen gen)
        {
            if (gen == null || Contexto == null) return false;
            if (docente <= 0 || docente == Contexto.IdTerceroVacante) return false;

            var grupo = ObtenerGrupo(gen);
            for (int i = 0; i < grupo.Count; i++)
            {
                var g = grupo[i];
                if (g.Estado == AgEstadoGen.Fijado && g.IdTercero != docente) return false;
                if (!Contexto.EsDocenteHabilitado(docente, g.IdPrograma, g.IdAsignatura, g.Semestre)) return false;
            }

            // Ocupaciones de este grupo
            List<AgCalendarioConflictos.Ocupacion> ocupGrupo = null;
            for (int i = 0; i < grupo.Count; i++)
            {
                if (_ocupacionesPorGen.TryGetValue(grupo[i].Id, out var oList))
                {
                    if (ocupGrupo == null) ocupGrupo = new List<AgCalendarioConflictos.Ocupacion>(oList);
                    else ocupGrupo.AddRange(oList);
                }
            }

            if (ocupGrupo == null || ocupGrupo.Count == 0)
            {
                var res = AgCalendarioConflictos.Crear(grupo, Contexto);
                if (res.IsFailure) return false;
                ocupGrupo = res.Value;
            }

            // Ocupaciones ajenas del docente
            if (_ocupacionesPorDocente.TryGetValue(docente, out var docOcup))
            {
                for (int i = 0; i < ocupGrupo.Count; i++)
                {
                    var o = ocupGrupo[i];
                    for (int j = 0; j < docOcup.Count; j++)
                    {
                        var a = docOcup[j];
                        if (AgAsignadorTercero.MismoGrupo(o.Gen, a.Gen)) continue;
                        if (AgCalendarioConflictos.SeCruzan(o, a)) return false;
                    }
                }
            }

            // Comprobar sesiones conservadas y agenda externa
            for (int i = 0; i < ocupGrupo.Count; i++)
            {
                var o = ocupGrupo[i];
                if (Contexto.SesionesConservadas != null && Contexto.SesionesConservadas.Count > 0)
                {
                    foreach (var s in Contexto.SesionesConservadas)
                    {
                        if (s.IdTercero == docente && s.Fecha.Date == o.Fecha && s.Inicio < o.Fin && o.Inicio < s.Fin)
                            return false;
                    }
                }
                if (Contexto.TerceroOcupado(docente, o.Fecha, new ConversorPDF.XLL.Academic.Domain.Models.FranjaHoraria(o.Inicio, o.Fin)))
                    return false;
            }

            return true;
        }

        public void ActualizarDocenteGrupo(AgGen gen, int nuevoDocente)
        {
            var grupo = ObtenerGrupo(gen);
            for (int i = 0; i < grupo.Count; i++)
            {
                var g = grupo[i];
                if (_ocupacionesPorGen.TryGetValue(g.Id, out var oList))
                {
                    if (g.IdTercero > 0 && _ocupacionesPorDocente.TryGetValue(g.IdTercero, out var viejas))
                    {
                        viejas.RemoveAll(x => x.Gen == g);
                    }

                    if (nuevoDocente > 0 && AgCalendarioConflictos.DocenteReal(g, Contexto))
                    {
                        if (!_ocupacionesPorDocente.TryGetValue(nuevoDocente, out var nuevas))
                        {
                            nuevas = new List<AgCalendarioConflictos.Ocupacion>();
                            _ocupacionesPorDocente[nuevoDocente] = nuevas;
                        }
                        nuevas.AddRange(oList);
                    }
                }
            }
        }
    }
}
