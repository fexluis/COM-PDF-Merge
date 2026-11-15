using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

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

        private static bool MismoGrupo(AgGen a, AgGen b) =>
            a != null && b != null
            && a.IdAula == b.IdAula && a.IdAsignatura == b.IdAsignatura
            && a.IdPrograma == b.IdPrograma && a.IdSede == b.IdSede
            && string.Equals(a.Grupo ?? "", b.Grupo ?? "", StringComparison.Ordinal);

        private static bool DisponibleGrupo(int docente, AgGen gen, AgIndividuo individuo, AgContexto contexto)
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
                if (contexto.TerceroOcupado(docente, o.Fecha, new ConversorPDF.XLL.Academic.Domain.Models.FranjaHoraria(o.Inicio, o.Fin)))
                    return false;
            }
            return true;
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

            // Si no hay candidatos disponibles en el pool, se marca como vacante
            gen.IdTercero = 0;
            gen.EsVacante = 1;
            gen.MotivoVacante = AgMotivoVacante.SIN_CANDIDATOS.Codigo();
            return false;
        }

        public bool DocenteDisponible(AgGen gen, AgIndividuo individuo, AgContexto contexto)
        {
            if (gen == null || contexto == null || gen.IdTercero <= 0) return false;
            return DisponibleGrupo(gen.IdTercero, gen, individuo, contexto);
        }

        public int BuscarAlterno(AgGen gen, AgContexto contexto, AgIndividuo individuo)
        {
            if (gen == null || contexto == null) return 0;

            if (contexto.PoolPorAsignatura.TryGetValue(gen.IdAsignatura, out var pool) && pool.Count > 0)
            {
                var habilitados = pool.Where(p => p.Activo == 1).OrderBy(p => p.Prioridad).ThenBy(p => p.VecesAsignada).ToList();
                foreach (var cand in habilitados)
                {
                    if (cand.IdTercero == gen.IdTercero) continue;

                    if (DisponibleGrupo(cand.IdTercero, gen, individuo, contexto))
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

        public string DiagnosticarMotivo(AgGen gen, AgContexto contexto)
        {
            if (gen == null || contexto == null) return AgMotivoVacante.SIN_CANDIDATOS.Codigo();

            if (!contexto.PoolPorAsignatura.TryGetValue(gen.IdAsignatura, out var pool) || pool.Count == 0 || !pool.Any(p => p.Activo == 1))
            {
                return AgMotivoVacante.SIN_CANDIDATOS.Codigo();
            }

            // Si existen candidatos en pool para la materia pero ninguno habilitado para este programa/semestre
            if (!pool.Any(p => p.Activo == 1 && contexto.EsDocenteHabilitado(p.IdTercero, gen.IdPrograma, gen.IdAsignatura, gen.Semestre)))
            {
                return AgMotivoVacante.NO_CUMPLE_REQUISITOS.Codigo();
            }

            // Si las fechas del gen caen fuera de la ventana semestral
            if (contexto.Ventana != null && gen.FechaInicioEstimada.HasValue && !contexto.Ventana.Contiene(gen.FechaInicioEstimada.Value))
            {
                return AgMotivoVacante.FUERA_DE_VENTANA.Codigo();
            }

            return AgMotivoVacante.SIN_DISPONIBILIDAD.Codigo();
        }
    }
}
