using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
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

            // 1. Validar habilitación en pool
            if (!contexto.EsDocenteHabilitado(gen.IdTercero, gen.IdPrograma, gen.IdAsignatura, gen.Semestre))
            {
                return false;
            }

            // 2. Validar superposición interna en el individuo
            if (individuo != null && individuo.Genes != null && gen.IdBloqueHorario > 0)
            {
                foreach (var other in individuo.Genes)
                {
                    if (ReferenceEquals(other, gen)) continue;
                    if (other.IdTercero == gen.IdTercero && other.IdBloqueHorario == gen.IdBloqueHorario)
                    {
                        return false;
                    }
                }
            }

            // 3. Validar ocupación externa (sesiones de producción o staging aprobadas)
            if (contexto.OcupacionExterna != null && contexto.OcupacionExterna.Count > 0 &&
                gen.FechaInicioEstimada.HasValue && gen.IdBloqueHorario > 0 &&
                contexto.BloquesPorId != null && contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out var bloque))
            {
                if (TimeSpan.TryParse(bloque.HoraInicio, out var tIni) && TimeSpan.TryParse(bloque.HoraFin, out var tFin) && tIni < tFin)
                {
                    var franja = new ConversorPDF.XLL.Academic.Domain.Models.FranjaHoraria(tIni, tFin);
                    if (contexto.TerceroOcupado(gen.IdTercero, gen.FechaInicioEstimada.Value, franja))
                    {
                        return false;
                    }
                }
            }

            return true;
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

                    bool tieneConflictoInterno = false;
                    if (individuo != null && individuo.Genes != null && gen.IdBloqueHorario > 0)
                    {
                        foreach (var other in individuo.Genes)
                        {
                            if (ReferenceEquals(other, gen)) continue;
                            if (other.IdTercero == cand.IdTercero && other.IdBloqueHorario == gen.IdBloqueHorario)
                            {
                                tieneConflictoInterno = true;
                                break;
                            }
                        }
                    }

                    if (!tieneConflictoInterno)
                    {
                        // Validar ocupación externa para el candidato alterno
                        if (contexto.OcupacionExterna != null && contexto.OcupacionExterna.Count > 0 &&
                            gen.FechaInicioEstimada.HasValue && gen.IdBloqueHorario > 0 &&
                            contexto.BloquesPorId != null && contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out var blk))
                        {
                            if (TimeSpan.TryParse(blk.HoraInicio, out var tI) && TimeSpan.TryParse(blk.HoraFin, out var tF) && tI < tF)
                            {
                                var fr = new ConversorPDF.XLL.Academic.Domain.Models.FranjaHoraria(tI, tF);
                                if (contexto.TerceroOcupado(cand.IdTercero, gen.FechaInicioEstimada.Value, fr))
                                {
                                    continue;
                                }
                            }
                        }

                        return cand.IdTercero;
                    }
                }
            }

            return 0;
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
