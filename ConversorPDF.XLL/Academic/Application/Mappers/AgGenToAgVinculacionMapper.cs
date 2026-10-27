using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Mappers
{
    /// <summary>
    /// Mapeador estático de AgGen / grupos de AgGen a AgVinculacion para la persistencia en staging.
    /// </summary>
    public static class AgGenToAgVinculacionMapper
    {
        public static AgVinculacion Map(AgGen gen, long idCorrida, long idIndividuo, string subunidad)
        {
            if (gen == null) return null;

            return new AgVinculacion
            {
                IdCorrida = idCorrida,
                IdIndividuo = idIndividuo,
                IdTercero = gen.IdTercero,
                IdPrograma = gen.IdPrograma,
                IdSede = gen.IdSede,
                IdAsignatura = gen.IdAsignatura,
                IdPeriodo = gen.IdPeriodo,
                Semestre = gen.Semestre ?? string.Empty,
                Grupo = gen.Grupo ?? "A",
                TotalHoras = gen.TotalHorasAsignatura,
                FechaInicio = gen.FechaInicioEstimada.HasValue ? gen.FechaInicioEstimada.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : string.Empty,
                FechaFin = gen.FechaFinEstimada.HasValue ? gen.FechaFinEstimada.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : string.Empty,
                Estado = AgEstadoVinculacion.Borrador,
                EsVacante = gen.EsVacante,
                MotivoVacante = gen.MotivoVacante ?? string.Empty,
                CandidatosPool = gen.CandidatosPool,
                Subunidad = subunidad ?? string.Empty
            };
        }

        /// <summary>
        /// Mapea un conjunto agrupado de genes de una misma asignatura/docente/aula/ventana a una única vinculación de staging.
        /// Asigna Estado = Borrador y calcula TotalHoras como la suma de las horas de los genes.
        /// </summary>
        public static AgVinculacion MapGrupo(List<AgGen> genes, long idCorrida, long idIndividuo, AgContexto contexto)
        {
            if (genes == null || genes.Count == 0) return null;

            var g0 = genes[0];
            string modalidad = ModalidadTutoriaCatalog.Presencial;
            if (contexto != null && contexto.AulasPorId != null && contexto.AulasPorId.TryGetValue(g0.IdAula, out var a))
            {
                modalidad = ModalidadTutoriaCatalog.NormalizarTutoria(a?.ModalidadVinculacion);
            }

            int idSede = g0.IdSede;
            if (idSede <= 0 && contexto != null)
            {
                if (contexto.AulasPorId != null && contexto.AulasPorId.TryGetValue(g0.IdAula, out var aSede) && aSede.IdSede > 0)
                {
                    idSede = aSede.IdSede;
                }
                else if (contexto.IdSede > 0)
                {
                    idSede = contexto.IdSede;
                }
            }

            int idPeriodo = (contexto?.Ventana != null && contexto.Ventana.IdPeriodo > 0)
                ? contexto.Ventana.IdPeriodo
                : (g0.IdPeriodo > 0 ? g0.IdPeriodo : (contexto != null ? contexto.IdPeriodo : 0));

            // Fase 1108 / Fase 1109: Preservar sub-ventana de oleada usando el rango (Min..Max) de los genes del grupo en lugar de sobrescribir con la ventana semestral
            var fiGrupo = genes.Where(g => g.FechaInicioEstimada.HasValue).Select(g => g.FechaInicioEstimada).Min();
            var ffGrupo = genes.Where(g => g.FechaFinEstimada.HasValue).Select(g => g.FechaFinEstimada).Max();

            string fechaInicio = fiGrupo?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                ?? (contexto?.Ventana != null ? contexto.Ventana.FechaInicio.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : string.Empty);

            string fechaFin = ffGrupo?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                ?? (contexto?.Ventana != null ? contexto.Ventana.FechaFin.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : string.Empty);

            int idPrograma = g0.IdPrograma;
            if (contexto?.AsignaturasPorId != null && contexto.AsignaturasPorId.TryGetValue(g0.IdAsignatura, out var asigRef) && asigRef.IdPrograma > 0)
            {
                idPrograma = asigRef.IdPrograma;
            }

            return new AgVinculacion
            {
                IdCorrida = idCorrida,
                IdIndividuo = idIndividuo,
                IdSede = idSede,
                IdTercero = (g0.EsVacante == 1 && contexto != null && contexto.IdTerceroVacante > 0) ? contexto.IdTerceroVacante : g0.IdTercero,
                IdPrograma = idPrograma,
                IdAsignatura = g0.IdAsignatura,
                IdPeriodo = idPeriodo,
                Semestre = g0.Semestre ?? string.Empty,
                Grupo = g0.Grupo ?? "A",
                ModalidadVinculacion = modalidad,
                TotalHoras = genes.Sum(x => x.HorasGen),
                FechaInicio = fechaInicio,
                FechaFin = fechaFin,
                Estado = AgEstadoVinculacion.Borrador,
                EsVacante = g0.EsVacante,
                MotivoVacante = g0.MotivoVacante ?? string.Empty,
                CandidatosPool = g0.CandidatosPool,
                Subunidad = contexto != null ? (contexto.Subunidad ?? string.Empty) : string.Empty,
                GenesOrigen = genes.Select(x => x.Id).ToList()
            };
        }
    }
}
