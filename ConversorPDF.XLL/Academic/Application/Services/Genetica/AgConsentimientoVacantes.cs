using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Calcula las vacantes estructurales y el universo de consentimiento a partir de la demanda real aula-asignatura (Fases 2052, 2058, 2059).
    /// Evita basar el consentimiento sobre el catálogo global y preserva la granularidad por sede sin emitir wildcards prematuros.
    /// </summary>
    internal static class AgConsentimientoVacantes
    {
        /// <summary>
        /// Evalúa la cobertura efectiva para cada relación demanda aula-asignatura y retorna el detalle de vacantes estructurales.
        /// </summary>
        /// <param name="ctx">Contexto global de optimización del AG.</param>
        /// <param name="asigs">Conjunto de identificadores de asignaturas donde TODAS sus cohortes demandadas carecen de docente.</param>
        /// <param name="claves">Conjunto de claves compuestas asignatura:sede para granularidad fina por sede.</param>
        /// <returns>Lista de información detallada de cada vacante estructural encontrada.</returns>
        internal static List<VacanteEstructuralInfo> Calcular(AgContexto ctx, out HashSet<int> asigs, out HashSet<string> claves)
        {
            asigs = new HashSet<int>();
            claves = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var detalle = new List<VacanteEstructuralInfo>();
            if (ctx?.AulaAsignaturas == null || ctx.AulasPorId == null || ctx.AsignaturasPorId == null) return detalle;

            var porAsignatura = ctx.AulaAsignaturas
                .Where(aa => ctx.AulasPorId.TryGetValue(aa.IdAula, out var au) && au != null
                          && ctx.AsignaturasPorId.ContainsKey(aa.IdAsignatura)
                          && (ctx.ManifiestoCobertura == null || ctx.PuedeGenerarAula(aa.IdAula)))
                .Select(aa =>
                {
                    var aula = ctx.AulasPorId[aa.IdAula];
                    var asig = ctx.AsignaturasPorId[aa.IdAsignatura];
                    return new { aa.IdAula, aa.IdAsignatura, Aula = aula, Asignatura = asig };
                })
                .GroupBy(x => x.IdAsignatura);

            foreach (var g in porAsignatura)
            {
                var evaluadas = g.Select(x =>
                {
                    int idProg = x.Asignatura.IdPrograma > 0 ? x.Asignatura.IdPrograma : x.Aula.IdPrograma;
                    var hab = ctx.PoolHabilitado(x.IdAsignatura, idProg, x.Aula.Semestre);
                    return new
                    {
                        x.IdAsignatura,
                        x.Aula,
                        x.Asignatura,
                        IdProg = idProg,
                        Hab = hab
                    };
                }).ToList();

                foreach (var e in evaluadas.Where(e => e.Hab.Count == 0))
                {
                    if (e.Aula.IdSede > 0)
                    {
                        claves.Add($"{g.Key}:{e.Aula.IdSede}");
                    }

                    // Fase 2058: Utilizar la sobrecarga con semestre de aula y programa de la cohorte
                    var motivo = AgFabricaIndividuos.ClasificarMotivo(ctx, e.Asignatura, e.Aula.Semestre, e.IdProg);
                    detalle.Add(new VacanteEstructuralInfo
                    {
                        IdAsignatura = e.IdAsignatura,
                        IdSede = e.Aula.IdSede,
                        Semestre = e.Aula.Semestre,
                        Grupo = e.Aula.Grupo,
                        IdAula = e.Aula.Id,
                        Motivo = motivo.Codigo()
                    });
                }

                // Fase 2059: Wildcard sólo si TODAS las cohortes demandadas de la materia quedan sin docente
                if (evaluadas.All(e => e.Hab.Count == 0))
                {
                    asigs.Add(g.Key);
                }
            }

            return detalle;
        }

        /// <summary>
        /// Determina de forma pura si el consentimiento actual en el contexto sigue vigente frente al universo demandado (Fase 2071).
        /// </summary>
        internal static bool EsVigente(AgContexto ctx, HashSet<int> asigs, HashSet<string> claves) =>
            ctx?.AsignaturasVacantesConsentidas != null
         && ctx?.ClavesVacantesConsentidas != null
         && ctx.AsignaturasVacantesConsentidas.SetEquals(asigs)
         && ctx.ClavesVacantesConsentidas.SetEquals(claves);
    }
}
