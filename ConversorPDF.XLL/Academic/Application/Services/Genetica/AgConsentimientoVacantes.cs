using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Calcula las vacantes estructurales y el universo de consentimiento a partir de la demanda real aula-asignatura (Fases 2052, 2058, 2059).
    /// Evita basar el consentimiento sobre el catálogo global y preserva la granularidad por sede sin emitir wildcards prematuros.
    /// Fase 2083: Helper de formateo legible agrupado por asignatura y sede.
    /// Fase 2084: Trazabilidad DEBUG [AG-CONSENT] para justificación de decisiones.
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
                .GroupBy(x => x.IdAsignatura)
                .ToList();

            int omitidasConDocente = 0;

            foreach (var g in porAsignatura)
            {
                var evaluadas = g.Select(x =>
                {
                    int idProg = x.Asignatura.IdPrograma > 0 ? x.Asignatura.IdPrograma : x.Aula.IdPrograma;
                    var hab = ctx.PoolHabilitado(x.IdAsignatura, idProg, x.Aula.Semestre);
                    return new
                    {
                        x.IdAula,
                        x.IdAsignatura,
                        x.Aula,
                        x.Asignatura,
                        IdProg = idProg,
                        Hab = hab
                    };
                }).ToList();

                var sinHab = evaluadas.Where(e => e.Hab.Count == 0).ToList();

                foreach (var e in sinHab)
                {
                    // Fase 2091: Registrar clave con sede 0 si IdSede <= 0 para proteger contra R16 permanente
                    claves.Add(e.Aula.IdSede > 0 ? $"{g.Key}:{e.Aula.IdSede}" : $"{g.Key}:0");
                    if (e.Aula.IdSede <= 0 && LoggerHelper.ModoVerboso)
                    {
                        LoggerHelper.Debug($"[AG-CONSENT] asig={e.IdAsignatura} aula={e.Aula.Id} IdSede=0 → registrado token {g.Key}:0 para proteger contra R16");
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

                string decision;
                // Fase 2059: Wildcard sólo si TODAS las cohortes demandadas de la materia quedan sin docente
                if (evaluadas.All(e => e.Hab.Count == 0))
                {
                    asigs.Add(g.Key);
                    decision = "WILDCARD";
                }
                else if (sinHab.Count > 0)
                {
                    decision = "PARCIAL";
                }
                else
                {
                    omitidasConDocente++;
                    decision = "NINGUNA (no entra al diálogo)";
                }

                if (LoggerHelper.ModoVerboso)
                {
                    var asigObj = ctx.AsignaturasPorId.TryGetValue(g.Key, out var aso) ? aso : null;
                    string nom = asigObj?.Nombre ?? "N/D";
                    var pool = ctx.PoolPorAsignatura != null && ctx.PoolPorAsignatura.TryGetValue(g.Key, out var pl) ? pl : null;
                    int poolFilas = pool?.Count ?? 0;
                    int activosReales = pool?.Count(p => p.Activo == 1 && (ctx.IdTerceroVacante <= 0 || p.IdTercero != ctx.IdTerceroVacante)) ?? 0;
                    int habProgSem = evaluadas.Sum(e => e.Hab.Count);
                    var motivoDiag = evaluadas.Count > 0
                        ? AgFabricaIndividuos.ClasificarMotivo(ctx, evaluadas[0].Asignatura, evaluadas[0].Aula.Semestre, evaluadas[0].IdProg).Codigo()
                        : "SIN_DATOS";
                    string cohortesStr = string.Join(", ", evaluadas.Select(e => $"{e.Aula.IdSede}:{e.Aula.Semestre}/{e.Aula.Grupo}"));

                    LoggerHelper.Debug($"[AG-CONSENT] asig={g.Key} nom=\"{nom}\" demanda={evaluadas.Count} cohortes=[{cohortesStr}] poolFilas={poolFilas} activosReales={activosReales} habilitadosProgSem={habProgSem} motivo={motivoDiag} decision={decision}");
                }
            }

            if (LoggerHelper.ModoVerboso)
            {
                LoggerHelper.Debug($"[AG-CONSENT] resumen: asignaturas={porAsignatura.Count} cohortes={detalle.Count} wildcards={asigs.Count} claves={claves.Count} omitidasConDocente={omitidasConDocente}");
            }

            return detalle;
        }

        /// <summary>
        /// Ordena semestres numéricamente (1..12) convirtiendo desde formato romano/arábigo (Fase 2083).
        /// Los valores no reconocidos se ubican al final.
        /// </summary>
        internal static int OrdenSemestre(string semestre)
        {
            if (string.IsNullOrWhiteSpace(semestre)) return int.MaxValue;
            string arabigo = AgFormatoAsignatura.SemestreArabigo(semestre);
            if (int.TryParse(arabigo, out int n)) return n;
            return int.MaxValue;
        }

        /// <summary>
        /// Formatea las vacantes estructurales agrupadas por asignatura y sede en formato legible (Fase 2083).
        /// Ejemplo: "37: sede VALLE · III, IV → SIN_CANDIDATOS".
        /// </summary>
        /// <param name="detalle">Colección de vacantes estructurales detectadas.</param>
        /// <param name="resolverSede">Función opcional para resolver el nombre o subunidad de la sede.</param>
        /// <returns>Cadena multilínea unida por viñeta para presentación en modal o log.</returns>
        internal static string FormatearDetalle(IEnumerable<VacanteEstructuralInfo> detalle, Func<int, string> resolverSede = null)
        {
            if (detalle == null) return string.Empty;
            var list = detalle.ToList();
            if (list.Count == 0) return string.Empty;

            Func<int, string> resolver = resolverSede ?? (id => id > 0 ? $"sede {id}" : "sede N/D");

            var lineas = list
                .GroupBy(d => d.IdAsignatura)
                .Select(g =>
                {
                    var sedesParte = string.Join("; ", g
                        .GroupBy(x => x.IdSede)
                        .OrderBy(sg => sg.Key)
                        .Select(sg =>
                        {
                            string etiquetaSede = resolver(sg.Key);
                            if (string.IsNullOrWhiteSpace(etiquetaSede))
                            {
                                etiquetaSede = sg.Key > 0 ? $"sede {sg.Key}" : "sede N/D";
                            }
                            var semestres = sg
                                .Select(x => AgFormatoAsignatura.SemestreRomano(x.Semestre))
                                .Distinct()
                                .OrderBy(OrdenSemestre);
                            string semestresStr = string.Join(", ", semestres);
                            return string.IsNullOrWhiteSpace(semestresStr)
                                ? etiquetaSede
                                : $"{etiquetaSede} · {semestresStr}";
                        }));

                    // Fase 2093: Motivo modal no vacío (con desempate alfabético determinista)
                    string motivo = g.Where(x => !string.IsNullOrWhiteSpace(x.Motivo))
                        .GroupBy(x => x.Motivo)
                        .OrderByDescending(m => m.Count())
                        .ThenBy(m => m.Key, StringComparer.OrdinalIgnoreCase)
                        .Select(m => m.Key)
                        .FirstOrDefault() ?? "SIN_MOTIVO";

                    return $"{g.Key}: {sedesParte} → {motivo}";
                });

            return string.Join("\n• ", lineas);
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
