using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Script.Serialization;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Genera un diagnóstico forense detallado y determinista a partir de las violaciones
    /// y genes de un individuo del Algoritmo Genético cuando es infactible (Fase 1994).
    /// </summary>
    internal static class AgDiagnosticoInfactibilidad
    {
        internal sealed class DetalleViolacionDto
        {
            public int[] indicesGenes { get; set; }
            public string fecha { get; set; }
            public IntervaloDto[] intervalos { get; set; }
        }

        internal sealed class IntervaloDto
        {
            public int IdAula { get; set; }
            public int IdAsignatura { get; set; }
            public int IdTercero { get; set; }
            public string inicio { get; set; }
            public string fin { get; set; }
        }

        private class ParConflicto
        {
            public string Clave { get; set; }
            public int IdTercero { get; set; }
            public string Fecha { get; set; }
            public string Franja { get; set; }
            public int IdAula1 { get; set; }
            public int IdAula2 { get; set; }
            public int IdAsig1 { get; set; }
            public int IdAsig2 { get; set; }
            public string Cohorte1 { get; set; }
            public string Cohorte2 { get; set; }
            public HashSet<string> Reglas { get; } = new HashSet<string>();
        }

        internal static IReadOnlyList<string> Describir(AgIndividuo individuo, AgContexto contexto)
        {
            if (individuo == null || individuo.HardViolations == 0)
            {
                return new List<string>();
            }

            var lineas = new List<string>();
            var serializer = new JavaScriptSerializer();

            int r1Count = 0;
            int r7Count = 0;
            int r7EntreGenes = 0;
            int r7Externa = 0;
            int detalleIlegible = 0;

            var paresDict = new Dictionary<string, ParConflicto>();
            var duras = (individuo.Violaciones ?? new List<AgViolacion>()).Where(v => v.EsDura).ToList();

            foreach (var v in duras)
            {
                if (v.Regla == AgReglaViolacion.R1_SuperposicionDocente) r1Count++;
                if (v.Regla == AgReglaViolacion.R7_TerceroOtroGrupoFechaHora) r7Count++;

                bool esParDocente = v.Regla == AgReglaViolacion.R1_SuperposicionDocente
                                 || v.Regla == AgReglaViolacion.R7_TerceroOtroGrupoFechaHora;
                if (!esParDocente) continue;

                if (string.IsNullOrWhiteSpace(v.DetalleJson) || v.DetalleJson == "{}") continue;

                DetalleViolacionDto det;
                try
                {
                    det = serializer.Deserialize<DetalleViolacionDto>(v.DetalleJson);
                }
                catch
                {
                    detalleIlegible++;
                    continue;
                }

                if (det?.intervalos == null || det.intervalos.Length == 0) continue;

                if (v.Regla == AgReglaViolacion.R7_TerceroOtroGrupoFechaHora)
                {
                    if (det.intervalos.Length == 1 || (v.Mensaje != null && v.Mensaje.StartsWith("Ocupación externa", StringComparison.OrdinalIgnoreCase)))
                    {
                        r7Externa++;
                        continue;
                    }
                    else
                    {
                        r7EntreGenes++;
                    }
                }

                if (det.intervalos.Length >= 2)
                {
                    var i1 = det.intervalos[0];
                    var i2 = det.intervalos[1];
                    int t = i1.IdTercero > 0 ? i1.IdTercero : i2.IdTercero;
                    string f = det.fecha ?? string.Empty;
                    string fr = $"{i1.inicio}-{i1.fin}";

                    // Clave normalizada no orientada
                    string k1 = $"{i1.IdAula}_{i1.IdAsignatura}";
                    string k2 = $"{i2.IdAula}_{i2.IdAsignatura}";
                    string parSubKey = string.CompareOrdinal(k1, k2) < 0 ? $"{k1}|{k2}" : $"{k2}|{k1}";
                    string clavePar = $"{t}|{f}|{fr}|{parSubKey}";

                    if (!paresDict.TryGetValue(clavePar, out var par))
                    {
                        par = new ParConflicto
                        {
                            Clave = clavePar,
                            IdTercero = t,
                            Fecha = f,
                            Franja = fr,
                            IdAula1 = i1.IdAula,
                            IdAula2 = i2.IdAula,
                            IdAsig1 = i1.IdAsignatura,
                            IdAsig2 = i2.IdAsignatura
                        };

                        if (det.indicesGenes != null && det.indicesGenes.Length >= 2 && individuo.Genes != null)
                        {
                            int idx1 = det.indicesGenes[0];
                            int idx2 = det.indicesGenes[1];
                            if (idx1 >= 0 && idx1 < individuo.Genes.Count)
                                par.Cohorte1 = $"{individuo.Genes[idx1].Semestre}/{individuo.Genes[idx1].Grupo}";
                            if (idx2 >= 0 && idx2 < individuo.Genes.Count)
                                par.Cohorte2 = $"{individuo.Genes[idx2].Semestre}/{individuo.Genes[idx2].Grupo}";
                        }

                        paresDict[clavePar] = par;
                    }

                    if (v.Regla == AgReglaViolacion.R1_SuperposicionDocente) par.Reglas.Add("R1");
                    if (v.Regla == AgReglaViolacion.R7_TerceroOtroGrupoFechaHora) par.Reglas.Add("R7");
                }
            }

            var paresList = paresDict.Values.Where(p => p.IdTercero > 0).ToList();
            var docentesUnicos = paresList.Select(p => p.IdTercero).Where(t => t > 0).Distinct().ToList();
            var asignaturasUnicas = paresList.SelectMany(p => new[] { p.IdAsig1, p.IdAsig2 }).Where(a => a > 0).Distinct().ToList();
            var aulasUnicas = paresList.SelectMany(p => new[] { p.IdAula1, p.IdAula2 }).Where(a => a > 0).Distinct().ToList();

            var durasPorRegla = duras
                .GroupBy(v => v.Regla)
                .Select(g => $"{g.Key} x{g.Count()}")
                .ToList();
            string reglasDurasStr = durasPorRegla.Count > 0 ? string.Join(", ", durasPorRegla) : "Ninguna";

            lineas.Add(string.Format(CultureInfo.InvariantCulture,
                "[AG-CAUSAS-DETALLE] HARD={0} · ReglasDuras=[{1}] · paresDocente={2} (R1={3}, R7={4}: entreGenes={5}, ocupacionExterna={6}) · docentes={7} · asignaturas={8} · aulas={9}",
                individuo.HardViolations, reglasDurasStr, paresList.Count, r1Count, r7Count, r7EntreGenes, r7Externa,
                docentesUnicos.Count, asignaturasUnicas.Count, aulasUnicas.Count));

            if (detalleIlegible > 0)
            {
                lineas.Add($"[AG-CAUSAS-DETALLE] DetalleIlegible={detalleIlegible} violaciones sin intervalo parseable.");
            }

            // Agrupamiento por docente con más pares
            var porDocente = paresList
                .GroupBy(p => p.IdTercero)
                .OrderByDescending(g => g.Count())
                .ThenBy(g => g.Key)
                .Take(5)
                .ToList();

            foreach (var gd in porDocente)
            {
                int idTercero = gd.Key;
                var asigsDoc = gd.SelectMany(p => new[] { p.IdAsig1, p.IdAsig2 }).Distinct().OrderBy(x => x).ToList();
                var aulasDoc = gd.SelectMany(p => new[] { p.IdAula1, p.IdAula2 }).Distinct().OrderBy(x => x).ToList();
                var cohortesDoc = gd.SelectMany(p => new[] { p.Cohorte1, p.Cohorte2 }).Where(c => !string.IsNullOrEmpty(c)).Distinct().OrderBy(x => x).ToList();
                var fechasDoc = gd.Select(p => p.Fecha).Where(f => !string.IsNullOrEmpty(f)).Distinct().OrderBy(x => x).ToList();
                var franjasDoc = gd.Select(p => p.Franja).Where(fr => !string.IsNullOrEmpty(fr)).Distinct().OrderBy(x => x).ToList();

                var nombresAsigs = asigsDoc.Select(id =>
                {
                    if (contexto != null && contexto.AsignaturasPorId != null && contexto.AsignaturasPorId.TryGetValue(id, out var aObj) && aObj != null)
                        return $"{id} ({aObj.Nombre})";
                    return id.ToString(CultureInfo.InvariantCulture);
                });

                var poolHab = asigsDoc.Select(id =>
                {
                    int hab = 0;
                    if (contexto != null && contexto.PoolPorAsignatura != null && contexto.PoolPorAsignatura.TryGetValue(id, out var pList) && pList != null)
                    {
                        int idCentinela = contexto.IdTerceroVacante > 0 ? contexto.IdTerceroVacante : -1;
                        var genAsig = (individuo.Genes ?? new List<AgGen>()).FirstOrDefault(x => x.IdAsignatura == id && x.IdPrograma > 0);
                        hab = pList.Count(p => p.Activo == 1 && p.IdTercero != idCentinela
                            && (genAsig == null || contexto.EsDocenteHabilitado(p.IdTercero, genAsig.IdPrograma, id, genAsig.Semestre)));
                    }
                    return $"({id})={hab}";
                });

                lineas.Add(string.Format(CultureInfo.InvariantCulture,
                    "[AG-CAUSAS-DETALLE] Docente {0}: pares={1} · asigs={2} · aulas={3} · cohortes={4} · fechas={5} · franjas={6} · poolHabilitado={7}",
                    idTercero, gd.Count(),
                    string.Join(",", nombresAsigs.Take(3)) + (nombresAsigs.Count() > 3 ? $" (+{nombresAsigs.Count() - 3})" : string.Empty),
                    string.Join(",", aulasDoc),
                    string.Join(" , ", cohortesDoc),
                    fechasDoc.Count,
                    string.Join(",", franjasDoc),
                    string.Join(" ", poolHab)));
            }

            // Muestra de hasta 3 pares más críticos
            foreach (var par in paresList.Take(3))
            {
                lineas.Add(string.Format(CultureInfo.InvariantCulture,
                    "[AG-CAUSAS-DETALLE] Par: {0} {1} · asig {2} aula {3} ({4}) ↔ asig {5} aula {6} ({7}) · tercero {8} · reglas {9}",
                    par.Fecha, par.Franja,
                    par.IdAsig1, par.IdAula1, par.Cohorte1 ?? "?",
                    par.IdAsig2, par.IdAula2, par.Cohorte2 ?? "?",
                    par.IdTercero, string.Join(",", par.Reglas)));
            }

            // Desglose de vacantes reclasificadas
            var genesVac = (individuo.Genes ?? new List<AgGen>()).Where(g => g.EsVacante == 1 || g.IdTercero <= 0).ToList();
            if (genesVac.Count > 0)
            {
                int cPoolVacio = 0;
                int cSinHabActiva = 0;
                int cSinHabProgSem = 0;
                int cGenIncompleto = 0;
                int cSinBloque = 0;
                int cSinDispHoraria = 0;
                int cSinPlanOleada = 0;

                foreach (var g in genesVac)
                {
                    var motivoCod = AgMotivoVacanteExtensions.Desde(g.MotivoVacante);
                    if (motivoCod == AgMotivoVacante.SIN_PLAN_OLEADA)
                    {
                        cSinPlanOleada++;
                        continue;
                    }
                    if (motivoCod == AgMotivoVacante.SIN_BLOQUE_NIVEL)
                    {
                        cSinBloque++;
                        continue;
                    }
                    if (g.IdPrograma <= 0 || string.IsNullOrWhiteSpace(g.Semestre))
                    {
                        cGenIncompleto++;
                        continue;
                    }
                    if (g.IdBloqueHorario <= 0)
                    {
                        cSinBloque++;
                        continue;
                    }

                    if (contexto == null || contexto.PoolPorAsignatura == null ||
                        !contexto.PoolPorAsignatura.TryGetValue(g.IdAsignatura, out var pList) || pList == null || pList.Count == 0)
                    {
                        cPoolVacio++;
                        continue;
                    }

                    var activos = pList.Where(p => p.Activo == 1 && p.IdTercero != contexto.IdTerceroVacante).ToList();
                    if (activos.Count == 0)
                    {
                        cSinHabActiva++;
                        continue;
                    }

                    bool habProgSem = activos.Any(p => contexto.EsDocenteHabilitado(p.IdTercero, g.IdPrograma, g.IdAsignatura, g.Semestre));
                    if (!habProgSem)
                    {
                        cSinHabProgSem++;
                        continue;
                    }

                    cSinDispHoraria++;
                }

                var porMotivoReg = genesVac
                    .GroupBy(g => string.IsNullOrEmpty(g.MotivoVacante) ? "SIN_MOTIVO" : g.MotivoVacante)
                    .Select(g => $"{g.Key}={g.Count()}")
                    .ToList();

                int cConsentidas = genesVac.Count(g => contexto != null && contexto.EstaConsentida(g.IdAsignatura, g.IdSede));
                int cLiberadas = genesVac.Count(g => (contexto == null || !contexto.EstaConsentida(g.IdAsignatura, g.IdSede)) && AgClasificadorVacante.Clasificar(g.MotivoVacante) == AgOrigenVacante.LiberadaAg);
                int cEstructuralesNc = genesVac.Count(g => (contexto == null || !contexto.EstaConsentida(g.IdAsignatura, g.IdSede)) && AgClasificadorVacante.Clasificar(g.MotivoVacante) == AgOrigenVacante.EstructuralDocente);
                int cNoDocente = genesVac.Count(g => (contexto == null || !contexto.EstaConsentida(g.IdAsignatura, g.IdSede)) && AgClasificadorVacante.Clasificar(g.MotivoVacante) == AgOrigenVacante.NoDocente);
                int cIndeterminado = genesVac.Count(g => (contexto == null || !contexto.EstaConsentida(g.IdAsignatura, g.IdSede)) && AgClasificadorVacante.Clasificar(g.MotivoVacante) == AgOrigenVacante.Indeterminado);

                lineas.Add(string.Format(CultureInfo.InvariantCulture,
                    "[AG-CAUSAS-DETALLE] Vacantes={0} (CONSENTIDAS={1}, LIBERADAS_POR_AG={2}, ESTRUCTURALES_NC={3}, NO_DOCENTE={4}, INDETERMINADAS={5}) · motivoRegistrado=[{6}] [POOL_VACIO={7}, SIN_HABILITACION_ACTIVA={8}, SIN_HABILITACION_PROGRAMA_SEMESTRE={9}, SIN_DISPONIBILIDAD_HORARIA_NO_VERIFICADA={10}, GEN_INCOMPLETO={11}, SIN_BLOQUE_NIVEL={12}, SIN_PLAN_OLEADA={13}]",
                    genesVac.Count, cConsentidas, cLiberadas, cEstructuralesNc, cNoDocente, cIndeterminado, string.Join(",", porMotivoReg),
                    cPoolVacio, cSinHabActiva, cSinHabProgSem, cSinDispHoraria, cGenIncompleto, cSinBloque, cSinPlanOleada));

                if (cSinHabProgSem > 0 || cPoolVacio > 0)
                {
                    lineas.Add(string.Format(CultureInfo.InvariantCulture,
                        "[AG-CAUSAS-DETALLE] Accion: {0} genes sin docente habilitado en catalogo/pool · {1} pares con docente en colision simultanea → verificar asignaciones o habilitar docentes alternos.",
                        cSinHabProgSem + cPoolVacio, paresList.Count));
                }
            }

            return lineas;
        }

        internal static string ResumenCorto(AgIndividuo individuo, AgContexto contexto, AgMotivoParada motivoParada, int hardViolations)
        {
            try
            {
                var desc = Describir(individuo, contexto);
                if (desc == null || desc.Count == 0)
                {
                    return $"Optimización sin factibilidad ({hardViolations} violaciones duras). Parada: {motivoParada}.";
                }

                var l0 = desc.FirstOrDefault(l => l.Contains("paresDocente="));
                var lDoc = desc.FirstOrDefault(l => l.Contains("Docente "));
                var lVac = desc.FirstOrDefault(l => l.Contains("Vacantes="));
                var lAcc = desc.FirstOrDefault(l => l.Contains("Accion:"));

                var duras = (individuo?.Violaciones ?? new List<AgViolacion>())
                    .Where(v => v.EsDura)
                    .GroupBy(v => v.Regla)
                    .Select(g => $"{g.Key} x{g.Count()}")
                    .ToList();
                string reglasDurasStr = duras.Count > 0 ? string.Join(", ", duras) : "Ninguna";

                var partes = new List<string>
                {
                    $"Optimización finalizada sin factibilidad (HARD={hardViolations}, Parada={motivoParada}).",
                    $"Reglas duras: [{reglasDurasStr}]"
                };

                if (!string.IsNullOrEmpty(l0)) partes.Add(l0.Replace("[AG-CAUSAS-DETALLE] ", string.Empty));
                if (!string.IsNullOrEmpty(lDoc)) partes.Add(lDoc.Replace("[AG-CAUSAS-DETALLE] ", string.Empty));
                if (!string.IsNullOrEmpty(lVac)) partes.Add(lVac.Replace("[AG-CAUSAS-DETALLE] ", string.Empty));
                if (!string.IsNullOrEmpty(lAcc)) partes.Add(lAcc.Replace("[AG-CAUSAS-DETALLE] ", string.Empty));
                partes.Add("No se materializar\u00e1n sesiones en staging.");

                partes.Add("Revise la bitácora [AG-CAUSAS-DETALLE] para el desglose forense por aula y asignatura.");
                return string.Join("\n", partes);
            }
            catch
            {
                return $"Optimización finalizada sin factibilidad ({hardViolations} violaciones duras). Revise la demanda y disponibilidad docente.";
            }
        }
    }
}
