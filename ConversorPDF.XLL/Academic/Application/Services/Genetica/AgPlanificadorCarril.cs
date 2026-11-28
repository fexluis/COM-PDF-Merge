using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Planificador determinista de cadena por carril para R6 (Fase 2371, Fase 2375, Fase 2377, Fase 2381, Fase 2383).
    /// Secuencia todas las asignaturas de un carril considerando capacidad real sobre fechas libres, colisiones externas,
    /// orden curricular estricto O(1), multi-gen independiente por día y anclas fijas basadas en fechas reales.
    /// </summary>
    internal static class AgPlanificadorCarril
    {
        private sealed class SujetoCarril
        {
            public int IdAsignatura { get; set; }
            public List<int> Indices { get; set; } = new List<int>();
            public int N { get; set; }
            public Dictionary<int, int> RequeridasPorDia { get; set; } = new Dictionary<int, int>();
            public bool EsFijo { get; set; }
            public DateTime? IniFijo { get; set; }
            public DateTime? FinFijo { get; set; }
        }

        internal static AgResultadoPlanCarril Planificar(
            AgIndividuo ind, AgContexto ctx, string cohorte, string carril, string ventanaAmpliada = null)
        {
            if (ind?.Genes == null || ctx == null)
            {
                return new AgResultadoPlanCarril { Motivo = "SIN_CARRIL" };
            }

            var genesCarril = ind.Genes
                .Select((g, i) => new { Gen = g, Indice = i })
                .Where(x => x.Gen.EsVacante == 0
                    && string.Equals(AgCalendarioConflictos.Cohorte(x.Gen), cohorte, StringComparison.Ordinal)
                    && string.Equals(AgCarriles.ClaveCarril(x.Gen), carril, StringComparison.Ordinal))
                .ToList();

            if (genesCarril.Count <= 1)
            {
                return new AgResultadoPlanCarril { Factible = true, Individuo = ind };
            }

            var inv = AgInventarioRanuras.Construir(ind, ctx, cohorte);
            var lane = inv?.Carriles?.FirstOrDefault(c => string.Equals(c.Clave, carril, StringComparison.Ordinal));
            if (lane == null || lane.FechasValidas == null || lane.FechasValidas.Count == 0)
            {
                return new AgResultadoPlanCarril { Motivo = "SIN_CARRIL" };
            }

            // FASE 2445: Si se suministra ventanaAmpliada ante fallo previo por capacidad o fechas, usar fechas válidas de la ventana ampliada
            var fechasValidas = lane.FechasValidas;
            if (!string.IsNullOrEmpty(ventanaAmpliada) && ctx.Ventana != null && ctx.Ventana.FechaInicio <= ctx.Ventana.FechaFin)
            {
                var diasAmp = AgDiasVentana.Dias(ventanaAmpliada);
                if (diasAmp.Count > 0)
                {
                    var validasAmp = new List<DateTime>();
                    for (var d = ctx.Ventana.FechaInicio.Date; d <= ctx.Ventana.FechaFin.Date; d = d.AddDays(1))
                    {
                        int diaIso = AgInventarioRanuras.Iso(d);
                        if (diasAmp.Contains(diaIso) && !ctx.EsFestivo(d))
                        {
                            validasAmp.Add(d);
                        }
                    }
                    if (validasAmp.Count > 0)
                    {
                        fechasValidas = validasAmp;
                    }
                }
            }

            int idAulaV = genesCarril[0].Gen.IdAula;

            var sujetos = genesCarril
                .GroupBy(x => x.Gen.IdAsignatura)
                .Select(grp =>
                {
                    var indices = grp.Select(x => x.Indice).OrderBy(i => ind.Genes[i].DiaSemana).ToList();
                    var reqDia = new Dictionary<int, int>();
                    foreach (var x in grp)
                    {
                        int dia = x.Gen.DiaSemana;
                        int nGen = (x.Gen.FechasGen != null && x.Gen.FechasGen.Count > 0)
                            ? x.Gen.FechasGen.Count
                            : (int)Math.Max(1, Math.Round(x.Gen.HorasGen / 4.0));
                        reqDia[dia] = reqDia.TryGetValue(dia, out int act) ? act + nGen : nGen;
                    }

                    int totalN = reqDia.Values.Sum();
                    bool fijo = grp.Any(x => x.Gen.EsFijo == 1);

                    DateTime? iniFijo = null;
                    DateTime? finFijo = null;
                    if (fijo)
                    {
                        foreach (var x in grp)
                        {
                            DateTime? iniG = x.Gen.FechaInicioEstimada ?? (x.Gen.FechasGen != null && x.Gen.FechasGen.Count > 0 ? (DateTime?)x.Gen.FechasGen.Min() : null);
                            DateTime? finG = x.Gen.FechaFinEstimada ?? (x.Gen.FechasGen != null && x.Gen.FechasGen.Count > 0 ? (DateTime?)x.Gen.FechasGen.Max() : null);
                            if (iniG.HasValue && (!iniFijo.HasValue || iniG.Value < iniFijo.Value)) iniFijo = iniG;
                            if (finG.HasValue && (!finFijo.HasValue || finG.Value > finFijo.Value)) finFijo = finG;
                        }
                    }

                    return new SujetoCarril
                    {
                        IdAsignatura = grp.Key,
                        Indices = indices,
                        N = totalN,
                        RequeridasPorDia = reqDia,
                        EsFijo = fijo,
                        IniFijo = iniFijo,
                        FinFijo = finFijo
                    };
                })
                .ToList();

            int reqTotal = sujetos.Sum(s => s.N);

            // Deducción previa de ocupación externa para medir capacidad real sobre fechas libres (Fase 2383)
            var indicesCadena = new HashSet<int>(genesCarril.Select(x => x.Indice));
            var ocupadasFuera = new HashSet<DateTime>();

            for (int i = 0; i < ind.Genes.Count; i++)
            {
                if (indicesCadena.Contains(i)) continue;
                var g = ind.Genes[i];
                if (g.EsVacante != 0) continue;

                bool mismaAula = g.IdAula == idAulaV;
                bool mismaCohorte = string.Equals(AgCalendarioConflictos.Cohorte(g), cohorte, StringComparison.Ordinal);
                if (!mismaAula && !mismaCohorte) continue;

                foreach (var f in AgInventarioRanuras.FechasDeGen(g))
                {
                    ocupadasFuera.Add(f.Date);
                }
            }

            int diasCarril = sujetos
                .SelectMany(s => s.RequeridasPorDia.Keys)
                .Distinct()
                .Count();
            int sesionesPorSemana = Math.Max(1, diasCarril > 0 ? diasCarril : 2);

            var libres = fechasValidas.Where(f => !ocupadasFuera.Contains(f.Date)).OrderBy(f => f).ToList();
            if (libres.Count == 0)
            {
                return new AgResultadoPlanCarril
                {
                    Motivo = "SIN_FECHAS",
                    SesionesRequeridas = reqTotal,
                    SesionesDisponibles = 0,
                    SesionesPorSemana = sesionesPorSemana
                };
            }

            int semanasDisp = sujetos.SelectMany(s => s.RequeridasPorDia.Keys)
                .Distinct()
                .Select(d => libres.Count(f => AgInventarioRanuras.Iso(f) == d))
                .DefaultIfEmpty(0)
                .Max();

            // Fase 2402, Fase 2406: Capacidad real del carril comparando sesiones requeridas totales contra fechas válidas del carril
            int capacidadTotal = fechasValidas.Count;
            if (reqTotal > capacidadTotal)
            {
                return new AgResultadoPlanCarril
                {
                    Motivo = "CAPACIDAD",
                    SesionesRequeridas = reqTotal,
                    SesionesDisponibles = capacidadTotal,
                    SesionesPorSemana = sesionesPorSemana
                };
            }

            foreach (var s in sujetos)
            {
                if (s.EsFijo) continue;
                foreach (var kvp in s.RequeridasPorDia)
                {
                    int dia = kvp.Key;
                    int nReq = kvp.Value;
                    int dispLibres = libres.Count(f => AgInventarioRanuras.Iso(f) == dia);
                    if (nReq > dispLibres)
                    {
                        int dispTotal = fechasValidas.Count(f => AgInventarioRanuras.Iso(f) == dia);
                        string mot = (nReq > dispTotal) ? "CAPACIDAD" : "SIN_FECHAS";
                        return new AgResultadoPlanCarril
                        {
                            Motivo = mot,
                            SesionesRequeridas = nReq,
                            SesionesDisponibles = dispLibres,
                            SesionesPorSemana = sesionesPorSemana
                        };
                    }
                }
            }

            var fijos = sujetos.Where(s => s.EsFijo).ToList();
            for (int i = 0; i < fijos.Count; i++)
            {
                for (int j = i + 1; j < fijos.Count; j++)
                {
                    if (fijos[i].IniFijo.HasValue && fijos[i].FinFijo.HasValue &&
                        fijos[j].IniFijo.HasValue && fijos[j].FinFijo.HasValue)
                    {
                        if (fijos[i].IniFijo.Value < fijos[j].FinFijo.Value && fijos[j].IniFijo.Value < fijos[i].FinFijo.Value)
                        {
                            return new AgResultadoPlanCarril
                            {
                                Motivo = "FIJO",
                                SesionesRequeridas = reqTotal,
                                SesionesDisponibles = semanasDisp,
                                SesionesPorSemana = sesionesPorSemana
                            };
                        }
                    }
                }
            }

            string ultimoMotivo = string.Empty;
            foreach (var orden in GenerarOrdenes(sujetos, idAulaV, ctx))
            {
                var clon = ind.Clonar();
                if (AsignarCadena(clon, sujetos, orden, libres, out var motivo))
                {
                    return new AgResultadoPlanCarril
                    {
                        Factible = true,
                        Individuo = clon,
                        SesionesRequeridas = reqTotal,
                        SesionesDisponibles = semanasDisp,
                        SesionesPorSemana = sesionesPorSemana
                    };
                }
                if (!string.IsNullOrEmpty(motivo))
                {
                    ultimoMotivo = motivo;
                }
            }

            return new AgResultadoPlanCarril
            {
                Motivo = !string.IsNullOrEmpty(ultimoMotivo) ? ultimoMotivo : "SIN_ORDEN",
                SesionesRequeridas = reqTotal,
                SesionesDisponibles = semanasDisp,
                SesionesPorSemana = sesionesPorSemana
            };
        }

        private static IEnumerable<List<SujetoCarril>> GenerarOrdenes(
            List<SujetoCarril> sujetos, int idAula, AgContexto ctx)
        {
            // Búsqueda O(1) de orden curricular vía diccionario indexado (Fase 2381)
            int OrdenCurricular(SujetoCarril s)
            {
                if (ctx?.AulaAsignaturaPorAulaAsig != null &&
                    ctx.AulaAsignaturaPorAulaAsig.TryGetValue((idAula, s.IdAsignatura), out var rel) &&
                    rel.OrdenCurricular > 0)
                {
                    return rel.OrdenCurricular;
                }
                if (ctx?.AulaAsignaturas != null)
                {
                    var relFallback = ctx.AulaAsignaturas.FirstOrDefault(aa => aa.IdAula == idAula && aa.IdAsignatura == s.IdAsignatura);
                    if (relFallback != null && relFallback.OrdenCurricular > 0) return relFallback.OrdenCurricular;
                }
                return int.MaxValue;
            }

            bool hayOrden = sujetos.Any(s => OrdenCurricular(s) != int.MaxValue);

            // 1. Orden curricular oficial (Fase 2377, Fase 2381)
            yield return sujetos.OrderBy(OrdenCurricular).ThenBy(s => s.IdAsignatura).ToList();

            // 2. Orden original en individuo
            yield return sujetos.ToList();

            // Permutaciones sólo permitidas si no existe orden curricular explícito que proteger (Fase 2381)
            if (!hayOrden)
            {
                yield return sujetos.OrderByDescending(s => s.N).ThenBy(s => s.IdAsignatura).ToList();
                yield return sujetos.OrderBy(s => s.N).ThenBy(s => s.IdAsignatura).ToList();

                if (sujetos.Count <= 5)
                {
                    foreach (var p in Permutaciones(sujetos))
                    {
                        yield return p;
                    }
                }
            }
        }

        private static IEnumerable<List<T>> Permutaciones<T>(List<T> lista)
        {
            if (lista.Count <= 1)
            {
                yield return lista;
                yield break;
            }

            for (int i = 0; i < lista.Count; i++)
            {
                var elemento = lista[i];
                var restantes = new List<T>(lista);
                restantes.RemoveAt(i);
                foreach (var sub in Permutaciones(restantes))
                {
                    var res = new List<T> { elemento };
                    res.AddRange(sub);
                    yield return res;
                }
            }
        }

        private static bool AsignarCadena(
            AgIndividuo clon,
            List<SujetoCarril> todosSujetos,
            List<SujetoCarril> orden,
            List<DateTime> libres,
            out string motivoFallo)
        {
            motivoFallo = string.Empty;
            var usadas = new HashSet<DateTime>();
            var rangosFijos = todosSujetos
                .Where(s => s.EsFijo && s.IniFijo.HasValue && s.FinFijo.HasValue)
                .Select(s => new { Ini = s.IniFijo.Value.Date, Fin = s.FinFijo.Value.Date })
                .ToList();

            DateTime cursor = libres[0].Date;

            foreach (var sujeto in orden)
            {
                if (sujeto.EsFijo)
                {
                    if (sujeto.FinFijo.HasValue && sujeto.FinFijo.Value.Date >= cursor)
                    {
                        cursor = sujeto.FinFijo.Value.Date.AddDays(1);
                    }
                    continue;
                }

                DateTime cursorS = cursor;
                bool exitoSujeto = true;
                var fechasDelSujeto = new List<DateTime>();

                foreach (int idx in sujeto.Indices.OrderBy(i => clon.Genes[i].DiaSemana))
                {
                    if (idx < 0 || idx >= clon.Genes.Count)
                    {
                        motivoFallo = "SIN_ORDEN";
                        return false;
                    }
                    var g = clon.Genes[idx];
                    int n = (g.FechasGen != null && g.FechasGen.Count > 0)
                        ? g.FechasGen.Count
                        : (int)Math.Max(1, Math.Round(g.HorasGen / 4.0));

                    var fechasGen = new List<DateTime>();
                    foreach (var f in libres.Where(d => d.Date >= cursorS))
                    {
                        if (AgInventarioRanuras.Iso(f) != g.DiaSemana) continue;
                        if (usadas.Contains(f.Date)) continue;

                        bool solapaFijo = false;
                        for (int rf = 0; rf < rangosFijos.Count; rf++)
                        {
                            if (f.Date >= rangosFijos[rf].Ini && f.Date <= rangosFijos[rf].Fin)
                            {
                                solapaFijo = true;
                                break;
                            }
                        }
                        if (solapaFijo) continue;

                        fechasGen.Add(f.Date);
                        if (fechasGen.Count == n) break;
                    }

                    if (fechasGen.Count < n)
                    {
                        motivoFallo = "SIN_FECHAS";
                        exitoSujeto = false;
                        break;
                    }

                    foreach (var f in fechasGen)
                    {
                        usadas.Add(f);
                        fechasDelSujeto.Add(f);
                    }

                    g.FechasGen = fechasGen;
                    g.FechaInicioEstimada = fechasGen.Min();
                    g.FechaFinEstimada = fechasGen.Max();
                    cursorS = fechasGen[fechasGen.Count - 1].AddDays(1);
                }

                if (!exitoSujeto) return false;
                if (fechasDelSujeto.Count > 0)
                {
                    cursor = fechasDelSujeto.Max().AddDays(1);
                }
            }

            return true;
        }
    }
}
