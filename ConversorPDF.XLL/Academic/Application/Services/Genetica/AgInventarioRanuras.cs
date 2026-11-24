using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Modelo de inventario de ranuras libres, válidas, ocupadas y liberables por cohorte y aula para la reparación R5 (Fase 2324).
    /// </summary>
    internal sealed class AgInventarioRanuras
    {
        public int IdAula { get; set; }
        public int M { get; set; }
        public IReadOnlyList<int> DiasVentana { get; set; }
        public IReadOnlyList<DateTime> FechasValidas { get; set; }
        public IReadOnlyList<DateTime> LibresReales { get; set; }
        public IReadOnlyList<DateTime> Liberables { get; set; }
        public IReadOnlyDictionary<DateTime, int> Ocupadas { get; set; }
        public int ColaMuerta { get; set; }
        public int HuecosIntermedios { get; set; }
        public double BalancePorDia { get; set; }

        /// <summary>Inventario por carril (aula + patrón) de la cohorte (Fase 2366).</summary>
        public IReadOnlyList<AgInventarioCarril> Carriles { get; set; } = new List<AgInventarioCarril>();

        public double Ocupacion => (FechasValidas != null && FechasValidas.Count > 0 && Ocupadas != null)
            ? (double)Ocupadas.Count / FechasValidas.Count
            : 0.0;

        /// <summary>Día ISO (domingo = 7) (Fase 2366).</summary>
        internal static int Iso(DateTime d)
        {
            return d.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)d.DayOfWeek;
        }

        /// <summary>Fechas de sesión de un gen, o proyección semanal si no las tiene (Fase 2366).</summary>
        internal static IEnumerable<DateTime> FechasDeGen(AgGen g)
        {
            if (g == null) return Array.Empty<DateTime>();
            if (g.FechasGen != null && g.FechasGen.Count > 0) return g.FechasGen.Select(d => d.Date);
            if (!g.FechaInicioEstimada.HasValue || !g.FechaFinEstimada.HasValue) return Array.Empty<DateTime>();
            return GenerarFechasRango(g.FechaInicioEstimada.Value, g.FechaFinEstimada.Value, g.DiaSemana, null);
        }

        public static AgInventarioRanuras Construir(AgIndividuo ind, AgContexto ctx, string cohorte)
        {
            if (ind == null || ctx == null)
            {
                return new AgInventarioRanuras
                {
                    DiasVentana = new List<int>(),
                    FechasValidas = new List<DateTime>(),
                    LibresReales = new List<DateTime>(),
                    Liberables = new List<DateTime>(),
                    Ocupadas = new Dictionary<DateTime, int>()
                };
            }

            var genesCohorte = (ind.Genes != null)
                ? ind.Genes.Where(g => g.EsVacante == 0 && string.Equals(AgCalendarioConflictos.Cohorte(g), cohorte ?? string.Empty, StringComparison.Ordinal)).ToList()
                : new List<AgGen>();

            var idsAulas = genesCohorte.Select(g => g.IdAula).Where(id => id > 0).Distinct().ToList();
            if (idsAulas.Count == 0 && ctx.Aulas != null && ctx.Aulas.Count > 0)
            {
                var primeraAula = ctx.Aulas.FirstOrDefault();
                if (primeraAula != null)
                {
                    idsAulas.Add(primeraAula.Id);
                }
            }
            int idAula = idsAulas.FirstOrDefault();

            // Determinar unión de días de ventana para las aulas de la cohorte (Fase 2340)
            var diasVentanaSet = new HashSet<int>();
            foreach (var aId in idsAulas)
            {
                var aObj = (ctx.AulasPorId != null && ctx.AulasPorId.TryGetValue(aId, out var a))
                    ? a
                    : (ctx.Aulas != null ? ctx.Aulas.FirstOrDefault(au => au.Id == aId) : null);
                if (aObj != null && !string.IsNullOrWhiteSpace(aObj.DiasVentana))
                {
                    foreach (var d in AgDiasVentana.Dias(aObj.DiasVentana)) diasVentanaSet.Add(d);
                }
            }

            if (diasVentanaSet.Count == 0)
            {
                var aula = (ctx.AulasPorId != null && ctx.AulasPorId.TryGetValue(idAula, out var a))
                    ? a
                    : (ctx.Aulas != null ? ctx.Aulas.FirstOrDefault(au => au.Id == idAula) : null);
                if (aula != null && !string.IsNullOrWhiteSpace(aula.DiasVentana))
                {
                    foreach (var d in AgDiasVentana.Dias(aula.DiasVentana)) diasVentanaSet.Add(d);
                }
                else
                {
                    foreach (var d in new[] { 1, 2, 3, 4, 5 }) diasVentanaSet.Add(d);
                }
            }
            var diasVentana = diasVentanaSet.OrderBy(d => d).ToList();

            int m = AgValidadorReglas.MaxCarrilesCohorte(genesCohorte, ctx);
            if (m <= 0) m = 2;

            var fechasValidas = new List<DateTime>();
            if (ctx.Ventana != null && ctx.Ventana.FechaInicio <= ctx.Ventana.FechaFin)
            {
                for (var d = ctx.Ventana.FechaInicio.Date; d <= ctx.Ventana.FechaFin.Date; d = d.AddDays(1))
                {
                    int diaIso = (d.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)d.DayOfWeek;
                    if (diasVentana.Contains(diaIso) && !ctx.EsFestivo(d))
                    {
                        fechasValidas.Add(d);
                    }
                }
            }

            var ocupadasDict = new Dictionary<DateTime, int>();
            foreach (var g in genesCohorte)
            {
                var fechasGen = (g.FechasGen != null && g.FechasGen.Count > 0)
                    ? g.FechasGen
                    : (g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue
                        ? GenerarFechasRango(g.FechaInicioEstimada.Value, g.FechaFinEstimada.Value, g.DiaSemana, ctx)
                        : new List<DateTime>());

                foreach (var f in fechasGen.Select(x => x.Date))
                {
                    if (ocupadasDict.TryGetValue(f, out var count))
                    {
                        ocupadasDict[f] = count + 1;
                    }
                    else
                    {
                        ocupadasDict[f] = 1;
                    }
                }
            }

            if (ctx.SesionesConservadas != null)
            {
                foreach (var sc in ctx.SesionesConservadas)
                {
                    if (idsAulas.Contains(sc.Clave.IdAula))
                    {
                        var f = sc.Fecha.Date;
                        if (ocupadasDict.TryGetValue(f, out var count))
                        {
                            ocupadasDict[f] = count + 1;
                        }
                        else
                        {
                            ocupadasDict[f] = 1;
                        }
                    }
                }
            }

            var libresReales = fechasValidas
                .Where(f => !ocupadasDict.ContainsKey(f))
                .ToList();

            int huecos = 0;
            int cola = 0;
            if (ocupadasDict.Count > 0 && fechasValidas.Count > 0)
            {
                DateTime minOcupada = ocupadasDict.Keys.Min();
                DateTime maxOcupada = ocupadasDict.Keys.Max();

                foreach (var f in fechasValidas)
                {
                    bool estaOcupada = ocupadasDict.ContainsKey(f);
                    if (!estaOcupada)
                    {
                        if (f > maxOcupada)
                        {
                            cola++;
                        }
                        else if (f > minOcupada && f < maxOcupada)
                        {
                            huecos++;
                        }
                    }
                }
            }

            var liberables = new List<DateTime>();
            foreach (var g in genesCohorte)
            {
                if (g.EsFijo == 1) continue;

                string origen = null;
                if (ctx.AulaAsignaturaPorAulaAsig != null && ctx.AulaAsignaturaPorAulaAsig.TryGetValue((g.IdAula, g.IdAsignatura), out var aa))
                {
                    origen = aa.OrigenPatron;
                }
                else if (ctx.AulaAsignaturasPorId != null && ctx.AulaAsignaturasPorId.TryGetValue(g.IdAsignatura, out var aa2))
                {
                    origen = aa2.OrigenPatron;
                }

                if (string.Equals(origen, "LEGADO", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(origen, "FIJO", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (g.FechasGen != null)
                {
                    foreach (var f in g.FechasGen.Select(x => x.Date))
                    {
                        if (!liberables.Contains(f))
                        {
                            liberables.Add(f);
                        }
                    }
                }
            }
            liberables.Sort();

            double balance = 0.0;
            if (diasVentana.Count > 0 && ocupadasDict.Count > 0)
            {
                int totalSesiones = ocupadasDict.Values.Sum();
                double idealPorDia = (double)totalSesiones / diasVentana.Count;
                double sumDesviacion = 0.0;
                foreach (var dia in diasVentana)
                {
                    int enDia = ocupadasDict
                        .Where(kvp => ((kvp.Key.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)kvp.Key.DayOfWeek) == dia)
                        .Sum(kvp => kvp.Value);
                    sumDesviacion += Math.Abs(enDia - idealPorDia);
                }
                balance = sumDesviacion;
            }

            // Fase 2366: inventario por carril (aula + patrón) con fechas válidas/ocupadas/libres.
            var carriles = new List<AgInventarioCarril>();
            foreach (var grp in genesCohorte.GroupBy(g => AgCarriles.ClaveCarril(g), StringComparer.Ordinal))
            {
                var muestra = grp.FirstOrDefault();
                if (muestra == null) continue;
                var diasCarril = AgDiasVentana.Dias(muestra.DiasPatron);
                if (diasCarril.Count == 0)
                {
                    diasCarril = grp.Select(g => g.DiaSemana).Where(d => d >= 1 && d <= 7).Distinct().OrderBy(d => d).ToList();
                }
                if (diasCarril.Count == 0) continue;

                var aulaCarril = (ctx.AulasPorId != null && ctx.AulasPorId.TryGetValue(muestra.IdAula, out var aCarril))
                    ? aCarril
                    : (ctx.Aulas != null ? ctx.Aulas.FirstOrDefault(au => au.Id == muestra.IdAula) : null);
                var diasAula = aulaCarril != null ? AgDiasVentana.Dias(aulaCarril.DiasVentana) : new List<int>();

                var ocupadasCarril = new HashSet<DateTime>(grp.SelectMany(FechasDeGen));
                var validasCarril = fechasValidas
                    .Where(f => diasCarril.Contains(Iso(f)) && (diasAula.Count == 0 || diasAula.Contains(Iso(f))))
                    .ToList();

                carriles.Add(new AgInventarioCarril
                {
                    Clave = grp.Key,
                    IdAula = muestra.IdAula,
                    DiasPatron = muestra.DiasPatron ?? string.Empty,
                    Dias = diasCarril,
                    FechasValidas = validasCarril,
                    FechasOcupadas = ocupadasCarril.OrderBy(d => d).ToList(),
                    FechasLibres = validasCarril.Where(f => !ocupadasCarril.Contains(f)).ToList()
                });
            }

            return new AgInventarioRanuras
            {
                IdAula = idAula,
                M = m,
                DiasVentana = diasVentana,
                FechasValidas = fechasValidas,
                LibresReales = libresReales,
                Liberables = liberables,
                Ocupadas = ocupadasDict,
                ColaMuerta = cola,
                HuecosIntermedios = huecos,
                BalancePorDia = balance,
                Carriles = carriles
            };
        }

        private static List<DateTime> GenerarFechasRango(DateTime ini, DateTime fin, int diaSemana, AgContexto ctx)
        {
            var res = new List<DateTime>();
            for (var d = ini.Date; d <= fin.Date; d = d.AddDays(1))
            {
                int diaIso = (d.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)d.DayOfWeek;
                if (diaIso == diaSemana && (ctx == null || !ctx.EsFestivo(d)))
                {
                    res.Add(d);
                }
            }
            return res;
        }
    }
}
