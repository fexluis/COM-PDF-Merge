using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Partición disjunta de la ventana semanal en bandas de días contiguos y resolución determinista
    /// de la asignación asignatura → banda (Fase 1224).
    /// Objetivo lexicográfico: (infactibles, patrones violados, makespan, desperdicio, -nº bandas).
    /// Internal para prevenir registro espurio de UDFs en Excel-DNA 1.8.0.
    /// </summary>
    internal static class AgBandasDias
    {
        /// <summary>Tope de combinaciones para búsqueda exhaustiva (k^n). Por encima se usa Best-Fit Decreasing.</summary>
        public const int LimiteExhaustivo = 4096;

        /// <summary>Resultado puro de la resolución: bandas, asignación y métricas del plan.</summary>
        internal sealed class Asignacion
        {
            public IReadOnlyList<IReadOnlyList<int>> Bandas { get; }
            public IReadOnlyList<int> BandaPorIndice { get; }
            public IReadOnlyList<int> SlotPorBanda { get; }
            public IReadOnlyList<int> SemanasPorBanda { get; }
            public IReadOnlyList<int> ExcesoPorBanda { get; }
            public int Makespan { get; }
            public double Ocupacion { get; }
            public int ViolacionesPatron { get; }

            public Asignacion(
                IReadOnlyList<IReadOnlyList<int>> bandas,
                int[] bandaPorIndice,
                int[] slotPorBanda,
                int[] semanasPorBanda,
                int[] excesoPorBanda,
                int makespan,
                double ocupacion,
                int violacionesPatron)
            {
                Bandas = bandas ?? new List<IReadOnlyList<int>>();
                BandaPorIndice = bandaPorIndice ?? new int[0];
                SlotPorBanda = slotPorBanda ?? new int[0];
                SemanasPorBanda = semanasPorBanda ?? new int[0];
                ExcesoPorBanda = excesoPorBanda ?? new int[0];
                Makespan = makespan;
                Ocupacion = ocupacion;
                ViolacionesPatron = violacionesPatron;
            }

            public bool Factible => ExcesoPorBanda.All(e => e <= 0);
        }

        /// <summary>
        /// Particiones contiguas, disjuntas y exhaustivas de la ventana: 2^(n-1) formas de cortar.
        /// VIE-SAB (2 días) → [{5},{6}] y [{5,6}].
        /// </summary>
        public static IReadOnlyList<IReadOnlyList<IReadOnlyList<int>>> Particiones(IReadOnlyList<int> ventana)
        {
            var resultado = new List<IReadOnlyList<IReadOnlyList<int>>>();
            var dias = (ventana ?? new List<int>()).Distinct().OrderBy(d => d).ToList();
            if (dias.Count == 0) return resultado;

            int cortes = dias.Count - 1;
            int total = 1 << cortes;
            for (int mascara = 0; mascara < total; mascara++)
            {
                var bandas = new List<IReadOnlyList<int>>();
                var actual = new List<int>();
                for (int i = 0; i < dias.Count; i++)
                {
                    actual.Add(dias[i]);
                    bool corta = i < cortes && ((mascara >> i) & 1) == 1;
                    if (corta || i == dias.Count - 1)
                    {
                        bandas.Add(actual);
                        actual = new List<int>();
                    }
                }
                resultado.Add(bandas);
            }
            return resultado;
        }

        /// <summary>Ranuras cuya ocurrencia cae en alguno de los días de la banda (conserva orden canónico).</summary>
        public static IReadOnlyList<AgRanura> RanurasDe(IReadOnlyList<AgRanura> ranuras, IReadOnlyList<int> banda)
        {
            if (ranuras == null || ranuras.Count == 0 || banda == null || banda.Count == 0)
            {
                return new List<AgRanura>();
            }
            var dias = new HashSet<int>(banda);
            return ranuras.Where(r => dias.Contains(r.Dia)).ToList();
        }

        /// <summary>
        /// [Obsolete documental: legado F1226; el camino canónico es carriles + oleadas (F1279)].
        /// Resuelve la mejor asignación asignatura → banda para la ventana y la demanda dadas.
        /// </summary>
        public static Asignacion Resolver(
            IReadOnlyList<int> ventana,
            IReadOnlyList<AgRanura> ranuras,
            IReadOnlyList<int> demanda,
            IReadOnlyList<string> diasPatron,
            int semanasDisponibles,
            IReadOnlyDictionary<int, int> capacidadDia = null)   // Fase 1267
        {
            int n = demanda == null ? 0 : demanda.Count;
            int W = Math.Max(1, semanasDisponibles);

            if (ranuras == null || ranuras.Count == 0)
            {
                ranuras = (ventana ?? new List<int>()).Distinct().OrderBy(d => d)
                    .Select(d => new AgRanura(d, AgFranjaJornada.NOCHE, 4.0)).ToList();
            }

            if (n == 0)
            {
                return new Asignacion(new List<IReadOnlyList<int>>(), new int[0], new int[0], new int[0], new int[0], 0, 0.0, 0);
            }

            int demandaTotal = demanda.Sum();
            Asignacion mejor = null;
            Costo mejorCosto = null;

            var particiones = Particiones(ventana)
                .OrderByDescending(p => p.Count)                                        // F1226: más bandas = más paralelismo
                .ThenBy(p => string.Join("|", p.Select(b => string.Join(".", b))), StringComparer.Ordinal)
                .ToList();

            // Fase 1251: dos pasadas — exigirPatrones=true; si no hay solución factible, relajar a false
            bool tienePatrones = diasPatron != null && diasPatron.Any(p => !string.IsNullOrWhiteSpace(p));
            bool[] pasadas = tienePatrones ? new[] { true, false } : new[] { false };

            foreach (bool exigirPatrones in pasadas)
            {
                foreach (var particion in particiones)
                {
                    int k = particion.Count;
                    if (k == 0) continue;

                    var slots = new int[k];
                    bool viable = true;
                    for (int b = 0; b < k && viable; b++)
                    {
                        slots[b] = RanurasDe(ranuras, particion[b]).Count;
                        if (slots[b] <= 0) viable = false;
                    }
                    if (!viable) continue;

                    long combinaciones = 1;
                    for (int i = 0; i < n && combinaciones <= LimiteExhaustivo; i++) combinaciones *= k;

                    if (combinaciones <= LimiteExhaustivo)
                    {
                        var asignacion = new int[n];
                        for (long c = 0; c < combinaciones; c++)
                        {
                            long v = c;
                            for (int i = 0; i < n; i++) { asignacion[i] = (int)(v % k); v /= k; }
                            Evaluar(particion, slots, asignacion, demanda, diasPatron, W, demandaTotal, ref mejor, ref mejorCosto, exigirPatrones, capacidadDia);
                        }
                    }
                    else
                    {
                        var asignacion = BestFit(slots, demanda, W);
                        Evaluar(particion, slots, asignacion, demanda, diasPatron, W, demandaTotal, ref mejor, ref mejorCosto, exigirPatrones, capacidadDia);
                    }
                }

                if (mejor != null) break;
            }

            if (mejor == null)
            {
                // Fallback determinista: banda única con toda la ventana (nunca oculta infactibilidad)
                var dias = (ventana ?? new List<int>()).Distinct().OrderBy(d => d).ToList();
                if (dias.Count == 0)
                {
                    return new Asignacion(new List<IReadOnlyList<int>>(), new int[n], new int[0],
                                          new int[0], new int[0], 0, 0.0, 0);   // Fase 1257
                }
                var unica = new List<IReadOnlyList<int>> { dias };
                var slots = new[] { Math.Max(1, ranuras.Count) };
                Evaluar(unica, slots, new int[n], demanda, diasPatron, W, demandaTotal, ref mejor, ref mejorCosto, false, capacidadDia);
            }

            return mejor;
        }

        private static bool ParticionAlojaPatrones(
            IReadOnlyList<IReadOnlyList<int>> particion, int[] asignacion, IReadOnlyList<string> diasPatron)
        {
            if (diasPatron == null) return true;
            for (int i = 0; i < asignacion.Length && i < diasPatron.Count; i++)
            {
                var dias = AgCarriles.Ventana(diasPatron[i]);
                if (dias.Count == 0) continue;
                bool hayBandaValida = false, actualValida = true;
                for (int b = 0; b < particion.Count; b++)
                {
                    bool ok = dias.All(d => particion[b].Contains(d));
                    if (ok) hayBandaValida = true;
                    if (b == asignacion[i]) actualValida = ok;
                }
                if (hayBandaValida && !actualValida) return false;
            }
            return true;
        }

        private static void Evaluar(
            IReadOnlyList<IReadOnlyList<int>> particion,
            int[] slots,
            int[] asignacion,
            IReadOnlyList<int> demanda,
            IReadOnlyList<string> diasPatron,
            int W,
            int demandaTotal,
            ref Asignacion mejor,
            ref Costo mejorCosto,
            bool exigirPatrones,
            IReadOnlyDictionary<int, int> capacidadDia)
        {
            // Fase 1251: antes de calcular coste, descartar asignación si viola alojamiento de patrones factibles
            if (exigirPatrones && !ParticionAlojaPatrones(particion, asignacion, diasPatron)) return;

            int k = particion.Count;
            var demandaBanda = new int[k];
            for (int i = 0; i < asignacion.Length; i++) demandaBanda[asignacion[i]] += demanda[i];

            var semanas = new int[k];
            var exceso = new int[k];
            int makespan = 0, desperdicio = 0, excesoTotal = 0;

            for (int b = 0; b < k; b++)
            {
                int capacidad = slots[b] * W;
                if (capacidadDia != null)
                {
                    capacidad = 0;
                    foreach (var d in particion[b])
                    {
                        if (capacidadDia.TryGetValue(d, out var cd)) capacidad += cd;
                    }
                }

                exceso[b] = Math.Max(0, demandaBanda[b] - capacidad);
                excesoTotal += exceso[b];
                int divisor = capacidadDia != null && W > 0
                    ? Math.Max(1, (int)Math.Round((double)capacidad / W))
                    : Math.Max(1, slots[b]);
                int necesarias = (int)Math.Ceiling(demandaBanda[b] / (double)divisor);
                semanas[b] = Math.Max(1, Math.Min(W, necesarias));
                makespan = Math.Max(makespan, semanas[b]);
                desperdicio += Math.Max(0, semanas[b] * slots[b] - demandaBanda[b]);
            }

            // Fase 1265: pasada estricta exige factibilidad (excesoTotal == 0)
            if (exigirPatrones && excesoTotal > 0) return;

            // Fase 1231 / Fase 1226: un patrón sólo se conserva si sus días viven dentro de la banda elegida
            int violaciones = 0;
            for (int i = 0; i < asignacion.Length; i++)
            {
                string patron = diasPatron != null && i < diasPatron.Count ? diasPatron[i] : null;
                if (string.IsNullOrWhiteSpace(patron)) continue;
                var dias = AgCarriles.Ventana(patron);
                if (dias.Count == 0) continue;
                var diasBanda = particion[asignacion[i]];
                if (dias.Any(d => !diasBanda.Contains(d))) violaciones++;
            }

            var costo = new Costo(excesoTotal, violaciones, makespan, desperdicio, -k);
            if (mejor == null || costo.EsMenorQue(mejorCosto))
            {
                mejorCosto = costo;
                // Fase 1273: capacidad efectiva total para cálculo de ocupación (descuenta festivos)
                int slotsTotales = capacidadDia != null
                    ? Math.Max(1, capacidadDia.Values.Sum())
                    : Math.Max(1, slots.Sum() * W);
                mejor = new Asignacion(
                    particion,
                    (int[])asignacion.Clone(),
                    slots,
                    semanas,
                    exceso,
                    makespan,
                    Math.Min(1.0, demandaTotal / (double)slotsTotales),
                    violaciones);
            }
        }

        private static int[] BestFit(int[] slots, IReadOnlyList<int> demanda, int W)
        {
            int k = slots.Length;
            var carga = new int[k];
            var asignacion = new int[demanda.Count];
            var orden = Enumerable.Range(0, demanda.Count)
                .OrderByDescending(i => demanda[i])
                .ThenBy(i => i)
                .ToList();

            foreach (var i in orden)
            {
                int mejorBanda = 0;
                double mejorClave = double.MaxValue;
                for (int b = 0; b < k; b++)
                {
                    int nueva = carga[b] + demanda[i];
                    int capacidad = slots[b] * W;
                    double sobre = nueva > capacidad ? nueva - capacidad : 0;
                    double clave = Math.Ceiling(nueva / (double)Math.Max(1, slots[b])) + sobre * 10.0;
                    if (clave < mejorClave) { mejorClave = clave; mejorBanda = b; }
                }
                asignacion[i] = mejorBanda;
                carga[mejorBanda] += demanda[i];
            }
            return asignacion;
        }

        private sealed class Costo
        {
            private readonly int _exceso;
            private readonly int _violaciones;
            private readonly int _makespan;
            private readonly int _desperdicio;
            private readonly int _negBandas;

            public Costo(int exceso, int violaciones, int makespan, int desperdicio, int negBandas)
            {
                _exceso = exceso;
                _violaciones = violaciones;
                _makespan = makespan;
                _desperdicio = desperdicio;
                _negBandas = negBandas;
            }

            public bool EsMenorQue(Costo otro)
            {
                if (otro == null) return true;
                if (_exceso != otro._exceso) return _exceso < otro._exceso;
                if (_violaciones != otro._violaciones) return _violaciones < otro._violaciones;
                if (_makespan != otro._makespan) return _makespan < otro._makespan;
                if (_desperdicio != otro._desperdicio) return _desperdicio < otro._desperdicio;
                return _negBandas < otro._negBandas;
            }
        }
    }
}
