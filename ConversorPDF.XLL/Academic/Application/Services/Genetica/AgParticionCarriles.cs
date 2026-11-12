using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Resultado de la evaluación temporal y estructural de una partición de bandas/carriles (Fase 1581).
    /// </summary>
    internal sealed class AgEvaluacionParticion
    {
        public AgPlanOleadasResultado Resultado { get; }
        public int PatronesNoAlojables { get; }
        public IReadOnlyCollection<int> PatronesAlojables { get; }
        public int Makespan { get; }
        public int Bandas { get; }
        public int MaxRanurasPorBanda { get; }
        public string Clave { get; }

        public AgEvaluacionParticion(
            AgPlanOleadasResultado resultado,
            int patronesNoAlojables,
            IReadOnlyCollection<int> patronesAlojables,
            int makespan,
            int bandas,
            int maxRanurasPorBanda,
            string clave)
        {
            Resultado = resultado;
            PatronesNoAlojables = patronesNoAlojables;
            PatronesAlojables = patronesAlojables ?? new List<int>();
            Makespan = makespan;
            Bandas = bandas;
            MaxRanurasPorBanda = maxRanurasPorBanda;
            Clave = clave ?? string.Empty;
        }
    }

    /// <summary>
    /// Decisión de selección entre la partición canónica de referencia y una alternativa (Fase 1581).
    /// </summary>
    internal sealed class AgSeleccionParticion
    {
        public bool Cambiada { get; }
        public AgEvaluacionParticion Referencia { get; }
        public AgEvaluacionParticion Elegida { get; }
        public int CandidatasEvaluadas { get; }
        public int ParticionesDescartadas { get; }

        public AgSeleccionParticion(
            bool cambiada,
            AgEvaluacionParticion referencia,
            AgEvaluacionParticion elegida,
            int candidatasEvaluadas,
            int particionesDescartadas)
        {
            Cambiada = cambiada;
            Referencia = referencia;
            Elegida = elegida;
            CandidatasEvaluadas = candidatasEvaluadas;
            ParticionesDescartadas = particionesDescartadas;
        }
    }

    /// <summary>
    /// Selector determinista de partición de carriles consciente del patrón sembrado (Fase 1581).
    /// </summary>
    internal static class AgParticionCarriles
    {
        public static AgSeleccionParticion Seleccionar(
            IReadOnlyList<int> ventana,
            IReadOnlyList<AgRanura> ranuras,
            IReadOnlyList<string> patrones,
            IReadOnlyList<IReadOnlyList<AgRanura>> bandasCanonicas,
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar,
            Func<AgEvaluacionParticion, bool> esAdmisible = null)
        {
            if (evaluar == null) throw new ArgumentNullException(nameof(evaluar));

            // F1582: Rechazar configuraciones de carril o ranura inválidas antes de invocar el delegado
            if (ranuras == null || ranuras.Count == 0)
                throw new ArgumentException("Ranuras vacías.", nameof(ranuras));

            if (bandasCanonicas == null || bandasCanonicas.Count == 0 ||
                bandasCanonicas.Any(b => b == null || b.Count == 0))
                throw new ArgumentException("Carriles inválidos.", nameof(bandasCanonicas));

            var referencia = evaluar(bandasCanonicas);

            bool hayPatronValido = false;
            if (patrones != null)
            {
                for (int i = 0; i < patrones.Count; i++)
                {
                    if (AgCarriles.Ventana(patrones[i]).Count > 0)
                    {
                        hayPatronValido = true;
                        break;
                    }
                }
            }

            // F1583: No detenerse si PatronesNoAlojables == 0; la partición alternativa puede aportar
            // capacidad requerida para resolver infactibles en asignaturas de alta demanda.
            if (!hayPatronValido || referencia.Resultado.Infactibles.Count == 0)
            {
                return new AgSeleccionParticion(false, referencia, referencia, 0, 0);
            }

            int kActual = bandasCanonicas.Count;
            int kMin = kActual >= 2 ? 2 : 1;

            var particiones = AgBandasDias.Particiones(ventana);
            var vistas = new HashSet<string>(StringComparer.Ordinal);
            var candidatas = new List<AgEvaluacionParticion>();
            int descartadas = 0;
            int evaluadas = 0; // F1584: Contador exacto de llamadas reales al evaluador excluyendo la referencia

            var idsRef = referencia.Resultado.Planes.Select(p => p.IdAsignatura).ToHashSet();
            var alojablesRef = new HashSet<int>(referencia.PatronesAlojables);

            foreach (var p in particiones)
            {
                if (p.Count < kMin || p.Count > kActual)
                {
                    descartadas++;
                    continue;
                }

                var bandas = p.Select(b => (IReadOnlyList<AgRanura>)AgBandasDias.RanurasDe(ranuras, b)).ToList();
                if (bandas.Any(b => b.Count == 0))
                {
                    descartadas++;
                    continue;
                }

                string firma = string.Join("|", bandas.Select(b => string.Join(";",
                    b.Select(r => $"{r.Dia}:{r.Franja}:{r.Horas:0.###}"))));
                if (!vistas.Add(firma))
                {
                    descartadas++;
                    continue;
                }

                int noAloj = ContarPatronesNoAlojables(bandas, patrones);
                // F1583: Permitir igualdad estructural; la mejora efectiva se exige después sobre infactibles
                if (noAloj > referencia.PatronesNoAlojables)
                {
                    descartadas++;
                    continue;
                }

                // (b) Conserva alojables previos
                var alojablesBandas = ObtenerIndicesAlojables(bandas, patrones);
                if (!alojablesRef.IsSubsetOf(alojablesBandas))
                {
                    descartadas++;
                    continue;
                }

                // Invariante de conservación de ranuras
                if (!ConservaRanuras(ranuras, bandas))
                {
                    descartadas++;
                    continue;
                }

                var ev = evaluar(bandas);
                evaluadas++; // F1584

                // F1586: Filtro de fidelidad exacta de patrón sembrado
                if (esAdmisible != null && !esAdmisible(ev))
                {
                    descartadas++;
                    continue;
                }

                // (c) No perder asignaturas planificadas
                var idsEv = ev.Resultado.Planes.Select(pl => pl.IdAsignatura).ToHashSet();
                if (!idsRef.IsSubsetOf(idsEv))
                {
                    descartadas++;
                    continue;
                }

                // (d) Reducción estricta de infactibles
                if (ev.Resultado.Infactibles.Count >= referencia.Resultado.Infactibles.Count)
                {
                    descartadas++;
                    continue;
                }

                candidatas.Add(ev);
            }

            if (candidatas.Count == 0)
            {
                return new AgSeleccionParticion(false, referencia, referencia, evaluadas, descartadas);
            }

            var elegida = candidatas
                .OrderBy(ev => ev.Resultado.Infactibles.Count)
                .ThenBy(ev => ev.PatronesNoAlojables)
                .ThenBy(ev => ev.Makespan)
                .ThenByDescending(ev => ev.Bandas)
                .ThenBy(ev => ev.MaxRanurasPorBanda)
                .ThenBy(ev => ev.Clave, StringComparer.Ordinal)
                .First();

            return new AgSeleccionParticion(true, referencia, elegida, evaluadas, descartadas);
        }

        public static int ContarPatronesNoAlojables(
            IReadOnlyList<IReadOnlyList<AgRanura>> bandas,
            IReadOnlyList<string> patrones)
        {
            if (patrones == null || patrones.Count == 0) return 0;
            int count = 0;
            for (int i = 0; i < patrones.Count; i++)
            {
                var dias = AgCarriles.Ventana(patrones[i]);
                if (dias.Count == 0) continue;

                if (!AlojaPatron(bandas, dias))
                {
                    count++;
                }
            }
            return count;
        }

        public static HashSet<int> ObtenerIndicesAlojables(
            IReadOnlyList<IReadOnlyList<AgRanura>> bandas,
            IReadOnlyList<string> patrones)
        {
            var res = new HashSet<int>();
            if (patrones == null) return res;
            for (int i = 0; i < patrones.Count; i++)
            {
                var dias = AgCarriles.Ventana(patrones[i]);
                if (dias.Count == 0) continue;

                if (AlojaPatron(bandas, dias))
                {
                    res.Add(i);
                }
            }
            return res;
        }

        private static bool AlojaPatron(IReadOnlyList<IReadOnlyList<AgRanura>> bandas, IReadOnlyList<int> diasPatron)
        {
            if (bandas == null) return false;
            for (int c = 0; c < bandas.Count; c++)
            {
                var diasBanda = bandas[c].Select(r => r.Dia).Distinct().ToHashSet();
                if (diasPatron.All(d => diasBanda.Contains(d)))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool ConservaRanuras(
            IReadOnlyList<AgRanura> origen,
            IReadOnlyList<IReadOnlyList<AgRanura>> bandas)
        {
            if (origen == null || bandas == null) return false;
            var flatten = bandas.SelectMany(b => b).ToList();
            if (flatten.Count != origen.Count) return false;

            var mapaOrigen = origen.GroupBy(r => (r.Dia, r.Franja, r.Horas))
                .ToDictionary(g => g.Key, g => g.Count());
            var mapaBandas = flatten.GroupBy(r => (r.Dia, r.Franja, r.Horas))
                .ToDictionary(g => g.Key, g => g.Count());

            if (mapaOrigen.Count != mapaBandas.Count) return false;
            foreach (var kv in mapaOrigen)
            {
                if (!mapaBandas.TryGetValue(kv.Key, out var c) || c != kv.Value) return false;
            }
            return true;
        }
    }
}

