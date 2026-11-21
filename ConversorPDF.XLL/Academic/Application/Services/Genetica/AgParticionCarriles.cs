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
    /// Decisión de selección entre la partición canónica de referencia y una alternativa (Fase 1581, Fase 2220, Fase 2228, Fase 2229, Fase 2232).
    /// CandidatasEvaluadas: catálogo → opciones evaluadas; enumeración → particiones contiguas evaluadas.
    /// OpcionesEvaluadas: opciones de catálogo evaluadas (0 si no aplica).
    /// </summary>
    internal sealed class AgSeleccionParticion
    {
        public bool Cambiada { get; }
        public AgEvaluacionParticion Referencia { get; }
        public AgEvaluacionParticion Elegida { get; }
        public int CandidatasEvaluadas { get; }
        public int ParticionesDescartadas { get; }
        public string Origen { get; }
        public string OpcionTexto { get; }
        public string ParticionElegidaTexto { get; }
        public int OpcionesEvaluadas { get; }

        public AgSeleccionParticion(
            bool cambiada,
            AgEvaluacionParticion referencia,
            AgEvaluacionParticion elegida,
            int candidatasEvaluadas,
            int particionesDescartadas,
            string origen = "enumeracion",
            string opcionTexto = "",
            int opcionesEvaluadas = 0,
            string particionElegidaTexto = null)
        {
            Cambiada = cambiada;
            Referencia = referencia;
            Elegida = elegida;
            CandidatasEvaluadas = candidatasEvaluadas;
            ParticionesDescartadas = particionesDescartadas;
            Origen = origen ?? "enumeracion";
            OpcionTexto = opcionTexto ?? string.Empty;
            OpcionesEvaluadas = opcionesEvaluadas;
            ParticionElegidaTexto = particionElegidaTexto ?? (elegida != null ? elegida.Clave : string.Empty);
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
            Func<AgEvaluacionParticion, bool> esAdmisible = null,
            IReadOnlyList<AgOpcionParticion> opciones = null)
        {
            if (evaluar == null) throw new ArgumentNullException(nameof(evaluar));

            // F1582: Rechazar configuraciones de carril o ranura inválidas antes de invocar el delegado
            if (ranuras == null || ranuras.Count == 0)
                throw new ArgumentException("Ranuras vacías.", nameof(ranuras));

            if (bandasCanonicas == null || bandasCanonicas.Count == 0 ||
                bandasCanonicas.Any(b => b == null || b.Count == 0))
                throw new ArgumentException("Carriles inválidos.", nameof(bandasCanonicas));

            var referencia = evaluar(bandasCanonicas);

            bool catalogoEvaluado = false;
            string textoOpcionIntentada = string.Empty;
            int evaluadasOpciones = 0;

            // Fase 2180 / Fase 2187 / Fase 2188 / Fase 2223 / Fase 2224 / Fase 2225: Evaluación focalizada de catálogo con gate de mejora y fallback
            if (opciones != null && opciones.Count > 0)
            {
                var evaluadasCat = new List<(AgEvaluacionParticion Ev, string Texto)>();
                foreach (var op in opciones.Take(AgOpcionesParticion.MaxOpcionesDecision))
                {
                    if (op == null || !ConservaRanuras(ranuras, op.Carriles)) continue;
                    var ev = evaluar(op.Carriles);
                    evaluadasOpciones++;
                    if (esAdmisible != null && !esAdmisible(ev)) continue;
                    if (!AlojaPatrones(op.Carriles, patrones)) continue;
                    evaluadasCat.Add((ev, op.CarrilesTexto));
                }

                catalogoEvaluado = evaluadasOpciones > 0;
                textoOpcionIntentada = evaluadasCat.Count > 0
                    ? evaluadasCat[0].Texto
                    : (opciones.Count > 0 ? (opciones[0]?.CarrilesTexto ?? string.Empty) : string.Empty);

                var elegidaOpcion = evaluadasCat
                    .Where(x => x.Ev.PatronesNoAlojables <= referencia.PatronesNoAlojables)
                    .Where(x => x.Ev.Resultado.Infactibles.Count < referencia.Resultado.Infactibles.Count
                             || (x.Ev.Resultado.Infactibles.Count == referencia.Resultado.Infactibles.Count
                                 && x.Ev.Makespan < referencia.Makespan))
                    .OrderBy(x => x.Ev.Resultado.Infactibles.Count)
                    .ThenBy(x => x.Ev.Makespan)
                    .ThenBy(x => x.Texto, StringComparer.Ordinal)
                    .FirstOrDefault();

                if (elegidaOpcion.Ev != null)
                {
                    // FASE 2232: CandidatasEvaluadas en catálogo exitoso reporta evaluadasOpciones
                    return new AgSeleccionParticion(true, referencia, elegidaOpcion.Ev, evaluadasOpciones, 0,
                        "catalogo", elegidaOpcion.Texto, evaluadasOpciones, elegidaOpcion.Ev.Clave);
                }

                // FASE 2223: sin mejora en catálogo, continuar a la enumeración contigua de Particiones() (no retornar aquí)
            }

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

            // F1583 / F1736 / F2230: No detenerse si PatronesNoAlojables == 0 a menos que no se haya evaluado catálogo.
            // Si catalogoEvaluado es true, continuar a la enumeración contigua para completar el fallback.
            if (!hayPatronValido && referencia.Resultado.Infactibles.Count == 0 && !catalogoEvaluado)
            {
                return new AgSeleccionParticion(false, referencia, referencia, 0, 0,
                    "enumeracion", string.Empty, 0, referencia.Clave);
            }

            int kActual = bandasCanonicas.Count;
            int kMin = kActual >= 2 ? 2 : 1;

            var particiones = AgBandasDias.Particiones(ventana);
            var vistas = new HashSet<string>(StringComparer.Ordinal);
            var candidatas = new List<AgEvaluacionParticion>();
            int descartadas = 0;
            int evaluadas = 0; // F1584 / F2229: Contador de llamadas reales al evaluador en enumeración contigua

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

                // (d) Reducción estricta de infactibles o acortamiento de makespan con igual número de infactibles (F1736)
                bool menorInfactibles = ev.Resultado.Infactibles.Count < referencia.Resultado.Infactibles.Count;
                bool mejorMakespan = ev.Resultado.Infactibles.Count == referencia.Resultado.Infactibles.Count && ev.Makespan < referencia.Makespan;
                if (!menorInfactibles && !mejorMakespan)
                {
                    descartadas++;
                    continue;
                }

                candidatas.Add(ev);
            }

            if (candidatas.Count == 0)
            {
                return new AgSeleccionParticion(false, referencia, referencia, evaluadas, descartadas,
                    catalogoEvaluado ? "catalogo_fallback" : "enumeracion", textoOpcionIntentada, evaluadasOpciones, referencia.Clave);
            }

            var elegida = candidatas
                .OrderBy(ev => ev.Resultado.Infactibles.Count)
                .ThenBy(ev => ev.PatronesNoAlojables)
                .ThenBy(ev => ev.Makespan)
                .ThenByDescending(ev => ev.Bandas)
                .ThenBy(ev => ev.MaxRanurasPorBanda)
                .ThenBy(ev => ev.Clave, StringComparer.Ordinal)
                .First();

            return new AgSeleccionParticion(true, referencia, elegida, evaluadas, descartadas,
                catalogoEvaluado ? "catalogo_fallback" : "enumeracion", textoOpcionIntentada, evaluadasOpciones, elegida.Clave);
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

        private static bool AlojaPatrones(IReadOnlyList<IReadOnlyList<AgRanura>> bandas, IReadOnlyList<string> patrones)
        {
            if (patrones == null || patrones.Count == 0) return true;
            return ContarPatronesNoAlojables(bandas, patrones) == 0;
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

