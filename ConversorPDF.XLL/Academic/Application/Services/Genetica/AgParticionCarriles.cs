using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;

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
        public int BandasSingleton { get; }

        public AgEvaluacionParticion(
            AgPlanOleadasResultado resultado,
            int patronesNoAlojables,
            IReadOnlyCollection<int> patronesAlojables,
            int makespan,
            int bandas,
            int maxRanurasPorBanda,
            string clave,
            int bandasSingleton = -1)
        {
            Resultado = resultado;
            PatronesNoAlojables = patronesNoAlojables;
            PatronesAlojables = patronesAlojables ?? new List<int>();
            Makespan = makespan;
            Bandas = bandas;
            MaxRanurasPorBanda = maxRanurasPorBanda;
            Clave = clave ?? string.Empty;
            BandasSingleton = bandasSingleton >= 0
                ? bandasSingleton
                : (string.IsNullOrWhiteSpace(Clave) ? 0 : Clave.Split('|')
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .Count(p => {
                        var tokens = p.Split('.', ',')
                            .Where(x => !string.IsNullOrWhiteSpace(x) && x.Any(char.IsDigit))
                            .Distinct()
                            .ToList();
                        return tokens.Count == 1;
                    }));
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
        internal const bool PermitirEnumeracion = false; // Fase 2433: Política catálogo-only por defecto

        public static AgSeleccionParticion Seleccionar(
            IReadOnlyList<int> ventana,
            IReadOnlyList<AgRanura> ranuras,
            IReadOnlyList<string> patrones,
            IReadOnlyList<IReadOnlyList<AgRanura>> bandasCanonicas,
            Func<IReadOnlyList<IReadOnlyList<AgRanura>>, AgEvaluacionParticion> evaluar,
            Func<AgEvaluacionParticion, bool> esAdmisible = null,
            IReadOnlyList<AgOpcionParticion> opciones = null,
            bool permitirEnumeracion = true,
            string diasVentana = null)
        {
            if (evaluar == null) throw new ArgumentNullException(nameof(evaluar));

            // F1582: Rechazar configuraciones de carril o ranura inválidas antes de invocar el delegado
            if (ranuras == null || ranuras.Count == 0)
                throw new ArgumentException("Ranuras vacías.", nameof(ranuras));

            if (bandasCanonicas == null || bandasCanonicas.Count == 0 ||
                bandasCanonicas.Any(b => b == null || b.Count == 0))
                throw new ArgumentException("Carriles inválidos.", nameof(bandasCanonicas));

            var referencia = evaluar(bandasCanonicas);

            const bool PermitirSingleton = false; // Fase 2404: Prohibir bandas singleton en motor y catálogo si existe alternativa

            bool catalogoEvaluado = false;
            string textoOpcionIntentada = string.Empty;
            int evaluadasOpciones = 0;

            AgEvaluacionParticion mejorCatalogo = null;
            string textoMejorCatalogo = string.Empty;

            // Fase 2180 / Fase 2187 / Fase 2188 / Fase 2223 / Fase 2224 / Fase 2225 / Fase 2431: Evaluación de catálogo con comparación diferida
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

                // Fase 2404 / Fase 2412: Prohibir bandas singleton también en opciones de catálogo si existe alternativa
                var opcionesValidas = PermitirSingleton
                    ? evaluadasCat
                    : evaluadasCat.Where(x => x.Ev.BandasSingleton == 0).ToList();

                var elegidaOpcion = opcionesValidas
                    .Where(x => x.Ev.PatronesNoAlojables <= referencia.PatronesNoAlojables)
                    .Where(x => x.Ev.Resultado.Infactibles.Count <= referencia.Resultado.Infactibles.Count)
                    .Where(x => x.Ev.Resultado.Infactibles.Count < referencia.Resultado.Infactibles.Count
                             || (referencia.Resultado.Infactibles.Count == 0
                                 && (x.Ev.PatronesNoAlojables < referencia.PatronesNoAlojables
                                     || x.Ev.BandasSingleton < referencia.BandasSingleton
                                     || (x.Ev.BandasSingleton == referencia.BandasSingleton
                                         && (x.Ev.Makespan < referencia.Makespan
                                             || (x.Ev.Makespan == referencia.Makespan && x.Ev.Bandas < referencia.Bandas))))))
                    .OrderBy(x => x.Ev.Resultado.Infactibles.Count)
                    .ThenBy(x => x.Ev.BandasSingleton)
                    .ThenBy(x => x.Ev.Makespan)
                    .ThenBy(x => x.Ev.Bandas)
                    .FirstOrDefault();

                if (elegidaOpcion.Ev != null)
                {
                    mejorCatalogo = elegidaOpcion.Ev;
                    textoMejorCatalogo = elegidaOpcion.Texto;
                }

                // FASE 2431 / FASE 2433: Si no se permite enumeración, resolver de inmediato con catálogo o sin_catalogo
                if (!permitirEnumeracion)
                {
                    if (mejorCatalogo != null)
                    {
                        if (mejorCatalogo.BandasSingleton < referencia.BandasSingleton)
                        {
                            LoggerHelper.Debug(
                                $"[Oleadas] bandasSingleton={referencia.BandasSingleton}→{mejorCatalogo.BandasSingleton}; " +
                                $"motivo=CAMBIO_POR_SINGLETON; opcion='{textoMejorCatalogo}'");
                        }

                        return new AgSeleccionParticion(
                            true, referencia, mejorCatalogo,
                            evaluadasOpciones,
                            0,
                            "catalogo",
                            textoMejorCatalogo,
                            evaluadasOpciones,
                            mejorCatalogo.Clave);
                    }

                    // FASE 2443: Si la referencia canónica es válida pero no hay mejor catálogo, devolver "canonico", no "catalogo"
                    string etiquetaVentana = !string.IsNullOrWhiteSpace(diasVentana)
                        ? diasVentana.Trim()
                        : (ventana != null ? AgDiasVentana.AIso(string.Join(",", ventana)) : string.Empty);
                    bool puedeEscalar = !string.IsNullOrEmpty(etiquetaVentana) && AgEscaleraVentanas.PuedeEscalar(etiquetaVentana);
                    bool prohibirSingleton = ventana != null && ventana.Count == 5;
                    bool tienePatronFinSemana = patrones != null && patrones.Any(p =>
                        !string.IsNullOrWhiteSpace(p) &&
                        (p.IndexOf("6M", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         p.IndexOf("6T", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         string.Equals(p.Trim(), "JUE-DOM", StringComparison.OrdinalIgnoreCase)));

                    bool todosPatronesEnVentana = patrones == null || ventana == null || patrones.All(p => AgCarriles.Ventana(p).All(ventana.Contains));
                    bool sinCatalogo = puedeEscalar && (
                        (prohibirSingleton && referencia.BandasSingleton > 0) ||
                        tienePatronFinSemana ||
                        (todosPatronesEnVentana && referencia.PatronesNoAlojables > 0));

                    if (sinCatalogo)
                    {
                        return new AgSeleccionParticion(
                            false, referencia, referencia,
                            evaluadasOpciones,
                            0,
                            "sin_catalogo",
                            textoOpcionIntentada,
                            evaluadasOpciones,
                            referencia.Clave);
                    }

                    return new AgSeleccionParticion(
                        false, referencia, referencia,
                        evaluadasOpciones,
                        0,
                        "canonico",
                        textoOpcionIntentada,
                        evaluadasOpciones,
                        referencia.Clave);
                }

                // FASE 2431: sin cortocircuito aquí; se continúa a la enumeración contigua para comparar catálogo vs enumeración
            }

            // FASE 2433 / FASE 2443: Sin opciones y sin permitir enumeración, verificar si la referencia canónica es válida
            if (!permitirEnumeracion)
            {
                string etiquetaVentana = !string.IsNullOrWhiteSpace(diasVentana)
                    ? diasVentana.Trim()
                    : (ventana != null ? AgDiasVentana.AIso(string.Join(",", ventana)) : string.Empty);
                bool puedeEscalar = !string.IsNullOrEmpty(etiquetaVentana) && AgEscaleraVentanas.PuedeEscalar(etiquetaVentana);
                bool prohibirSingleton = ventana != null && ventana.Count == 5;
                bool tienePatronFinSemana = patrones != null && patrones.Any(p =>
                    !string.IsNullOrWhiteSpace(p) &&
                    (p.IndexOf("6M", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     p.IndexOf("6T", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     string.Equals(p.Trim(), "JUE-DOM", StringComparison.OrdinalIgnoreCase)));

                bool todosPatronesEnVentana = patrones == null || ventana == null || patrones.All(p => AgCarriles.Ventana(p).All(ventana.Contains));
                bool sinCatalogo = puedeEscalar && (
                    (prohibirSingleton && referencia.BandasSingleton > 0) ||
                    tienePatronFinSemana ||
                    (todosPatronesEnVentana && referencia.PatronesNoAlojables > 0));

                if (sinCatalogo)
                {
                    return new AgSeleccionParticion(
                        false, referencia, referencia,
                        0,
                        0,
                        "sin_catalogo",
                        string.Empty,
                        0,
                        referencia.Clave);
                }

                return new AgSeleccionParticion(
                    false, referencia, referencia,
                    0,
                    0,
                    "canonico",
                    string.Empty,
                    0,
                    referencia.Clave);
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
                string origenRef = referencia.BandasSingleton == 0 ? "catalogo" : "enumeracion";
                return new AgSeleccionParticion(false, referencia, referencia, 0, 0,
                    origenRef, string.Empty, 0, referencia.Clave);
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

            // Fase 2401: No descartar por singleton; sólo penalizar en el orden de desempate
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

                // (d) Reducción estricta de infactibles, singleton o acortamiento de makespan/bandas (F1736, F2395)
                bool menorInfactibles = ev.Resultado.Infactibles.Count < referencia.Resultado.Infactibles.Count;
                bool mejorMakespan = ev.Resultado.Infactibles.Count == referencia.Resultado.Infactibles.Count && ev.Makespan < referencia.Makespan;
                bool mejorSingleton = referencia.Resultado.Infactibles.Count == 0 && ev.BandasSingleton < referencia.BandasSingleton;
                bool menorBandas = referencia.Resultado.Infactibles.Count == 0
                    && ev.BandasSingleton == referencia.BandasSingleton
                    && ev.Makespan <= referencia.Makespan
                    && ev.Bandas < referencia.Bandas;
                if (!menorInfactibles && !mejorMakespan && !mejorSingleton && !menorBandas)
                {
                    descartadas++;
                    continue;
                }

                candidatas.Add(ev);
            }

            AgEvaluacionParticion mejorEnumeracion = null;
            if (candidatas.Count > 0)
            {
                // Fase 2404: Política explícita de singleton en el motor: prohibir bandas singleton si existe alternativa
                var elegibles = candidatas;
                if (!PermitirSingleton)
                {
                    var sinSingleton = candidatas
                        .Where(ev => ev.BandasSingleton == 0)
                        .ToList();

                    if (sinSingleton.Count > 0)
                    {
                        elegibles = sinSingleton;
                    }
                    else
                    {
                        LoggerHelper.WarningLimitado(
                            $"SINGLETON_NECESARIO|{ventana?.Count}|{referencia?.Clave}",
                            "[Oleadas] SINGLETON_NECESARIO: no hay partición sin bandas de 1 día; se conserva la mejor disponible.");
                    }
                }

                mejorEnumeracion = elegibles
                    .OrderBy(ev => ev.Resultado.Infactibles.Count)
                    .ThenBy(ev => ev.BandasSingleton)
                    .ThenBy(ev => ev.PatronesNoAlojables)
                    .ThenBy(ev => ev.Makespan)
                    .ThenBy(ev => ev.Bandas)
                    .ThenBy(ev => ev.MaxRanurasPorBanda)
                    .ThenBy(ev => ev.Clave, StringComparer.Ordinal)
                    .First();
            }

            // FASE 2431: Comparar catálogo y enumeración antes de decidir. En empate exacto (<= 0) gana catálogo.
            if (mejorCatalogo == null && mejorEnumeracion == null)
            {
                return new AgSeleccionParticion(false, referencia, referencia, evaluadas, descartadas,
                    catalogoEvaluado ? "catalogo_fallback" : "enumeracion", textoOpcionIntentada, evaluadasOpciones, referencia.Clave);
            }

            bool ganaCatalogo = mejorCatalogo != null &&
                (mejorEnumeracion == null || Comparar(mejorCatalogo, mejorEnumeracion) <= 0);

            if (ganaCatalogo)
            {
                if (mejorCatalogo.BandasSingleton < referencia.BandasSingleton)
                {
                    LoggerHelper.Debug(
                        $"[Oleadas] bandasSingleton={referencia.BandasSingleton}→{mejorCatalogo.BandasSingleton}; " +
                        $"motivo=CAMBIO_POR_SINGLETON; opcion='{textoMejorCatalogo}'");
                }

                // FASE 2220 / FASE 2432: CandidatasEvaluadas refleja ambas rutas evaluadas si hubo enumeración
                return new AgSeleccionParticion(
                    true, referencia, mejorCatalogo,
                    evaluadasOpciones + evaluadas,
                    descartadas,
                    "catalogo",
                    textoMejorCatalogo,
                    evaluadasOpciones,
                    mejorCatalogo.Clave);
            }

            return new AgSeleccionParticion(
                true, referencia, mejorEnumeracion,
                evaluadasOpciones + evaluadas,
                descartadas,
                catalogoEvaluado ? "catalogo_fallback" : "enumeracion",
                textoOpcionIntentada,
                evaluadasOpciones,
                mejorEnumeracion.Clave);
        }

        private static int Comparar(
            AgEvaluacionParticion a,
            AgEvaluacionParticion b)
        {
            if (ReferenceEquals(a, b)) return 0;
            if (a == null) return 1;
            if (b == null) return -1;

            int c = a.Resultado.Infactibles.Count.CompareTo(b.Resultado.Infactibles.Count);
            if (c != 0) return c;

            c = a.BandasSingleton.CompareTo(b.BandasSingleton);
            if (c != 0) return c;

            c = a.PatronesNoAlojables.CompareTo(b.PatronesNoAlojables);
            if (c != 0) return c;

            c = a.Makespan.CompareTo(b.Makespan);
            if (c != 0) return c;

            c = a.Bandas.CompareTo(b.Bandas);
            if (c != 0) return c;

            c = a.MaxRanurasPorBanda.CompareTo(b.MaxRanurasPorBanda);
            if (c != 0) return c;

            // FASE 2431: En empate métrico entre catálogo y enumeración, retorna 0 para que gane catálogo
            return 0;
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

        /// <summary>
        /// Determina si una opción de catálogo cubre exactamente las ranuras del contexto (Fase 2425).
        /// Evita que una lista no vacía pero inaplicable bloquee el fallback de catálogo.
        /// </summary>
        internal static bool EsOpcionAplicable(
            IReadOnlyList<AgRanura> ranuras, AgOpcionParticion opcion)
        {
            if (ranuras == null || ranuras.Count == 0) return false;
            if (opcion == null || opcion.Carriles == null || opcion.Carriles.Count == 0) return false;
            return ConservaRanuras(ranuras, opcion.Carriles);
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

