using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Compara una referencia por carriles con planes de sesiones fechadas para aprovechar ranuras y acortar el fin.
    /// La referencia reserva semanas por carril; la compactación puede compartir una semana entre materias.
    /// Carril identifica la banda elegida y Oleada su orden de asignación.
    /// </summary>
    internal sealed class AgPlanificadorOleadas
    {
        private readonly AgPlanificadorPatronSemanal _planificadorPatron;

        // Fase 1101 / Fase 1111: Contador sincronizado de ejecuciones efectivas para observabilidad y pruebas de memoización
        private int _construcciones;
        internal int Construcciones => System.Threading.Volatile.Read(ref _construcciones);

        public AgPlanificadorOleadas(AgPlanificadorPatronSemanal planificadorPatron = null)
        {
            _planificadorPatron = planificadorPatron ?? new AgPlanificadorPatronSemanal();
        }

        public Result<AgPlanOleadasResultado> Construir(
            IReadOnlyList<(int IdAsignatura, double Horas, int OrdenCurricular)> asignaturas,
            string diasVentana, DateTime fechaInicio, DateTime fechaFin, double unidadSesion,
            Func<DateTime, bool> esFestivo = null, IReadOnlyCollection<AgBloqueHorario> bloques = null,
            string jornada = null, IReadOnlyDictionary<int, string> diasPatronPorAsignatura = null, string diagnostico = null)
        {
            return ConstruirInterno(asignaturas, diasVentana, fechaInicio, fechaFin, unidadSesion,
                esFestivo, bloques, jornada, diasPatronPorAsignatura, false, System.Threading.CancellationToken.None, diagnostico);
        }

        internal Result<AgPlanOleadasResultado> ConstruirEvaluacion(
            IReadOnlyList<(int IdAsignatura, double Horas, int OrdenCurricular)> asignaturas,
            string diasVentana, DateTime fechaInicio, DateTime fechaFin, double unidadSesion,
            Func<DateTime, bool> esFestivo = null, IReadOnlyCollection<AgBloqueHorario> bloques = null,
            string jornada = null, IReadOnlyDictionary<int, string> diasPatronPorAsignatura = null,
            System.Threading.CancellationToken ct = default(System.Threading.CancellationToken), string diagnostico = null)
        {
            return ConstruirInterno(asignaturas, diasVentana, fechaInicio, fechaFin, unidadSesion,
                esFestivo, bloques, jornada, diasPatronPorAsignatura, true, ct, diagnostico);
        }

        private Result<AgPlanOleadasResultado> ConstruirInterno(
            IReadOnlyList<(int IdAsignatura, double Horas, int OrdenCurricular)> asignaturas,
            string diasVentana, DateTime fechaInicio, DateTime fechaFin, double unidadSesion,
            Func<DateTime, bool> esFestivo, IReadOnlyCollection<AgBloqueHorario> bloques,
            string jornada, IReadOnlyDictionary<int, string> patrones, bool permitirParcial,
            System.Threading.CancellationToken ct, string diagnostico)
        {
            ct.ThrowIfCancellationRequested();
            if (fechaFin < fechaInicio || unidadSesion <= 0 || double.IsNaN(unidadSesion) || double.IsInfinity(unidadSesion))
            {
                return Result.Failure<AgPlanOleadasResultado>("I1_ENTRADA_INVALIDA: fechas o unidad inválidas.");
            }
            if (AgCarriles.Ventana(diasVentana).Count == 0)
            {
                return Result.Failure<AgPlanOleadasResultado>("I1_VENTANA_NO_DEFINIDA: ventana sin días.");
            }
            if (asignaturas == null || asignaturas.Count == 0)
            {
                return Result.Success(new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>()));
            }
            if (asignaturas.GroupBy(a => a.IdAsignatura).Any(g => g.Count() > 1))
            {
                return Result.Failure<AgPlanOleadasResultado>("I1_ASIGNATURA_DUPLICADA: IdAsignatura repetido.");
            }
            if (asignaturas.Any(a => a.Horas <= 0 || double.IsNaN(a.Horas) || double.IsInfinity(a.Horas)))
            {
                return Result.Failure<AgPlanOleadasResultado>("I1_ENTRADA_INVALIDA: demanda inválida.");
            }
            var ranuras = AgRanuras.PorVentana(diasVentana, bloques, jornada, unidadSesion);
            if ((ranuras == null || ranuras.Count == 0) && bloques != null && bloques.Count > 0)
            {
                ranuras = AgRanuras.PorVentana(diasVentana, null, jornada, unidadSesion);
            }
            if (ranuras == null || ranuras.Count == 0)
            {
                return Result.Failure<AgPlanOleadasResultado>("I1_CONF_RANURA: ventana sin ranuras activas.");
            }
            var referencia = ConstruirReferencia(asignaturas, diasVentana, fechaInicio, fechaFin,
                unidadSesion, esFestivo, bloques, jornada, patrones);
            if (referencia.IsFailure && !referencia.Error.StartsWith("I5_", StringComparison.Ordinal))
            {
                LoggerHelper.Debug($"[AG-RANURAS-TMP] ELEGIR {diagnostico} plan=error_referencia motivo={referencia.Error} candidatoEjecutado=false");
                return referencia;
            }
            // También se buscan mejoras cuando la referencia ya es factible en una ventana ampliada.
            var alternativas = new List<AgPlanOleadasResultado>();
            AgPlanOleadasResultado referenciaValida = null;
            if (referencia.IsSuccess && AgPlanificadorRanuras.ValidarFechado(referencia.Value,
                asignaturas, ranuras, fechaInicio, fechaFin, unidadSesion, esFestivo, patrones, ct).IsSuccess)
            {
                referenciaValida = referencia.Value;
                alternativas.Add(referenciaValida);
            }
            bool tienePatrones = patrones != null && patrones.Values.Any(v => !string.IsNullOrWhiteSpace(v));
            bool hayAutomaticas = asignaturas.Any(a => patrones == null
                || !patrones.TryGetValue(a.IdAsignatura, out var p)
                || string.IsNullOrWhiteSpace(p));

            if (!permitirParcial && tienePatrones && !hayAutomaticas && referencia.IsSuccess && referencia.Value.Infactibles.Count == 0)
            {
                return referencia;
            }

            if (tienePatrones)
            {
                var completado = AgPlanificadorRanuras.CompletarReferencia(asignaturas, ranuras,
                    fechaInicio, fechaFin, unidadSesion, esFestivo, patrones, referenciaValida, ct);
                if (completado.IsSuccess) { alternativas.Add(completado.Value); }

                // Primero un pase barato sobre huecos. Búsqueda acotada solo cuando no lo resuelve.
                if (completado.IsFailure || completado.Value.Infactibles.Count > 0)
                {
                    var buscado = AgPlanificadorRanuras.Construir(asignaturas, ranuras,
                        fechaInicio.Date, fechaFin.Date, unidadSesion, esFestivo, patrones,
                        ct: ct, diagnostico: diagnostico);
                    if (AgPlanificadorRanuras.ValidarFechado(buscado, asignaturas, ranuras,
                        fechaInicio, fechaFin, unidadSesion, esFestivo, patrones, ct).IsSuccess)
                    {
                        alternativas.Add(buscado);
                    }
                }
            }
            else
            {
                var buscado = AgPlanificadorRanuras.Construir(asignaturas, ranuras,
                    fechaInicio.Date, fechaFin.Date, unidadSesion, esFestivo, patrones,
                    ct: ct, diagnostico: diagnostico);
                if (AgPlanificadorRanuras.ValidarFechado(buscado, asignaturas, ranuras,
                    fechaInicio, fechaFin, unidadSesion, esFestivo, patrones, ct).IsSuccess)
                {
                    alternativas.Add(buscado);
                }
            }

            var requeridas = asignaturas.ToDictionary(a => a.IdAsignatura,
                a => (int)Math.Ceiling(a.Horas / unidadSesion));
            var completasPrevias = referenciaValida == null ? new List<int>()
                : referenciaValida.Planes.Where(p => !referenciaValida.Infactibles.Contains(p.IdAsignatura))
                    .Select(p => p.IdAsignatura).ToList();
            var elegibles = alternativas.Where(p => completasPrevias.All(id =>
                p.Planes.Any(x => x.IdAsignatura == id && x.Sesiones.Count == requeridas[id]))).ToList();
            var completos = elegibles.Where(p => p.Infactibles.Count == 0).ToList();

            if (completos.Count > 0)
            {
                // Si la referencia ya era completa, solo preferir alternativa si termina estrictamente antes
                if (referenciaValida != null && referenciaValida.Infactibles.Count == 0)
                {
                    DateTime finRef = referenciaValida.Planes.SelectMany(p => p.Sesiones)
                        .Select(s => s.Fecha).DefaultIfEmpty(DateTime.MaxValue).Max();
                    var mejorAntes = completos.Where(p =>
                    {
                        DateTime finCand = p.Planes.SelectMany(x => x.Sesiones)
                            .Select(s => s.Fecha).DefaultIfEmpty(DateTime.MaxValue).Max();
                        return tienePatrones ? (finCand < finRef) : (finRef - finCand > TimeSpan.FromDays(7));
                    }).OrderBy(p => p.Planes.SelectMany(x => x.Sesiones).Select(s => s.Fecha).Max()).FirstOrDefault();

                    if (mejorAntes != null)
                    {
                        LoggerHelper.Debug($"[AG-RANURAS-TMP] ELEGIR {diagnostico} plan=alternativa_completa_adelantada " +
                            $"sesiones={mejorAntes.Planes.Sum(p => p.Sesiones.Count)}");
                        return Result.Success(mejorAntes);
                    }

                    LoggerHelper.Debug($"[AG-RANURAS-TMP] ELEGIR {diagnostico} plan=referencia_completa_canonico " +
                        $"sesiones={referenciaValida.Planes.Sum(p => p.Sesiones.Count)}");
                    return Result.Success(referenciaValida);
                }

                // Si la referencia no era completa, elegir el completo que termine más temprano
                var mejorCompleto = completos.OrderBy(p => p.Planes.SelectMany(x => x.Sesiones)
                    .Select(s => s.Fecha).DefaultIfEmpty(DateTime.MaxValue).Max()).First();
                LoggerHelper.Debug($"[AG-RANURAS-TMP] ELEGIR {diagnostico} plan=completo_rescate " +
                    $"sesiones={mejorCompleto.Planes.Sum(p => p.Sesiones.Count)}");
                return Result.Success(mejorCompleto);
            }

            // No hay plan 100% completo
            if (!permitirParcial)
            {
                return referencia;
            }

            var elegido = elegibles.OrderByDescending(p => p.Planes.Sum(x => x.Sesiones.Count))
                .ThenBy(p => p.Infactibles.Count)
                .ThenBy(p => p.Planes.SelectMany(x => x.Sesiones).Select(s => s.Fecha)
                    .DefaultIfEmpty(DateTime.MaxValue).Max()).FirstOrDefault();

            if (elegido != null)
            {
                LoggerHelper.Debug($"[AG-RANURAS-TMP] ELEGIR {diagnostico} plan=parcial_optimo " +
                    $"sesiones={elegido.Planes.Sum(p => p.Sesiones.Count)} infactibles={elegido.Infactibles.Count}");
                return Result.Success(elegido);
            }

            return referencia;
        }

        /// <summary>
        /// Construye el plan de carril y oleada para una cohorte/aula (Fases 1057, 1088, 1092, 1095, 1226).
        /// Orden canónico de asignaturas: (OrdenCurricular, IdAsignatura).
        /// </summary>
        private Result<AgPlanOleadasResultado> ConstruirReferencia(
            IReadOnlyList<(int IdAsignatura, double Horas, int OrdenCurricular)> asignaturas,
            string diasVentana,
            DateTime fechaInicio,
            DateTime fechaFin,
            double unidadSesion,
            Func<DateTime, bool> esFestivo = null,
            IReadOnlyCollection<AgBloqueHorario> bloques = null,
            string jornada = null,
            IReadOnlyDictionary<int, string> diasPatronPorAsignatura = null)   // Fase 1226
        {
            System.Threading.Interlocked.Increment(ref _construcciones);

            if (fechaFin < fechaInicio)
            {
                return Result.Failure<AgPlanOleadasResultado>(
                    "I1_FECHAS_INVALIDAS: FechaFin es anterior a FechaInicio.");
            }

            var ventana = AgCarriles.Ventana(diasVentana);
            if (ventana == null || ventana.Count == 0)
            {
                return Result.Failure<AgPlanOleadasResultado>(
                    "I1_VENTANA_NO_DEFINIDA: Aula sin ventana semanal declarada.");
            }

            if (asignaturas == null || asignaturas.Count == 0)
            {
                return Result.Success(new AgPlanOleadasResultado(new List<AgPlanOleada>(), new List<int>()));
            }

            // F1585: Rechazar IDs duplicados antes de planificar para prevenir desalineación y fallo en ToDictionary
            if (asignaturas.GroupBy(a => a.IdAsignatura).Any(g => g.Count() > 1))
            {
                return Result.Failure<AgPlanOleadasResultado>(
                    "I1_ASIGNATURA_DUPLICADA: IdAsignatura repetido.");
            }

            double unidad = unidadSesion > 0 ? unidadSesion : 4.0;
            // Fase 1154 / Fase 1155 / Fase 1166 / Fase 1174 / Fase 1178 / Fase 1183: Resolución canónica de ranuras (Día, Franja, Horas)
            // F1178: el catálogo es una restricción de calidad, no un candado de factibilidad.
            var ranuras = AgRanuras.PorVentana(ventana, bloques, jornada, unidad);
            if ((ranuras == null || ranuras.Count == 0) && bloques != null && bloques.Count > 0)
            {
                var faltantes = AgRanuras.PorVentana(ventana, null, jornada, unidad)
                    .Select(r => $"{r.Dia}:{r.Franja}({r.Horas:0.#}h)");
                LoggerHelper.WarningLimitado(
                    $"I1_CONF_BLOQUE_HORARIO|{diasVentana}|{jornada}|{unidad}",
                    $"[Oleadas] I1_CONF_BLOQUE_HORARIO: ventana '{diasVentana}' sin bloque activo de {unidad}h " +
                    $"(jornada '{jornada}'); faltan [{string.Join(", ", faltantes)}]. Se planifica con ranuras de compatibilidad.");
                ranuras = AgRanuras.PorVentana(ventana, null, jornada, unidad);
            }
            if (ranuras == null || ranuras.Count == 0)
            {
                LoggerHelper.WarningLimitado(
                    $"RanurasCompatibilidad|{diasVentana}|{unidad}",
                    $"[Oleadas] Sin ranuras resolubles para '{diasVentana}' ({unidad}h); se usa NOCHE por compatibilidad.");
                ranuras = ventana.Select(d => new AgRanura(d, AgFranjaJornada.NOCHE, unidad)).ToList();
            }   // Fase 1263: llave restituida

            // 1. Obtener los slices semanales disponibles en la ventana semestral
            var semanasValidas = new List<(DateTime Inicio, DateTime Fin)>();
            for (var c = fechaInicio.Date; c <= fechaFin.Date;)
            {
                bool hayDiaVentana = false;
                for (int i = 0; i < 7 && !hayDiaVentana; i++)
                {
                    if (c > DateTime.MaxValue.Date.AddDays(-i)) break;
                    var d = c.AddDays(i);
                    if (d > fechaFin.Date) break;
                    int iso = d.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)d.DayOfWeek;
                    if (ventana.Contains(iso) && (esFestivo == null || !esFestivo(d)))
                    {
                        hayDiaVentana = true;
                    }
                }

                if (hayDiaVentana)
                {
                    var finSem = c > DateTime.MaxValue.Date.AddDays(-6) || c.AddDays(6) > fechaFin.Date ? fechaFin.Date : c.AddDays(6);
                    semanasValidas.Add((c, finSem));
                }

                if (c > DateTime.MaxValue.Date.AddDays(-7) || c.AddDays(7) > fechaFin.Date) break;
                c = c.AddDays(7);
            }

            int W = semanasValidas.Count;

            if (W == 0)
            {
                return Result.Failure<AgPlanOleadasResultado>(
                    $"I5_SIN_SEMANAS_HABILES: 0 semanas válidas entre {fechaInicio:yyyy-MM-dd} y {fechaFin:yyyy-MM-dd} " +
                    $"para ventana '{diasVentana}' (festivos o días fuera de catálogo).");   // Fase 1245
            }

            // 2. Orden canónico de asignaturas: (COALESCE(OrdenCurricular, 0), IdAsignatura)
            var ordenadas = asignaturas
                .OrderBy(a => a.OrdenCurricular > 0 ? a.OrdenCurricular : 0)
                .ThenBy(a => a.IdAsignatura)
                .ToList();

            var demanda = ordenadas.Select(a => (int)Math.Ceiling(a.Horas / unidad)).ToList();
            int demandaTotal = demanda.Sum();

            // Fase 1267 / Fase 1273: capacidad efectiva por día (festivos descontados) — antes del guard
            var capacidadDia = new Dictionary<int, int>();
            foreach (var r in ranuras)
            {
                int oc = AgPlanificadorTemporal.ContarOcurrencias(r.Dia, fechaInicio.Date, fechaFin.Date, esFestivo);
                capacidadDia.TryGetValue(r.Dia, out var acc);
                capacidadDia[r.Dia] = acc + oc;
            }

            int capacidadTotal = capacidadDia.Values.Sum();
            if (demandaTotal > capacidadTotal)
            {
                return Result.Failure<AgPlanOleadasResultado>(
                    $"I5_SIN_CAPACIDAD: demanda {demandaTotal} sesiones > capacidad efectiva {capacidadTotal} sesiones " +
                    $"({ranuras.Count} ranuras/semana x {W} semanas, festivos descontados) en ventana '{diasVentana}'.");
            }

            // 3. Fase 1279 / Fase 1581: carriles y oleadas secuenciales evaluadas por partición.
            //    Fase 1581: cuando la partición canónica fragmenta patrones sembrados, se evalúan particiones
            //    alternativas y se acepta la que minimice infactibles reales sin perder planes previos.
            var ocurrenciasPorDia = new Dictionary<int, int>();
            foreach (var r in ranuras)
            {
                if (!ocurrenciasPorDia.ContainsKey(r.Dia))
                {
                    ocurrenciasPorDia[r.Dia] = AgPlanificadorTemporal.ContarOcurrencias(
                        r.Dia, fechaInicio.Date, fechaFin.Date, esFestivo);
                }
            }

            var ctx = new ContextoEvaluacion
            {
                Ordenadas = ordenadas,
                Demanda = demanda,
                Patrones = ordenadas.Select(a => diasPatronPorAsignatura != null
                             && diasPatronPorAsignatura.TryGetValue(a.IdAsignatura, out var p)
                             && !string.IsNullOrWhiteSpace(p) ? p : string.Empty).ToList(),
                SemanasValidas = semanasValidas,
                W = W,
                Unidad = unidad,
                EsFestivo = esFestivo,
                OcurrenciasPorDia = ocurrenciasPorDia,
                DemandaTotal = demandaTotal,
                CapacidadTotal = capacidadTotal
            };

            int carrilesCanonicos = AgCarriles.Cantidad(ranuras);
            var canonicas = Enumerable.Range(0, carrilesCanonicos)
                .Select(c => (IReadOnlyList<AgRanura>)AgCarriles.DelCarril(ranuras, c))
                .ToList();

            var seleccion = AgParticionCarriles.Seleccionar(
                ventana,
                ranuras,
                ctx.Patrones,
                canonicas,
                bandas => EvaluarParticion(ctx, diasVentana, bandas),
                ev => EsParticionFiel(ctx, ev));

            if (seleccion.Cambiada)
            {
                var r = seleccion.Referencia;
                var e = seleccion.Elegida;
                LoggerHelper.Info(
                    $"[Oleadas] Partición por patrón: ventana='{diasVentana}'; " +
                    $"bandas={string.Join("|", e.Clave.Split('|').Select(b => "{" + b.Replace('.', ',') + "}"))}; " +
                    $"noAlojables={r.PatronesNoAlojables}→{e.PatronesNoAlojables}; " +
                    $"infactibles={r.Resultado.Infactibles.Count}→{e.Resultado.Infactibles.Count}; " +
                    $"makespan={r.Makespan}→{e.Makespan}; candidatasEvaluadas={seleccion.CandidatasEvaluadas}.");
            }

            return Result.Success(seleccion.Elegida.Resultado);
        }

        private sealed class ContextoEvaluacion
        {
            public List<(int IdAsignatura, double Horas, int OrdenCurricular)> Ordenadas;
            public List<int> Demanda;
            public List<string> Patrones;
            public List<(DateTime Inicio, DateTime Fin)> SemanasValidas;
            public int W;
            public double Unidad;
            public Func<DateTime, bool> EsFestivo;
            public Dictionary<int, int> OcurrenciasPorDia;
            public int DemandaTotal;
            public int CapacidadTotal;
        }

        private AgEvaluacionParticion EvaluarParticion(
            ContextoEvaluacion ctx,
            string diasVentana,
            IReadOnlyList<IReadOnlyList<AgRanura>> bandas)
        {
            // F1582: Defensa interna contra estado inválido de semanas o bandas
            if (ctx.W <= 0 || ctx.SemanasValidas.Count != ctx.W)
                throw new InvalidOperationException("Calendario de evaluación inválido.");

            if (bandas == null || bandas.Count == 0 || bandas.Any(b => b == null || b.Count == 0))
                throw new ArgumentException("Partición inválida.", nameof(bandas));

            int carriles = bandas.Count;
            var ranurasCarril = new List<IReadOnlyList<AgRanura>>(carriles);
            var diasCarril = new List<IReadOnlyList<int>>(carriles);
            var slotsCarril = new int[carriles];
            var capacidadCarril = new int[carriles];

            for (int c = 0; c < carriles; c++)
            {
                var rc = bandas[c];
                ranurasCarril.Add(rc);
                diasCarril.Add(rc.Select(r => r.Dia).Distinct().ToList());
                slotsCarril[c] = Math.Max(1, rc.Count);
                capacidadCarril[c] = rc.Sum(r => ctx.OcurrenciasPorDia[r.Dia]);
            }

            // 3.1 Pasada 1: el patrón sembrado fija el carril cuando cabe íntegro y hay capacidad
            // F1600: Busca entre todos los carriles compatibles con capacidad efectiva disponible,
            // en lugar de elegir primero el más específico sin verificar capacidad.
            var carrilDe = new int[ctx.Ordenadas.Count];
            for (int i = 0; i < carrilDe.Length; i++) carrilDe[i] = -1;

            var carga = new int[carriles];
            int renormalizados = 0;
            for (int i = 0; i < ctx.Ordenadas.Count; i++)
            {
                var diasPatron = AgCarriles.Ventana(ctx.Patrones[i]);
                if (diasPatron.Count == 0) continue;

                int mejor = -1;
                int mejorTam = int.MaxValue;
                for (int c = 0; c < carriles; c++)
                {
                    if (diasPatron.Any(d => !diasCarril[c].Contains(d)))
                        continue;

                    if (carga[c] + ctx.Demanda[i] > capacidadCarril[c])
                        continue;

                    if (diasCarril[c].Count < mejorTam)
                    {
                        mejor = c;
                        mejorTam = diasCarril[c].Count;
                    }
                }

                if (mejor >= 0)
                {
                    carrilDe[i] = mejor;
                    carga[mejor] += ctx.Demanda[i];
                }
                else
                {
                    renormalizados++;
                }
            }

            // 3.2 Pasada 2: reparto greedy por utilización relativa manteniendo restricciones sembradas
            // F1600: Asignaturas con patrón sembrado respetan contención de días y capacidad;
            // si ningún carril compatible tiene capacidad, carrilDe[i] permanece -1 para ser diagnosticado
            // en Pasada 3 en lugar de desbordarse a un carril incompatible arbitrario.
            for (int i = 0; i < ctx.Ordenadas.Count; i++)
            {
                if (carrilDe[i] >= 0) continue;

                var diasPatron = AgCarriles.Ventana(ctx.Patrones[i]);

                int elegido = -1;
                double mejorScore = double.MaxValue;
                for (int c = 0; c < carriles; c++)
                {
                    if (diasPatron.Any(d => !diasCarril[c].Contains(d)))
                        continue;

                    if (diasPatron.Count > 0 &&
                        carga[c] + ctx.Demanda[i] > capacidadCarril[c])
                        continue;

                    int proyectada = carga[c] + ctx.Demanda[i];
                    double uso = proyectada / (double)Math.Max(1, capacidadCarril[c]);
                    double score = (proyectada > capacidadCarril[c] ? 10.0 : 0.0) + uso;
                    if (elegido < 0
                        || score < mejorScore - 1e-9
                        || (Math.Abs(score - mejorScore) <= 1e-9 && carga[c] < carga[elegido]))
                    {
                        elegido = c;
                        mejorScore = score;
                    }
                }

                if (elegido < 0)
                    continue;

                carrilDe[i] = elegido;
                carga[elegido] += ctx.Demanda[i];
            }

            // 3.3 Pasada 3: cursor secuencial por carril
            var cursor = new int[carriles];
            var oleada = new int[carriles];
            var planes = new List<AgPlanOleada>(ctx.Ordenadas.Count);
            var infactibles = new List<int>();
            var detalleInfactibles = new List<string>();

            for (int i = 0; i < ctx.Ordenadas.Count; i++)
            {
                var asig = ctx.Ordenadas[i];
                string accion = AgEscaleraVentanas.PuedeEscalar(diasVentana)
                    ? $" Escale la ventana a '{AgEscaleraVentanas.Siguiente(diasVentana)}' (AgEscaleraVentanas)."
                    : string.Empty;

                int c = carrilDe[i];
                // F1600 / F1601: Manejar asignación pendiente antes de indexar slotsCarril[c]
                // Distingue incompatibilidad estructural (R8) de saturación de capacidad del carril (I5)
                // y expone capacidades reservadas de carriles compatibles en un solo detalle por asignatura.
                if (c < 0)
                {
                    var dias = AgCarriles.Ventana(ctx.Patrones[i]);
                    bool alojable = diasCarril.Any(b => dias.All(b.Contains));

                    string codigo = alojable
                        ? "I5_CAPACIDAD_CARRIL_PATRON"
                        : "R8_DiaFueraDePatron";

                    string capacidades = string.Join(" | ",
                        Enumerable.Range(0, carriles)
                            .Where(j => dias.All(d => diasCarril[j].Contains(d)))
                            .Select(j =>
                                $"c{j}[{string.Join(",", diasCarril[j])}]" +
                                $" capacidad={capacidadCarril[j]}" +
                                $" reservada={carga[j]}"));

                    string compatStr = !string.IsNullOrEmpty(capacidades)
                        ? $"; compatibles: {capacidades}"
                        : string.Empty;

                    infactibles.Add(asig.IdAsignatura);
                    detalleInfactibles.Add(
                        $"Asig {asig.IdAsignatura}: patrón '{ctx.Patrones[i]}'; " +
                        $"demanda={ctx.Demanda[i]} sesiones; " +
                        $"sin carril compatible disponible ({codigo}){compatStr}.{accion}");
                    continue;
                }

                var diasFijos = AgCarriles.Ventana(ctx.Patrones[i]);
                var ranurasEfectivas = ranurasCarril[c].Where(r => diasFijos.Count == 0 || diasFijos.Contains(r.Dia)).ToList();
                int SRequeridas = ctx.Demanda[i];
                int kBase = Math.Max(1, (int)Math.Ceiling(SRequeridas / (double)slotsCarril[c]));

                // F1581: R8 verificado antes de construir el patrón para evitar trabajo descartado
                var diasPatronAsig = AgCarriles.Ventana(ctx.Patrones[i]);
                if (diasPatronAsig.Count > 0 && diasPatronAsig.Any(d => !diasCarril[c].Contains(d)))
                {
                    infactibles.Add(asig.IdAsignatura);
                    detalleInfactibles.Add(
                        $"Asig {asig.IdAsignatura}: patrón '{ctx.Patrones[i]}' no cabe en carril {c} " +
                        $"[{string.Join(",", diasCarril[c])}] (R8_DiaFueraDePatron).{accion}");
                    continue;
                }

                DateTime ini = ctx.SemanasValidas[0].Inicio;
                DateTime fin = ctx.SemanasValidas[ctx.W - 1].Fin;
                int k = kBase;
                Result<AgPatronSemanal> resPatron = Result.Failure<AgPatronSemanal>("sin semanas disponibles");

                for (int extra = 0; ; extra++)
                {
                    k = kBase + extra;
                    if (cursor[c] + k > ctx.W) break;
                    ini = ctx.SemanasValidas[cursor[c]].Inicio;
                    fin = ctx.SemanasValidas[cursor[c] + k - 1].Fin;
                    resPatron = _planificadorPatron.Construir(
                        ranurasCarril[c], (int)Math.Ceiling(asig.Horas), ini, fin, ctx.EsFestivo, ctx.Unidad);
                    if (resPatron.IsSuccess) break;
                }

                if (resPatron.IsFailure)
                {
                    infactibles.Add(asig.IdAsignatura);
                    detalleInfactibles.Add(
                        $"Asig {asig.IdAsignatura}: {asig.Horas:0.#}h = {SRequeridas} sesiones; carril {c} " +
                        $"[{string.Join(",", diasCarril[c])}] desde semana {cursor[c] + 1}/{ctx.W}; " +
                        $"{Math.Max(0, capacidadCarril[c] - cursor[c] * slotsCarril[c])} sesiones libres en el carril; " +
                        $"{resPatron.Error}{accion}");
                    continue;
                }

                int S = resPatron.Value.SesionesPorSemana != null ? resPatron.Value.SesionesPorSemana.Sum() : SRequeridas;

                planes.Add(new AgPlanOleada(
                    asig.IdAsignatura,
                    c,
                    oleada[c],
                    ranurasEfectivas.Select(r => r.Dia).Distinct().ToList(),
                    ini,
                    fin,
                    k,
                    S > 0 ? S : SRequeridas,
                    slotsCarril[c] * k,
                    ranurasCarril[c],
                    AgPlanificadorRanuras.Materializar(asig.IdAsignatura, ranurasEfectivas, SRequeridas, ini, fin, ctx.EsFestivo)));

                cursor[c] += k;
                oleada[c]++;
            }

            string ventanaSugerida = infactibles.Count > 0 ? AgEscaleraVentanas.Siguiente(diasVentana) : string.Empty;
            double ocupacion = ctx.CapacidadTotal > 0
                ? Math.Round(100.0 * ctx.DemandaTotal / ctx.CapacidadTotal, 1)
                : 0.0;

            var resResultado = new AgPlanOleadasResultado(
                planes,
                infactibles,
                detalleInfactibles,
                diasCarril,
                ventanaSugerida,
                renormalizados,
                ocupacion);

            int noAloj = AgParticionCarriles.ContarPatronesNoAlojables(bandas, ctx.Patrones);
            var alojables = AgParticionCarriles.ObtenerIndicesAlojables(bandas, ctx.Patrones);
            int makespan = cursor.Length > 0 ? cursor.Max() : 0;
            int maxRanuras = bandas.Count > 0 ? bandas.Max(b => b.Count) : 0;
            string clave = string.Join("|", bandas.Select(b =>
                string.Join(".", b.Select(r => r.Dia).Distinct().OrderBy(d => d))));

            return new AgEvaluacionParticion(
                resResultado,
                noAloj,
                alojables,
                makespan,
                bandas.Count,
                maxRanuras,
                clave);
        }

        /// <summary>
        /// F1586: Verifica que ninguna asignatura con patrón sembrado vea ampliados sus días permitidos
        /// por haber sido asignada a una banda con días adicionales a los de su patrón.
        /// </summary>
        private static bool EsParticionFiel(ContextoEvaluacion ctx, AgEvaluacionParticion ev)
        {
            if (ev?.Resultado?.Planes == null || ev.Resultado.Planes.Count == 0) return true;
            var porId = ev.Resultado.Planes.ToDictionary(p => p.IdAsignatura);

            for (int i = 0; i < ctx.Ordenadas.Count; i++)
            {
                var dias = AgCarriles.Ventana(ctx.Patrones[i]);
                if (dias.Count == 0 || !porId.TryGetValue(ctx.Ordenadas[i].IdAsignatura, out var plan))
                    continue;

                if (!new HashSet<int>(dias).SetEquals(plan.RanurasPatron.Select(r => r.Dia)))
                    return false;
            }
            return true;
        }
    }
}
