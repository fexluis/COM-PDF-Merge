using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
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
                        return finCand < finRef;
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

            // Fase 2109 / Fase 2114: la capacidad agregada insuficiente (o el calendario sin semanas) es
            // infactibilidad evaluable, no error estructural. En modo parcial se publica el mejor plan parcial
            // determinista para que el escalador compare el escalón y ofrezca reclasificación.
            if (AgEvaluadorDemandaVentana.EsInfactibilidadDeCapacidad(referencia.Error))
            {
                var parcial = AgPlanificadorRanuras.Construir(asignaturas, ranuras, fechaInicio.Date, fechaFin.Date,
                    unidadSesion, esFestivo, patrones, ct: ct, diagnostico: diagnostico);
                if (parcial != null && parcial.Planes.Count > 0)
                {
                    LoggerHelper.Warning($"[Oleadas] Infactibilidad evaluable: {referencia.Error} | " +
                        $"ubicadas={parcial.Planes.Sum(p => p.Sesiones.Count)} pendientes={parcial.Infactibles.Count}");
                    return Result.Success(parcial);
                }
            }

            return referencia;
        }

        /// <summary>
        /// Construye el plan de carril y oleada para una cohorte/aula (Fases 1057, 1088, 1092, 1095, 1226).
        /// Prioridad de asignación: horas descendentes, orden curricular normalizado e identificador ascendentes.
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

            // 1. Obtener los slices semanales disponibles en la ventana semestral (Fase 2129: helper canónico compartido)
            var semanasValidas = ConstruirSemanasValidas(ventana, fechaInicio, fechaFin, esFestivo);

            int W = semanasValidas.Count;

            if (W == 0)
            {
                return Result.Failure<AgPlanOleadasResultado>(
                    $"I5_SIN_SEMANAS_HABILES: 0 semanas válidas entre {fechaInicio:yyyy-MM-dd} y {fechaFin:yyyy-MM-dd} " +
                    $"para ventana '{diasVentana}' (festivos o días fuera de catálogo).");   // Fase 1245
            }

            // 2. La prioridad por horas rige la reserva de capacidad y las oleadas.
            var ordenadas = AgOrdenAsignaturas.PorHoras(asignaturas);

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
                SemanasValidas = semanasValidas.ToList(),
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

            // 3.1 Reserva en un único recorrido por prioridad de horas.
            // Los patrones exigen contención y capacidad; las automáticas equilibran utilización.
            var carrilDe = new int[ctx.Ordenadas.Count];
            for (int i = 0; i < carrilDe.Length; i++) carrilDe[i] = -1;

            var carga = new int[carriles];
            int renormalizados = 0;
            for (int i = 0; i < ctx.Ordenadas.Count; i++)
            {
                var diasPatron = AgCarriles.Ventana(ctx.Patrones[i]);
                if (diasPatron.Count == 0)
                {
                    int elegido = -1;
                    double mejorScore = double.MaxValue;
                    for (int c = 0; c < carriles; c++)
                    {
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
                    carrilDe[i] = elegido;
                    carga[elegido] += ctx.Demanda[i];
                    continue;
                }

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

                int demandaRestanteCarril = 0;
                for (int next = i + 1; next < ctx.Ordenadas.Count; next++)
                {
                    if (carrilDe[next] == c)
                    {
                        demandaRestanteCarril += ctx.Demanda[next];
                    }
                }
                int semanasMinimasRestantes = (int)Math.Ceiling(demandaRestanteCarril / (double)slotsCarril[c]);

                var sesiones = AgPlanificadorRanuras.Materializar(asig.IdAsignatura, ranurasEfectivas, SRequeridas, ini, fin, ctx.EsFestivo);
                while (sesiones.Count < SRequeridas && cursor[c] + k + semanasMinimasRestantes < ctx.W)
                {
                    k++;
                    fin = ctx.SemanasValidas[cursor[c] + k - 1].Fin;
                    sesiones = AgPlanificadorRanuras.Materializar(asig.IdAsignatura, ranurasEfectivas, SRequeridas, ini, fin, ctx.EsFestivo);
                }

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
                    sesiones));

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

        /// <summary>
        /// Obtiene los slices semanales disponibles en la ventana entre fechaInicio y fechaFin
        /// que contienen al menos un día hábil perteneciente a los días de la ventana (Fase 2129).
        /// </summary>
        internal static IReadOnlyList<(DateTime Inicio, DateTime Fin)> ConstruirSemanasValidas(
            IReadOnlyList<int> diasVentana, DateTime fechaInicio, DateTime fechaFin, Func<DateTime, bool> esFestivo)
        {
            var semanasValidas = new List<(DateTime Inicio, DateTime Fin)>();
            if (diasVentana == null || diasVentana.Count == 0 || fechaFin < fechaInicio)
            {
                return semanasValidas;
            }

            var diasSet = new HashSet<int>(diasVentana);
            for (var c = fechaInicio.Date; c <= fechaFin.Date;)
            {
                bool hayDiaVentana = false;
                for (int i = 0; i < 7 && !hayDiaVentana; i++)
                {
                    if (c > DateTime.MaxValue.Date.AddDays(-i)) break;
                    var d = c.AddDays(i);
                    if (d > fechaFin.Date) break;
                    int iso = d.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)d.DayOfWeek;
                    if (diasSet.Contains(iso) && (esFestivo == null || !esFestivo(d)))
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

            return semanasValidas;
        }

        /// <summary>
        /// Resultado de la operación de realineación determinista de oleadas por carril (Fase 2129).
        /// </summary>
        internal sealed class ResultadoRealineacion
        {
            internal bool Exito;
            internal bool HuboCambios;
            internal int CarrilesEvaluados;
            internal int CarrilesReparados;
            internal int AsignaturasDesplazadas;
            internal int Fallos;
            internal readonly List<string> Detalle = new List<string>();
            internal string Error = string.Empty;
        }

        private sealed class AsignaturaCarril
        {
            public int IdAsignatura;
            public List<AgGen> Genes;
            public DateTime Ini;
            public DateTime Fin;
            public int OrdenCurricular;
            public IReadOnlyList<int> Dias;
            public double Unidad;
            public double TotalHoras;
            public int SesionesRequeridas;
            public IReadOnlyList<AgRanura> Ranuras;
            public int Slots;
        }

        private sealed class AccionGen
        {
            public AgGen Gen;
            public DateTime FechaInicio;
            public DateTime FechaFin;
            public List<DateTime> FechasGen;
            public double HorasGen;
            public string PatronSemanal;
        }

        /// <summary>
        /// Repara el individuo campeón secuenciando las oleadas solapadas dentro de cada carril efectivo
        /// mediante forward shifting determinista sobre las semanas hábiles de la ventana (Fase 2129).
        /// </summary>
        internal static ResultadoRealineacion RealinearOleadas(AgIndividuo ind, AgContexto ctx)
        {
            if (ind?.Genes == null || ind.Genes.Count == 0)
            {
                return new ResultadoRealineacion { Exito = true, HuboCambios = false };
            }
            if (ctx?.Ventana == null)
            {
                return new ResultadoRealineacion { Exito = false, Error = "Ventana semestral no resuelta." };
            }

            var res = new ResultadoRealineacion { Exito = true };
            var planificadorPatron = new AgPlanificadorPatronSemanal();

            // 1) Carriles efectivos por cohorte y carril canónico (Fase 2140 / Fase 2163)
            var carriles = ind.Genes
                .Where(g => g.EsVacante == 0
                         && g.FechaInicioEstimada.HasValue
                         && g.FechaFinEstimada.HasValue)
                .GroupBy(g => (Cohorte: AgCalendarioConflictos.Cohorte(g), Carril: AgCarriles.ClaveCarril(g)))
                .OrderBy(k => k.Key.Cohorte, StringComparer.Ordinal)
                .ThenBy(k => k.Key.Carril, StringComparer.Ordinal)
                .ToList();

            foreach (var carrilGrp in carriles)
            {
                res.CarrilesEvaluados++;
                var genes = carrilGrp.ToList();
                var porAsig = genes.GroupBy(g => g.IdAsignatura).ToList();
                if (porAsig.Count <= 1) continue;

                int idAulaCarril = genes.FirstOrDefault()?.IdAula ?? 0;
                var aula = ctx.Aulas?.FirstOrDefault(a => a.Id == idAulaCarril)
                    ?? (ctx.AulasPorId != null && ctx.AulasPorId.TryGetValue(idAulaCarril, out var aPorId) ? aPorId : null);

                var asigs = new List<AsignaturaCarril>();
                foreach (var gAsig in porAsig)
                {
                    var asigGenes = gAsig.ToList();
                    var primerGen = asigGenes.FirstOrDefault();
                    if (primerGen == null) continue;

                    int ordenCurricular = 999;
                    if (ctx.AulaAsignaturaPorAulaAsig != null && primerGen.IdAula > 0 &&
                        ctx.AulaAsignaturaPorAulaAsig.TryGetValue((primerGen.IdAula, primerGen.IdAsignatura), out var aaExacta))
                    {
                        ordenCurricular = (aaExacta != null && aaExacta.OrdenCurricular > 0) ? aaExacta.OrdenCurricular : 999;
                    }
                    else if (ctx.AulaAsignaturasPorId != null &&
                             ctx.AulaAsignaturasPorId.TryGetValue(primerGen.IdAsignatura, out var aaDef))
                    {
                        ordenCurricular = (aaDef != null && aaDef.OrdenCurricular > 0) ? aaDef.OrdenCurricular : 999;
                    }

                    var prog = ctx.Programas?.FirstOrDefault(p => p.Id == primerGen.IdPrograma || (aula != null && p.Id == aula.IdPrograma));
                    double unidad = aula != null ? AgUnidadesSemanales.HorasSesionAula(aula, AgNivelAcademicoCatalogo.Desde(prog?.Modalidad)) : 4.0;
                    if (unidad <= 0) unidad = 4.0;

                    double totalHoras = asigGenes.Max(g => g.TotalHorasAsignatura);
                    if (totalHoras <= 0) totalHoras = asigGenes.Sum(g => g.HorasGen);
                    int sesionesRequeridas = (int)Math.Ceiling(totalHoras / unidad);

                    // Parseo de días
                    IReadOnlyList<int> dias = AgCarriles.Ventana(primerGen.DiasPatron);
                    if (dias.Count == 0)
                    {
                        dias = asigGenes.Select(g => g.DiaSemana).Distinct().OrderBy(d => d).ToList();
                    }
                    if (dias.Count == 0 && aula != null)
                    {
                        dias = AgCarriles.Ventana(aula.DiasVentana);
                    }
                    if (dias.Count == 0)
                    {
                        dias = new List<int> { 1, 2, 3, 4, 5, 6 };
                    }

                    // Fase 2141: Ranuras derivadas de los genes con fallback a catálogo/jornada
                    var ranurasGen = asigGenes
                        .Select(g => (g.DiaSemana, Franja: ctx.BloquesPorId != null && g.IdBloqueHorario > 0 && ctx.BloquesPorId.TryGetValue(g.IdBloqueHorario, out var b) ? b.Franja : null))
                        .Where(x => x.DiaSemana > 0)
                        .Select(x => new AgRanura(x.DiaSemana, x.Franja ?? (AgRanuras.FranjasDe(x.DiaSemana, aula?.Jornada, unidad).FirstOrDefault() ?? AgFranjaJornada.MANANA), unidad))
                        .GroupBy(r => (r.Dia, r.Franja)).Select(g => g.First()).ToList();

                    IReadOnlyList<AgRanura> ranuras = ranurasGen;
                    if (ranuras.Count == 0)
                    {
                        ranuras = AgRanuras.PorVentana(dias, ctx.Bloques, aula?.Jornada, unidad);
                    }
                    if (ranuras == null || ranuras.Count == 0)
                    {
                        ranuras = dias.Select(d => new AgRanura(d, AgRanuras.FranjasDe(d, aula?.Jornada, unidad).FirstOrDefault() ?? AgFranjaJornada.NOCHE, unidad)).ToList();
                    }

                    var diasEfectivos = ranuras.Select(r => r.Dia).Distinct().OrderBy(d => d).ToList();
                    if (diasEfectivos.Count == 0) diasEfectivos = dias.ToList();

                    asigs.Add(new AsignaturaCarril
                    {
                        IdAsignatura = gAsig.Key,
                        Genes = asigGenes,
                        Ini = asigGenes.Min(g => g.FechaInicioEstimada.Value),
                        Fin = asigGenes.Max(g => g.FechaFinEstimada.Value),
                        OrdenCurricular = ordenCurricular,
                        Dias = diasEfectivos,
                        Unidad = unidad,
                        TotalHoras = totalHoras,
                        SesionesRequeridas = sesionesRequeridas,
                        Ranuras = ranuras,
                        Slots = Math.Max(1, ranuras.Count)
                    });
                }

                if (asigs.Count <= 1) continue;

                var orden = asigs
                    .OrderBy(a => a.OrdenCurricular > 0 ? a.OrdenCurricular : 999)
                    .ThenBy(a => a.Ini)
                    .ThenBy(a => a.IdAsignatura)
                    .ToList();

                if (!HaySolape(orden)) continue;

                var diasCarril = asigs.SelectMany(a => a.Dias).Distinct().OrderBy(d => d).ToList();
                if (diasCarril.Count == 0 && aula != null)
                {
                    diasCarril = AgCarriles.Ventana(aula.DiasVentana).ToList();
                }
                if (diasCarril.Count == 0)
                {
                    diasCarril = new List<int> { 1, 2, 3, 4, 5, 6 };
                }

                var slices = ConstruirSemanasValidas(diasCarril, ctx.Ventana.FechaInicio, ctx.Ventana.FechaFin, ctx.EsFestivo);
                if (slices.Count == 0)
                {
                    res.Fallos++;
                    res.Detalle.Add($"cohorte={carrilGrp.Key.Cohorte} carril={carrilGrp.Key.Carril} motivo=I5_SIN_SEMANAS");
                    continue;
                }

                // 2) Fase de cálculo (no muta todavía)
                var acciones = new List<AccionGen>();
                var prevFin = orden[0].Fin;
                int cursor = 0;
                bool carrilOk = true;
                string motivo = string.Empty;

                for (int i = 1; i < orden.Count && carrilOk; i++)
                {
                    var actual = orden[i];

                    // Fase 2149: Avanzar cursor buscando un slice que tenga días útiles estrictamente posteriores a prevFin
                    while (cursor < slices.Count && !TieneDiaUtilPosterior(slices[cursor], diasCarril, prevFin, ctx.EsFestivo))
                    {
                        cursor++;
                    }
                    if (cursor >= slices.Count)
                    {
                        carrilOk = false;
                        motivo = $"NoCabe(cursor={cursor}/{slices.Count})";
                        break;
                    }

                    DateTime ini = slices[cursor].Inicio;
                    if (ini <= prevFin)
                    {
                        for (var d = slices[cursor].Inicio.Date; d <= slices[cursor].Fin.Date; d = d.AddDays(1))
                        {
                            int iso = d.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)d.DayOfWeek;
                            if (d > prevFin.Date && diasCarril.Contains(iso) && !ctx.EsFestivo(d))
                            {
                                ini = d;
                                break;
                            }
                        }
                    }

                    int k = Math.Max(1, (int)Math.Ceiling(actual.SesionesRequeridas / (double)actual.Slots));
                    DateTime fin = default(DateTime);
                    AgPatronSemanal patronExitoso = null;

                    while (true)
                    {
                        if (cursor + k > slices.Count)
                        {
                            carrilOk = false;
                            motivo = $"NoCabe(cursor={cursor}+k={k}/{slices.Count})";
                            break;
                        }
                        fin = slices[cursor + k - 1].Fin;
                        var resPatron = planificadorPatron.Construir(
                            actual.Dias, (int)actual.TotalHoras, ini, fin, ctx.EsFestivo, actual.Unidad);
                        if (resPatron.IsSuccess)
                        {
                            patronExitoso = resPatron.Value;
                            break;
                        }
                        k++;
                    }
                    if (!carrilOk) break;

                    var sesiones = AgPlanificadorRanuras.Materializar(
                        actual.IdAsignatura, actual.Ranuras, actual.SesionesRequeridas, ini, fin, ctx.EsFestivo);
                    if (sesiones == null || sesiones.Count < actual.SesionesRequeridas)
                    {
                        carrilOk = false;
                        motivo = $"Capacidad(sesiones={sesiones?.Count ?? 0}/{actual.SesionesRequeridas})";
                        break;
                    }

                    // Acciones por gen (Fase 2138: preservación de franja y sin colapso de sesiones en el mismo día)
                    var accionesAsig = new List<AccionGen>();
                    string patronStr = patronExitoso != null && patronExitoso.SesionesPorSemana != null
                        ? string.Join(",", patronExitoso.SesionesPorSemana)
                        : string.Empty;

                    bool sinRanura = false;
                    foreach (var gen in actual.Genes)
                    {
                        if (gen.DiaSemana <= 0)
                        {
                            sinRanura = true;
                            motivo = "DiaDesconocido";
                            break;
                        }

                        string franjaGen = ctx.BloquesPorId != null && gen.IdBloqueHorario > 0
                            && ctx.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out var bGen) ? bGen.Franja : null;

                        var fechas = sesiones
                            .Where(s => s.Ranura.Dia == gen.DiaSemana
                                     && (franjaGen == null || string.Equals(s.Ranura.Franja, franjaGen, StringComparison.OrdinalIgnoreCase)))
                            .Select(s => s.Fecha)
                            .Distinct()
                            .OrderBy(d => d)
                            .ToList();

                        int sesionesDia = sesiones.Count(s => s.Ranura.Dia == gen.DiaSemana
                            && (franjaGen == null || string.Equals(s.Ranura.Franja, franjaGen, StringComparison.OrdinalIgnoreCase)));
                        if (franjaGen == null && sesionesDia > fechas.Count)
                        {
                            sinRanura = true;
                            motivo = "FranjaAmbigua";
                            break;
                        }

                        if (fechas.Count == 0)
                        {
                            sinRanura = true;
                            break;
                        }

                        accionesAsig.Add(new AccionGen
                        {
                            Gen = gen,
                            FechaInicio = ini,
                            FechaFin = fin,
                            FechasGen = fechas,
                            HorasGen = fechas.Count * actual.Unidad,
                            PatronSemanal = patronStr
                        });
                    }

                    if (sinRanura)
                    {
                        carrilOk = false;
                        if (string.IsNullOrEmpty(motivo)) motivo = "SinRanura";
                        break;
                    }

                    // Fase 2139 / Fase 2158: Validación estricta de Invariantes I1 e I2 sin reparto residual arbitrario
                    bool validaInvariantes = true;
                    if (Math.Abs(actual.TotalHoras - actual.SesionesRequeridas * actual.Unidad) > 0.001)
                    {
                        validaInvariantes = false;
                        motivo = $"I2_DemandaNoMultiplo({actual.TotalHoras:0.#}h/{actual.Unidad:0.#}h)";
                    }

                    if (ini < ctx.Ventana.FechaInicio || fin > ctx.Ventana.FechaFin) validaInvariantes = false;
                    foreach (var a in accionesAsig)
                    {
                        if (a.FechasGen.Count != a.FechasGen.Distinct().Count())
                        {
                            validaInvariantes = false;
                            motivo = "FechasDuplicadas";
                            break;
                        }
                        if (Math.Abs(a.FechasGen.Count * actual.Unidad - a.HorasGen) > 0.001)
                        {
                            validaInvariantes = false;
                            motivo = "I2";
                            break;
                        }
                        if (a.FechasGen.Any(f => f < ctx.Ventana.FechaInicio || f > ctx.Ventana.FechaFin))
                        {
                            validaInvariantes = false;
                            break;
                        }
                        if (a.FechasGen.Any(f => ctx.EsFestivo(f)))
                        {
                            validaInvariantes = false;
                            break;
                        }
                        // Fase 2161: Validación estricta de día de semana en cada gen sin excepción para gen único
                        if (a.Gen.DiaSemana <= 0)
                        {
                            validaInvariantes = false;
                            motivo = "DiaDesconocido";
                            break;
                        }
                        if (a.FechasGen.Any(f =>
                        {
                            int diaIso = f.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)f.DayOfWeek;
                            return diaIso != a.Gen.DiaSemana;
                        }))
                        {
                            validaInvariantes = false;
                            motivo = "DiaInvalido";
                            break;
                        }
                    }
                    if (Math.Abs(accionesAsig.Sum(a => a.HorasGen) - actual.TotalHoras) > 0.001)
                    {
                        validaInvariantes = false;
                    }
                    if (accionesAsig.Sum(a => a.FechasGen.Count) != actual.SesionesRequeridas)
                    {
                        validaInvariantes = false;
                    }

                    if (!validaInvariantes)
                    {
                        carrilOk = false;
                        if (string.IsNullOrEmpty(motivo)) motivo = "I1/I2";
                        break;
                    }

                    acciones.AddRange(accionesAsig);
                    prevFin = fin;
                    cursor += k;
                }

                if (!carrilOk)
                {
                    res.Fallos++;
                    res.Detalle.Add($"cohorte={carrilGrp.Key.Cohorte} carril={carrilGrp.Key.Carril} motivo={motivo}");
                    continue;
                }

                // 3) Aplicar en el carril
                foreach (var a in acciones)
                {
                    a.Gen.FechaInicioEstimada = a.FechaInicio;
                    a.Gen.FechaFinEstimada = a.FechaFin;
                    a.Gen.FechasGen = a.FechasGen;
                    a.Gen.HorasGen = a.HorasGen;
                    a.Gen.PatronSemanal = a.PatronSemanal;
                }
                res.CarrilesReparados++;
                res.AsignaturasDesplazadas += Math.Max(0, orden.Count - 1);
            }

            res.HuboCambios = res.CarrilesReparados > 0;
            return res;
        }

        private static bool TieneDiaUtilPosterior((DateTime Inicio, DateTime Fin) s, IReadOnlyList<int> dias, DateTime prevFin, Func<DateTime, bool> festivo)
        {
            var diasSet = new HashSet<int>(dias);
            for (var d = s.Inicio.Date; d <= s.Fin.Date; d = d.AddDays(1))
            {
                if (d > prevFin.Date)
                {
                    int iso = d.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)d.DayOfWeek;
                    if (diasSet.Contains(iso) && (festivo == null || !festivo(d)))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private static bool HaySolape(IReadOnlyList<AsignaturaCarril> orden)
        {
            if (orden == null || orden.Count <= 1) return false;

            for (int idx = 1; idx < orden.Count; idx++)
            {
                if (orden[idx].Ini < orden[idx - 1].Fin) return true;
            }

            for (int i = 0; i < orden.Count; i++)
            {
                for (int j = i + 1; j < orden.Count; j++)
                {
                    if (orden[i].Ini <= orden[j].Fin && orden[j].Ini <= orden[i].Fin)
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
