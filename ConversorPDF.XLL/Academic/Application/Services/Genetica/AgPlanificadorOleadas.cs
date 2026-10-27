using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Servicio determinista de planificación por carriles disjuntos y oleadas secuenciales (Fase 1057 → Fase 1279).
    /// Reparte las asignaturas entre los carriles de AgCarriles (2 ranuras disjuntas por carril) y las secuencia
    /// dentro de cada carril con un cursor temporal: cada carril tiene como máximo una asignatura activa por semana.
    /// "Carril" en AgPlanOleada transporta el índice de carril; "Oleada" es el orden secuencial dentro del carril.
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

        /// <summary>
        /// Construye el plan de carril y oleada para una cohorte/aula (Fases 1057, 1088, 1092, 1095, 1226).
        /// Orden canónico de asignaturas: (OrdenCurricular, IdAsignatura).
        /// </summary>
        public Result<AgPlanOleadasResultado> Construir(
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
            for (var c = fechaInicio.Date; c <= fechaFin.Date; c = c.AddDays(7))
            {
                bool hayDiaVentana = false;
                for (int i = 0; i < 7 && !hayDiaVentana; i++)
                {
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
                    var finSem = c.AddDays(6) > fechaFin.Date ? fechaFin.Date : c.AddDays(6);
                    semanasValidas.Add((c, finSem));
                }
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

            // 3. Fase 1279: carriles disjuntos + oleadas secuenciales dentro del carril.
            //    Sustituye la partición en bandas paralelas de F1226, que asignaba a todas las asignaturas
            //    la misma sub-ventana (iniBanda = fechaInicio y Oleada = 0), provocando apilamiento y R5/R13 masivos.
            int carriles = AgCarriles.Cantidad(ranuras);
            var ranurasCarril = new List<IReadOnlyList<AgRanura>>(carriles);
            var diasCarril = new List<IReadOnlyList<int>>(carriles);
            var slotsCarril = new int[carriles];
            var capacidadCarril = new int[carriles];
            for (int c = 0; c < carriles; c++)
            {
                var rc = AgCarriles.DelCarril(ranuras, c);
                ranurasCarril.Add(rc);
                diasCarril.Add(rc.Select(r => r.Dia).Distinct().ToList());
                slotsCarril[c] = Math.Max(1, rc.Count);
                // Fase 1285: capacidad efectiva por ranura (festivos descontados sin doble conteo en días multirranura)
                capacidadCarril[c] = rc.Sum(r =>
                    AgPlanificadorTemporal.ContarOcurrencias(r.Dia, fechaInicio.Date, fechaFin.Date, esFestivo));
            }

            var patrones = ordenadas
                .Select(a => diasPatronPorAsignatura != null
                             && diasPatronPorAsignatura.TryGetValue(a.IdAsignatura, out var p)
                             && !string.IsNullOrWhiteSpace(p) ? p : string.Empty)
                .ToList();

            // 3.1 Pasada 1: el patrón sembrado (dias_patron) fija el carril cuando identifica exactamente uno.
            var carrilDe = new int[ordenadas.Count];
            for (int i = 0; i < carrilDe.Length; i++) carrilDe[i] = -1;

            var carga = new int[carriles];
            int renormalizados = 0;
            for (int i = 0; i < ordenadas.Count; i++)
            {
                var diasPatron = AgCarriles.Ventana(patrones[i]);
                if (diasPatron.Count == 0) continue;

                int mejor = -1;
                int mejorTam = int.MaxValue;
                for (int c = 0; c < carriles; c++)
                {
                    if (diasPatron.Any(d => !diasCarril[c].Contains(d))) continue;
                    if (diasCarril[c].Count < mejorTam)   // F1279: carril más específico ante ambigüedad SÁB ⊆ {JUE,VIE,SÁB}
                    {
                        mejor = c;
                        mejorTam = diasCarril[c].Count;
                    }
                }

                if (mejor >= 0 && carga[mejor] + demanda[i] <= capacidadCarril[mejor])
                {
                    carrilDe[i] = mejor;
                    carga[mejor] += demanda[i];
                }
                else
                {
                    renormalizados++;   // patrón no compatible con la capacidad del carril: se resuelve por carga
                }
            }

            // 3.2 Pasada 2: reparto greedy por utilización relativa (equilibra makespan entre carriles)
            for (int i = 0; i < ordenadas.Count; i++)
            {
                if (carrilDe[i] >= 0) continue;

                int elegido = -1;
                double mejorScore = double.MaxValue;
                for (int c = 0; c < carriles; c++)
                {
                    int proyectada = carga[c] + demanda[i];
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

                if (elegido < 0) elegido = 0;
                carrilDe[i] = elegido;
                carga[elegido] += demanda[i];
            }

            // 3.3 Pasada 3: cursor secuencial por carril (la oleada real)
            var cursor = new int[carriles];
            var oleada = new int[carriles];
            var planes = new List<AgPlanOleada>(ordenadas.Count);
            var infactibles = new List<int>();
            var detalleInfactibles = new List<string>();

            for (int i = 0; i < ordenadas.Count; i++)
            {
                var asig = ordenadas[i];
                int c = carrilDe[i];
                int SRequeridas = demanda[i];
                int kBase = Math.Max(1, (int)Math.Ceiling(SRequeridas / (double)slotsCarril[c]));

                DateTime ini = fechaInicio.Date;
                DateTime fin = fechaFin.Date;
                int k = kBase;
                Result<AgPatronSemanal> resPatron = Result.Failure<AgPatronSemanal>("sin semanas disponibles");

                for (int extra = 0; ; extra++)
                {
                    k = kBase + extra;
                    if (cursor[c] + k > W) break;
                    ini = semanasValidas[cursor[c]].Inicio;
                    fin = semanasValidas[cursor[c] + k - 1].Fin;
                    resPatron = _planificadorPatron.Construir(
                        ranurasCarril[c], (int)Math.Ceiling(asig.Horas), ini, fin, esFestivo, unidad);
                    if (resPatron.IsSuccess) break;
                }

                string accion = AgEscaleraVentanas.PuedeEscalar(diasVentana)
                    ? $" Escale la ventana a '{AgEscaleraVentanas.Siguiente(diasVentana)}' (AgEscaleraVentanas)."
                    : string.Empty;

                // R8: el patrón sembrado no cabe en el carril asignado (F1256, mismo contrato que v2)
                var diasPatronAsig = AgCarriles.Ventana(patrones[i]);
                if (diasPatronAsig.Count > 0 && diasPatronAsig.Any(d => !diasCarril[c].Contains(d)))
                {
                    infactibles.Add(asig.IdAsignatura);
                    detalleInfactibles.Add(
                        $"Asig {asig.IdAsignatura}: patrón '{patrones[i]}' no cabe en carril {c} " +
                        $"[{string.Join(",", diasCarril[c])}] (R8_DiaFueraDePatron).{accion}");
                    continue;
                }

                if (resPatron.IsFailure)
                {
                    infactibles.Add(asig.IdAsignatura);
                    detalleInfactibles.Add(
                        $"Asig {asig.IdAsignatura}: {asig.Horas:0.#}h = {SRequeridas} sesiones; carril {c} " +
                        $"[{string.Join(",", diasCarril[c])}] desde semana {cursor[c] + 1}/{W}; " +
                        $"{Math.Max(0, capacidadCarril[c] - cursor[c] * slotsCarril[c])} sesiones libres en el carril; " +
                        $"{resPatron.Error}{accion}");
                    continue;
                }

                int S = resPatron.Value.SesionesPorSemana != null ? resPatron.Value.SesionesPorSemana.Sum() : SRequeridas;

                planes.Add(new AgPlanOleada(
                    asig.IdAsignatura,
                    c,
                    oleada[c],
                    new List<int>(diasCarril[c]),
                    ini,
                    fin,
                    k,
                    S > 0 ? S : SRequeridas,
                    slotsCarril[c] * k,
                    ranurasCarril[c]));

                cursor[c] += k;
                oleada[c]++;
            }

            string ventanaSugerida = infactibles.Count > 0 ? AgEscaleraVentanas.Siguiente(diasVentana) : string.Empty;
            double ocupacion = capacidadTotal > 0
                ? Math.Round(100.0 * demandaTotal / capacidadTotal, 1)
                : 0.0;

            return Result.Success(new AgPlanOleadasResultado(
                planes,
                infactibles,
                detalleInfactibles,
                diasCarril,
                ventanaSugerida,
                renormalizados,
                ocupacion));
        }
    }
}
