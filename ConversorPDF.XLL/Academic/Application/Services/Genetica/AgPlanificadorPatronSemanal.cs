using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Planificador determinista de patrones semanales y distribución de sesiones multi-día (Fase 381 / Fase 413 / Fase 415).
    /// Deriva la distribución exacta de sesiones por semana y proyecta las fechas reales en el calendario desde FechaInicio
    /// considerando los días K declarados, semanas disponibles y omitiendo festivos, reproduciendo fielmente los
    /// histogramas empíricos observados en el Excel de referencia.
    /// </summary>
    public sealed class AgPlanificadorPatronSemanal
    {
        public const int TopeDiasSemana = 5;

        /// <summary>
        /// Fuente única canónica de días K (ISO 1=Lunes .. 7=Domingo) para una asignación de aula (Fase 439 / Fase 949 / Fase 1008).
        /// Si no está definida o está vacía, lee la ventana del aula (aula.DiasVentana). Si tampoco existe, retorna el default { 2, 3 }.
        /// </summary>
        /// <param name="aa">Configuración de aula-asignatura o null.</param>
        /// <param name="aula">Aula asociada o null.</param>
        /// <param name="carril">Índice de carril/asignatura para alternancia de ventanas multidía (Fase 1008).</param>
        /// <returns>Lista inmutable de días de la semana ISO válidos [1..7].</returns>
        public static IReadOnlyList<int> DiasKDe(AgAulaAsignatura aa, AgAula aula = null, int carril = 0)
        {
            if (aa != null && !string.IsNullOrWhiteSpace(aa.DiasPatron))
            {
                var dias = aa.DiasPatron.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(s => int.TryParse(s.Trim(), out int d) ? d : 0)
                                        .Where(d => d >= 1 && d <= 7)
                                        .Distinct()
                                        .ToList();
                if (dias.Count > 0)
                {
                    // Fase 1010 / Fase 1018 / Fase 1058: intersectar con la ventana del aula y normalizar si supera DiasPorCarril
                    var ventana = aula != null && !string.IsNullOrWhiteSpace(aula.DiasVentana) ? AgCarriles.Ventana(aula.DiasVentana) : null;
                    if (ventana != null && ventana.Count > 0)
                    {
                        // Fase 1058: si el patrón guardado tiene más de 2 días (legado V73), normalizar al carril del índice
                        if (dias.Count > AgCarriles.DiasPorCarril)
                        {
                            LoggerHelper.WarningLimitado(
                                $"PatronMultidiaNormalizado|{aa?.IdAula}|{aa?.IdAsignatura}|{aa.DiasPatron}",
                                $"[Patron] Aula {aa?.IdAula} asig {aa?.IdAsignatura}: patrón legado [{aa.DiasPatron}] con {dias.Count} días normalizado al carril de {AgCarriles.DiasPorCarril} días.");
                            return AgCarriles.DelCarril(ventana, carril);
                        }

                        // Fase 1074 / Fase 1156: aceptar si forma carril de 2 días consecutivos o 1 día proveniente de ranura (I2)
                        var dentro = dias.Where(d => ventana.Contains(d)).OrderBy(d => d).ToList();
                        bool esCarril = (dentro.Count == AgCarriles.DiasPorCarril && dentro.Zip(dentro.Skip(1), (a, b) => b - a).All(gap => gap == 1))
                                     || (dentro.Count == 1 && ventana.Contains(dentro[0]));
                        if (esCarril) return dentro;

                        LoggerHelper.WarningLimitado(
                            $"PatronNoEsCarril|{aa?.IdAula}|{aa?.IdAsignatura}|{aa.DiasPatron}",
                            $"[Patron] Aula {aa?.IdAula} asig {aa?.IdAsignatura}: patrón [{aa.DiasPatron}] no es carril disjunto de '{aula.DiasVentana}'; se usa carril {AgCarriles.Indice(carril, AgCarriles.Cantidad(ventana))}.");
                        return AgCarriles.DelCarril(ventana, carril);
                    }
                    return dias;
                }
            }

            if (aula != null && !string.IsNullOrWhiteSpace(aula.DiasVentana))
            {
                var diasVentana = AgCarriles.Ventana(aula.DiasVentana);
                if (diasVentana != null && diasVentana.Count > 0)
                {
                    // Fase 1020 / Fase 1058: carriles disjuntos delegados en AgCarriles
                    int carriles = AgCarriles.Cantidad(diasVentana);
                    if (Math.Abs(carril) >= carriles)
                    {
                        LoggerHelper.WarningLimitado(
                            $"CarrilExcedido|{aula?.Id}|{carril}|{carriles}",
                            $"[Patron] Aula {aula?.Id}: carril {carril} excede la capacidad ({carriles}) para ventana '{aula?.DiasVentana}'; se reutiliza el carril {AgCarriles.Indice(carril, carriles)} (riesgo R2/R4).");
                    }
                    if (diasVentana.Count % AgCarriles.DiasPorCarril != 0)
                    {
                        LoggerHelper.WarningLimitado(
                            $"VentanaImpar|{aula?.Id}|{aula?.DiasVentana}",
                            $"[Patron] Aula {aula?.Id}: ventana '{aula?.DiasVentana}' con {diasVentana.Count} días; {carriles} carril(es) por días; el último día se cubre por ranuras (modelo F1154).");
                    }
                    return AgCarriles.DelCarril(diasVentana, carril);
                }
            }

            // Fase 1093: Traza clara cuando un aula carece de ventana declarada (I1)
            LoggerHelper.WarningLimitado(
                $"PatronSinVentana|{aula?.Id}",
                "[Patron] Aula sin ventana declarada; se usa {2,3} sólo por compatibilidad (I1 no satisfecho).");
            return new List<int> { 2, 3 };
        }

        /// <summary>
        /// Resuelve la lista ordenada de ranuras asignadas al carril especificado (Fase 1156 / Fase 1159).
        /// Si existen bloques y ventana de aula, utiliza el modelo canónico AgRanuras.
        /// </summary>
        internal static IReadOnlyList<AgRanura> RanurasDe(
            AgAulaAsignatura aa,
            AgAula aula = null,
            int carril = 0,
            IReadOnlyCollection<AgBloqueHorario> bloques = null,
            double unidad = 4.0)
        {
            if (aula != null && !string.IsNullOrWhiteSpace(aula.DiasVentana))
            {
                var ranurasVentana = AgRanuras.PorVentana(aula.DiasVentana, bloques, aula.Jornada, unidad);
                if (ranurasVentana.Count > 0)
                {
                    return AgCarriles.DelCarril(ranurasVentana, carril);
                }
            }

            var dias = DiasKDe(aa, aula, carril);
            return dias.Select(d => new AgRanura(
                    d,
                    AgRanuras.FranjasDe(d, aula?.Jornada, unidad).FirstOrDefault() ?? AgFranjaJornada.NOCHE,
                    unidad))
                .ToList();
        }

        /// <summary>
        /// Proyecta las fechas calendario reales de una ranura dada su cantidad requerida de sesiones (Fase 1156).
        /// Omite festivos y toma estrictamente las ocurrencias del día ISO de la ranura.
        /// </summary>
        public static IReadOnlyList<DateTime> FechasDeRanura(
            int diaIso,
            int sesiones,
            DateTime fechaInicio,
            DateTime fechaFin,
            Func<DateTime, bool> esFestivo = null)
        {
            var fechas = new List<DateTime>();
            if (sesiones <= 0) return fechas;

            var cursor = fechaInicio.Date;
            while (cursor <= fechaFin.Date && fechas.Count < sesiones)
            {
                int d = cursor.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)cursor.DayOfWeek;
                if (d == diaIso && (esFestivo == null || !esFestivo(cursor)))
                {
                    fechas.Add(cursor);
                }
                if (cursor == fechaFin.Date) { break; }
                cursor = cursor.AddDays(1);
            }
            return fechas;
        }

        /// <summary>
        /// Fase 984 / Fase 1019 / Fase 1156: resuelve días K desde un patrón textual explícito ('1,3' / '4,5'), intersectado
        /// con la ventana del aula si se provee. Sin patrón cae al comportamiento canónico con carril.
        /// </summary>
        public static IReadOnlyList<int> DiasKDe(string diasPatron, string diasVentana = null, AgAula aula = null, int carril = 0)
        {
            if (!string.IsNullOrWhiteSpace(diasPatron))
            {
                var pedidos = diasPatron.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(s => int.TryParse(s.Trim(), out int d) ? d : 0)
                                        .Where(d => d >= 1 && d <= 7)
                                        .Distinct()
                                        .ToList();

                var ventana = !string.IsNullOrWhiteSpace(diasVentana)
                    ? AgDiasVentana.Dias(diasVentana)
                    : (aula != null && !string.IsNullOrWhiteSpace(aula.DiasVentana) ? AgDiasVentana.Dias(aula.DiasVentana) : null);

                if (ventana != null && ventana.Count > 0)
                {
                    // Fase 1084 / Fase 1156: aceptar si forma carril de 2 días consecutivos o 1 día proveniente de ranura (I2)
                    var dentro = pedidos.Where(d => ventana.Contains(d)).OrderBy(d => d).ToList();
                    bool esCarril = (dentro.Count == AgCarriles.DiasPorCarril && dentro.Zip(dentro.Skip(1), (a, b) => b - a).All(gap => gap == 1))
                                 || (dentro.Count == 1 && ventana.Contains(dentro[0]));
                    if (esCarril) return dentro;

                    LoggerHelper.WarningLimitado(
                        $"PatronTextoNoEsCarril|{diasPatron}|{aula?.DiasVentana}",
                        $"[Patron] Patrón textual [{diasPatron}] no es carril disjunto de '{aula?.DiasVentana}'; se resuelve carril sobre la ventana.");
                }
                else if (pedidos.Count > 0)
                {
                    return pedidos;
                }
            }

            return DiasKDe((AgAulaAsignatura)null, aula, carril);
        }

        /// <summary>
        /// Calcula el número de semanas disponibles entre fechaInicio y fechaFin que contienen al menos un día hábil en diasK (Fase 426).
        /// </summary>
        public static int SemanasDisponibles(
            IReadOnlyList<int> diasK,
            DateTime ini,
            DateTime fin,
            Func<DateTime, bool> esFestivo = null)
        {
            if (fin < ini || diasK == null || diasK.Count == 0) return 0;
            int semanas = 0;
            for (var c = ini.Date; c <= fin.Date; c = c.AddDays(7))
            {
                bool hay = false;
                for (int i = 0; i < 7 && !hay; i++)
                {
                    var d = c.AddDays(i);
                    if (d > fin.Date) break;
                    int iso = d.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)d.DayOfWeek;
                    if (diasK.Contains(iso) && (esFestivo == null || !esFestivo(d)))
                    {
                        hay = true;
                    }
                }
                if (hay) semanas++;
            }
            return semanas;
        }

        private static int Iso(DayOfWeek dow)
        {
            return dow == DayOfWeek.Sunday ? 7 : (int)dow;
        }

        /// <summary>
        /// Construye el patrón semanal recorriendo el calendario real desde fechaInicio hasta fechaFin.
        /// </summary>
        /// <param name="diasK">Días de la semana permitidos en convención ISO (1=Lunes .. 7=Domingo).</param>
        /// <param name="totalHoras">Total de horas requerido de la asignatura.</param>
        /// <param name="fechaInicio">Fecha inicial de la proyección.</param>
        /// <param name="fechaFin">Fecha máxima de la ventana.</param>
        /// <param name="esFestivo">Predicado de días festivos.</param>
        /// <param name="unidadSesion">Duración de sesión de la unidad curricular (4.0h pregrado, 8.0h posgrado).</param>
        /// <returns>Resultado con el patrón semanal o fallo si las sesiones no caben en la ventana.</returns>
        public Result<AgPatronSemanal> Construir(
            IReadOnlyList<int> diasK,
            int totalHoras,
            DateTime fechaInicio,
            DateTime fechaFin,
            Func<DateTime, bool> esFestivo = null,
            double unidadSesion = 4.0)
        {
            if (diasK == null || diasK.Count == 0)
            {
                return Result.Failure<AgPatronSemanal>("No se especificaron días de la semana para el patrón.");
            }

            // Fase 983 / D4': sesiones según la unidad real (4 h pregrado, 8 h posgrado)
            double unidad = unidadSesion > 0 ? unidadSesion : 4.0;
            int S = (int)Math.Ceiling(totalHoras / unidad);
            if (S <= 0)
            {
                return Result.Success(new AgPatronSemanal(diasK, new int[0], new DateTime[0]));
            }

            // Fase 427: Tope real acotado por los días declarados en K
            int nCap = Math.Min(diasK.Count, TopeDiasSemana);
            int W = SemanasDisponibles(diasK, fechaInicio, fechaFin, esFestivo);
            if (W == 0)
            {
                return Result.Failure<AgPatronSemanal>("Infactible: la ventana no contiene días hábiles para los días K declarados.");
            }

            // Fase 1141: Factibilidad en días hábiles reales (una semana con 1 día hábil no equivale a 2 ranuras)
            int capacidadDias = 0;
            for (int i = 0; i < diasK.Count; i++)
            {
                capacidadDias += AgPlanificadorTemporal.ContarOcurrencias(diasK[i], fechaInicio, fechaFin, esFestivo);
            }

            if (S > capacidadDias)
            {
                return Result.Failure<AgPatronSemanal>(
                    $"Infactible: {S} sesiones exceden {capacidadDias} días hábiles reales de [{string.Join(",", diasK)}] " +
                    $"entre {fechaInicio:yyyy-MM-dd} y {fechaFin:yyyy-MM-dd} (festivos descontados).");
            }

            int semanasMin = (int)Math.Ceiling(S / (double)nCap);
            if (semanasMin > W)
            {
                return Result.Failure<AgPatronSemanal>(
                    $"Infactible: {S} sesiones no caben en {W} semanas × {nCap} días.");
            }

            int semanas = Math.Min(W, Math.Max(semanasMin, (int)Math.Ceiling(S / (double)nCap)));

            int baseD = S / semanas;
            int resto = S % semanas;
            var porSemana = new List<int>();

            // Fase 425: Escalar SOLO si el patrón uniforme no cabe en W
            bool requiereEscalado = (int)Math.Ceiling(S / (double)Math.Max(1, baseD)) > W;

            if (!requiereEscalado)
            {
                // Preservar la distribución uniforme estándar observada en el 90% del Excel
                for (int w = 0; w < semanas; w++)
                {
                    porSemana.Add(w < semanas - resto ? baseD : baseD + 1);
                }
            }
            else
            {
                // Fase 424 / Fase 430: Reparto monótono creciente por redondeo acumulativo telescópico
                for (int w = 1; w <= semanas; w++)
                {
                    int acum = (int)Math.Ceiling(S * w / (double)semanas);
                    int prev = (int)Math.Ceiling(S * (w - 1) / (double)semanas);
                    porSemana.Add(acum - prev);
                }

                // Fase 430: Orden ascendente estable (preserva Sum() == S)
                porSemana.Sort();
                while (porSemana[0] < 1 || porSemana[porSemana.Count - 1] > nCap)
                {
                    break;
                }
            }

            if (porSemana.Any(v => v < 1 || v > nCap))
            {
                return Result.Failure<AgPatronSemanal>($"Infactible: vector fuera de rango [1..{nCap}].");
            }

            // Fase 423: Guarda de conservación estricta de suma total de sesiones (Invariante I1/I2)
            if (porSemana.Sum() != S)
            {
                return Result.Failure<AgPatronSemanal>(
                    $"Infactible: Σ semanas ({porSemana.Sum()}) ≠ sesiones requeridas ({S}).");
            }

            // Fase 427: Días efectivos son estricta y únicamente los de diasK
            var dias = new List<int>(diasK);

            // Generar fechas reales para cada semana w tomando porSemana[w] de los n días declarados
            var fechas = new List<DateTime>();
            var cursor = fechaInicio.Date;
            for (int w = 0; w < semanas; w++)
            {
                int usados = 0;
                while (cursor <= fechaFin.Date && usados < porSemana[w])
                {
                    if (dias.Contains(Iso(cursor.DayOfWeek)) && (esFestivo == null || !esFestivo(cursor)))
                    {
                        fechas.Add(cursor);
                        usados++;
                    }
                    cursor = cursor.AddDays(1);
                }

                // Fase 433: Diagnóstico claro con causa y acción correctiva ante festivos o insuficiencia de franjas
                if (usados < porSemana[w])
                {
                    return Result.Failure<AgPatronSemanal>(
                        $"Infactible: semana {w + 1} sólo {usados}/{porSemana[w]} sesiones (festivos o días K no disponibles). Ajuste ventana o agregue franjas.");
                }
            }

            return Result.Success(new AgPatronSemanal(dias, porSemana, fechas));
        }

        /// <summary>
        /// Construye el patrón semanal a partir de ranuras (Fase 1156 / Fase 1159 / Fase 1164).
        /// Distribuye las sesiones equitativamente entre las ranuras del carril
        /// y genera las fechas calendario permitiendo 2 sesiones el mismo sábado con franjas distintas.
        /// </summary>
        internal Result<AgPatronSemanal> Construir(
            IReadOnlyList<AgRanura> ranuras,
            int totalHoras,
            DateTime fechaInicio,
            DateTime fechaFin,
            Func<DateTime, bool> esFestivo = null,
            double unidadSesion = 4.0)
        {
            if (ranuras == null || ranuras.Count == 0)
            {
                return Result.Failure<AgPatronSemanal>("No se especificaron ranuras para el patrón.");
            }

            double unidad = unidadSesion > 0 ? unidadSesion : 4.0;
            int S = (int)Math.Ceiling(totalHoras / unidad);
            if (S <= 0)
            {
                return Result.Success(new AgPatronSemanal(ranuras.Select(r => r.Dia).ToList(), new int[0], new DateTime[0]));
            }

            int nRanuras = ranuras.Count;
            var fechasDisponibles = new List<DateTime>[nRanuras];
            int totalDisponibles = 0;
            for (int i = 0; i < nRanuras; i++)
            {
                fechasDisponibles[i] = FechasDeRanura(ranuras[i].Dia, 999, fechaInicio, fechaFin, esFestivo).ToList();
                totalDisponibles += fechasDisponibles[i].Count;
            }

            if (S > totalDisponibles)
            {
                return Result.Failure<AgPatronSemanal>(
                    $"Infactible: {S} sesiones exceden {totalDisponibles} días hábiles disponibles entre {fechaInicio:yyyy-MM-dd} y {fechaFin:yyyy-MM-dd}.");
            }

            // Distribuir sesiones adaptativamente considerando días hábiles efectivos de cada ranura
            int[] sesionesPorRanura = new int[nRanuras];
            int remanente = S;
            for (int i = 0; i < nRanuras; i++)
            {
                int restantes = nRanuras - i;
                int target = (int)Math.Ceiling(remanente / (double)restantes);
                int minCapRestantes = 0;
                for (int j = i + 1; j < nRanuras; j++) minCapRestantes += fechasDisponibles[j].Count;

                int asignadas = Math.Min(target, fechasDisponibles[i].Count);
                if (remanente - asignadas > minCapRestantes)
                {
                    asignadas = remanente - minCapRestantes;
                }
                sesionesPorRanura[i] = asignadas;
                remanente -= asignadas;
            }

            var todasFechas = new List<DateTime>();
            for (int i = 0; i < nRanuras; i++)
            {
                todasFechas.AddRange(fechasDisponibles[i].Take(sesionesPorRanura[i]));
            }

            todasFechas.Sort();
            var diasSemana = ranuras.Select(r => r.Dia).Distinct().ToList();

            // Fase 1164: Agrupar fechas por semana calendario para vector SesionesPorSemana fidedigno
            var porSemana = new List<int>();
            var ini = fechaInicio.Date;
            while (ini <= fechaFin.Date && porSemana.Count < 60)
            {
                int finDias = 6;
                var finSem = ini.AddDays(Math.Min(finDias, (fechaFin.Date - ini).Days));
                int c = todasFechas.Count(f => f >= ini && f <= finSem);
                if (c > 0) porSemana.Add(c);
                if ((fechaFin.Date - ini).Days < 7) { break; }
                ini = ini.AddDays(7);
            }

            return Result.Success(new AgPatronSemanal(diasSemana, porSemana, todasFechas));
        }
    }
}
