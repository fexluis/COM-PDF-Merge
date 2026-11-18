using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Helper estático para garantizar la coherencia de docente y vacante en todos los genes
    /// que pertenecen al mismo grupo de asignatura (ClaveGrupo) a lo largo del cromosoma.
    /// </summary>
    public static class AgCoherenciaGrupo
    {
        // Fase 577 & Fase 585: Deduplicación en bucle caliente delimitada por IdCorrida para evitar miles de escrituras I/O repetidas
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, byte> _incidenciasCoherencia =
            new System.Collections.Concurrent.ConcurrentDictionary<string, byte>();

        // Fase 585: Helper de deduplicación con alcance de corrida
        private static bool TryIncidencia(AgContexto ctx, string clave)
            => _incidenciasCoherencia.TryAdd($"{ctx?.IdCorrida ?? 0}|{clave}", 0);

        /// <summary>
        /// Reinicia las claves de deduplicación de incidencias al iniciar cada corrida (Fase 852).
        /// </summary>
        public static void ReiniciarIncidencias() => _incidenciasCoherencia.Clear();

        // Fase 585 / Fase 590: Cálculo de capacidad real de sesiones = ocurrencias del día propio del gen en su ventana
        private static int SesionesMaximas(AgGen g, AgContexto ctx)
        {
            DateTime c0 = (g.FechaInicioEstimada ?? (ctx?.Ventana != null ? ctx.Ventana.FechaInicio : DateTime.Today)).Date;
            DateTime t0 = (g.FechaFinEstimada ?? (ctx?.Ventana != null ? ctx.Ventana.FechaFin : c0.AddMonths(4))).Date;
            if (t0 < c0) return 0;
            int diaIso = (g.DiaSemana >= 1 && g.DiaSemana <= 7) ? g.DiaSemana : 1;
            Func<DateTime, bool> esFestivo = null;
            if (ctx != null)
            {
                esFestivo = ctx.EsFestivo;
            }
            return AgPlanificadorTemporal.ContarOcurrencias(diaIso, c0, t0, esFestivo);
        }

        /// <summary>
        /// Sincroniza la lista de fechas discretas (FechasGen) con las sesiones requeridas por HorasGen / unidad,
        /// controlando colisiones por aula, fecha e intervalo horario e informando insuficiencia explícita sin reducir demanda.
        /// </summary>
        private static void SincronizarFechasGen(
            AgGen g, double unidad, int targetSesiones, AgContexto ctx, IReadOnlyCollection<AgGen> genes)
        {
            if (g == null || g.FechasGen == null) { return; }
            g.HorasGen = Math.Max(0, targetSesiones) * unidad;
            DateTime inicio = (g.FechaInicioEstimada ?? ctx?.Ventana?.FechaInicio ?? DateTime.Today).Date;
            DateTime fin = (g.FechaFinEstimada ?? ctx?.Ventana?.FechaFin ?? inicio.AddMonths(4)).Date;
            if (ctx?.Ventana != null)
            {
                inicio = inicio < ctx.Ventana.FechaInicio.Date ? ctx.Ventana.FechaInicio.Date : inicio;
                fin = fin > ctx.Ventana.FechaFin.Date ? ctx.Ventana.FechaFin.Date : fin;
            }
            int dia = g.DiaSemana >= 1 && g.DiaSemana <= 7 ? g.DiaSemana : 1;
            AgBloqueHorario bloque;
            if (ctx?.BloquesPorId != null && ctx.BloquesPorId.TryGetValue(g.IdBloqueHorario, out bloque))
            {
                dia = bloque.DiaSemanaNumero;
            }
            DayOfWeek dow = dia == 7 ? DayOfWeek.Sunday : (DayOfWeek)dia;
            Func<DateTime, bool> disponible = fecha => fecha >= inicio && fecha <= fin
                && fecha.DayOfWeek == dow && (ctx == null || !ctx.EsFestivo(fecha))
                && !genes.Any(otro => !ReferenceEquals(otro, g) && g.IdAula > 0 && otro.IdAula == g.IdAula
                    && otro.HorasGen > 0 && otro.FechasGen != null
                    && otro.FechasGen.Any(f => f.Date == fecha)
                    && BloquesSeSolapan(g, otro, ctx));

            // Conservar fechas válidas; buscar también huecos anteriores a la última sesión.
            g.FechasGen = g.FechasGen.Select(f => f.Date).Distinct().OrderBy(f => f)
                .Where(disponible).Take(Math.Max(0, targetSesiones)).ToList();
            for (DateTime fecha = inicio; fecha <= fin && g.FechasGen.Count < targetSesiones;)
            {
                if (!g.FechasGen.Contains(fecha) && disponible(fecha)) { g.FechasGen.Add(fecha); }
                if (fecha == fin) { break; }
                fecha = fecha.AddDays(1);
            }
            g.FechasGen.Sort();
            if (g.FechasGen.Count < targetSesiones)
            {
                g.Estado = AgEstadoGen.Incompleto;
                g.EsVacante = 1;
                if (string.IsNullOrEmpty(g.MotivoVacante))
                {
                    g.MotivoVacante = AgMotivoVacante.SIN_PLAN_OLEADA.Codigo();
                }
                if (ctx?.IdTerceroVacante > 0) { g.IdTercero = ctx.IdTerceroVacante; }
                LoggerHelper.WarningLimitado($"InsuficienciaFechas|{ctx?.IdCorrida}|{g.IdAula}|{g.IdAsignatura}",
                    $"[Coherencia] Gen {g.Id}: aula {g.IdAula} asig {g.IdAsignatura} solo ubicó {g.FechasGen.Count}/{targetSesiones} sesiones; conserva {g.HorasGen:0.#} h pendientes de cobertura.");
            }
        }

        private static bool BloquesSeSolapan(AgGen primero, AgGen segundo, AgContexto ctx)
        {
            if (ctx?.BloquesPorId == null
                || !ctx.BloquesPorId.TryGetValue(primero.IdBloqueHorario, out var a)
                || !ctx.BloquesPorId.TryGetValue(segundo.IdBloqueHorario, out var b)
                || !a.EsHorarioValido
                || !b.EsHorarioValido)
            {
                return true;
            }
            return a.TimeSpanInicio < b.TimeSpanFin && b.TimeSpanInicio < a.TimeSpanFin;
        }

        /// <summary>
        /// Unifica el docente y estado de vacante por ClaveGrupo (utilizando la moda de los docentes reales asignados,
        /// o priorizando genes con estado Fijado) en todo el cromosoma del individuo.
        /// </summary>
        /// <param name="ind">Grupos que se van a reparar.</param>
        /// <param name="ctx">Calendario y bloques vigentes.</param>
        /// <param name="ocupacion">Genes de todo el individuo si se repara solo un grupo; por defecto se usa ind.</param>
        public static void Sellar(AgIndividuo ind, AgContexto ctx, IReadOnlyCollection<AgGen> ocupacion = null)
        {
            if (ind?.Genes == null) return;
            var genesOcupacion = ocupacion ?? ind.Genes;

            long t0 = 0;
            if (AgTelemetriaSubetapas.Activo)
            {
                AgTelemetriaSubetapas.ConteoSellados++;
                t0 = System.Diagnostics.Stopwatch.GetTimestamp();
            }

            try
            {
                foreach (var grp in ind.Genes.Where(g => ctx?.ManifiestoCobertura == null || ctx.PuedeGenerarAula(g.IdAula))
                    .GroupBy(g => string.IsNullOrWhiteSpace(g.ClaveGrupo)
                             ? $"{g.IdAsignatura}|{g.IdAula}|{g.Grupo}" : g.ClaveGrupo))
                {
                var genes = grp.ToList();

                // Fase 1114 / Fase 1118: invariante de duración por nivel (-1 = no resoluble, no asumir pregrado)
                double UnidadNivel(AgGen g)
                {
                    if (g.IdAula > 0)
                    {
                        if (ctx?.UnidadHorasPorAulaId != null && ctx.UnidadHorasPorAulaId.TryGetValue(g.IdAula, out double u))
                        {
                            return u;
                        }
                        if (ctx?.AulasPorId != null && ctx.AulasPorId.TryGetValue(g.IdAula, out var aN))
                        {
                            Programa progN = null;
                            if (ctx.ProgramasPorId != null && ctx.ProgramasPorId.TryGetValue(aN.IdPrograma, out var pN))
                            {
                                progN = pN;
                            }
                            else
                            {
                                progN = ctx.Programas?.FirstOrDefault(p => p.Id == aN.IdPrograma);
                            }
                            return AgUnidadesSemanales.HorasSesionAula(aN, AgNivelAcademicoCatalogo.Desde(progN?.Modalidad));
                        }
                    }
                    return -1.0;
                }

                foreach (var g in genes.Where(x => x.Estado != AgEstadoGen.Fijado))
                {
                    double uNivel = UnidadNivel(g);
                    if (uNivel <= 0) continue;                        // nivel no verificable (F1118)
                    if (g.HorasSemanales <= 0) continue;              // duración desconocida: no es violación I6 (F1117)
                    if (g.HorasSemanales <= uNivel + 0.001) continue; // sólo corregir SOBRE-duración (8h→4h) (F1117)

                    g.HorasSemanales = uNivel;
                    if (!g.EsFijoTemporal)
                    {
                        g.FechasGen?.Clear();
                    }

                    // Fase 1121 / Fase 1125 / Fase 1127 / Fase 1135 / Fase 1160: selección unificada de bloque con preferencia por franja de la ranura
                    AgAula aJ = null;
                    if (g.IdAula > 0 && ctx?.AulasPorId != null) ctx.AulasPorId.TryGetValue(g.IdAula, out aJ);
                    string franjaGen = null;
                    if (ctx?.BloquesPorId != null && g.IdBloqueHorario > 0
                        && ctx.BloquesPorId.TryGetValue(g.IdBloqueHorario, out var bGen))
                    {
                        franjaGen = bGen.Franja;
                    }

                    var bloqueNivel = AgSeleccionBloque.ParaDia(aJ, g.DiaSemana, uNivel, ctx, franjaGen);

                    if (bloqueNivel != null)
                    {
                        g.AsignarBloque(bloqueNivel);          // reasigna franja del mismo día con duración correcta
                    }
                    else
                    {
                        g.IdBloqueHorario = 0;
                        g.Estado = AgEstadoGen.Incompleto;
                        g.EsVacante = 1;
                        if (ctx?.IdTerceroVacante > 0) g.IdTercero = ctx.IdTerceroVacante;
                        g.MotivoVacante = AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo();
                        LoggerHelper.WarningLimitado(
                            $"NivelDuracionSinBloque|{ctx?.IdCorrida}|{g.IdAula}|{g.DiaSemana}|{uNivel}",
                            $"[Coherencia] Gen {g.Id}: sin bloque de {uNivel}h para día {g.DiaSemana}; cuarentena (SIN_BLOQUE_NIVEL).");
                    }
                }

                // Fase 1294, Fase 1301, Fase 1307 & Fase 1308: En sábado (día 6), desdoblar franjas duplicadas asegurando unicidad
                int genesSabadoTotal = genes.Count(g => g.DiaSemana == 6);
                var genesSabado = genes
                    .Where(g => g.DiaSemana == 6 && g.Estado != AgEstadoGen.Fijado)
                    .OrderBy(g => g.IndiceDiaPatron)
                    .ThenBy(g => g.Id)
                    .ToList();

                if (genesSabadoTotal > 1 && genesSabado.Count > 0)
                {
                    // Fase 1308: resolver franja con fallback al catálogo; null = desconocida (no mutar a ciegas)
                    string FranjaResuelta(AgGen gg)
                    {
                        if (gg.IdBloqueHorario <= 0) return null;
                        if (ctx?.BloquesPorId != null && ctx.BloquesPorId.TryGetValue(gg.IdBloqueHorario, out var bb) && bb != null)
                            return bb.Franja ?? string.Empty;
                        return ctx?.Bloques?.FirstOrDefault(x => x.Id == gg.IdBloqueHorario)?.Franja;
                    }

                    var franjasOcupadas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    // Fase 1305: reservar franjas de genes bloqueados para no reasignar sobre ellos
                    foreach (var gf in genes.Where(g => g.DiaSemana == 6 && g.Estado == AgEstadoGen.Fijado))
                    {
                        var fr = FranjaResuelta(gf);
                        if (fr != null) franjasOcupadas.Add(fr);
                    }

                    foreach (var g in genesSabado)
                    {
                        string f = FranjaResuelta(g);
                        if (f == null) continue;                 // bloque no resoluble: no mutar a ciegas
                        if (franjasOcupadas.Add(f)) continue;    // franja libre

                        AgAula aSab = null;
                        if (g.IdAula > 0 && ctx?.AulasPorId != null) ctx.AulasPorId.TryGetValue(g.IdAula, out aSab);
                        double duracion = g.HorasSemanales > 0 ? g.HorasSemanales : 4.0;
                        string alterna = AgRanuras.FranjasDe(6, aSab?.Jornada, duracion)
                            .FirstOrDefault(x => !franjasOcupadas.Contains(x));

                        if (string.IsNullOrEmpty(alterna))
                        {
                            LoggerHelper.WarningLimitado($"SabadoSinFranjaLibre|{ctx?.IdCorrida}|{g.ClaveGrupo}|{duracion}",
                                $"[Coherencia] Gen {g.Id}: día 6 sin franja libre ({franjasOcupadas.Count} ocupada(s)); se mantiene franja actual.");
                            continue;
                        }

                        var bloqueAlt = AgSeleccionBloque.ParaDia(aSab, 6, duracion, ctx, alterna);
                        // Fase 1312: ParaDia degrada a lista[0] si la franja pedida no existe en catálogo -> validar
                        if (bloqueAlt != null && string.Equals(bloqueAlt.Franja, alterna, StringComparison.OrdinalIgnoreCase))
                        {
                            g.AsignarBloque(bloqueAlt);
                            franjasOcupadas.Add(bloqueAlt.Franja);
                        }
                        else
                        {
                            LoggerHelper.WarningLimitado($"SabadoSinAlterna|{ctx?.IdCorrida}|{g.ClaveGrupo}|{duracion}|{alterna}",
                                $"[Coherencia] Gen {g.Id}: catálogo sin bloque día 6 franja '{alterna}' ({duracion}h); franja duplicada persiste.");
                        }
                    }
                }

                // Fase 562, 563, 566 & Fase 1802: Reparto canónico y determinista de horas por grupo.
                // Si el grupo proviene de un plan fechado (FechasGen pobladas) y la suma ya cubre el total esperado,
                // no se redistribuye aunque algún gen tenga HorasGen <= 0.0 (sesiones concentradas en ranuras específicas).
                // La redistribución uniforme solo se ejecuta si la suma difiere del esperado, o en genes heredados sin plan fechado.
                double esperado = genes.Max(g => g.TotalHorasAsignatura);
                if (esperado > 0 && genes.Count > 0)
                {
                    double suma = genes.Sum(g => g.HorasGen);
                    bool hayPlanFechado = genes.Any(g => g.FechasGen != null && g.FechasGen.Count > 0);
                    bool hayPendientes = !hayPlanFechado && genes.Any(g => g.Estado != AgEstadoGen.Fijado && g.HorasGen <= 0.0);
                    if (Math.Abs(suma - esperado) > 0.001 || hayPendientes)
                    {
                        var fijados = genes.Where(g => g.Estado == AgEstadoGen.Fijado).ToList();
                        var ajustables = genes.Where(g => g.Estado != AgEstadoGen.Fijado)
                                              .OrderBy(g => g.IndiceDiaPatron)
                                              .ToList();

                        double horasFijas = fijados.Sum(g => g.HorasGen);
                        double horasRestantes = Math.Max(0.0, esperado - horasFijas);

                        // Fase 968: unidad canónica del grupo (4 h pregrado / 8 h posgrado)
                        double ObtenerUnidad(AgGen g)
                        {
                            if (g.IdAula > 0 && ctx != null && ctx.AulasPorId != null
                                && ctx.AulasPorId.TryGetValue(g.IdAula, out var a))
                            {
                                var prog = ctx.Programas?.FirstOrDefault(p => p.Id == a.IdPrograma);
                                return AgUnidadesSemanales.HorasSesionAula(a, AgNivelAcademicoCatalogo.Desde(prog?.Modalidad));
                            }
                            return 4.0;
                        }

                        double unidad = genes.Select(ObtenerUnidad).GroupBy(u => u).OrderByDescending(x => x.Count()).First().Key;

                        // Fase 972 / 974: degradar a 4 h si ni el total ni el remanente son múltiplos de la unidad (D4': 4 h siempre es divisible)
                        if (Math.Abs(esperado / unidad - Math.Round(esperado / unidad)) > 0.001
                            || Math.Abs(horasRestantes / unidad - Math.Round(horasRestantes / unidad)) > 0.001)
                        {
                            unidad = 4.0;
                        }

                        int totalSesiones = (int)Math.Round(horasRestantes / unidad);

                        // Fase 579 & 585: advertir cuando el residuo fijado impide Σ == esperado (dedupe por corrida)
                        double sumaProyectada = horasFijas + totalSesiones * unidad;
                        if (Math.Abs(sumaProyectada - esperado) > 0.001 && TryIncidencia(ctx, $"{grp.Key}|fijado|{horasFijas}"))
                        {
                            LoggerHelper.Warning($"[Coherencia] Grupo '{grp.Key}': fijados={horasFijas} h no cuantizadas; Σ proyectada {sumaProyectada} ≠ {esperado}.");
                        }

                        if (ajustables.Count == 0)
                        {
                            // FASE 896 & 901 / FASE 968: Reescalar fijados si todos los genes son fijos y la suma difiere del esperado (I1)
                            // Cuantización dinámica según unidad del aula/programa (4h u 8h) evitando reducir a 0 horas.
                            if (fijados.Count > 0 && suma > 0)
                            {
                                double factor = esperado / suma;
                                foreach (var g in fijados)
                                {
                                    double u = ObtenerUnidad(g);
                                    double cuantizado = Math.Round(g.HorasGen * factor / u) * u;
                                    g.HorasGen = cuantizado < u ? u : cuantizado;   // nunca 0 si esperado > 0
                                    if (g.FechasGen != null && g.FechasGen.Count > 0)
                                    {
                                        int targetSes = (int)Math.Round(g.HorasGen / u);
                                        SincronizarFechasGen(g, u, targetSes, ctx, genesOcupacion);
                                    }
                                }

                                double delta = esperado - fijados.Sum(g => g.HorasGen);
                                if (Math.Abs(delta) > 0.001)
                                {
                                    var mayor = fijados.OrderByDescending(g => g.HorasGen).First();
                                    double u = ObtenerUnidad(mayor);
                                    double nuevo = Math.Round((mayor.HorasGen + delta) / u) * u;
                                    mayor.HorasGen = Math.Max(u, nuevo);
                                    if (mayor.FechasGen != null && mayor.FechasGen.Count > 0)
                                    {
                                        int targetSes = (int)Math.Round(mayor.HorasGen / u);
                                        SincronizarFechasGen(mayor, u, targetSes, ctx, genesOcupacion);
                                    }
                                }

                                LoggerHelper.Warning($"[Coherencia] Grupo '{grp.Key}': fijados reescalados a {esperado:0.#} h (antes {suma:0.#}) para satisfacer I1.");
                            }
                            else if (TryIncidencia(ctx, $"{grp.Key}|sinAjustables|{esperado}"))
                            {
                                LoggerHelper.Warning($"[Coherencia] Grupo '{grp.Key}': sin genes ajustables; Σ={suma} ≠ esperado={esperado}.");
                            }
                        }
                        else
                        {
                            if (totalSesiones < ajustables.Count)
                            {
                                // Fase 572, 574, 577 & 585: consolidar y trazar una sola vez por corrida
                                if (TryIncidencia(ctx, $"{grp.Key}|consolidar|{totalSesiones}"))
                                {
                                    LoggerHelper.Debug($"[Coherencia] Grupo '{grp.Key}': {totalSesiones} sesiones < {ajustables.Count} genes; consolidando carga.");
                                }
                                for (int i = 0; i < ajustables.Count; i++)
                                {
                                    ajustables[i].HorasGen = i < totalSesiones ? unidad : 0.0;
                                    if (ajustables[i].FechasGen != null && ajustables[i].FechasGen.Count > 0)
                                    {
                                        int targetSes = i < totalSesiones ? 1 : 0;
                                        SincronizarFechasGen(ajustables[i], unidad, targetSes, ctx, genesOcupacion);
                                    }
                                }
                            }
                            else
                            {
                                int baseSesiones = totalSesiones / ajustables.Count;
                                int extraSesiones = totalSesiones % ajustables.Count;

                                for (int i = 0; i < ajustables.Count; i++)
                                {
                                    var g = ajustables[i];
                                    int sesionesAsignadas = baseSesiones + (i < extraSesiones ? 1 : 0);
                                    g.HorasGen = sesionesAsignadas * unidad;

                                    // Si tiene FechasGen precalculadas, sincronizar cardinalidad e invariante cantidadFechas * unidad == HorasGen
                                    if (g.FechasGen != null && g.FechasGen.Count > 0)
                                    {
                                        SincronizarFechasGen(g, unidad, sesionesAsignadas, ctx, genesOcupacion);
                                    }
                                }
                            }

                            // Fase 585: Reparto multi-gen del residuo entre todos los genes con holgura en su ventana
                            double nuevaSuma = genes.Sum(g => g.HorasGen);
                            double deltaResidual = esperado - nuevaSuma;
                            if (Math.Abs(deltaResidual) > 0.001)
                            {
                                int deltaSes = (int)Math.Round(deltaResidual / unidad);
                                foreach (var g in ajustables
                                    .OrderByDescending(g => SesionesMaximas(g, ctx) - (int)Math.Round(g.HorasGen / unidad)))
                                {
                                    if (deltaSes == 0) break;
                                    int actual = (int)Math.Round(g.HorasGen / unidad);
                                    int objetivo = Math.Max(0, Math.Min(SesionesMaximas(g, ctx), actual + deltaSes));
                                    int aplicado = objetivo - actual;
                                    if (aplicado == 0) continue;
                                    g.HorasGen = objetivo * unidad;
                                    if (g.FechasGen != null && g.FechasGen.Count > 0)
                                    {
                                        SincronizarFechasGen(g, unidad, objetivo, ctx, genesOcupacion);
                                    }
                                    deltaSes -= aplicado;
                                }
                                if (deltaSes != 0 && TryIncidencia(ctx, $"{grp.Key}|residual|{deltaResidual}"))
                                {
                                    LoggerHelper.Warning($"[Coherencia] Grupo '{grp.Key}': residuo {deltaResidual:0.#} h no reconciliable (faltan {deltaSes} sesiones de capacidad).");
                                }
                            }
                        }
                    }
                }

                // Fase 524 & Fase 597: si hay genes fijados, el docente del gen fijado es canónico (dedupe por corrida)
                if (genes.Count(g => g.Estado == AgEstadoGen.Fijado) > 1 &&
                    TryIncidencia(ctx, $"{grp.Key}|multiFijado|{genes.Count(g => g.Estado == AgEstadoGen.Fijado)}"))
                {
                    LoggerHelper.Warning($"[Coherencia] Grupo '{grp.Key}' con múltiples genes fijados; se prioriza el primero por orden de patrón.");
                }

                var fijo = genes.FirstOrDefault(g => g.Estado == AgEstadoGen.Fijado);
                if (fijo != null)
                {
                    foreach (var g in genes.Where(g => g.Estado != AgEstadoGen.Fijado))
                    {
                        g.IdTercero = fijo.IdTercero;
                        g.EsVacante = fijo.EsVacante;
                        g.MotivoVacante = fijo.EsVacante == 1 ? fijo.MotivoVacante : string.Empty;
                    }
                    continue;
                }

                // Fase 1091 / Fase 1104: Preservar motivos no asignables (SIN_PLAN_OLEADA y SIN_BLOQUE_NIVEL) aun si algún gen del grupo tiene docente asignado (e.g. cruce mixto)
                var motivoNoAsignable = new[]
                {
                    AgMotivoVacante.SIN_PLAN_OLEADA.Codigo(),
                    AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo()
                };
                var genNoAsignable = genes.FirstOrDefault(g => motivoNoAsignable.Contains(g.MotivoVacante));
                if (genNoAsignable != null)
                {
                    foreach (var g in genes.Where(g => g.Estado != AgEstadoGen.Fijado))
                    {
                        g.EsVacante = 1;
                        g.IdTercero = ctx?.IdTerceroVacante ?? 0;
                        // Fase 1896-B: no propagar motivo ajeno a genes que ya tengan causa propia
                        if (string.IsNullOrEmpty(g.MotivoVacante))
                        {
                            g.MotivoVacante = genNoAsignable.MotivoVacante;
                        }
                    }
                    continue;
                }

                var reales = genes.Where(g => g.EsVacante == 0 && g.IdTercero > 0).ToList();
                int doc;
                int vac;
                string motivo;

                if (reales.Count == 0)
                {
                    doc = ctx?.IdTerceroVacante ?? 0;
                    vac = 1;
                    // Fase 1904: no dejar vacío el motivo de una vacante consolidada
                    motivo = genes.Select(g => g.MotivoVacante).FirstOrDefault(m => !string.IsNullOrWhiteSpace(m))
                          ?? AgMotivoVacante.SIN_DISPONIBILIDAD.Codigo();
                }
                else
                {
                    doc = reales.GroupBy(g => g.IdTercero).OrderByDescending(x => x.Count()).First().Key;
                    vac = 0;
                    motivo = string.Empty;
                }

                foreach (var g in genes.Where(g => g.Estado != AgEstadoGen.Fijado))
                {
                    g.IdTercero = doc;
                    g.EsVacante = vac;
                    g.MotivoVacante = motivo;
                }
            }
        }
        finally
        {
            if (AgTelemetriaSubetapas.Activo)
            {
                AgTelemetriaSubetapas.TiempoSelladoTicks += (System.Diagnostics.Stopwatch.GetTimestamp() - t0);
            }
        }
    }
    }
}
