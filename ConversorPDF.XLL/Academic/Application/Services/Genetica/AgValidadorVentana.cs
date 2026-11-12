using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Validador pre-vuelo de factibilidad que comprueba la suficiencia de recursos y límites de la ventana semestral.
    /// </summary>
    public class AgValidadorVentana : IAgValidadorVentana
    {
        public Task<Result<InformeFactibilidad>> ValidarFactibilidadAsync(AgContexto contexto, CancellationToken ct)
        {
            return Task.FromResult(ValidarFactibilidad(contexto));
        }

        public Result<InformeFactibilidad> ValidarFactibilidad(AgContexto contexto)
        {
            if (contexto == null) return Result.Failure<InformeFactibilidad>("El contexto del AG no puede ser nulo.");

            var informe = new InformeFactibilidad();

            // 1. Validar presencia y validez de la ventana
            if (contexto.Ventana == null)
            {
                informe.AgregarBloqueo("No se ha definido la ventana temporal del semestre (α).");
                return Result.Success(informe);
            }

            if (contexto.Ventana.FechaFin < contexto.Ventana.FechaInicio)
            {
                informe.AgregarBloqueo($"La fecha final de la ventana ({contexto.Ventana.FechaFin:yyyy-MM-dd}) es anterior a la fecha inicial ({contexto.Ventana.FechaInicio:yyyy-MM-dd}).");
                return Result.Success(informe);
            }

            // 2. Validar disponibilidad de aulas y bloques
            var aulasDisponibles = contexto.Aulas.Where(a => a.Estado == AgEstadoAula.Disponible).ToList();
            if (aulasDisponibles.Count == 0)
            {
                informe.AgregarBloqueo("No hay aulas disponibles configuradas para la sede/subunidad.");
            }

            var bloquesActivos = contexto.Bloques.Where(b => b.Activo == 1).ToList();
            if (bloquesActivos.Count == 0)
            {
                informe.AgregarBloqueo("No hay bloques horarios activos configurados.");
            }

            // 3. Validar demanda vs oferta (Fase 380 / Fase 396 / Fase 401)
            double horasRequeridas = contexto.Asignaturas.Sum(a => a.Horas);
            informe.TotalHorasRequeridas = horasRequeridas;
            informe.TotalCursosProgramar = contexto.Asignaturas.Count;
            informe.TotalDocentesDisponibles = contexto.Docentes.Count;

            var ini = contexto.Ventana.FechaInicio.Date;
            var fin = contexto.Ventana.FechaFin.Date;

            // Fase 418: Mismo alcance para oferta que para la demanda evaluada (aulasDisponibles para semestre/grupo)
            var aulasOferta = aulasDisponibles;

            double ofertaSemestre = aulasOferta.Sum(_ => bloquesActivos.Sum(b =>
                b.DuracionHoras * AgPlanificadorTemporal.ContarOcurrencias(b.DiaSemana, ini, fin, contexto.EsFestivo)));
            informe.TotalCapacidadHorariaAulas = ofertaSemestre;

            if (horasRequeridas > ofertaSemestre)
            {
                informe.AgregarBloqueo($"Capacidad insuficiente: Se requieren {horasRequeridas:0.#} horas pero la oferta real de aulas por ocurrencias en la ventana es de {ofertaSemestre:0.#} horas.");
            }

            // Factibilidad por asignación individual (Fase 380-B / Fase 419 / Fase 434 / Fase 439 / Fase 586 / Fase 591)
            foreach (var asig in contexto.Asignaturas)
            {
                AgAulaAsignatura aa = null;
                if (contexto.AulaAsignaturasPorId != null)
                {
                    contexto.AulaAsignaturasPorId.TryGetValue(asig.Id, out aa);
                }

                // Fase 439 / Fase 608: los días K se resuelven exclusivamente dentro de CapacidadDe(configs|aa)
                // (variable local diasKAsig retirada: ya no gobierna la factibilidad)
                // Fase 602: la cobertura de ventana se resuelve vía capacidadReal multi-aula (se retira early-continue mono-aula)

                double hAsig = (aa != null && aa.HorasObjetivo > 0)
                    ? aa.HorasObjetivo
                    : (asig.IntHoraria > 0 ? (double)asig.IntHoraria : (asig.Horas > 0 ? (double)asig.Horas : (asig.Creditos > 0 ? asig.Creditos * 16.0 : 64.0)));

                AgAula AulaDe(int idAula) =>
                    contexto.AulasPorId != null && contexto.AulasPorId.TryGetValue(idAula, out var au) ? au : null;

                // Fase 980: helper para resolver la unidad de sesión efectiva de un aula
                double UnidadDe(AgAula au)
                {
                    if (au == null) return 4.0;
                    var progAu = contexto.Programas?.FirstOrDefault(p => p.Id == au.IdPrograma);
                    return AgUnidadesSemanales.HorasSesionAula(au, AgNivelAcademicoCatalogo.Desde(progAu?.Modalidad));
                }

                // Fase 586 / Fase 591 / Fase 601 / Fase 959 / Fase 981 / Fase 982: capacidad real en sesiones (acotada por franjas reales en catálogo)
                int CapacidadDe(AgAulaAsignatura a)
                {
                    var au = AulaDe(a?.IdAula ?? 0);
                    double unidadAu = UnidadDe(au);

                    // Fase 1123 / Fase 1158 / Fase 1167 / Fase 1181: capacidad por carril; retornar capMax (la asignación concreta la evalúa el planificador)
                    var ranurasAu = au != null ? AgRanuras.PorVentana(au.DiasVentana, contexto.Bloques, au.Jornada, unidadAu) : null;
                    if (ranurasAu != null && ranurasAu.Count > 0)
                    {
                        int carrilesR = AgCarriles.Cantidad(ranurasAu);
                        int capMin = int.MaxValue;
                        int capMax = 0;
                        for (int c = 0; c < carrilesR; c++)
                        {
                            int cap = AgCarriles.DelCarril(ranurasAu, c)
                                .Sum(r => AgPlanificadorTemporal.ContarOcurrencias(r.Dia, ini, fin, contexto.EsFestivo));
                            capMin = Math.Min(capMin, cap);
                            capMax = Math.Max(capMax, cap);
                        }
                        if (carrilesR > 1)
                        {
                            LoggerHelper.WarningLimitado(
                                $"CapacidadMulticarrilRanuras|{a?.IdAula}",
                                $"[ValidadorVentana] Aula {a?.IdAula}: {carrilesR} carriles; capacidad por carril min={capMin}, max={capMax} ranuras.");
                        }
                        return capMax;   // F1181: la asignatura puede caer en cualquier carril; la infactibilidad real la reporta AgPlanificadorOleadas (F1173)
                    }

                    // Fase 1012 / Fase 1014 / Fase 1190: fallback con días K y franjas reales por día
                    var diasCarril = AgPlanificadorPatronSemanal.DiasKDe(a, au, 0);
                    var ventanaIso = au != null ? AgDiasVentana.Dias(au.DiasVentana) : null;
                    // Fase 1094 / Fase 1124: carriles canónicos mediante AgCarriles.Cantidad
                    int carriles = (ventanaIso != null && ventanaIso.Count > 0)
                        ? AgCarriles.Cantidad(ventanaIso)
                        : 1;
                    if (carriles > 1)
                    {
                        // Fase 1014: el motor reparte carriles; la capacidad debe advertir el supuesto usado
                        LoggerHelper.WarningLimitado(
                            $"CapacidadMulticarril|{a?.IdAula}",
                            $"[ValidadorVentana] Aula {a?.IdAula}: {carriles} carriles; capacidad calculada con el carril 0 ([{string.Join(",", diasCarril)}]).");
                    }
                    return diasCarril.Sum(d =>
                        AgPlanificadorTemporal.ContarOcurrencias(d, ini, fin, contexto.EsFestivo)
                        * Math.Max(1, contexto.Bloques == null
                            ? 1
                            : contexto.Bloques
                                .Where(b => b.Activo == 1 && b.DiaSemanaNumero == d
                                            && Math.Abs(b.DuracionHoras - unidadAu) < 0.001)
                                .Select(b => (b.Franja ?? string.Empty).Trim().ToUpperInvariant())
                                .Distinct()
                                .Count()));
                }

                // Fase 586 / Fase 591 / Fase 601: máximo entre configuraciones de aula; sin configuraciones
                // se conserva el patrón del índice o el por defecto {2,3}
                var configs = (contexto.AulaAsignaturas != null && contexto.AulaAsignaturas.Count > 0)
                    ? contexto.AulaAsignaturas.Where(x => x.IdAsignatura == asig.Id).ToList()
                    : new List<AgAulaAsignatura>();

                if (configs.Count == 0 && aa != null) configs.Add(aa);

                int capacidadReal = configs.Count > 0 ? configs.Select(CapacidadDe).Max() : CapacidadDe(null);

                // Fase 980: la unidad se resuelve con el aula ganadora (misma que gobierna capacidadReal),
                // no con la asignación 'aa' que puede ser nula en primera corrida.
                AgAulaAsignatura mejorCfg = configs.Count > 0 ? configs.OrderByDescending(CapacidadDe).First() : null;
                var aulaGanadora = AulaDe(mejorCfg?.IdAula ?? aa?.IdAula ?? 0);

                double unidadAsig = UnidadDe(aulaGanadora);
                int sesionesRequeridas = (int)Math.Ceiling(hAsig / unidadAsig);

                // Fase 603 / Fase 959: reporte detallado identificando el aula ganadora con días K efectivos
                if (sesionesRequeridas > capacidadReal)
                {
                    var ganadora = mejorCfg;
                    informe.AgregarBloqueo(
                        $"{EtiquetaSede(contexto, AulaDe(ganadora?.IdAula ?? 0))} Asignatura '{asig.Nombre}' (Id={asig.Id}): {sesionesRequeridas} sesiones > capacidad real {capacidadReal} " +
                        $"(aula={ganadora?.IdAula ?? 0}, días=[{string.Join(",", AgPlanificadorPatronSemanal.DiasKDe(ganadora, AulaDe(ganadora?.IdAula ?? 0)))}], ventana {ini:yyyy-MM-dd}..{fin:yyyy-MM-dd}).");
                }

                // Fase 938 / Fase 954 / Fase 979: Validar contención D14 con días efectivos de cada configuración
                foreach (var cfg in configs)
                {
                    var auCfg = AulaDe(cfg.IdAula);
                    if (auCfg != null && !string.IsNullOrWhiteSpace(auCfg.DiasVentana))
                    {
                        // Fase 1012 / Fase 1014: días efectivos del carril base y advertencia de supuesto multicarril
                        var diasEfectivos = AgPlanificadorPatronSemanal.DiasKDe(cfg, auCfg, 0);
                        var ventanaIsoCfg = AgDiasVentana.Dias(auCfg.DiasVentana);
                        // Fase 1094: Usar regla canónica de división entera (AgCarriles.Cantidad) en vez de Math.Ceiling
                        int carrilesCfg = AgCarriles.Cantidad(ventanaIsoCfg);
                        if (carrilesCfg > 1)
                        {
                            LoggerHelper.WarningLimitado(
                                $"ValidarVentanaMulticarril|{cfg?.IdAula}",
                                $"[ValidadorVentana] Aula {cfg?.IdAula}: {carrilesCfg} carriles; validación evaluada con el carril 0 ([{string.Join(",", diasEfectivos)}]).");
                        }

                        // Fase 979 / D4': validar con la duración efectiva del nivel, no con horas_dia crudo
                        var progCfg = contexto.Programas?.FirstOrDefault(p => p.Id == auCfg.IdPrograma);
                        int horasEfectivasCfg = (int)AgUnidadesSemanales.HorasSesionAula(auCfg, AgNivelAcademicoCatalogo.Desde(progCfg?.Modalidad));

                        var v = Validar(string.Join(",", diasEfectivos), auCfg.DiasVentana, auCfg.Jornada, horasEfectivasCfg, contexto.Bloques);
                        if (v.IsFailure)
                        {
                            informe.AgregarBloqueo($"{EtiquetaSede(contexto, auCfg)} Asignatura '{asig.Nombre}' / Aula '{auCfg.AulaRef}': {v.Error}");
                        }
                    }
                }
            }

            // Fase 1059 (I5): Validar factibilidad aritmética de carriles y oleadas por aula/cohorte
            if (contexto.Aulas != null && contexto.AulaAsignaturas != null && contexto.AulaAsignaturas.Count > 0)
            {
                var asigPorId = (contexto.AsignaturasPorId != null && contexto.AsignaturasPorId.Count > 0)
                    ? contexto.AsignaturasPorId
                    : (contexto.Asignaturas?.ToDictionary(a => a.Id) ?? new Dictionary<int, Asignatura>());
                foreach (var au in contexto.Aulas)
                {
                    var asigsAula = contexto.AulaAsignaturas.Where(aa => aa.IdAula == au.Id && aa.IdAsignatura > 0).ToList();
                    if (asigsAula.Count == 0) continue;

                    var ventanaAu = AgCarriles.Ventana(au.DiasVentana);
                    // Fase 1093: Bloquear factibilidad explícitamente si ag_aula carece de dias_ventana (I1)
                    if (ventanaAu.Count == 0)
                    {
                        informe.AgregarBloqueo(
                            $"[I1_VENTANA_VACIA] {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}': sin dias_ventana. Ejecute 'Generar Aulas' con la selección vigente.");
                        continue;
                    }

                    var progAu = contexto.Programas?.FirstOrDefault(p => p.Id == au.IdPrograma);
                    var nivelAu = AgNivelAcademicoCatalogo.Desde(progAu?.Modalidad);
                    // Fase 1114: Bloqueo explícito si aula de pregrado tiene configuradas 8 h/día (unidad canónica es 4 h)
                    if (au.HorasDia == 8 && nivelAu == AgNivelAcademico.Pregrado)
                    {
                        informe.AgregarBloqueo(
                            $"[D4_PREGRADO_8H] {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}': 8 h/día no permitido en PREGRADO (unidad canónica 4 h). " +
                            "Seleccione 4h en el árbol y ejecute 'Generar Aulas'.");
                    }
                    double unidadAu = AgUnidadesSemanales.HorasSesionAula(au, nivelAu);

                    // Fase 1142 / Fase 1158: Capacidad real medida en ranuras efectivas (SAB cuenta 2 veces si tiene M+T)
                    var ranurasAu = AgRanuras.PorVentana(au.DiasVentana, contexto.Bloques, au.Jornada, unidadAu);
                    int carrilesAu = ranurasAu.Count > 0 ? AgCarriles.Cantidad(ranurasAu) : AgCarriles.Cantidad(ventanaAu);
                    int capacidadTotalRanuras = 0;
                    if (ranurasAu.Count > 0)
                    {
                        for (int c = 0; c < carrilesAu; c++)
                        {
                            var rCarril = AgCarriles.DelCarril(ranurasAu, c);
                            for (int r = 0; r < rCarril.Count; r++)
                            {
                                capacidadTotalRanuras += AgPlanificadorTemporal.ContarOcurrencias(rCarril[r].Dia, ini, fin, contexto.EsFestivo);
                            }
                        }
                    }
                    else
                    {
                        for (int c = 0; c < carrilesAu; c++)
                        {
                            var diasCarrilAu = AgCarriles.DelCarril(ventanaAu, c);
                            for (int d = 0; d < diasCarrilAu.Count; d++)
                            {
                                capacidadTotalRanuras += AgPlanificadorTemporal.ContarOcurrencias(diasCarrilAu[d], ini, fin, contexto.EsFestivo);
                            }
                        }
                    }
                    
                    // Fase 1105 / Fase 1110: Calcular demanda real resolviendo horas_objetivo o int_horaria; omitir desconocidas (<=0) con advertencia
                    int totalSesionesAula = 0;
                    int materiasComputadas = 0;
                    foreach (var aa in asigsAula)
                    {
                        double horasAu = aa.HorasObjetivo > 0
                            ? aa.HorasObjetivo
                            : (asigPorId.TryGetValue(aa.IdAsignatura, out var asigAu) && asigAu.IntHoraria > 0 ? asigAu.IntHoraria : 0.0);

                        if (horasAu <= 0.0)
                        {
                            informe.AgregarAdvertencia($"[I5_HORAS_DESCONOCIDAS] {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}' asig {aa.IdAsignatura}: horas_objetivo=0; no se computa en la demanda.");
                            continue;
                        }
                        totalSesionesAula += (int)Math.Ceiling(horasAu / unidadAu);
                        materiasComputadas++;
                    }

                    if (totalSesionesAula > capacidadTotalRanuras)
                    {
                        string escalon = AgEscaleraVentanas.Siguiente(au.DiasVentana);
                        informe.AgregarBloqueo(
                            $"[I5_INFACTIBLE] {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}': demanda total de {totalSesionesAula} sesiones " +
                            $"({materiasComputadas} materias computadas de {asigsAula.Count}) excede la capacidad real de " +
                            $"{capacidadTotalRanuras} ranuras ({ranurasAu.Count} ranuras/semana, festivos descontados) " +
                            $"en ventana {ini:yyyy-MM-dd}..{fin:yyyy-MM-dd}." +
                            (string.IsNullOrEmpty(escalon)
                                ? " Amplíe la ventana semestral o reduzca asignaturas simultáneas."
                                : $" Escale la ventana a '{escalon}' o amplíe el semestre."));
                        continue;
                    }

                    // Fase 1122 / Fase 1129: Validar sub-ventana de oleada por asignatura dentro del aula (filtrando por alcance real del motor)
                    var semAula = AgFormatoAsignatura.SemestreRomano(au.Semestre);
                    var itemsPlan = asigsAula
                        .Where(aa => asigPorId.TryGetValue(aa.IdAsignatura, out var aF)
                                     && (aF.IdPrograma <= 0 || au.IdPrograma <= 0 || aF.IdPrograma == au.IdPrograma)
                                     && (string.IsNullOrWhiteSpace(semAula) || semAula == "SIN ASIGNAR"
                                         || string.Equals(AgFormatoAsignatura.SemestreRomano(aF.Semestre), semAula, StringComparison.OrdinalIgnoreCase)))
                        .Select(aa => (IdAsignatura: aa.IdAsignatura,
                                       Horas: aa.HorasObjetivo > 0
                                           ? aa.HorasObjetivo
                                           : (asigPorId.TryGetValue(aa.IdAsignatura, out var asigAu) && asigAu.IntHoraria > 0 ? asigAu.IntHoraria : 0.0),
                                       OrdenCurricular: aa.OrdenCurricular))
                        .Where(x => x.Horas > 0)
                        .ToList();

                    var patronesAula = itemsPlan
                        .Select(x => (x.IdAsignatura,
                                      Patron: contexto.AulaAsignaturaPorAulaAsig != null &&
                                              contexto.AulaAsignaturaPorAulaAsig.TryGetValue((au.Id, x.IdAsignatura), out var aaP)
                                                  ? aaP?.DiasPatron : null))
                        .Where(t => !string.IsNullOrWhiteSpace(t.Patron))
                        .GroupBy(t => t.IdAsignatura)
                        .ToDictionary(g => g.Key, g => g.First().Patron);

                    var planRes = new AgPlanificadorOleadas().Construir(itemsPlan, au.DiasVentana, ini, fin, unidadAu,
                        contexto.EsFestivo, contexto.Bloques, au.Jornada, patronesAula);   // Fase 1252
                    if (planRes.IsFailure)
                    {
                        // F1594: Distinguir I1_ASIGNATURA_DUPLICADA de I5_INFACTIBLE_OLEADA
                        string codigo = planRes.Error.StartsWith("I1_ASIGNATURA_DUPLICADA:", StringComparison.Ordinal)
                            ? "I1_ASIGNATURA_DUPLICADA"
                            : "I5_INFACTIBLE_OLEADA";

                        informe.AgregarBloqueo($"[{codigo}] {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}': {planRes.Error}");
                    }
                    else if (planRes.Value.Infactibles != null && planRes.Value.Infactibles.Count > 0)
                    {
                        string detalle = planRes.Value.DetalleInfactibles != null && planRes.Value.DetalleInfactibles.Count > 0
                            ? " " + string.Join(" | ", planRes.Value.DetalleInfactibles) + "."
                            : string.Empty;

                        string sugerida = planRes.Value.VentanaSugerida;
                        if (!string.IsNullOrEmpty(sugerida)) detalle += $" Escale la ventana a '{sugerida}'.";

                        // F1600 / F1601: Distinguir incompatibilidad R8 de saturación I5_CAPACIDAD_CARRIL_PATRON y agotamiento temporal I5_INFACTIBLE_OLEADA
                        bool esR8 = planRes.Value.DetalleInfactibles != null &&
                                    planRes.Value.DetalleInfactibles.Any(x => x.Contains("R8_DiaFueraDePatron"));

                        bool esCapacidadPatron = planRes.Value.DetalleInfactibles != null &&
                                    planRes.Value.DetalleInfactibles.Any(x => x.Contains("I5_CAPACIDAD_CARRIL_PATRON"));

                        string etiquetaBloqueo = esR8
                            ? "[R8_DiaFueraDePatron]"
                            : (esCapacidadPatron ? "[I5_CAPACIDAD_CARRIL_PATRON]" : "[I5_INFACTIBLE_OLEADA]");

                        informe.AgregarBloqueo(
                            esR8
                                ? $"{etiquetaBloqueo} {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}': asignaturas [{string.Join(",", planRes.Value.Infactibles)}] " +
                                  $"con día de patrón fuera de la banda asignada.{detalle}"
                                : $"{etiquetaBloqueo} {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}': asignaturas [{string.Join(",", planRes.Value.Infactibles)}] " +
                                  $"no caben en su sub-ventana de oleada.{detalle} Amplíe la ventana, reduzca horas o aumente carriles.");   // Fase 1269 / Fase 1601
                    }

                    if (capacidadTotalRanuras > 0)
                    {
                        double ocupacion = totalSesionesAula / (double)capacidadTotalRanuras * 100.0;
                        if (ocupacion >= 88.0)
                        {
                            informe.AgregarAdvertencia(
                                $"[I5_HOLGURA_BAJA] {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}': ocupación {ocupacion:0.#}% " +
                                $"({totalSesionesAula}/{capacidadTotalRanuras} sesiones). Sin margen para festivos " +
                                $"ni reprogramaciones; considere ampliar la ventana (AgEscaleraVentanas).");   // Fase 1230
                        }
                    }
                }
            }

            // 4. Validar cobertura del Pool docente para cada asignatura (Fase 295)
            int conDocente = 0;
            foreach (var asig in contexto.Asignaturas)
            {
                bool tieneHabilitados = false;
                if (contexto.PoolPorAsignatura != null && contexto.PoolPorAsignatura.TryGetValue(asig.Id, out var habilitados) && habilitados.Any(h => h.Activo == 1))
                {
                    tieneHabilitados = true;
                    conDocente++;
                }

                if (!tieneHabilitados)
                {
                    informe.AsignaturasVacantes[asig.Id] = AgMotivoVacante.SIN_CANDIDATOS;
                    informe.AgregarAdvertencia($"La asignatura '{asig.Nombre}' (Id={asig.Id}) no tiene ningún docente habilitado en el Pool.");
                }
            }

            if (contexto.Asignaturas.Count > 0 && conDocente == 0 && contexto.Docentes.Count == 0)
            {
                informe.AgregarBloqueo($"Ninguna de las {contexto.Asignaturas.Count} asignaturas a programar tiene docentes habilitados en el Pool ni en catálogo general.");
            }

            return Result.Success(informe);
        }

        /// <summary>
        /// Fase 1579: Resuelve una etiqueta legible de sede para los bloqueos/advertencias del pre-vuelo,
        /// de modo que el log identifique la sede además de la referencia del aula (AulaRef).
        /// Prioriza el nombre de la sede en catálogo, luego su subunidad y, en último término, la del aula.
        /// </summary>
        private static string EtiquetaSede(AgContexto contexto, AgAula aula)
        {
            if (aula == null) return "Sede 'N/D'";

            Sede sede = null;
            if (aula.IdSede > 0)
            {
                if (contexto?.SedesPorId != null)
                {
                    contexto.SedesPorId.TryGetValue(aula.IdSede, out sede);
                }

                if (sede == null && contexto?.Sedes != null)
                {
                    sede = contexto.Sedes.FirstOrDefault(s => s != null && s.Id == aula.IdSede);
                }
            }

            string nombre = sede?.Nombre;
            if (string.IsNullOrWhiteSpace(nombre)) nombre = sede?.Subunidad;
            if (string.IsNullOrWhiteSpace(nombre)) nombre = aula.Subunidad;
            if (string.IsNullOrWhiteSpace(nombre)) nombre = contexto?.Subunidad;
            if (string.IsNullOrWhiteSpace(nombre))
            {
                nombre = aula.IdSede > 0 ? $"Id {aula.IdSede}" : "N/D";
            }

            return $"Sede '{nombre.Trim()}'";
        }

        /// <summary>
        /// Valida la completitud de la tupla de decisión del aula (diasVentana, jornada, horasDia)
        /// y la contención de los días del patrón dentro de la ventana del aula (Fase 1054).
        /// </summary>
        public static Result Validar(string diasPatron, string diasVentana, string jornada, int horasDia, IEnumerable<AgBloqueHorario> bloques)
        {
            if (string.IsNullOrWhiteSpace(diasVentana))
            {
                return Result.Failure("I1_VENTANA_VACIA: No se especificó la ventana de días del aula.");
            }

            if (string.IsNullOrWhiteSpace(jornada))
            {
                return Result.Failure("I1_JORNADA_VACIA: No se especificó la jornada del aula.");
            }

            if (horasDia != 4 && horasDia != 8)
            {
                return Result.Failure($"I6_HORAS_DIA_INVALIDO: Horas por día ({horasDia}) inválido. Debe ser 4 u 8 horas.");
            }

            var permitidos = AgDiasVentana.Dias(diasVentana);
            if (permitidos == null || permitidos.Count == 0)
            {
                return Result.Failure($"I1_VENTANA_INVALIDA: La ventana '{diasVentana}' no contiene días válidos [1..7].");
            }

            // 1. Validar contención de días si se definió patrón
            if (!string.IsNullOrWhiteSpace(diasPatron))
            {
                var pedidos = diasPatron.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => int.TryParse(s, out int d) ? d : 0)
                    .Where(d => d >= 1 && d <= 7)
                    .ToList();

                var noContenidos = pedidos.Where(d => !permitidos.Contains(d)).ToList();
                if (noContenidos.Count > 0)
                {
                    return Result.Failure($"Los días del patrón [{string.Join(",", pedidos)}] no están contenidos en la ventana seleccionada '{diasVentana}' [{string.Join(",", permitidos)}]. Días no contenidos: [{string.Join(",", noContenidos)}].");
                }
            }

            // 2. Validar existencia de bloque si se definió jornada y horas_dia
            var fj = AgFranjaJornada.Resolve(jornada, horasDia);
            if (!fj.HasValue)
            {
                return Result.Failure($"La combinación de jornada '{jornada}' y {horasDia} horas/día no es válida.");
            }

            if (bloques != null)
            {
                bool existe = bloques.Any(b => b.Activo == 1
                    && string.Equals(b.Franja, fj.Value.Franja, StringComparison.OrdinalIgnoreCase)
                    && Math.Abs(b.DuracionHoras - fj.Value.DuracionHoras) < 0.001);

                if (!existe)
                {
                    return Result.Failure($"No existe ningún bloque horario activo para la jornada '{jornada}' ({fj.Value.Franja}) con duración {horasDia} horas en el catálogo.");
                }
            }

            return Result.Success();
        }
    }
}
