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

            // 2. Validar disponibilidad de aulas y bloques (Fase 1846: incluye aulas autorizadas PuedeGenerarAula)
            var aulasProgramables = contexto.Aulas.Where(a => a.Estado == AgEstadoAula.Disponible || contexto.PuedeGenerarAula(a.Id)).ToList();
            if (aulasProgramables.Count == 0)
            {
                informe.AgregarBloqueo("No hay aulas programables configuradas para la sede/subunidad.");
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

            // Fase 418 & Fase 1846: Mismo alcance para oferta que para la demanda evaluada (aulas programables para semestre/grupo)
            var aulasOferta = aulasProgramables;

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
                    // La capacidad hipotética no invalida un horario canónico completo ya verificado.
                    if (contexto.ManifiestoCobertura != null &&
                        contexto.ManifiestoCobertura.Relaciones.Where(r => r.Clave.IdAula == au.Id)
                            .All(r => Math.Abs(contexto.SesionesConservadas.Where(s => s.Clave.Valor == r.Clave.Valor)
                                .Sum(s => s.Horas) - r.HorasEfectivas) <= 0.001))
                        continue;
                    // FASE 1859 & 1865: Contabilizar relaciones existentes antes de comprobaciones estructurales
                    int totalRelaciones = (contexto.AulaAsignaturas ?? new List<AgAulaAsignatura>())
                        .Count(aa => aa.IdAula == au.Id && aa.IdAsignatura > 0);
                    if (totalRelaciones == 0) continue;

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

                    // FASE 1856 & 1861: Unificar definición de demanda sin fallback ficticio de 64h (preserva I5_HORAS_DESCONOCIDAS)
                    var demandasAu = AgEvaluadorDemandaVentana.Demandas(contexto, au, fallbackPermitido: false);
                    if (demandasAu.Count == 0) continue;

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
                    
                    // FASE 1856: Calcular demanda real unificada con Demandas; omitir desconocidas (<=0) con advertencia
                    int totalSesionesAula = 0;
                    int materiasComputadas = 0;
                    foreach (var d in demandasAu)
                    {
                        if (d.Horas <= 0.0)
                        {
                            informe.AgregarAdvertencia($"[I5_HORAS_DESCONOCIDAS] {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}' asig {d.IdAsignatura}: horas_objetivo=0; no se computa en la demanda.");
                            continue;
                        }
                        totalSesionesAula += (int)Math.Ceiling(d.Horas / unidadAu);
                        materiasComputadas++;
                    }

                    if (totalSesionesAula > capacidadTotalRanuras)
                    {
                        string escalon = AgEscaleraVentanas.Siguiente(au.DiasVentana);
                        string consejoEscalon;
                        if (string.IsNullOrEmpty(escalon))
                        {
                            consejoEscalon = " Amplíe la ventana semestral o reduzca asignaturas simultáneas.";
                        }
                        else
                        {
                            // FASE 1856: Validar ganancia real antes de sugerir el escalón
                            var ranurasEscalon = AgRanuras.PorVentana(escalon, contexto.Bloques, au.Jornada, unidadAu);
                            int capEscalon = 0;
                            if (ranurasEscalon.Count > 0)
                            {
                                int carrilesEsc = AgCarriles.Cantidad(ranurasEscalon);
                                for (int c = 0; c < carrilesEsc; c++)
                                {
                                    var rCarril = AgCarriles.DelCarril(ranurasEscalon, c);
                                    for (int r = 0; r < rCarril.Count; r++)
                                    {
                                        capEscalon += AgPlanificadorTemporal.ContarOcurrencias(rCarril[r].Dia, ini, fin, contexto.EsFestivo);
                                    }
                                }
                            }

                            if (capEscalon > capacidadTotalRanuras)
                            {
                                consejoEscalon = $" Escale la ventana a '{escalon}' (capacidad {capEscalon}) o amplíe el semestre.";
                            }
                            else
                            {
                                consejoEscalon = $" El siguiente escalón '{escalon}' no añade ranuras útiles (no existe bloque activo de {unidadAu:0} h para la jornada '{au.Jornada}' en ag_bloque_horario). Configure bloques para esa jornada o amplíe el semestre.";
                            }
                        }

                        informe.AgregarBloqueo(
                            $"[I5_INFACTIBLE] {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}': demanda total de {totalSesionesAula} sesiones " +
                            $"({materiasComputadas} materias computadas de {totalRelaciones}) excede la capacidad real de " +
                            $"{capacidadTotalRanuras} ranuras ({ranurasAu.Count} ranuras/semana, festivos descontados) " +
                            $"en ventana {ini:yyyy-MM-dd}..{fin:yyyy-MM-dd}." +
                            consejoEscalon);
                        continue;
                    }

                    // Fase 1763 & Fase 1768: Evaluación de saturación de ventana/carril y diagnóstico I5_DIA_SATURADO
                    // Fase 1772: no bloquear si todas las materias del aula tienen otra aula viable
                    bool TieneCapacidadAlternativa(AgAulaAsignatura cfg, int requeridas)
                    {
                        var auAlt = contexto.Aulas?.FirstOrDefault(x => x.Id == cfg.IdAula);
                        if (auAlt == null) return false;
                        var progAlt = contexto.Programas?.FirstOrDefault(p => p.Id == auAlt.IdPrograma);
                        var nivelAlt = AgNivelAcademicoCatalogo.Desde(progAlt?.Modalidad);
                        double unidadAlt = AgUnidadesSemanales.HorasSesionAula(auAlt, nivelAlt);
                        var ranurasAlt = AgRanuras.PorVentana(auAlt.DiasVentana, contexto.Bloques, auAlt.Jornada, unidadAlt);
                        int capacidadAlt = ranurasAlt.Sum(r =>
                            AgPlanificadorTemporal.ContarOcurrencias(r.Dia, ini, fin, contexto.EsFestivo));
                        return capacidadAlt >= requeridas;
                    }

                    // FASE 1859: Iteración canónica sobre relaciones del aula en contexto
                    bool todasConAlternativa = true;
                    foreach (var rel in (contexto.AulaAsignaturas ?? new List<AgAulaAsignatura>()).Where(aa => aa.IdAula == au.Id))
                    {
                        double horasAlt = rel.HorasObjetivo > 0
                            ? rel.HorasObjetivo
                            : (asigPorId.TryGetValue(rel.IdAsignatura, out var asigAlt) && asigAlt.IntHoraria > 0 ? asigAlt.IntHoraria : 0.0);
                        if (horasAlt <= 0) continue;

                        int reqAlt = (int)Math.Ceiling(horasAlt / unidadAu);
                        var alternativas = (contexto.AulaAsignaturas ?? new List<AgAulaAsignatura>())
                            .Where(x => x.IdAsignatura == rel.IdAsignatura && x.IdAula != au.Id)
                            .ToList();
                        if (!alternativas.Any(c => TieneCapacidadAlternativa(c, reqAlt)))
                        {
                            todasConAlternativa = false;
                            break;
                        }
                    }

                    if (!todasConAlternativa)
                    {
                        var evalDemanda = AgEvaluadorDemandaVentana.Evaluar(au.DiasVentana, au.Jornada, unidadAu, contexto, au, ini, fin);
                        if (evalDemanda.HaySaturacion || evalDemanda.SaturacionCarriles)
                        {
                            informe.AgregarBloqueo(
                                $"[I5_DIA_SATURADO] {EtiquetaSede(contexto, au)} Aula '{au.AulaRef}': {evalDemanda.Motivo}; " +
                                $"{evalDemanda.SemanasRequeridas} semanas; fin proyectado " +
                                (evalDemanda.FechaFinProyectada.HasValue ? evalDemanda.FechaFinProyectada.Value.ToString("yyyy-MM-dd") : "sin proyección disponible") + ".");
                            continue; // Suprimir el diagnóstico genérico de oleada que aconsejaría escalar erróneamente
                        }
                    }

                    // Fase 1122 / Fase 1129: Validar sub-ventana de oleada por asignatura dentro del aula (filtrando por alcance real del motor)
                    var itemsPlan = AgEvaluadorDemandaVentana.Demandas(contexto, au);

                    var patronesAula = itemsPlan
                        .Select(x => (x.IdAsignatura,
                                      Patron: contexto.AulaAsignaturaPorAulaAsig != null &&
                                              contexto.AulaAsignaturaPorAulaAsig.TryGetValue((au.Id, x.IdAsignatura), out var aaP)
                                              && !string.Equals(aaP.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase)
                                                  ? aaP.DiasPatron : null))
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

            // Fase 2046 / Fase 2090 / Fase 2096 / Fase 2100 / Fase 2105: Poblar informe con el déficit docente estructural considerando concurrencia por jornada en un solo pase O(N)
            if (contexto.AulaAsignaturas != null && contexto.AulasPorId != null)
            {
                var deficitBase = ConstruirDeficitBase(contexto);
                Dictionary<int, List<int>> sedesPorAsig = null;
                if (LoggerHelper.ModoVerboso)
                {
                    sedesPorAsig = contexto.AulaAsignaturas
                        .GroupBy(aa => aa.IdAsignatura)
                        .ToDictionary(
                            g => g.Key,
                            g => g.Select(aa => contexto.AulasPorId.TryGetValue(aa.IdAula, out var a) && a != null ? a.IdSede : 0).Distinct().ToList()
                        );
                }

                foreach (var kvp in deficitBase)
                {
                    int idAsig = kvp.Key;
                    var info = kvp.Value;
                    int deficit = Math.Max(0, info.MaxConcurrentes - info.Docs);
                    if (deficit > 0)
                    {
                        informe.DeficitDocentePorAsignatura[idAsig] = deficit;
                    }

                    if (LoggerHelper.ModoVerboso && sedesPorAsig != null)
                    {
                        sedesPorAsig.TryGetValue(idAsig, out var sedesDistinct);
                        var estadoConsentidas = (sedesDistinct ?? new List<int>())
                            .Select(s => $"{s}:{(contexto.EstaConsentida(idAsig, s) ? "consentida" : "no-consentida")}");
                        LoggerHelper.Debug($"[AG-TOPE] PreVuelo asig={idAsig} cohortes={info.Cohortes} maxConcurrentes={info.MaxConcurrentes} docs={info.Docs} deficit={deficit} sedes=[{string.Join(", ", estadoConsentidas)}]");
                    }
                }
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

        /// <summary>
        /// Calcula el déficit estructural de docentes para una asignatura específica evaluando la concurrencia por jornada (Fase 2090, 2096, 2100).
        /// Una sola fuente canónica de verdad compartida por el diagnóstico pre-vuelo y el cálculo de fallback.
        /// </summary>
        /// <param name="c">Contexto de optimización.</param>
        /// <param name="idAsignatura">Identificador de la asignatura.</param>
        /// <param name="cohortes">Total de cohortes (Semestre, Grupo, Sede) no consentidas demandadas.</param>
        /// <param name="docs">Docentes habilitados activos en el pool.</param>
        /// <returns>Déficit estructural no negativo (maxConcurrentes - docs).</returns>
        internal static int CalcularDeficitPorAsignatura(AgContexto c, int idAsignatura, out int cohortes, out int docs)
        {
            cohortes = 0;
            docs = 0;
            if (c?.AulaAsignaturas == null || c.AulasPorId == null) return 0;

            var dict = ConstruirDeficitBase(c);
            if (dict.TryGetValue(idAsignatura, out var tupla))
            {
                cohortes = tupla.Cohortes;
                docs = tupla.Docs;
                return Math.Max(0, tupla.MaxConcurrentes - docs);
            }

            return 0;
        }

        /// <summary>
        /// Construye en un solo recorrido O(N) el diccionario base de cohortes, concurrencia por jornada y docentes por asignatura (Fase 2100).
        /// </summary>
        internal static Dictionary<int, (int Cohortes, int MaxConcurrentes, int Docs)> ConstruirDeficitBase(AgContexto c)
        {
            var resultado = new Dictionary<int, (int Cohortes, int MaxConcurrentes, int Docs)>();
            if (c?.AulaAsignaturas == null || c.AulasPorId == null) return resultado;

            foreach (var grupo in c.AulaAsignaturas.GroupBy(aa => aa.IdAsignatura))
            {
                int idAsig = grupo.Key;
                int docs = 0;
                if (c.PoolPorAsignatura != null && c.PoolPorAsignatura.TryGetValue(idAsig, out var pool))
                {
                    docs = pool.Count(p => p.Activo == 1 && (c.IdTerceroVacante <= 0 || p.IdTercero != c.IdTerceroVacante));
                }

                var demanda = grupo
                    .Where(aa => c.AulasPorId.TryGetValue(aa.IdAula, out var a) && a != null
                              && !c.EstaConsentida(idAsig, a.IdSede))
                    .Select(aa =>
                    {
                        var a = c.AulasPorId[aa.IdAula];
                        return (Jornada: (a.Jornada ?? string.Empty).Trim().ToUpperInvariant(),
                                Semestre: a.Semestre ?? string.Empty,
                                Grupo: a.Grupo ?? string.Empty,
                                Sede: a.IdSede);
                    })
                    .Distinct()
                    .ToList();

                int cohortes = demanda.Count;
                int maxConcurrentes = demanda
                    .GroupBy(d => d.Jornada)
                    .Select(g => g.Select(d => (d.Semestre, d.Grupo, d.Sede)).Distinct().Count())
                    .DefaultIfEmpty(0)
                    .Max();

                resultado[idAsig] = (cohortes, maxConcurrentes, docs);
            }

            return resultado;
        }

        /// <summary>
        /// Calcula el déficit estructural de docentes frente a grupos/aulas concurrentes por cohorte y jornada (Fase 2021, 2026, 2031, 2036, 2090, 2096, 2100).
        /// Σ por asignatura no consentida de max(0, maxConcurrentes - docentesHabilitados).
        /// Retorna -1 si el contexto está incompleto para desactivar el tope de forma segura.
        /// </summary>
        public static int CalcularDeficitEstructural(AgContexto contexto, InformeFactibilidad informe = null)
        {
            if (informe?.DeficitDocentePorAsignatura != null && informe.DeficitDocentePorAsignatura.Count > 0)
            {
                return informe.DeficitDocentePorAsignatura.Values.Where(v => v > 0).Sum();
            }

            if (contexto?.Asignaturas == null || contexto.Asignaturas.Count == 0
                || contexto.AulaAsignaturas == null || contexto.BloquesPorId == null)
            {
                return -1; // contexto incompleto: tope desactivado, no 0
            }

            int deficitTotal = 0;
            var deficitBase = ConstruirDeficitBase(contexto);

            foreach (var kvp in deficitBase)
            {
                int idAsig = kvp.Key;
                var info = kvp.Value;
                int deficit = Math.Max(0, info.MaxConcurrentes - info.Docs);
                deficitTotal += deficit;

                if (LoggerHelper.ModoVerboso)
                {
                    LoggerHelper.Debug($"[AG-TOPE] Fallback asig={idAsig} cohortes={info.Cohortes} maxConcurrentes={info.MaxConcurrentes} docs={info.Docs} deficit={deficit}");
                }
            }

            if (LoggerHelper.ModoVerboso)
            {
                LoggerHelper.Debug($"[AG-TOPE] fallback total calculado: {deficitTotal}");
            }
            LoggerHelper.Warning($"[AG-TOPE] fallback usado: {deficitTotal}");
            return deficitTotal;
        }
    }
}
