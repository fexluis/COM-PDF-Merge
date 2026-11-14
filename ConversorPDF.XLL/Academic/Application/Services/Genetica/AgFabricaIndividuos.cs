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
    /// Fábrica para la generación y mutación constructiva de individuos cromosómicos iniciales.
    /// </summary>
    public class AgFabricaIndividuos
    {
        private readonly Random _random;
        private readonly AgPlanificadorPatronSemanal _planificadorPatron;
        private readonly AgPlanificadorOleadas _planificadorOleadas;
        private readonly Dictionary<int, int> _rotacionDocente = new Dictionary<int, int>(); // Fase 533: reparto balanceado
        // Fase 1090 & Fase 1102: Memoización de planes de oleadas por corrida, aula y parámetros
        private readonly Dictionary<string, Result<AgPlanOleadasResultado>> _cachePlanes =
            new Dictionary<string, Result<AgPlanOleadasResultado>>(StringComparer.Ordinal);
        private long _ultimaCorrida = -1;

        public AgFabricaIndividuos()
            : this(null, null, null)
        {
        }

        public AgFabricaIndividuos(Random random)
            : this(random, null, null)
        {
        }

        public AgFabricaIndividuos(Random random, AgPlanificadorPatronSemanal planificadorPatron)
            : this(random, planificadorPatron, null)
        {
        }

        internal AgFabricaIndividuos(Random random, AgPlanificadorPatronSemanal planificadorPatron, AgPlanificadorOleadas planificadorOleadas)
        {
            _random = random ?? new Random();
            _planificadorPatron = planificadorPatron ?? new AgPlanificadorPatronSemanal();
            _planificadorOleadas = planificadorOleadas ?? new AgPlanificadorOleadas(_planificadorPatron);
        }

        public AgIndividuo CrearAleatorio(AgContexto contexto, long idCorrida, int generacion = 0, AgDiagnosticoTraza traza = null)
        {
            if (contexto == null) throw new ArgumentNullException(nameof(contexto));

            // Fase 1102: Reiniciar caché de planes si cambia la corrida para evitar contaminación entre ejecuciones
            if (_ultimaCorrida != idCorrida)
            {
                _cachePlanes.Clear();
                _ultimaCorrida = idCorrida;
            }

            // AG-RANURAS-TMP: detalle de genes solo en la primera semilla de esta corrida.
            bool diagnosticoPrimeraSemilla = LoggerHelper.ModoVerboso && _cachePlanes.Count == 0;
            _rotacionDocente.Clear(); // Fase 533: reiniciar rotación por individuo
            var trazaEfectiva = traza ?? contexto.Traza;
            int descartadasSemestre = 0;

            var individuo = new AgIndividuo
            {
                IdCorrida = idCorrida,
                Generacion = generacion,
                Estado = AgEstadoIndividuo.NoEvaluado,
                Subunidad = contexto.Subunidad
            };

            var aulasValidas = contexto.Aulas.Where(a => a.Estado == AgEstadoAula.Disponible).ToList();
            var bloquesValidos = contexto.Bloques.Where(b => b.Activo == 1).ToList();

            int fallbackAulaId = contexto.IdAulaCorrida > 0
                ? contexto.IdAulaCorrida
                : (aulasValidas.Count > 0 ? aulasValidas[0].Id : 0);

            // Fase 475: Derivar grupo del AulaRef cuando ag_aula.grupo sea nulo
            string GrupoDeAula(AgAula a)
            {
                if (a == null) return "A";
                if (!string.IsNullOrWhiteSpace(a.Grupo)) return a.Grupo;
                var partes = (a.AulaRef ?? string.Empty).Split('|');
                return partes.Length >= 5 && partes[4].Length > 0 ? partes[4] : "A";
            }

            // Fase 478: Derivar semestre del AulaRef cuando ag_aula.semestre sea nulo
            string SemestreDeAula(AgAula a)
            {
                if (a == null) return string.Empty;
                if (!string.IsNullOrWhiteSpace(a.Semestre)) return a.Semestre;
                var partes = (a.AulaRef ?? string.Empty).Split('|');
                return partes.Length >= 4 ? partes[3] : string.Empty;
            }

            // Derivar programa del AulaRef cuando ag_aula.idprograma sea nulo/cero
            int ProgramaDeAula(AgAula a)
            {
                if (a == null) return contexto.IdPrograma;
                if (a.IdPrograma > 0) return a.IdPrograma;
                var partes = (a.AulaRef ?? string.Empty).Split('|');
                if (partes.Length >= 3 && int.TryParse(partes[2], out int idProg) && idProg > 0)
                {
                    return idProg;
                }
                return contexto.IdPrograma;
            }

            // Fase 454, Fase 475 & Fase 478: Agrupar por cohorte representativa considerando AulaRef / Semestre + Grupo + Programa
            // Fase 952 / Fase 953: Preservar carriles concurrentes usando AgClaveCohorte y tupla (Cohorte, Jornada, DiasVentana)
            // Fase 1003: Deduplicar por Id de aula para conservar carriles concurrentes de la misma cohorte
            // Fase 1033: DesdeAula evita colisión de registro de Excel-DNA 1.8.0
            var gruposAula = contexto.IdAulaCorrida > 0
                ? aulasValidas.Where(a => a.Id == contexto.IdAulaCorrida).ToList()
                : (aulasValidas.Count > 0
                    // Fase 1003: deduplicar por Id de aula conserva carriles concurrentes de la misma cohorte
                    ? aulasValidas.GroupBy(a => a.Id).Select(g => g.First()).ToList()
                    : new List<AgAula> { null });

            foreach (var aula in gruposAula)
            {
                string semAulaRom = AgFormatoAsignatura.SemestreRomano(SemestreDeAula(aula));
                int idProgAula = ProgramaDeAula(aula);

                // Restringir a las asignaturas específicas asignadas a este aula en ag_aula_asignatura (si existen)
                HashSet<int> asigsValidasAula = null;
                if (aula != null && contexto.AulaAsignaturas != null && contexto.AulaAsignaturas.Count > 0)
                {
                    var ids = contexto.AulaAsignaturas.Where(x => x.IdAula == aula.Id && x.IdAsignatura > 0)
                                                      .Select(x => x.IdAsignatura)
                                                      .ToList();
                    if (ids.Count > 0)
                    {
                        asigsValidasAula = new HashSet<int>(ids);
                    }
                }

                // Ronda 206 / Fase 1057: Pre-filtrar y ordenar canónicamente por (OrdenCurricular, Id)
                var asigsCandidatas = new List<Asignatura>();
                foreach (var asig in contexto.Asignaturas)
                {
                    // 1. Filtrar por asignación explícita del aula si ag_aula_asignatura contiene registros para esta aula
                    if (asigsValidasAula != null && !asigsValidasAula.Contains(asig.Id))
                    {
                        continue;
                    }

                    // 2. Filtrar por programa: una asignatura no puede programarse en un aula de otro programa (ej. APT-52 vs APT-53)
                    int idProgAsig = asig.IdPrograma;
                    if (idProgAula > 0 && idProgAsig > 0 && idProgAula != idProgAsig)
                    {
                        continue;
                    }

                    // 3. Fase 478 & Fase 481: Cruzar asignatura con aula por semestre (ignorar asignaturas de otros semestres para esta cohorte/aula)
                    if (!string.IsNullOrWhiteSpace(semAulaRom) && semAulaRom != "SIN ASIGNAR")
                    {
                        if (!string.IsNullOrWhiteSpace(asig.Semestre))
                        {
                            string semAsigRom = AgFormatoAsignatura.SemestreRomano(asig.Semestre);
                            if (semAsigRom != "SIN ASIGNAR" && !string.Equals(semAulaRom, semAsigRom, StringComparison.OrdinalIgnoreCase))
                            {
                                descartadasSemestre++;
                                continue;
                            }
                        }
                    }

                    asigsCandidatas.Add(asig);
                }

                int ObtenerOrdenCurricular(int idAsignatura)
                {
                    if (aula != null && contexto.AulaAsignaturaPorAulaAsig != null &&
                        contexto.AulaAsignaturaPorAulaAsig.TryGetValue((aula.Id, idAsignatura), out var aaExacta))
                    {
                        return aaExacta.OrdenCurricular;
                    }
                    if (contexto.AulaAsignaturasPorId != null &&
                        contexto.AulaAsignaturasPorId.TryGetValue(idAsignatura, out var aaDef))
                    {
                        return aaDef.OrdenCurricular;
                    }
                    return 0;
                }

                double ObtenerHorasObjetivo(Asignatura a)
                {
                    AgAulaAsignatura aaObj = (aula != null && contexto.AulaAsignaturaPorAulaAsig != null)
                        ? (contexto.AulaAsignaturaPorAulaAsig.TryGetValue((aula.Id, a.Id), out var exact) ? exact : null)
                        : (contexto.AulaAsignaturasPorId != null && contexto.AulaAsignaturasPorId.TryGetValue(a.Id, out var def) ? def : null);

                    return AgEvaluadorDemandaVentana.HorasObjetivo(a, aaObj);
                }

                var asigsOrdenadas = asigsCandidatas
                    .OrderBy(a => ObtenerOrdenCurricular(a.Id))
                    .ThenBy(a => a.Id)
                    .ToList();

                DateTime fInicioVentana = contexto.Ventana != null ? contexto.Ventana.FechaInicio : DateTime.Today;
                DateTime fFinVentana = contexto.Ventana != null ? contexto.Ventana.FechaFin : DateTime.Today.AddMonths(4);
                var progAulaNivel = aula != null ? contexto.Programas?.FirstOrDefault(p => p.Id == aula.IdPrograma) : null;
                var nivelAulaNivel = AgNivelAcademicoCatalogo.Desde(progAulaNivel?.Modalidad);
                double unidadSesionAula = AgUnidadesSemanales.HorasSesionAula(aula, nivelAulaNivel);

                var itemsPlan = asigsOrdenadas
                    .Select(a => (IdAsignatura: a.Id, Horas: ObtenerHorasObjetivo(a), OrdenCurricular: ObtenerOrdenCurricular(a.Id)))
                    .ToList();

                Dictionary<int, AgPlanOleada> mapaPlanes = null;
                bool planInfactible = false;
                var asigsInfactibles = new HashSet<int>();

                string ObtenerDiasPatron(int idAsig)
                {
                    AgAulaAsignatura aaPatron = (aula != null && contexto.AulaAsignaturaPorAulaAsig != null)
                        ? (contexto.AulaAsignaturaPorAulaAsig.TryGetValue((aula.Id, idAsig), out var exact) ? exact : null)
                        : (contexto.AulaAsignaturasPorId != null && contexto.AulaAsignaturasPorId.TryGetValue(idAsig, out var def) ? def : null);
                    if (aaPatron != null && string.Equals(aaPatron.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase))
                    {
                        return string.Empty;
                    }
                    return aaPatron?.DiasPatron ?? string.Empty;
                }

                if (aula != null && !string.IsNullOrWhiteSpace(aula.DiasVentana) && itemsPlan.Count > 0)
                {
                    // Fase 1090, 1102, Fase 1170 & Fase 1241: Memoización de plan por corrida, periodo, aula, parámetros y huella de patrones
                    string huellaPatrones = string.Join(";", asigsOrdenadas
                        .Select(a => $"{a.Id}={ObtenerDiasPatron(a.Id)}")
                        .Where(s => !s.EndsWith("=")));

                    string clavePlan = $"{idCorrida}|{contexto.Ventana?.IdPeriodo}|{aula.Id}|{aula.DiasVentana}|{aula.Jornada}|" +
                                       $"{contexto.Bloques?.Count ?? 0}|" +
                                       $"{fInicioVentana:yyyyMMdd}|{fFinVentana:yyyyMMdd}|{unidadSesionAula}|{itemsPlan.Count}|" +
                                       $"{huellaPatrones}|{AgHuellaPlanVentana.Calcular(contexto, aula)}";
                    if (!_cachePlanes.TryGetValue(clavePlan, out var planRes))
                    {
                        if (LoggerHelper.ModoVerboso)
                        {
                            LoggerHelper.Debug($"[AG-RANURAS-TMP] FABRICA_CACHE_MISS corrida={idCorrida} aula={aula.Id} ref={aula.AulaRef} dias={aula.DiasVentana} jornada={aula.Jornada} plazo={fInicioVentana:yyyy-MM-dd}..{fFinVentana:yyyy-MM-dd} " +
                                "demanda=" + string.Join(";", itemsPlan.Select(p => $"{p.IdAsignatura}:{p.Horas}h:orden={p.OrdenCurricular}")) + " patrones=" + huellaPatrones);
                        }
                        var patronesPlan = new Dictionary<int, string>();
                        foreach (var it in itemsPlan)
                        {
                            string dp = ObtenerDiasPatron(it.IdAsignatura);
                            if (!string.IsNullOrWhiteSpace(dp)) patronesPlan[it.IdAsignatura] = dp;
                        }

                        planRes = _planificadorOleadas.Construir(
                            itemsPlan,
                            aula.DiasVentana,
                            fInicioVentana,
                            fFinVentana,
                            unidadSesionAula,
                            contexto.EsFestivo,
                            contexto.Bloques,      // Fase 1170
                            aula.Jornada,          // Fase 1170
                            patronesPlan, $"corrida={idCorrida} aula={aula.Id} ref={aula.AulaRef} origen=fabrica");
                        _cachePlanes[clavePlan] = planRes;
                    }

                    if (planRes.IsSuccess)
                    {
                        // Fase 1092: Preservar planes de las asignaturas factibles
                        mapaPlanes = planRes.Value.Planes.ToDictionary(p => p.IdAsignatura);
                        if (planRes.Value.Infactibles.Count > 0)
                        {
                            foreach (var idInf in planRes.Value.Infactibles)
                            {
                                asigsInfactibles.Add(idInf);
                            }
                            LoggerHelper.WarningLimitado(
                                $"PlanOleadasParcial|{contexto?.IdCorrida}|{aula.Id}",
                                $"[AgFabricaIndividuos] Aula {aula.Id} ({aula.AulaRef}): {planRes.Value.Infactibles.Count} asignaturas infactibles aisladas ({string.Join(",", planRes.Value.Infactibles)}); {mapaPlanes.Count} factibles conservadas.");
                        }
                    }
                    else
                    {
                        planInfactible = true;
                        LoggerHelper.WarningLimitado(
                            $"PlanOleadasFallo|{contexto?.IdCorrida}|{aula.Id}",
                            $"[AgFabricaIndividuos] Plan oleadas infactible para aula {aula.Id} ({aula.AulaRef}): {planRes.Error}");
                    }
                }
                else if (aula != null && string.IsNullOrWhiteSpace(aula.DiasVentana) && itemsPlan.Count > 0)
                {
                    // Fase 1093: Si el aula carece de DiasVentana, marcar cohorte como infactible en vez de planificar con default {2,3}
                    planInfactible = true;
                    LoggerHelper.WarningLimitado(
                        $"AulaSinVentana|{contexto?.IdCorrida}|{aula.Id}",
                        $"[AgFabricaIndividuos] Aula {aula.Id} ({aula.AulaRef}): sin DiasVentana declarados; cohorte en cuarentena SIN_PLAN_OLEADA (I1).");
                }

                int idxAsig = 0;
                foreach (var asig in asigsOrdenadas)
                {
                    // Seleccionar docente del pool habilitado (Fase 295 / Fase 462)
                    int idDocente = 0;
                    if (contexto.PoolPorAsignatura != null && contexto.PoolPorAsignatura.TryGetValue(asig.Id, out var poolAsig) && poolAsig.Count > 0)
                    {
                        var habilitados = poolAsig.Where(p => p.Activo == 1 && (contexto.IdTerceroVacante <= 0 || p.IdTercero != contexto.IdTerceroVacante)).ToList();
                        if (habilitados.Count > 0)
                        {
                            var orden = habilitados
                                .OrderBy(p => p.VecesAsignada)
                                .ThenBy(p => p.IdTercero)
                                .ToList();
                            if (!_rotacionDocente.TryGetValue(asig.Id, out int rot))
                                rot = _random.Next(orden.Count);
                            var elegido = orden[rot % orden.Count];
                            _rotacionDocente[asig.Id] = rot + 1;
                            idDocente = elegido.IdTercero;
                        }
                    }

                    int esVacante = idDocente > 0 ? 0 : 1;
                    string motivoVacante = string.Empty;
                    int candidatosPool = (contexto.PoolPorAsignatura != null && contexto.PoolPorAsignatura.TryGetValue(asig.Id, out var plCount))
                        ? plCount.Count(p => p.Activo == 1 && (contexto.IdTerceroVacante <= 0 || p.IdTercero != contexto.IdTerceroVacante))
                        : 0;

                    if (esVacante == 1)
                    {
                        idDocente = contexto.IdTerceroVacante;
                        var motivo = ClasificarMotivo(contexto, asig);
                        motivoVacante = motivo.Codigo();
                        candidatosPool = 0;
                    }

                    // Fase 1077, Fase 1085 & Fase 1099: aislar vacante por asignatura sin plan factible usando SIN_PLAN_OLEADA (I5)
                    bool cohorteSinPlan = planInfactible && asigsInfactibles.Count == 0;
                    bool esAsigSinPlan = asigsInfactibles.Contains(asig.Id) || cohorteSinPlan;
                    if (esAsigSinPlan)
                    {
                        esVacante = 1;
                        motivoVacante = AgMotivoVacante.SIN_PLAN_OLEADA.Codigo();
                        if (contexto.IdTerceroVacante > 0)
                        {
                            idDocente = contexto.IdTerceroVacante;
                        }
                    }

                    int idAula = aula != null
                        ? aula.Id
                        : (contexto.IdAulaCorrida > 0 ? contexto.IdAulaCorrida : fallbackAulaId);

                    int idSede = (aula != null && aula.IdSede > 0) ? aula.IdSede : asig.IdSede;
                    if (idSede <= 0)
                    {
                        if (contexto.AulasPorId != null && contexto.AulasPorId.TryGetValue(idAula, out var aSede) && aSede.IdSede > 0)
                        {
                            idSede = aSede.IdSede;
                        }
                        else if (contexto.IdSede > 0)
                        {
                            idSede = contexto.IdSede;
                        }
                    }

                    AgBloqueHorario bloque = bloquesValidos.Count > 0 ? bloquesValidos[_random.Next(bloquesValidos.Count)] : (contexto.Bloques.Count > 0 ? contexto.Bloques[0] : null);

                    DateTime fechaInicio = fInicioVentana;
                    DateTime fechaFin = fFinVentana;

                    AgAulaAsignatura aa = (aula != null && contexto.AulaAsignaturaPorAulaAsig != null)
                        ? (contexto.AulaAsignaturaPorAulaAsig.TryGetValue((aula.Id, asig.Id), out var aaExacta) ? aaExacta : null)
                        : (contexto.AulaAsignaturasPorId != null && contexto.AulaAsignaturasPorId.TryGetValue(asig.Id, out var aaDef) ? aaDef : null);

                    double totalHoras = ObtenerHorasObjetivo(asig);

                    string grupoEfectivo = (aula != null)
                        ? GrupoDeAula(aula)
                        : (!string.IsNullOrWhiteSpace(contexto.GrupoActual) ? contexto.GrupoActual : "A");
                    string claveGrupo = $"{asig.Id}|{idAula}|{grupoEfectivo}";

                    List<int> diasK;
                    IReadOnlyList<AgRanura> ranurasPatron = null;
                    IReadOnlyList<AgSesionPlanificada> sesionesPlanificadas = null;
                    if (mapaPlanes != null && mapaPlanes.TryGetValue(asig.Id, out var planOleada))
                    {
                        fechaInicio = planOleada.FechaInicio;
                        fechaFin = planOleada.FechaFin;
                        diasK = planOleada.DiasPatron.ToList();
                        ranurasPatron = planOleada.RanurasPatron;
                        sesionesPlanificadas = planOleada.Sesiones;
                    }
                    else
                    {
                        diasK = AgPlanificadorPatronSemanal.DiasKDe(aa, aula, idxAsig).ToList();
                    }

                    var progAula = aula != null ? contexto.Programas?.FirstOrDefault(p => p.Id == aula.IdPrograma) : null;
                    var nivelAula = AgNivelAcademicoCatalogo.Desde(progAula?.Modalidad);
                    int horasEfectivas = (int)AgUnidadesSemanales.HorasSesionAula(aula, nivelAula);

                    // Fase 1157: Resolver ranuras directamente si no vinieron en el plan
                    if (ranurasPatron == null || ranurasPatron.Count == 0)
                    {
                        ranurasPatron = AgPlanificadorPatronSemanal.RanurasDe(aa, aula, idxAsig, contexto.Bloques, horasEfectivas);
                    }

                    if (ranurasPatron != null && ranurasPatron.Count > 0)
                    {
                        int nRan = ranurasPatron.Count;
                        int totalSesiones = (int)Math.Ceiling(totalHoras / (double)horasEfectivas);
                        int[] sesionesRan = new int[nRan];
                        if (nRan == 1)
                        {
                            sesionesRan[0] = totalSesiones;
                        }
                        else
                        {
                            int s0 = (int)Math.Ceiling(totalSesiones / (double)nRan);
                            sesionesRan[0] = s0;
                            for (int rIdx = 1; rIdx < nRan; rIdx++)
                            {
                                int rem = totalSesiones - sesionesRan.Take(rIdx).Sum();
                                sesionesRan[rIdx] = (int)Math.Ceiling(rem / (double)(nRan - rIdx));
                            }
                        }

                        // Fase 1168: Distinct para evitar dias_patron duplicados (ej. "6,6")
                        string strDiasPatron = string.Join(",", ranurasPatron.Select(r => r.Dia).Distinct());
                        for (int i = 0; i < nRan; i++)
                        {
                            var ranura = ranurasPatron[i];
                            int diaIso = ranura.Dia;
                            var permitidasRanura = AgRanuras.FranjasDe(diaIso, aula?.Jornada, ranura.Horas);
                            // Fase 1293 / Fase 1309: exigir franja exacta SOLO en día 6 con 2+ franjas reales; L-V mantiene equivalencias
                            bool franjaExactaObligatoria = diaIso == 6 && permitidasRanura.Count > 1;

                            Func<AgBloqueHorario, bool> coincideRanura = b =>
                                b.DiaSemanaNumero == diaIso
                                && Math.Abs(b.DuracionHoras - ranura.Horas) < 0.001
                                && ((permitidasRanura.Count == 0 || franjaExactaObligatoria)
                                    ? string.Equals(b.Franja, ranura.Franja, StringComparison.OrdinalIgnoreCase)
                                    : permitidasRanura.Any(f => string.Equals(f, b.Franja, StringComparison.OrdinalIgnoreCase)));

                            AgBloqueHorario bloqueRanura = null;
                            if (bloquesValidos != null)
                            {
                                bloqueRanura = bloquesValidos.Where(coincideRanura).OrderBy(b => b.Id).FirstOrDefault();
                            }
                            if (bloqueRanura == null && contexto.Bloques != null)
                            {
                                string subAula = (aula?.Subunidad ?? string.Empty).Trim();
                                bloqueRanura = contexto.Bloques
                                    .Where(coincideRanura)
                                    .OrderByDescending(b => string.Equals((b.Subunidad ?? string.Empty).Trim(), subAula, StringComparison.OrdinalIgnoreCase))
                                    .ThenBy(b => b.Id)
                                    .FirstOrDefault();
                            }

                            // Fase 1300: catálogo sin bloque de la franja exacta -> conservar comportamiento legacy y avisar
                            if (bloqueRanura == null && franjaExactaObligatoria)
                            {
                                Func<AgBloqueHorario, bool> coincideLegacy = b =>
                                    b.DiaSemanaNumero == diaIso
                                    && Math.Abs(b.DuracionHoras - ranura.Horas) < 0.001
                                    && permitidasRanura.Any(f => string.Equals(f, b.Franja, StringComparison.OrdinalIgnoreCase));

                                bloqueRanura = (bloquesValidos ?? new List<AgBloqueHorario>()).Where(coincideLegacy).OrderBy(b => b.Id).FirstOrDefault()
                                            ?? (contexto.Bloques ?? new List<AgBloqueHorario>()).Where(coincideLegacy).OrderBy(b => b.Id).FirstOrDefault();

                                LoggerHelper.WarningLimitado(
                                    $"FranjaSinBloque|{contexto?.IdCorrida}|{idAula}|{diaIso}|{ranura.Franja}",
                                    bloqueRanura != null
                                        ? $"[Fabrica] Sin bloque {ranura.Horas}h franja '{ranura.Franja}' día {diaIso} (aula {idAula}); se aplica fallback legacy."
                                        : $"[Fabrica] Sin bloque {ranura.Horas}h franja '{ranura.Franja}' día {diaIso} (aula {idAula}); gen quedará en cuarentena (I6).");
                            }

                            var gen = new AgGen
                            {
                                IdAsignatura = asig.Id,
                                IdPrograma = asig.IdPrograma > 0 ? asig.IdPrograma : ((aula != null && aula.IdPrograma > 0) ? aula.IdPrograma : contexto.IdPrograma),
                                IdSede = idSede,
                                Semestre = (aula != null && !string.IsNullOrWhiteSpace(aula.Semestre)) ? aula.Semestre : (asig.Semestre ?? string.Empty),
                                Grupo = grupoEfectivo,
                                IdTercero = idDocente,
                                IdAula = idAula,
                                IndiceDiaPatron = i,
                                PatronSemanal = totalSesiones.ToString(),
                                DiasPatron = strDiasPatron,
                                ClaveGrupo = claveGrupo,
                                TotalHorasAsignatura = totalHoras,
                                FechaInicioEstimada = fechaInicio,
                                FechaFinEstimada = fechaFin,
                                EsFijo = 0,
                                EsVacante = esVacante,
                                MotivoVacante = motivoVacante,
                                CandidatosPool = candidatosPool,
                                Estado = AgEstadoGen.Valido,
                                Subunidad = contexto.Subunidad
                            };

                            if (aula != null) gen.AsignarAula(aula);
                            gen.IdSede = idSede;
                            gen.Grupo = grupoEfectivo;
                            gen.ClaveGrupo = claveGrupo;

                            if (bloqueRanura != null)
                            {
                                gen.AsignarBloque(bloqueRanura);
                                gen.HorasSemanales = bloqueRanura.DuracionHoras > 0 ? bloqueRanura.DuracionHoras : (double)horasEfectivas;
                                gen.EsFijoTemporal = true;
                            }
                            else
                            {
                                // Fase 1071 / Fase 1157: Sin bloque de la ranura -> cuarentena (I6)
                                gen.IdBloqueHorario = 0;
                                gen.DiaSemana = diaIso;
                                gen.HorasSemanales = horasEfectivas;
                                gen.EsVacante = 1;
                                gen.Estado = AgEstadoGen.Incompleto;
                                if (contexto.IdTerceroVacante > 0) gen.IdTercero = contexto.IdTerceroVacante;
                                gen.MotivoVacante = AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo();
                                LoggerHelper.WarningLimitado(
                                    $"BloqueRanuraInexistente|{contexto?.IdCorrida}|{gen.IdAula}|{diaIso}|{ranura.Franja}|{horasEfectivas}",
                                    $"[Fabrica] Aula {gen.IdAula}: sin bloque de {horasEfectivas}h para día {diaIso} franja '{ranura.Franja}'; gen en cuarentena (I6).");
                            }

                            // Fechas específicas precalculadas por ranura
                            gen.FechasGen = sesionesPlanificadas != null && sesionesPlanificadas.Count > 0
                                ? sesionesPlanificadas.Where(s => s.Ranura.Equals(ranura)).Select(s => s.Fecha).OrderBy(d => d).ToList()
                                : AgPlanificadorPatronSemanal.FechasDeRanura(diaIso, sesionesRan[i], fechaInicio, fechaFin, contexto.EsFestivo).ToList();
                            gen.HorasGen = gen.FechasGen.Count * horasEfectivas;
                            if (diagnosticoPrimeraSemilla)
                            {
                                LoggerHelper.Debug($"[AG-RANURAS-TMP] FABRICA_FECHAS corrida={idCorrida} aula={gen.IdAula} asig={gen.IdAsignatura} dia={diaIso} franja={ranura.Franja} fuente={(sesionesPlanificadas != null && sesionesPlanificadas.Count > 0 ? "plan_fechado" : "calendario_legacy")} fechas={gen.FechasGen.Count}");
                            }

                            // Fase 1099: la cuarentena I1 anula bloque y estado para impedir materialización indebida
                            if (cohorteSinPlan)
                            {
                                gen.IdBloqueHorario = 0;
                                gen.DiasPatron = string.Empty;
                                gen.Estado = AgEstadoGen.Incompleto;
                                gen.MotivoVacante = AgMotivoVacante.SIN_PLAN_OLEADA.Codigo();
                            }

                            individuo.Genes.Add(gen);
                        }

                        // Fase 416 / Fase 432: Verificación constructiva de Invariante I1 (Σ HorasGen == totalHoras)
                        double sumaG = individuo.Genes.Where(g => g.ClaveGrupo == claveGrupo).Sum(g => g.HorasGen);
                        if (Math.Abs(sumaG - totalHoras) > 0.001)
                        {
                            LoggerHelper.WarningLimitado(
                                $"InvarianteI1RanurasAjuste|{claveGrupo}",
                                $"[Fabrica] {claveGrupo}: Σ HorasGen={sumaG:0.#} != TotalHoras={totalHoras:0.#}; ajustando...");
                            double delta = totalHoras - sumaG;
                            foreach (var g in individuo.Genes.Where(g => g.ClaveGrupo == claveGrupo).OrderByDescending(g => g.HorasGen))
                            {
                                if (Math.Abs(delta) < 0.0001) break;
                                double nuevo = Math.Max(0.0, g.HorasGen + delta);
                                delta -= nuevo - g.HorasGen;
                                g.HorasGen = nuevo;
                            }
                        }
                    }
                    else if (diasK.Count == 1)
                    {
                        var gen = new AgGen
                        {
                            IdAsignatura = asig.Id,
                            IdPrograma = asig.IdPrograma > 0 ? asig.IdPrograma : ((aula != null && aula.IdPrograma > 0) ? aula.IdPrograma : contexto.IdPrograma),
                            IdSede = idSede,
                            Semestre = (aula != null && !string.IsNullOrWhiteSpace(aula.Semestre)) ? aula.Semestre : (asig.Semestre ?? string.Empty),
                            Grupo = grupoEfectivo,
                            IdTercero = idDocente,
                            IdAula = idAula,
                            TotalHorasAsignatura = totalHoras,
                            HorasGen = totalHoras,
                            ClaveGrupo = claveGrupo,
                            FechaInicioEstimada = fechaInicio,
                            FechaFinEstimada = fechaFin,
                            EsFijo = 0,
                            EsVacante = esVacante,
                            MotivoVacante = motivoVacante,
                            CandidatosPool = candidatosPool,
                            Estado = AgEstadoGen.Valido,
                            Subunidad = contexto.Subunidad
                        };

                        if (aula != null) gen.AsignarAula(aula);
                        gen.IdSede = idSede;
                        gen.Grupo = grupoEfectivo;
                        gen.ClaveGrupo = claveGrupo;

                        int diaIsoSingle = diasK[0];
                        AgBloqueHorario bloqueEfectivo = null;
                        if (aula != null && !string.IsNullOrWhiteSpace(aula.Jornada))
                        {
                            var fj = AgFranjaJornada.Resolve(aula.Jornada, horasEfectivas, diaIsoSingle);
                            if (fj.HasValue)
                            {
                                bloqueEfectivo = bloquesValidos.FirstOrDefault(b => b.DiaSemanaNumero == diaIsoSingle
                                    && string.Equals(b.Franja, fj.Value.Franja, StringComparison.OrdinalIgnoreCase)
                                    && Math.Abs(b.DuracionHoras - fj.Value.DuracionHoras) < 0.001);
                                if (bloqueEfectivo != null)
                                {
                                    gen.EsFijoTemporal = true;
                                }
                            }
                        }

                        if (bloqueEfectivo == null)
                        {
                            // Ronda 206: Exigir estrictamente la duración de sesión requerida (horasEfectivas), sin degradación a 8h en pregrado
                            var franjasSingle = bloquesValidos.Where(b => b.DiaSemanaNumero == diaIsoSingle && Math.Abs(b.DuracionHoras - horasEfectivas) < 0.001).ToList();
                            if (franjasSingle.Count == 0 && contexto.Bloques != null)
                            {
                                franjasSingle = contexto.Bloques.Where(b => b.DiaSemanaNumero == diaIsoSingle && Math.Abs(b.DuracionHoras - horasEfectivas) < 0.001).ToList();
                            }
                            if (franjasSingle.Count > 0)
                            {
                                bloqueEfectivo = franjasSingle[Math.Abs(asig.Id) % franjasSingle.Count];
                            }
                        }

                        if (bloqueEfectivo != null)
                        {
                            gen.AsignarBloque(bloqueEfectivo);
                            gen.HorasSemanales = bloqueEfectivo.DuracionHoras > 0 ? bloqueEfectivo.DuracionHoras : (double)horasEfectivas;
                        }
                        else
                        {
                            // Fase 1083: sin bloque de la duración del nivel → cuarentena; nunca bloque ajeno (I6)
                            gen.IdBloqueHorario = 0;
                            gen.DiaSemana = diaIsoSingle;
                            gen.HorasSemanales = horasEfectivas;
                            gen.EsVacante = 1;
                            gen.Estado = AgEstadoGen.Incompleto;
                            if (contexto.IdTerceroVacante > 0) gen.IdTercero = contexto.IdTerceroVacante;
                            gen.MotivoVacante = AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo();
                            LoggerHelper.WarningLimitado(
                                $"BloqueNivelInexistenteSingle|{contexto?.IdCorrida}|{gen.IdAula}|{diaIsoSingle}|{horasEfectivas}",
                                $"[Fabrica] Aula {gen.IdAula}: sin bloque de {horasEfectivas}h para día {diaIsoSingle}; gen en cuarentena (I6).");
                        }

                        // Fase 1099: la cuarentena I1 anula bloque y estado para impedir materialización indebida
                        if (cohorteSinPlan)
                        {
                            gen.IdBloqueHorario = 0;
                            gen.DiasPatron = string.Empty;
                            gen.Estado = AgEstadoGen.Incompleto;
                            gen.MotivoVacante = AgMotivoVacante.SIN_PLAN_OLEADA.Codigo();
                        }

                        individuo.Genes.Add(gen);
                    }
                    else
                    {
                        var progPatron = aula != null ? contexto.Programas?.FirstOrDefault(p => p.Id == aula.IdPrograma) : null;
                        double unidadPatron = aula != null
                            ? AgUnidadesSemanales.HorasSesionAula(aula, AgNivelAcademicoCatalogo.Desde(progPatron?.Modalidad))
                            : 4.0;

                        var resPatron = _planificadorPatron.Construir(diasK, (int)Math.Ceiling(totalHoras), fechaInicio, fechaFin, contexto.EsFestivo, unidadPatron);

                        // Ronda 206 / Fase 1058: Si falla la construcción del patrón, NO degradar silenciosamente a día único (genera sobre-sesiones/inviabilidad).
                        if (resPatron.IsFailure)
                        {
                            esVacante = 1;
                            motivoVacante = AgMotivoVacante.SIN_DISPONIBILIDAD.Codigo();
                            if (contexto.IdTerceroVacante > 0)
                            {
                                idDocente = contexto.IdTerceroVacante;
                            }
                        }

                        var patron = resPatron.IsSuccess ? resPatron.Value : null;
                        string strPatronSemanal = patron != null ? string.Join(",", patron.SesionesPorSemana) : string.Empty;
                        string strDiasPatron = string.Join(",", diasK);

                        for (int i = 0; i < diasK.Count; i++)
                        {
                            int diaIso = diasK[i];
                            AgBloqueHorario bloqueDia = null;
                            if (aula != null && !string.IsNullOrWhiteSpace(aula.Jornada))
                            {
                                var fj = AgFranjaJornada.Resolve(aula.Jornada, horasEfectivas, diaIso);
                                if (fj.HasValue)
                                {
                                    bloqueDia = bloquesValidos.FirstOrDefault(b => b.DiaSemanaNumero == diaIso
                                        && string.Equals(b.Franja, fj.Value.Franja, StringComparison.OrdinalIgnoreCase)
                                        && Math.Abs(b.DuracionHoras - fj.Value.DuracionHoras) < 0.001);
                                }
                            }

                            if (bloqueDia == null)
                            {
                                // Ronda 206: Exigir estrictamente la duración de sesión requerida (horasEfectivas), sin degradar a cualquier bloque
                                var franjasDia = bloquesValidos.Where(b => b.DiaSemanaNumero == diaIso && Math.Abs(b.DuracionHoras - horasEfectivas) < 0.001).ToList();
                                if (franjasDia.Count == 0 && contexto.Bloques != null)
                                {
                                    franjasDia = contexto.Bloques.Where(b => b.DiaSemanaNumero == diaIso && Math.Abs(b.DuracionHoras - horasEfectivas) < 0.001).ToList();
                                }

                                bloqueDia = franjasDia.Count > 0
                                    ? franjasDia[Math.Abs(asig.Id + i) % franjasDia.Count]
                                    : null;
                            }

                            var gen = new AgGen
                            {
                                IdAsignatura = asig.Id,
                                IdPrograma = asig.IdPrograma > 0 ? asig.IdPrograma : ((aula != null && aula.IdPrograma > 0) ? aula.IdPrograma : contexto.IdPrograma),
                                IdSede = idSede,
                                Semestre = (aula != null && !string.IsNullOrWhiteSpace(aula.Semestre)) ? aula.Semestre : (asig.Semestre ?? string.Empty),
                                Grupo = grupoEfectivo,
                                IdTercero = idDocente,
                                IdAula = idAula,
                                IndiceDiaPatron = i,
                                PatronSemanal = strPatronSemanal,
                                DiasPatron = strDiasPatron,
                                ClaveGrupo = claveGrupo,
                                TotalHorasAsignatura = totalHoras,
                                FechaInicioEstimada = fechaInicio,
                                FechaFinEstimada = fechaFin,
                                EsFijo = 0,
                                EsVacante = esVacante,
                                MotivoVacante = motivoVacante,
                                CandidatosPool = candidatosPool,
                                Estado = AgEstadoGen.Valido,
                                Subunidad = contexto.Subunidad
                            };

                            if (aula != null) gen.AsignarAula(aula);
                            gen.IdSede = idSede;
                            gen.Grupo = grupoEfectivo;
                            gen.ClaveGrupo = claveGrupo;

                            if (bloqueDia != null)
                            {
                                gen.AsignarBloque(bloqueDia);
                                gen.HorasSemanales = bloqueDia.DuracionHoras > 0 ? bloqueDia.DuracionHoras : (double)horasEfectivas;
                                if (aula != null && !string.IsNullOrWhiteSpace(aula.Jornada))
                                {
                                    gen.EsFijoTemporal = true;
                                }
                            }
                            else
                            {
                                // Fase 1071: Sin bloque de la duración del nivel: NO asignar bloque ajeno ni fallback (I6)
                                gen.IdBloqueHorario = 0;
                                gen.DiaSemana = diaIso;
                                gen.HorasSemanales = horasEfectivas;
                                gen.EsVacante = 1;
                                gen.Estado = AgEstadoGen.Incompleto;
                                if (contexto.IdTerceroVacante > 0) gen.IdTercero = contexto.IdTerceroVacante;
                                gen.MotivoVacante = AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo();
                                LoggerHelper.WarningLimitado(
                                    $"BloqueNivelInexistente|{contexto?.IdCorrida}|{gen.IdAula}|{diaIso}|{horasEfectivas}",
                                    $"[Fabrica] Aula {gen.IdAula}: sin bloque de {horasEfectivas}h para día {diaIso}; gen en cuarentena (I6).");
                            }

                            // Fase 414: Asignar fechas específicas precalculadas por el planificador a este gen
                            if (patron != null && patron.FechasValidas != null)
                            {
                                gen.FechasGen = patron.FechasValidas
                                    .Where(f => (f.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)f.DayOfWeek) == diaIso)
                                    .ToList();
                            }

                            // Fase 416 & F651 & F977: Calcular horas que corresponden a este gen usando la duración de su sesión
                            int sesionesEsteDia = (gen.FechasGen != null && gen.FechasGen.Count > 0)
                                ? gen.FechasGen.Count
                                : gen.SesionesDelGen();

                            double unidadDia = gen.HorasSemanales > 0 ? gen.HorasSemanales : (double)horasEfectivas;
                            double cargaDia = totalHoras / Math.Max(1, diasK.Count);
                            cargaDia = Math.Max(unidadDia, Math.Round(cargaDia / unidadDia) * unidadDia);

                            gen.HorasGen = sesionesEsteDia > 0
                                ? sesionesEsteDia * unidadDia
                                : cargaDia;

                            // Fase 1099: la cuarentena I1 anula bloque y estado para impedir materialización indebida
                            if (cohorteSinPlan)
                            {
                                gen.IdBloqueHorario = 0;
                                gen.DiasPatron = string.Empty;
                                gen.Estado = AgEstadoGen.Incompleto;
                                gen.MotivoVacante = AgMotivoVacante.SIN_PLAN_OLEADA.Codigo();
                            }

                            individuo.Genes.Add(gen);
                        }

                        // Fase 416 / Fase 432: Verificación y ajuste constructivo de Invariante I1 (Σ HorasGen == totalHoras)
                        double sumaGenes = individuo.Genes.Where(g => g.ClaveGrupo == claveGrupo).Sum(g => g.HorasGen);
                        if (Math.Abs(sumaGenes - totalHoras) > 0.001)
                        {
                            LoggerHelper.WarningLimitado(
                                $"I1Ajuste|{contexto?.IdCorrida}|{claveGrupo}",
                                $"[I1] Asig {asig.Id}: ΣHorasGen={sumaGenes} ≠ {totalHoras}; ajustando residuo de forma acotada");
                            double delta = totalHoras - sumaGenes;
                            foreach (var g in individuo.Genes.Where(g => g.ClaveGrupo == claveGrupo)
                                                             .OrderByDescending(g => g.HorasGen))
                            {
                                if (Math.Abs(delta) < 0.0001) break;
                                double nuevo = Math.Max(0.0, g.HorasGen + delta);
                                delta -= nuevo - g.HorasGen;
                                g.HorasGen = nuevo;
                            }
                        }
                    }

                    // Fase 1008: avanzar carril para la siguiente asignatura de la cohorte
                    idxAsig++;
                }
            }
            // Fase 481 & Fase 515: Trazabilidad de incidencias de semestre descartadas (deduplicada en AgDiagnosticoTraza)
            if (trazaEfectiva != null)
            {
                trazaEfectiva.Conteo("AsignaturasDescartadasPorSemestre", descartadasSemestre, 0);
                if (individuo.Genes.Count == 0 && contexto.Asignaturas != null && contexto.Asignaturas.Count > 0)
                {
                    trazaEfectiva.Advertencia("GenesVaciosPorSemestre", $"{descartadasSemestre} asignaturas descartadas por discrepancia de semestre entre aulas y pensum.");
                }
            }
            else if (descartadasSemestre > 0)
            {
                LoggerHelper.Debug($"[AgFabricaIndividuos] AsignaturasDescartadasPorSemestre: {descartadasSemestre}");
            }

            if (diagnosticoPrimeraSemilla)
            {
                foreach (var g in individuo.Genes)
                {
                    LoggerHelper.Debug($"[AG-RANURAS-TMP] SEMILLA corrida={idCorrida} aula={g.IdAula} asig={g.IdAsignatura} dia={g.DiaSemana} bloque={g.IdBloqueHorario} horas={g.HorasGen} fechas={g.FechasGen?.Count ?? 0} primera={g.FechasGen?.Select(d => (DateTime?)d).Min():yyyy-MM-dd} ultima={g.FechasGen?.Select(d => (DateTime?)d).Max():yyyy-MM-dd} rango={g.FechaInicioEstimada:yyyy-MM-dd}..{g.FechaFinEstimada:yyyy-MM-dd} estado={g.Estado}");
                }
            }
            return individuo;
        }

        /// <summary>
        /// Clasifica el motivo por el cual una asignatura queda en condición de vacante (Fase 319).
        /// Prioridad descendente: POOL_VACIO / SIN_CANDIDATOS -> NO_CUMPLE_REQUISITOS -> SIN_DISPONIBILIDAD.
        /// </summary>
        public static AgMotivoVacante ClasificarMotivo(AgContexto c, Asignatura a)
        {
            if (c == null || a == null) return AgMotivoVacante.SIN_CANDIDATOS;

            if (c.PoolPorAsignatura == null || !c.PoolPorAsignatura.TryGetValue(a.Id, out var pool) || pool.Count == 0)
            {
                return (c.Docentes?.Count ?? 0) == 0 ? AgMotivoVacante.POOL_VACIO : AgMotivoVacante.SIN_CANDIDATOS;
            }

            if (!pool.Any(p => p.Activo == 1))
            {
                return AgMotivoVacante.NO_CUMPLE_REQUISITOS;
            }

            // Si existen en pool con activo=1 pero ninguno habilitado para el programa/semestre
            if (!pool.Any(p => p.Activo == 1 && c.EsDocenteHabilitado(p.IdTercero, a.IdPrograma, a.Id, a.Semestre)))
            {
                return AgMotivoVacante.NO_CUMPLE_REQUISITOS;
            }

            return AgMotivoVacante.SIN_DISPONIBILIDAD;
        }
    }
}
