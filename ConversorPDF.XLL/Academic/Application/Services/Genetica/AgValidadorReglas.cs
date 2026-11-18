using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Evaluador integral de reglas de restricción duras (R1-R7, W1) y blandas (S1-S5) para el AG.
    /// </summary>
    public class AgValidadorReglas : IAgValidadorReglas
    {
        // Fase 1899: Familia de motivos atribuibles a la disponibilidad docente (excluye vacantes estructurales de oleadas/bloques)
        private static readonly HashSet<string> MotivosDocentes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            AgMotivoVacante.SIN_CANDIDATOS.Codigo(),
            AgMotivoVacante.SIN_DISPONIBILIDAD.Codigo(),
            AgMotivoVacante.NO_CUMPLE_REQUISITOS.Codigo(),
            AgMotivoVacante.FUERA_DE_VENTANA.Codigo()
        };

        public void Evaluar(AgIndividuo individuo, AgContexto contexto)
        {
            if (individuo == null) throw new ArgumentNullException(nameof(individuo));
            if (contexto == null) throw new ArgumentNullException(nameof(contexto));

            individuo.Violaciones.Clear();
            int hardCount = 0;
            double softScore = 0.0;

            AgAula AulaDe(int idAula) =>
                idAula > 0 && contexto.AulasPorId != null && contexto.AulasPorId.TryGetValue(idAula, out var a) ? a : null;

            var genes = individuo.Genes;
            if (genes == null || genes.Count == 0)
            {
                int faltantes = contexto.Asignaturas != null ? contexto.Asignaturas.Count : 0;

                // Fase 1525: registrar el detalle de las asignaturas faltantes (antes sólo se guardaba el contador,
                // por lo que ag_violacion quedaba vacío y la UI mostraba HARD 0 con hard_violations > 0).
                if (contexto.Asignaturas != null)
                {
                    foreach (var asig in contexto.Asignaturas)
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R5_HabilitacionDocentePool,
                            esDura: true,
                            mensaje: $"La asignatura requerida '{asig.Nombre}' (Id={asig.Id}) no se encuentra programada en el cromosoma."));
                    }
                }

                individuo.HardViolations = faltantes;
                individuo.SoftViolations = 0;
                individuo.Estado = faltantes == 0 ? AgEstadoIndividuo.Factible : AgEstadoIndividuo.Infactible;
                return;
            }

            // --- Cobertura de asignaturas requeridas (Fase 1572: acotada al alcance de las aulas cubiertas) ---
            if (contexto.Asignaturas != null && contexto.Asignaturas.Count > 0)
            {
                var aulasCubiertas = new HashSet<int>(genes.Where(g => g.IdAula > 0).Select(g => g.IdAula));
                var plan = contexto.AulaAsignaturas ?? new List<AgAulaAsignatura>();
                var idsAsignadas = new HashSet<int>(genes.Select(g => g.IdAsignatura));

                foreach (var asig in contexto.Asignaturas)
                {
                    bool exigida = plan.Count == 0
                        || plan.Any(aa => aulasCubiertas.Contains(aa.IdAula) && aa.IdAsignatura == asig.Id);

                    if (exigida && !idsAsignadas.Contains(asig.Id))
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R5_HabilitacionDocentePool,
                            esDura: true,
                            mensaje: $"La asignatura requerida '{asig.Nombre}' (Id={asig.Id}) no se encuentra programada en el cromosoma."));
                        hardCount++;
                    }
                }
            }

            // Mapa para desplazamientos inter-sede (S2)
            var docenteDias = new Dictionary<string, HashSet<int>>(); // Docente-Dia -> Set de Sedes

            for (int i = 0; i < genes.Count; i++)
            {
                var gen = genes[i];

                // --- R5: Habilitación Docente en Pool (Fase 320: R5 relajada, sólo con docente real) ---
                if (gen.EsVacante == 0 && gen.IdTercero > 0)
                {
                    if (!contexto.EsDocenteHabilitado(gen.IdTercero, gen.IdPrograma, gen.IdAsignatura, gen.Semestre))
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R5_HabilitacionDocentePool,
                            esDura: true,
                            mensaje: $"El docente {gen.IdTercero} no está habilitado en el pool para la asignatura {gen.IdAsignatura} (Semestre: {gen.Semestre}).",
                            idGen: gen.Id));
                        hardCount++;
                    }
                }
                else if (gen.EsVacante == 1)
                {
                    // Fase 1896-C / Fase 1899 / Fase 1904: vacío = causa desconocida => se trata como docente (conservador)
                    bool motivoDocente = string.IsNullOrWhiteSpace(gen.MotivoVacante)
                                      || MotivosDocentes.Contains(gen.MotivoVacante);
                    if (motivoDocente)
                    {
                        bool vacanteAutorizada = contexto.AsignaturasVacantesConsentidas != null
                            && contexto.AsignaturasVacantesConsentidas.Contains(gen.IdAsignatura)
                            && string.Equals(gen.MotivoVacante, AgMotivoVacante.SIN_CANDIDATOS.Codigo(), StringComparison.OrdinalIgnoreCase);

                        if (!vacanteAutorizada)
                        {
                            individuo.Violaciones.Add(new AgViolacion(
                                AgReglaViolacion.R5_HabilitacionDocentePool,
                                esDura: true,
                                mensaje: $"Asignatura {gen.IdAsignatura} en vacante no autorizada (motivo: '{gen.MotivoVacante}').",
                                idGen: gen.Id));
                            hardCount++;
                        }
                    }
                }

                // --- R4: Compatibilidad Aula (Fase 1571: acepta Disponible y Aprobado) ---
                if (gen.IdAula > 0 && contexto.AulasPorId.TryGetValue(gen.IdAula, out var aula))
                {
                    bool aulaUsable = aula.Estado == AgEstadoAula.Disponible || aula.Estado == AgEstadoAula.Aprobado;
                    if (!aulaUsable)
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R4_CompatibilidadAula,
                            esDura: true,
                            mensaje: $"El aula '{aula.Codigo}' no está disponible (Estado: {aula.Estado}).",
                            idGen: gen.Id));
                        hardCount++;
                    }

                    // Fase 487/489 & Fase 664: Coherencia aula-cohorte. "SIN ASIGNAR"/vacío = comodín. Coherencia estricta de programa.
                    string semAula = AgFormatoAsignatura.SemestreRomano(aula.Semestre);
                    string semGen = AgFormatoAsignatura.SemestreRomano(gen.Semestre);
                    bool semOk = semAula == "SIN ASIGNAR" || semGen == "SIN ASIGNAR"
                              || string.Equals(semAula, semGen, StringComparison.OrdinalIgnoreCase);
                    bool grpOk = string.IsNullOrWhiteSpace(aula.Grupo) || string.IsNullOrWhiteSpace(gen.Grupo)
                              || string.Equals(aula.Grupo, gen.Grupo, StringComparison.OrdinalIgnoreCase);
                    bool progOk = aula.IdPrograma <= 0 || gen.IdPrograma <= 0
                              || aula.IdPrograma == gen.IdPrograma;
                    if (!semOk || !grpOk || !progOk)
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R4_CompatibilidadAula,
                            esDura: true,
                            mensaje: $"Aula {aula.Codigo} (prog={aula.IdPrograma},sem={aula.Semestre},grp={aula.Grupo}) fuera de cohorte (prog={gen.IdPrograma},sem={gen.Semestre},grp={gen.Grupo}).",
                            idGen: gen.Id));
                        hardCount++;
                    }
                }

                // Coherencia asignatura-programa (Fase 664)
                if (contexto.AsignaturasPorId != null && contexto.AsignaturasPorId.TryGetValue(gen.IdAsignatura, out var asigEval))
                {
                    if (asigEval.IdPrograma > 0 && gen.IdPrograma > 0 && asigEval.IdPrograma != gen.IdPrograma)
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R4_CompatibilidadAula,
                            esDura: true,
                            mensaje: $"Gen {gen.Id}: Asignatura '{asigEval.Nombre}' (prog={asigEval.IdPrograma}) cruzada con programa {gen.IdPrograma}.",
                            idGen: gen.Id));
                        hardCount++;
                    }
                }

                // --- W1: Ventana Semestral ---
                if (contexto.Ventana != null)
                {
                    if (gen.FechaInicioEstimada.HasValue && !contexto.Ventana.Contiene(gen.FechaInicioEstimada.Value))
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.W1_VentanaSemestralFueraDeRango,
                            esDura: true,
                            mensaje: $"La fecha de inicio {gen.FechaInicioEstimada.Value:yyyy-MM-dd} está fuera de la ventana semestral [{contexto.Ventana.FechaInicio:yyyy-MM-dd} - {contexto.Ventana.FechaFin:yyyy-MM-dd}].",
                            idGen: gen.Id));
                        hardCount++;
                    }
                    if (gen.FechaFinEstimada.HasValue && !contexto.Ventana.Contiene(gen.FechaFinEstimada.Value))
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.W1_VentanaSemestralFueraDeRango,
                            esDura: true,
                            mensaje: $"La fecha de fin {gen.FechaFinEstimada.Value:yyyy-MM-dd} está fuera de la ventana semestral [{contexto.Ventana.FechaInicio:yyyy-MM-dd} - {contexto.Ventana.FechaFin:yyyy-MM-dd}].",
                            idGen: gen.Id));
                        hardCount++;
                    }
                }

                // --- R9: Centinela VACANTE (Fase 464) ---
                if (contexto.IdTerceroVacante > 0)
                {
                    bool tieneCentinela = gen.IdTercero == contexto.IdTerceroVacante;
                    if (gen.EsVacante == 1 && !tieneCentinela)
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R9_VacanteSinCentinela,
                            esDura: true,
                            mensaje: $"Gen {gen.Id} vacante con idTercero={gen.IdTercero} ≠ VACANTE({contexto.IdTerceroVacante}).",
                            idGen: gen.Id));
                        hardCount++;
                    }
                    if (gen.EsVacante == 0 && tieneCentinela)
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R9_CentinelaEnNoVacante,
                            esDura: true,
                            mensaje: $"Gen {gen.Id} no vacante asignado al centinela VACANTE ({contexto.IdTerceroVacante}).",
                            idGen: gen.Id));
                        hardCount++;
                    }
                }

                // FASE 897 & 900 / Fase 1100 / Fase 1106: Bloque horario no asignado (0) en gen con carga horaria (excluye SIN_BLOQUE_NIVEL y SIN_PLAN_OLEADA ya tipificados en R7)
                if (gen.HorasGen > 0.0 && gen.IdBloqueHorario <= 0)
                {
                    if (gen.MotivoVacante != AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo() &&
                        gen.MotivoVacante != AgMotivoVacante.SIN_PLAN_OLEADA.Codigo())
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R14_BloqueHorarioNoAsignado,
                            esDura: true,
                            mensaje: $"Gen {gen.Id}: bloque horario no asignado ({gen.IdBloqueHorario}).",
                            idGen: gen.Id));
                        hardCount++;
                    }
                }

                // S2 / S4: Docente en el día (ignorar vacante o centinela)
                if (gen.IdBloqueHorario > 0 && gen.EsVacante == 0 && gen.IdTercero > 0 &&
                    (contexto.IdTerceroVacante <= 0 || gen.IdTercero != contexto.IdTerceroVacante) &&
                    contexto.BloquesPorId != null && contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out var blk))
                {
                    string docDiaKey = $"{gen.IdTercero}_{blk.DiaSemanaNumero}";
                    if (!docenteDias.TryGetValue(docDiaKey, out var sedesDoc))
                    {
                        sedesDoc = new HashSet<int>();
                        docenteDias[docDiaKey] = sedesDoc;
                    }
                    sedesDoc.Add(gen.IdSede);
                }
            }

            // Fase 523 & Fase 544: coherencia de docente por ClaveGrupo (R10)
            foreach (var gr in genes.Where(g => !string.IsNullOrWhiteSpace(g.ClaveGrupo)).GroupBy(g => g.ClaveGrupo))
            {
                if (gr.Select(g => g.IdTercero).Distinct().Count() > 1)
                {
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R10_CoherenciaDocenteGrupo,
                        esDura: true,
                        mensaje: $"Grupo '{gr.Key}': docentes distintos ({string.Join(",", gr.Select(g => g.IdTercero).Distinct())})."));
                    hardCount++;
                }
            }

            // Clave canónica de cohorte
            string Cohorte(AgGen g) => AgCalendarioConflictos.Cohorte(g);

            string FranjaDe(AgGen g) =>
                (contexto.BloquesPorId != null && g.IdBloqueHorario > 0
                 && contexto.BloquesPorId.TryGetValue(g.IdBloqueHorario, out var b) && b != null)
                    ? (b.Franja ?? string.Empty).Trim().ToUpperInvariant() : string.Empty;

            // --- R5: Máximo de asignaturas activas concurrentes por cohorte por día real (Fase 1056 / Fase 1087 / Fase 1167 / Fase 1570) ---
            var porGrupo = genes.Where(g => g.EsVacante == 0).GroupBy(x => Cohorte(x));
            foreach (var g in porGrupo)
            {
                // Fase 1087 / Fase 1167: Calibrar maxCarriles usando el modelo de ranuras por aula
                int maxCarriles = g.Select(x => AulaDe(x.IdAula))
                    .Where(a => a != null && !string.IsNullOrWhiteSpace(a.DiasVentana))
                    .Select(a => AgCarriles.Cantidad(AgRanuras.PorVentana(a.DiasVentana, contexto.Bloques, a.Jornada,
                        AgUnidadesSemanales.HorasSesionAula(a, AgNivelAcademicoCatalogo.Desde(contexto.Programas?.FirstOrDefault(p => p.Id == a.IdPrograma)?.Modalidad)))))
                    .DefaultIfEmpty(2)
                    .Max();

                // Mapear intervalos [min(FechaInicioEstimada), max(FechaFinEstimada)] por asignatura
                var asigsConFechas = g.Where(x => x.FechaInicioEstimada.HasValue && x.FechaFinEstimada.HasValue)
                                      .GroupBy(x => x.IdAsignatura)
                                      .Select(grp => new
                                      {
                                          IdAsignatura = grp.Key,
                                          Ini = grp.Min(x => x.FechaInicioEstimada.Value.Date),
                                          Fin = grp.Max(x => x.FechaFinEstimada.Value.Date)
                                      })
                                      .ToList();

                if (asigsConFechas.Count > 0)
                {
                    var minGlobal = asigsConFechas.Min(x => x.Ini);
                    var maxGlobal = asigsConFechas.Max(x => x.Fin);

                    DateTime? tramoIni = null;
                    DateTime? tramoFin = null;
                    int maxSobrecargaTramo = 0;

                    for (var dia = minGlobal; dia <= maxGlobal; dia = dia.AddDays(1))
                    {
                        int activas = asigsConFechas.Count(a => a.Ini <= dia && dia <= a.Fin);
                        if (activas > maxCarriles)
                        {
                            if (!tramoIni.HasValue)
                            {
                                tramoIni = dia;
                                maxSobrecargaTramo = activas;
                            }
                            else if (activas > maxSobrecargaTramo)
                            {
                                maxSobrecargaTramo = activas;
                            }
                            tramoFin = dia;
                        }
                        else
                        {
                            if (tramoIni.HasValue)
                            {
                                hardCount++;
                                individuo.Violaciones.Add(new AgViolacion(
                                    AgReglaViolacion.R5_MaxAsignaturasSemana,
                                    esDura: true,
                                    mensaje: $"La cohorte '{g.Key}' tiene {maxSobrecargaTramo} asignaturas concurrentes entre {tramoIni.Value:yyyy-MM-dd} y {tramoFin.Value:yyyy-MM-dd} (máximo permitido: {maxCarriles} carriles)."));
                                tramoIni = null;
                                tramoFin = null;
                                maxSobrecargaTramo = 0;
                            }
                        }
                    }

                    if (tramoIni.HasValue)
                    {
                        hardCount++;
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R5_MaxAsignaturasSemana,
                            esDura: true,
                            mensaje: $"La cohorte '{g.Key}' tiene {maxSobrecargaTramo} asignaturas concurrentes entre {tramoIni.Value:yyyy-MM-dd} y {tramoFin.Value:yyyy-MM-dd} (máximo permitido: {maxCarriles} carriles)."));
                    }
                }
            }

            // --- R6: Secuencia curricular por carril dentro de la cohorte (Fase 1056 / Fase 1075) ---
            foreach (var g in porGrupo)
            {
                // Fase 1075: Agrupar por IdAula y carril (DiasPatron) para no penalizar aulas distintas de la misma cohorte
                var porCarril = g.Where(x => x.FechaInicioEstimada.HasValue && x.FechaFinEstimada.HasValue)
                                 .GroupBy(x => $"{x.IdAula}|{(string.IsNullOrWhiteSpace(x.DiasPatron) ? "DEF" : x.DiasPatron)}");

                foreach (var carrilGrp in porCarril)
                {
                    var porAsignatura = carrilGrp.GroupBy(x => x.IdAsignatura).ToList();
                    if (porAsignatura.Count <= 1) continue;

                    var orden = porAsignatura
                        .OrderBy(asigGrp =>
                        {
                            var x = asigGrp.First();
                            AgAulaAsignatura aa = null;
                            if (contexto.AulaAsignaturaPorAulaAsig != null && x.IdAula > 0 &&
                                contexto.AulaAsignaturaPorAulaAsig.TryGetValue((x.IdAula, x.IdAsignatura), out var aaExacta))
                            {
                                aa = aaExacta;
                            }
                            else if (contexto.AulaAsignaturasPorId != null &&
                                     contexto.AulaAsignaturasPorId.TryGetValue(x.IdAsignatura, out var aaDef))
                            {
                                aa = aaDef;
                            }
                            return (aa != null && aa.OrdenCurricular > 0) ? aa.OrdenCurricular : 999;
                        })
                        .ThenBy(asigGrp => asigGrp.Min(x => x.FechaInicioEstimada.Value))
                        .ToList();

                    for (int idx = 1; idx < orden.Count; idx++)
                    {
                        var previaFin = orden[idx - 1].Max(x => x.FechaFinEstimada.Value);
                        var actualIni = orden[idx].Min(x => x.FechaInicioEstimada.Value);

                        if (actualIni < previaFin)
                        {
                            hardCount++;
                            individuo.Violaciones.Add(new AgViolacion(
                                AgReglaViolacion.R6_SecuenciaAsignaturas,
                                esDura: true,
                                mensaje: $"Solape de secuencia curricular en carril '{carrilGrp.Key}' de cohorte '{g.Key}': Asignatura {orden[idx].Key} inicia ({actualIni:yyyy-MM-dd}) antes de que finalice la previa {orden[idx - 1].Key} ({previaFin:yyyy-MM-dd}).",
                                idGen: orden[idx].First().Id));
                        }
                    }
                }
            }

            // --- R11: Duración de sesión según nivel académico (Fase 1056 / Fase 1076 / D4 / I6) ---
            foreach (var gen in genes.Where(x => x.EsVacante == 0 && x.IdBloqueHorario > 0))
            {
                var aulaGen = AulaDe(gen.IdAula);
                var progGen = contexto.Programas?.FirstOrDefault(p => p.Id == gen.IdPrograma);
                var nivel = AgNivelAcademicoCatalogo.Desde(progGen?.Modalidad);
                double duracionEsperada = AgUnidadesSemanales.HorasSesionAula(aulaGen, nivel);

                double duracionBloque = gen.HorasSemanales;
                AgBloqueHorario bGen = null;
                if (gen.IdBloqueHorario > 0 && contexto.BloquesPorId != null &&
                    contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out bGen))
                {
                    duracionBloque = bGen.DuracionHoras;
                }

                if (duracionBloque <= 0.001) continue;

                if (Math.Abs(duracionBloque - duracionEsperada) > 0.001)
                {
                    hardCount++;
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R11_DuracionSesionInvalida,
                        esDura: true,
                        mensaje: $"Gen {gen.Id} (Asig {gen.IdAsignatura}): duración de sesión ({duracionBloque}h) no corresponde a la requerida ({duracionEsperada}h) para nivel {nivel}.",
                        idGen: gen.Id));
                }
                else if (bGen != null && aulaGen != null)
                {
                    // Fase 1133 / Fase 1161: Validar que la franja corresponda a las ranuras válidas del día
                    var franjasDia = AgRanuras.FranjasDe(gen.DiaSemana, aulaGen.Jornada, duracionEsperada);
                    if (franjasDia.Count > 0 &&
                        !franjasDia.Any(f => string.Equals(f, bGen.Franja, StringComparison.OrdinalIgnoreCase)))
                    {
                        hardCount++;
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R11_DuracionSesionInvalida,
                            esDura: true,
                            mensaje: $"Gen {gen.Id}: franja {bGen.Franja} no es ranura del día {gen.DiaSemana} para aula {aulaGen.Id} (jornada {aulaGen.Jornada}; permitidas {string.Join("/", franjasDia)}).",
                            idGen: gen.Id));
                    }
                }
            }

            // --- R13: Oleadas no solapadas dentro del mismo carril (Fase 1056 / Fase 1075 / I4) ---
            foreach (var g in porGrupo)
            {
                // Fase 1075: Agrupar por IdAula y carril (DiasPatron)
                var porCarril = g.Where(x => x.FechaInicioEstimada.HasValue && x.FechaFinEstimada.HasValue)
                                 .GroupBy(x => $"{x.IdAula}|{(string.IsNullOrWhiteSpace(x.DiasPatron) ? "DEF" : x.DiasPatron)}");

                foreach (var carrilGrp in porCarril)
                {
                    var asigs = carrilGrp.GroupBy(x => x.IdAsignatura)
                        .Select(grp => new
                        {
                            IdAsignatura = grp.Key,
                            FechaInicio = grp.Min(x => x.FechaInicioEstimada.Value),
                            FechaFin = grp.Max(x => x.FechaFinEstimada.Value),
                            FirstGenId = grp.First().Id
                        })
                        .ToList();

                    for (int i = 0; i < asigs.Count; i++)
                    {
                        for (int j = i + 1; j < asigs.Count; j++)
                        {
                            bool solapanFechas = asigs[i].FechaInicio <= asigs[j].FechaFin && asigs[j].FechaInicio <= asigs[i].FechaFin;
                            if (solapanFechas)
                            {
                                hardCount++;
                                individuo.Violaciones.Add(new AgViolacion(
                                    AgReglaViolacion.R13_OleadasDispersas,
                                    esDura: true,
                                    mensaje: $"Solape de oleadas en carril '{carrilGrp.Key}' de cohorte '{g.Key}': Asignatura {asigs[i].IdAsignatura} ({asigs[i].FechaInicio:yyyy-MM-dd}..{asigs[i].FechaFin:yyyy-MM-dd}) solapa con Asignatura {asigs[j].IdAsignatura} ({asigs[j].FechaInicio:yyyy-MM-dd}..{asigs[j].FechaFin:yyyy-MM-dd}).",
                                    idGen: asigs[j].FirstGenId));
                            }
                        }
                    }
                }
            }

            // --- Fase 1077 & Fase 1085: Asignatura sin plan de oleada factible (I5) genera violación dura R7_CoberturaHorariaIncompleta ---
            foreach (var gen in genes.Where(g => g.EsVacante == 1 && g.MotivoVacante == AgMotivoVacante.SIN_PLAN_OLEADA.Codigo()))
            {
                hardCount++;
                individuo.Violaciones.Add(new AgViolacion(
                    AgReglaViolacion.R7_CoberturaHorariaIncompleta,
                    esDura: true,
                    mensaje: $"Asignatura {gen.IdAsignatura} (aula {gen.IdAula}) sin plan de carril/oleada factible (I5).",
                    idGen: gen.Id));
            }

            // --- Fase 1103: Asignatura sin bloque de duración del nivel (4h pregrado) genera violación dura R7_CoberturaHorariaIncompleta ---
            foreach (var gen in genes.Where(g => g.EsVacante == 1 &&
                         g.MotivoVacante == AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo()))
            {
                hardCount++;
                individuo.Violaciones.Add(new AgViolacion(
                    AgReglaViolacion.R7_CoberturaHorariaIncompleta,
                    esDura: true,
                    mensaje: $"Asignatura {gen.IdAsignatura} (aula {gen.IdAula}, día {gen.DiaSemana}): sin bloque de duración del nivel (4h pregrado).",
                    idGen: gen.Id));
            }

            // --- Fase 456 & Fase 1298: R8_DiasDuplicadosGrupo y R8_DiaFueraDePatron (Multi-Día) ---
            foreach (var gr in genes.Where(g => !string.IsNullOrWhiteSpace(g.ClaveGrupo)).GroupBy(g => g.ClaveGrupo))
            {
                // Fase 1298: R8 debe contar pares (Día, Franja); el sábado 4h admite 2 bloques (MANANA + TARDE)
                var pares = gr.Select(g => new { Dia = g.DiaSemana, Franja = FranjaDe(g) }).ToList();
                if (pares.Distinct().Count() != pares.Count)
                {
                    hardCount++;
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R8_DiasDuplicadosGrupo,
                        esDura: true,
                        mensaje: $"Grupo '{gr.Key}': bloques repetidos en el patrón ({string.Join(", ", pares.Select(p => $"D{p.Dia}/{p.Franja}"))}).",
                        idGen: gr.First().Id));
                }
                // Fase 988: cada carril valida su propio patrón (un grupo puede tener varias oleadas)
                var genFuera = gr.Where(g => !string.IsNullOrWhiteSpace(g.DiasPatron))
                                 .FirstOrDefault(g =>
                                 {
                                     var permitidosGen = g.DiasPatron.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                         .Select(s => int.TryParse(s, out int d) ? d : 0)
                                         .Where(d => d >= 1 && d <= 7)
                                         .ToList();
                                     return permitidosGen.Count > 0 && !permitidosGen.Contains(g.DiaSemana);
                                 });

                if (genFuera != null)
                {
                    hardCount++;
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R8_DiaFueraDePatron,
                        esDura: true,
                        mensaje: $"Grupo '{gr.Key}': día {genFuera.DiaSemana} fuera del patrón de su carril '{genFuera.DiasPatron}'.",
                        idGen: genFuera.Id));
                }
            }

            // --- Fase 1809: Evaluación temporal autoritativa con AgCalendarioConflictos (R1, R2, R3, R4, R7) ---
            if (contexto.BloquesPorId != null && contexto.BloquesPorId.Count > 0)
            {
                var genesProgramables = genes.Where(g => g.MotivoVacante != AgMotivoVacante.SIN_BLOQUE_NIVEL.Codigo()
                                                      && g.MotivoVacante != AgMotivoVacante.SIN_PLAN_OLEADA.Codigo()).ToList();
                var proyeccion = AgCalendarioConflictos.Crear(genesProgramables, contexto);
                if (proyeccion.IsFailure)
                {
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R14_BloqueHorarioNoAsignado, true,
                        "Calendario no verificable: " + proyeccion.Error));
                    individuo.HardViolations = hardCount + 1;
                    individuo.SoftViolations = softScore;
                    individuo.Estado = AgEstadoIndividuo.Infactible;
                    return;
                }
                var ocupaciones = proyeccion.Value;
                // Conservadas no son genes: ocupación fija con intervalo real, sin autocruces productivos.
                foreach (var o in ocupaciones)
                    foreach (var fija in contexto.SesionesConservadas.Where(s => s.Fecha.Date == o.Fecha &&
                        s.Inicio < o.Fin && o.Inicio < s.Fin))
                    {
                        if (fija.Clave.IdAula == o.Gen.IdAula ||
                            (AgCalendarioConflictos.DocenteReal(o.Gen, contexto) && fija.IdTercero == o.Gen.IdTercero))
                        {
                            individuo.Violaciones.Add(new AgViolacion(
                                fija.Clave.IdAula == o.Gen.IdAula ? AgReglaViolacion.R2_SuperposicionAula
                                    : AgReglaViolacion.R1_SuperposicionDocente,
                                esDura: true,
                                mensaje: $"Cruce con sesión CONSERVADA #{fija.Id} el {o.Fecha:yyyy-MM-dd} ({fija.Inicio}-{fija.Fin}).",
                                idGen: o.Gen.Id));
                            hardCount++;
                        }
                    }
                var serializador = new JavaScriptSerializer();

                void Registrar(AgReglaViolacion regla, IEnumerable<AgCalendarioConflictos.Ocupacion> implicadas, string mensaje)
                {
                    var lista = implicadas.ToList();
                    var indices = lista.Select(o => genes.IndexOf(o.Gen)).Distinct().ToArray();
                    string detalle = serializador.Serialize(new
                    {
                        indicesGenes = indices,
                        fecha = lista[0].Fecha.ToString("yyyy-MM-dd"),
                        intervalos = lista.Select(o => new
                        {
                            o.Gen.IdAula,
                            o.Gen.IdAsignatura,
                            o.Gen.IdTercero,
                            inicio = o.Inicio.ToString(),
                            fin = o.Fin.ToString()
                        }).ToArray()
                    });
                    individuo.Violaciones.Add(new AgViolacion(regla, true, mensaje, detalleJson: detalle));
                    hardCount++;
                }

                foreach (var dia in ocupaciones.GroupBy(o => o.Fecha))
                {
                    var lista = dia.ToList();
                    for (int i = 0; i < lista.Count; i++)
                        for (int j = i + 1; j < lista.Count; j++)
                        {
                            var a = lista[i]; var b = lista[j];
                            if (!AgCalendarioConflictos.SeCruzan(a, b)) continue;
                            var par = new[] { a, b };
                            string mensaje = $"Cruce real {dia.Key:yyyy-MM-dd}: asignaturas {a.Gen.IdAsignatura}/{b.Gen.IdAsignatura}.";
                            bool docente = AgCalendarioConflictos.DocenteReal(a.Gen, contexto)
                                && AgCalendarioConflictos.DocenteReal(b.Gen, contexto)
                                && a.Gen.IdTercero == b.Gen.IdTercero;
                            bool mismaCohorte = AgCalendarioConflictos.Cohorte(a.Gen) == AgCalendarioConflictos.Cohorte(b.Gen);
                            if (docente)
                            {
                                Registrar(AgReglaViolacion.R1_SuperposicionDocente, par, mensaje);
                                if (!mismaCohorte)
                                    Registrar(AgReglaViolacion.R7_TerceroOtroGrupoFechaHora, par, mensaje);
                            }
                            if (a.Gen.IdAula > 0 && a.Gen.IdAula == b.Gen.IdAula)
                                Registrar(AgReglaViolacion.R2_SuperposicionAula, par, mensaje);
                            if (mismaCohorte)
                                Registrar(AgReglaViolacion.R3_SuperposicionGrupo, par, mensaje);
                        }
                }

                // R4
                foreach (var dia in ocupaciones.Where(o => o.Gen.EsVacante == 0).GroupBy(o => new { Cohorte = AgCalendarioConflictos.Cohorte(o.Gen), o.Fecha }))
                {
                    foreach (var franja in dia.GroupBy(o => o.Franja.Trim().ToUpperInvariant()))
                        if (franja.Count() > 1)
                            Registrar(AgReglaViolacion.R4_DosAsignaturasMismoDia, franja, $"Más de una sesión en la franja {franja.Key} el {dia.Key.Fecha:yyyy-MM-dd}.");
                    int tope = dia.Key.Fecha.DayOfWeek == DayOfWeek.Saturday ? 2 : 1;
                    if (dia.Select(o => o.Franja.Trim().ToUpperInvariant()).Distinct().Count() > tope)
                        Registrar(AgReglaViolacion.R4_DosAsignaturasMismoDia, dia, $"Tope diario excedido el {dia.Key.Fecha:yyyy-MM-dd}.");
                }

                // R7 Externa
                foreach (var o in ocupaciones.Where(o => AgCalendarioConflictos.DocenteReal(o.Gen, contexto)))
                    if (contexto.TerceroOcupado(o.Gen.IdTercero, o.Fecha, new ConversorPDF.XLL.Academic.Domain.Models.FranjaHoraria(o.Inicio, o.Fin)))
                        Registrar(AgReglaViolacion.R7_TerceroOtroGrupoFechaHora, new[] { o }, $"Ocupación externa del docente {o.Gen.IdTercero} el {o.Fecha:yyyy-MM-dd}.");
            }



            // --- Evaluar S2: Desplazamientos inter-sede en el mismo día ---
            foreach (var kvp in docenteDias)
            {
                if (kvp.Value.Count > 1)
                {
                    // El docente tiene que ir a 2 sedes distintas en el mismo día
                    softScore += contexto.Pesos.PesoS2_DesplazamientoInterSede * (kvp.Value.Count - 1);
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.S2_DesplazamientoInterSede,
                        esDura: false,
                        mensaje: $"El docente tiene clases en {kvp.Value.Count} sedes distintas en el mismo día."));
                }
            }

            // --- Fase 320: Penalización suave por Asignaturas Vacantes (R5 relajada) ---
            int vacantesCount = genes.Count(g => g.EsVacante == 1 || g.IdTercero <= 0);
            if (vacantesCount > 0)
            {
                double penalizacionVacante = vacantesCount * (contexto.Pesos?.PesoVacante ?? 50.0);
                softScore += penalizacionVacante;
                individuo.Violaciones.Add(new AgViolacion(
                    AgReglaViolacion.R5_HabilitacionDocentePool,
                    esDura: false,
                    mensaje: $"La solución cuenta con {vacantesCount} asignaturas en condición de VACANTE (penalización suave: {penalizacionVacante:F1})."));
            }

            // Fase 967 / F957 / F970 / F971: solape por intervalo matemático (mismo día, horas cruzadas)
            // Fase 971: particionar por día (DiaSemanaNumero) antes de evaluar pares para mitigar complejidad O(n²) y evitar congelamiento en poblaciones reales.
            if (contexto.BloquesPorId != null && contexto.BloquesPorId.Count > 0)
            {
                var conBloque = genes
                    .Where(g => g.IdBloqueHorario > 0 && contexto.BloquesPorId.ContainsKey(g.IdBloqueHorario))
                    .Select(g => new { Gen = g, Bloque = contexto.BloquesPorId[g.IdBloqueHorario] })
                    .ToList();

                foreach (var dia in conBloque.GroupBy(x => x.Bloque.DiaSemanaNumero))
                {
                    var lista = dia.ToList();
                    for (int i = 0; i < lista.Count; i++)
                    {
                        var g1 = lista[i].Gen;
                        var b1 = lista[i].Bloque;

                        for (int j = i + 1; j < lista.Count; j++)
                        {
                            var g2 = lista[j].Gen;
                            var b2 = lista[j].Bloque;

                            // Solape en conflicto si comparten ámbito o recurso (misma aula, mismo docente o misma cohorte/grupo)
                            // Fase 970: discriminar cohorte por IdPrograma e IdSede para evitar falsos solapes entre programas o sedes distintas
                            bool mismoRecurso = g1.IdAula == g2.IdAula
                                || (g1.IdTercero > 0 && g1.IdTercero == g2.IdTercero && g1.EsVacante == 0 && g2.EsVacante == 0)
                                || (!string.IsNullOrWhiteSpace(g1.ClaveGrupo) && g1.ClaveGrupo == g2.ClaveGrupo)
                                || (g1.IdPrograma == g2.IdPrograma && g1.IdSede == g2.IdSede
                                    && !string.IsNullOrWhiteSpace(g1.Grupo) && !string.IsNullOrWhiteSpace(g1.Semestre)
                                    && g1.Grupo == g2.Grupo && g1.Semestre == g2.Semestre);

                            if (!mismoRecurso) continue;

                            // Fase 1569: el mismo bloque ya lo cubren R1/R2/R3; además validar solape temporal entre ventanas
                            if (g1.IdBloqueHorario != g2.IdBloqueHorario
                                && AgBloqueSolape.SeSolapan(b1, b2)
                                && AgVentanaSolape.SeSolapan(g1.FechaInicioEstimada, g1.FechaFinEstimada,
                                                             g2.FechaInicioEstimada, g2.FechaFinEstimada))
                            {
                                hardCount++;
                                individuo.Violaciones.Add(new AgViolacion(
                                    AgReglaViolacion.R14_BloqueSolapadoPorIntervalo,
                                    esDura: true,
                                    mensaje: $"Bloques solapados ({b1.Franja} {b1.HoraInicio}-{b1.HoraFin} vs {b2.Franja} {b2.HoraInicio}-{b2.HoraFin}) en gen {g1.Id} y gen {g2.Id}.",
                                    idGen: g1.Id));
                                break;
                            }
                        }
                    }
                }
            }

            individuo.HardViolations = hardCount;
            individuo.SoftViolations = softScore;
            individuo.Estado = hardCount == 0 ? AgEstadoIndividuo.Factible : AgEstadoIndividuo.Infactible;
        }
    }
}
