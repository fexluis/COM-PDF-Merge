using System;
using System.Collections.Generic;
using System.Linq;
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
                individuo.HardViolations = faltantes;
                individuo.SoftViolations = 0;
                individuo.Estado = faltantes == 0 ? AgEstadoIndividuo.Factible : AgEstadoIndividuo.Infactible;
                return;
            }

            // --- Cobertura de asignaturas requeridas ---
            if (contexto.Asignaturas != null && contexto.Asignaturas.Count > 0)
            {
                var idsAsignadas = new HashSet<int>(genes.Select(g => g.IdAsignatura));
                foreach (var asig in contexto.Asignaturas)
                {
                    if (!idsAsignadas.Contains(asig.Id))
                    {
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R5_HabilitacionDocentePool,
                            esDura: true,
                            mensaje: $"La asignatura requerida '{asig.Nombre}' (Id={asig.Id}) no se encuentra programada en el cromosoma."));
                        hardCount++;
                    }
                }
            }

            // Mapas para detección eficiente de colisiones
            var docenteBloques = new Dictionary<string, List<AgGen>>();
            var aulaBloques = new Dictionary<string, List<AgGen>>();
            var grupoBloques = new Dictionary<string, List<AgGen>>();
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

                // --- R4: Compatibilidad Aula ---
                if (gen.IdAula > 0 && contexto.AulasPorId.TryGetValue(gen.IdAula, out var aula))
                {
                    if (aula.Estado != AgEstadoAula.Disponible)
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
                    string semGen  = AgFormatoAsignatura.SemestreRomano(gen.Semestre);
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

                // Indexar para colisiones horarias
                if (gen.IdBloqueHorario > 0)
                {
                    // R1: Docente superpuesto (ignorar si es vacante o centinela)
                    if (gen.EsVacante == 0 && gen.IdTercero > 0 && (contexto.IdTerceroVacante <= 0 || gen.IdTercero != contexto.IdTerceroVacante))
                    {
                        string docKey = $"{gen.IdTercero}_{gen.IdBloqueHorario}";
                        if (!docenteBloques.TryGetValue(docKey, out var lDoc))
                        {
                            lDoc = new List<AgGen>();
                            docenteBloques[docKey] = lDoc;
                        }
                        lDoc.Add(gen);
                    }

                    // R2: Aula superpuesta
                    if (gen.IdAula > 0)
                    {
                        string aulaKey = $"{gen.IdAula}_{gen.IdBloqueHorario}";
                        if (!aulaBloques.TryGetValue(aulaKey, out var lAula))
                        {
                            lAula = new List<AgGen>();
                            aulaBloques[aulaKey] = lAula;
                        }
                        lAula.Add(gen);
                    }

                    // R3: Grupo de estudiantes superpuesto (Fase 532: incluir IdSede para no colisionar grupos homónimos en distintas sedes)
                    string grpKey = $"{gen.IdSede}_{gen.IdPrograma}_{gen.Semestre}_{gen.Grupo}_{gen.IdBloqueHorario}";
                    if (!grupoBloques.TryGetValue(grpKey, out var lGrp))
                    {
                        lGrp = new List<AgGen>();
                        grupoBloques[grpKey] = lGrp;
                    }
                    lGrp.Add(gen);

                    // S2 / S4: Docente en el día (ignorar vacante o centinela)
                    if (gen.EsVacante == 0 && gen.IdTercero > 0 && (contexto.IdTerceroVacante <= 0 || gen.IdTercero != contexto.IdTerceroVacante) && contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out var blk))
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

            // --- Evaluar colisiones R1 (Docente) ---
            foreach (var kvp in docenteBloques)
            {
                if (kvp.Value.Count > 1)
                {
                    int cruces = kvp.Value.Count - 1;
                    hardCount += cruces;
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R1_SuperposicionDocente,
                        esDura: true,
                        mensaje: $"El docente tiene {kvp.Value.Count} asignaturas asignadas simultáneamente en el bloque {kvp.Value[0].IdBloqueHorario}."));
                }
            }

            // --- Evaluar colisiones R2 (Aula) ---
            foreach (var kvp in aulaBloques)
            {
                if (kvp.Value.Count > 1)
                {
                    int cruces = kvp.Value.Count - 1;
                    hardCount += cruces;
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R2_SuperposicionAula,
                        esDura: true,
                        mensaje: $"El aula {kvp.Value[0].IdAula} está asignada a {kvp.Value.Count} cursos en el mismo bloque horario {kvp.Value[0].IdBloqueHorario}."));
                }
            }

            // --- Evaluar colisiones R3 (Grupo) ---
            foreach (var kvp in grupoBloques)
            {
                if (kvp.Value.Count > 1)
                {
                    int cruces = kvp.Value.Count - 1;
                    hardCount += cruces;
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R3_SuperposicionGrupo,
                        esDura: true,
                        mensaje: $"El grupo tiene {kvp.Value.Count} clases en simultáneo en el bloque horario {kvp.Value[0].IdBloqueHorario}."));
                }
            }

            // F1027 / F939 / F953: clave canónica de cohorte (evita colisiones falsas entre sedes/programas homónimos)
            string Cohorte(AgGen g)
            {
                string sem = AgFormatoAsignatura.SemestreRomano(g.Semestre);
                if (string.IsNullOrWhiteSpace(sem) || sem == "SIN ASIGNAR")
                {
                    sem = (g.Semestre ?? string.Empty).Trim();
                }
                return $"{g.IdSede}|{g.IdPrograma}|{sem}|{(g.Grupo ?? string.Empty).Trim()}";
            }

            // --- R4: Unicidad por (Cohorte, Día, Franja) y tope de bloques diarios (Fase 1162) ---
            string FranjaDe(AgGen g) =>
                (contexto.BloquesPorId != null && g.IdBloqueHorario > 0
                 && contexto.BloquesPorId.TryGetValue(g.IdBloqueHorario, out var b) && b != null)
                    ? (b.Franja ?? string.Empty).Trim().ToUpperInvariant() : string.Empty;

            var porGrupoDiaFranja = genes.Where(g => g.EsVacante == 0)
                                         .GroupBy(g => new { Cohorte = Cohorte(g), g.DiaSemana, Franja = FranjaDe(g) });

            foreach (var gdf in porGrupoDiaFranja.Where(x => x.Count() > 1))
            {
                hardCount += (gdf.Count() - 1);
                individuo.Violaciones.Add(new AgViolacion(
                    AgReglaViolacion.R4_DosAsignaturasMismoDia,
                    esDura: true,
                    mensaje: $"La cohorte '{gdf.Key.Cohorte}' tiene {gdf.Count()} asignaturas el día {gdf.Key.DiaSemana} en la misma franja '{gdf.Key.Franja}'."));
            }

            // Tope por día: 2 en SÁBADO (M+T si 4h), 1 en el resto (Fase 1162 / Fase 1175)
            foreach (var gd in genes.Where(g => g.EsVacante == 0).GroupBy(g => new { Cohorte = Cohorte(g), g.DiaSemana }))
            {
                int tope = gd.Key.DiaSemana == 6 ? 2 : 1;
                var franjas = gd.Select(FranjaDe).Distinct().ToList();
                if (franjas.Count > tope)
                {
                    hardCount++;
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R4_DosAsignaturasMismoDia,
                        esDura: true,
                        mensaje: $"La cohorte '{gd.Key.Cohorte}' tiene {franjas.Count} bloques el día {gd.Key.DiaSemana} ({string.Join("/", franjas)}); tope {tope}."));
                }
            }


            // --- R5: Máximo de asignaturas activas concurrentes por semana ISO por cohorte (Fase 1056 / Fase 1087 / Fase 1167) ---
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

                // Mapear semanas activas por asignatura
                var asigsConFechas = g.Where(x => x.FechaInicioEstimada.HasValue && x.FechaFinEstimada.HasValue)
                                      .GroupBy(x => x.IdAsignatura)
                                      .ToList();

                var semanasAsig = new Dictionary<int, HashSet<int>>();
                foreach (var asigGrp in asigsConFechas)
                {
                    var fIni = asigGrp.Min(x => x.FechaInicioEstimada.Value);
                    var fFin = asigGrp.Max(x => x.FechaFinEstimada.Value);
                    var semanas = new HashSet<int>();
                    for (var d = fIni.Date; d <= fFin.Date; d = d.AddDays(7))
                    {
                        semanas.Add(AgSemanaIso.Clave(d));
                    }
                    semanas.Add(AgSemanaIso.Clave(fFin.Date));
                    semanasAsig[asigGrp.Key] = semanas;
                }

                // Evaluar concurrencia por cada semana ISO activa
                var todasSemanas = semanasAsig.Values.SelectMany(s => s).Distinct().ToList();
                foreach (var sem in todasSemanas)
                {
                    int concurrentes = semanasAsig.Count(kvp => kvp.Value.Contains(sem));
                    if (concurrentes > maxCarriles)
                    {
                        hardCount += (concurrentes - maxCarriles);
                        individuo.Violaciones.Add(new AgViolacion(
                            AgReglaViolacion.R5_MaxAsignaturasSemana,
                            esDura: true,
                            mensaje: $"La cohorte '{g.Key}' tiene {concurrentes} asignaturas activas en la semana ISO {sem} (máximo permitido: {maxCarriles} carriles)."));
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

            // --- R7: Tercero en dos grupos a la misma fecha/bloque horario ---
            var docenteBloqueGrupos = genes.Where(g => g.EsVacante == 0 && g.IdTercero > 0 && g.IdBloqueHorario > 0)
                                           .GroupBy(g => new { g.IdTercero, g.IdBloqueHorario });
            foreach (var col in docenteBloqueGrupos)
            {
                int distinctGrupos = col.Select(Cohorte).Distinct().Count();
                if (distinctGrupos > 1)
                {
                    hardCount += (distinctGrupos - 1);
                    individuo.Violaciones.Add(new AgViolacion(
                        AgReglaViolacion.R7_TerceroOtroGrupoFechaHora,
                        esDura: true,
                        mensaje: $"El docente {col.Key.IdTercero} está asignado a {distinctGrupos} cohortes distintas en el mismo bloque horario {col.Key.IdBloqueHorario}."));
                }
            }

            // --- R7 (continuación): Ocupación externa del docente (sesiones de producción o staging aprobadas) ---
            if (contexto.OcupacionExterna != null && contexto.OcupacionExterna.Count > 0)
            {
                foreach (var gen in genes.Where(g => g.EsVacante == 0 && g.IdTercero > 0 && g.IdBloqueHorario > 0))
                {
                    if (gen.FechaInicioEstimada.HasValue && contexto.BloquesPorId != null && contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out var bloque))
                    {
                        if (TimeSpan.TryParse(bloque.HoraInicio, out var tIni) && TimeSpan.TryParse(bloque.HoraFin, out var tFin) && tIni < tFin)
                        {
                            var franja = new ConversorPDF.XLL.Academic.Domain.Models.FranjaHoraria(tIni, tFin);
                            if (contexto.TerceroOcupado(gen.IdTercero, gen.FechaInicioEstimada.Value, franja))
                            {
                                hardCount++;
                                individuo.Violaciones.Add(new AgViolacion(
                                    AgReglaViolacion.R7_TerceroOtroGrupoFechaHora,
                                    esDura: true,
                                    mensaje: $"El docente {gen.IdTercero} ya tiene una sesión externa en producción el {gen.FechaInicioEstimada.Value:yyyy-MM-dd} en la franja {bloque.HoraInicio} - {bloque.HoraFin}.",
                                    idGen: gen.Id));
                            }
                        }
                    }
                }
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

                            if (AgBloqueSolape.SeSolapan(b1, b2))
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
