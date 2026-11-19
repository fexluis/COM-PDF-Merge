using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Compuerta centralizada y estricta de factibilidad y promoción para soluciones del Algoritmo Genético.
    /// Valida que la solución no presente restricciones duras (HARD == 0), no tenga vacantes
    /// (salvo autorización explícita justificada), que las lecturas a base de datos hayan sido exitosas,
    /// que no existan sesiones de genes desconocidos, y que exista cobertura exacta multiconjunto
    /// entre las sesiones proyectadas del plan/contexto sellado y las materializadas en staging.
    /// Valida que cada sesión coincida en inicio/fin con el bloque autorizado y que sus horas lectivas
    /// correspondan a la DuracionHoras del bloque (admitiendo bloques con pausas intermedias).
    /// </summary>
    public static class AgCompuertaPromocion
    {
        /// <summary>
        /// Valida la cobertura exacta multiconjunto entre la proyección de genes del plan sellado y las sesiones de staging.
        /// </summary>
        /// <param name="genes">Genes del individuo evaluado.</param>
        /// <param name="sesiones">Sesiones leídas desde ag_sesiones (staging).</param>
        /// <param name="contexto">Contexto sellado del AG con el catálogo de bloques y parámetros de calendario.</param>
        /// <param name="permitirVacantes">Indica si se permite la existencia de vacantes justificadas.</param>
        /// <returns>Result indicando éxito o la causa detallada del déficit o inconsistencia de cobertura.</returns>
        public static Result ValidarCobertura(
            IReadOnlyCollection<AgGen> genes,
            IReadOnlyCollection<AgSesion> sesiones,
            AgContexto contexto = null,
            bool permitirVacantes = false)
        {
            if (genes == null || genes.Count == 0)
            {
                return Result.Failure("No hay genes en el individuo para evaluar cobertura.");
            }

            if (sesiones == null || sesiones.Count == 0)
            {
                return Result.Failure("No se encontraron sesiones en staging.");
            }

            // 1. Rechazar sesiones asociadas a genes desconocidos que no pertenezcan al individuo
            var idsGenes = new HashSet<long>(genes.Select(g => (long)g.Id));
            var sesionesDesconocidas = sesiones.Where(s => !idsGenes.Contains(s.IdGen)).ToList();
            if (sesionesDesconocidas.Count > 0)
            {
                return Result.Failure($"Se encontraron {sesionesDesconocidas.Count} sesiones en staging asociadas a genes desconocidos que no pertenecen al individuo.");
            }

            // 2. Rechazar sesiones sin fecha asignada
            if (sesiones.Any(s => !s.FechaSesion.HasValue))
            {
                return Result.Failure("Existen sesiones en staging sin fecha asignada.");
            }

            // 3. Si no se permiten vacantes, rechazar sesiones marcadas como vacantes
            if (!permitirVacantes && sesiones.Any(s => s.EsVacante == 1))
            {
                return Result.Failure("Existen sesiones en staging marcadas como vacantes.");
            }

            // 4. Corrección [P2]: Validar carga horaria por gen usando estrictamente g.HorasGen.
            // Rechazar valores negativos o no finitos; permitir genes de cero horas sin sesiones; rechazar sesiones inesperadas.
            var sesionesPorGen = sesiones.GroupBy(s => s.IdGen).ToDictionary(g => g.Key, g => g.ToList());
            foreach (var g in genes)
            {
                double horasEsperadas = g.HorasGen;
                if (double.IsNaN(horasEsperadas) || double.IsInfinity(horasEsperadas) || horasEsperadas < 0.0)
                {
                    return Result.Failure($"El gen #{g.Id} (Asignatura {g.IdAsignatura}) tiene un valor de HorasGen no válido ({horasEsperadas}).");
                }

                sesionesPorGen.TryGetValue(g.Id, out var sesGen);
                int cantSesiones = sesGen?.Count ?? 0;

                if (horasEsperadas <= 0.001)
                {
                    if (cantSesiones > 0)
                    {
                        return Result.Failure($"El gen #{g.Id} (Asignatura {g.IdAsignatura}) tiene 0 horas asignadas pero presenta {cantSesiones} sesiones inesperadas en staging.");
                    }
                    continue;
                }

                // Corrección [P1]: Exigir fechas autorizadas en el gen (FechasGen) para evitar que el expansor
                // recurra al generador semanal legacy divergente.
                if (g.FechasGen == null || g.FechasGen.Count == 0)
                {
                    return Result.Failure($"El gen #{g.Id} (Asignatura {g.IdAsignatura}) carece de fechas autorizadas (FechasGen) en el plan sellado.");
                }

                if (cantSesiones == 0)
                {
                    return Result.Failure($"El gen #{g.Id} (Asignatura {g.IdAsignatura}) tiene {horasEsperadas:0.##}h asignadas pero no tiene sesiones en staging.");
                }

                double horasStaging = sesGen.Sum(s => s.Horas);
                if (Math.Abs(horasStaging - horasEsperadas) > 0.01)
                {
                    return Result.Failure($"Inconsistencia horaria en gen #{g.Id}: staging={horasStaging:0.##}h vs esperado={horasEsperadas:0.##}h.");
                }
            }

            // 5. Exigir plan/contexto sellado con catálogo de bloques para validar contra el bloque autorizado y proyección canónica
            if (contexto == null || contexto.BloquesPorId == null || contexto.BloquesPorId.Count == 0)
            {
                return Result.Failure("Se requiere el contexto/plan sellado con el catálogo de bloques para validar la cobertura de la corrida.");
            }

            // 6. Corrección [P1]: Validación horaria con referencia al bloque autorizado.
            // Comparar inicio/fin con el bloque autorizado; horas lectivas de la sesión contra DuracionHoras del bloque.
            // Esto previene sustituciones de franja (ej: 13:00-17:00 por 08:00-12:00) y admite bloques legítimos de 8h lectivas
            // cuyo intervalo de reloj sea mayor debido a pausas intermedias (ej: almuerzo).
            foreach (var g in genes)
            {
                if (g.HorasGen <= 0.001) continue;

                if (!contexto.BloquesPorId.TryGetValue(g.IdBloqueHorario, out var bloqueAutorizado) || bloqueAutorizado == null)
                {
                    return Result.Failure($"No se encontró el bloque horario autorizado #{g.IdBloqueHorario} para el gen #{g.Id}.");
                }

                if (!TimeSpan.TryParse(bloqueAutorizado.HoraInicio, CultureInfo.InvariantCulture, out var bloqueIni) ||
                    !TimeSpan.TryParse(bloqueAutorizado.HoraFin, CultureInfo.InvariantCulture, out var bloqueFin))
                {
                    return Result.Failure($"El bloque horario #{bloqueAutorizado.Id} tiene horas de inicio/fin inválidas ({bloqueAutorizado.HoraInicio} a {bloqueAutorizado.HoraFin}).");
                }

                sesionesPorGen.TryGetValue(g.Id, out var sesGen);
                if (sesGen != null)
                {
                    foreach (var s in sesGen)
                    {
                        if (s.IdAula != g.IdAula)
                        {
                            return Result.Failure($"Inconsistencia de recurso aula en sesión #{s.Id} (gen #{g.Id}): staging={s.IdAula} vs gen={g.IdAula}.");
                        }

                        if (!permitirVacantes)
                        {
                            if (s.IdTercero != g.IdTercero)
                            {
                                return Result.Failure($"Inconsistencia de recurso docente en sesión #{s.Id} (gen #{g.Id}): staging={s.IdTercero} vs gen={g.IdTercero}.");
                            }
                        }
                        else
                        {
                            if (g.EsVacante == 0 && s.IdTercero != g.IdTercero)
                            {
                                return Result.Failure($"Inconsistencia de recurso docente en sesión #{s.Id} (gen #{g.Id}): staging={s.IdTercero} vs gen={g.IdTercero}.");
                            }
                        }

                        if (!TryParseIntervalo(s.HorarioSesion, out var iniSes, out var finSes))
                        {
                            return Result.Failure($"La sesión #{s.Id} del gen #{g.Id} no tiene un formato de horario reconocible ('{s.HorarioSesion}').");
                        }

                        // Comparar hora de inicio y fin con el bloque autorizado
                        if (iniSes != bloqueIni || finSes != bloqueFin)
                        {
                            return Result.Failure(
                                $"Inconsistencia de intervalo en sesión #{s.Id} (gen #{g.Id}): " +
                                $"horario registrado '{s.HorarioSesion}' ({iniSes:hh\\:mm} a {finSes:hh\\:mm}) no coincide con el bloque autorizado #{bloqueAutorizado.Id} ({bloqueAutorizado.HoraInicio} a {bloqueAutorizado.HoraFin}).");
                        }

                        // Comparar horas lectivas de la sesión contra DuracionHoras del bloque autorizado
                        if (Math.Abs(s.Horas - bloqueAutorizado.DuracionHoras) > 0.01)
                        {
                            return Result.Failure(
                                $"Inconsistencia de horas lectivas en sesión #{s.Id} (gen #{g.Id}): " +
                                $"{s.Horas:0.##}h no coincide con la duración lectiva autorizada del bloque #{bloqueAutorizado.Id} ({bloqueAutorizado.DuracionHoras:0.##}h).");
                        }
                    }
                }
            }

            // 7. Proyección canónica por calendario de conflictos del plan sellado
            var calendario = AgCalendarioConflictos.Crear(genes, contexto);
            if (calendario.IsFailure)
            {
                return Result.Failure($"Fallo al proyectar calendario de conflictos para validar cobertura: {calendario.Error}");
            }

            // Clave multiconjunto completa: GenId | IdAula | IdTercero | Fecha | Intervalo | Horas
            string ClaveMulticonjunto(long gId, int idAula, int idTercero, DateTime fecha, TimeSpan ini, TimeSpan fin, double horas)
            {
                int terc = (permitirVacantes && idTercero <= 0) ? 0 : idTercero;
                return string.Format(CultureInfo.InvariantCulture,
                    "{0}|{1}|{2}|{3:yyyy-MM-dd}|{4:hh\\:mm}-{5:hh\\:mm}|{6:0.##}",
                    gId, idAula, terc, fecha.Date, ini, fin, horas);
            }

            var previstas = new List<string>();
            foreach (var o in calendario.Value)
            {
                double duracion = o.Gen.HorasSemanales;
                if (contexto.BloquesPorId.TryGetValue(o.Gen.IdBloqueHorario, out var bl) && bl.DuracionHoras > 0)
                {
                    duracion = bl.DuracionHoras;
                }
                previstas.Add(ClaveMulticonjunto(o.Gen.Id, o.Gen.IdAula, o.Gen.IdTercero, o.Fecha, o.Inicio, o.Fin, duracion));
            }

            var efectivas = new List<string>();
            foreach (var s in sesiones)
            {
                if (!TryParseIntervalo(s.HorarioSesion, out var iniSes, out var finSes))
                {
                    return Result.Failure($"La sesión #{s.Id} en staging tiene un formato de horario no reconocido: '{s.HorarioSesion}'.");
                }
                efectivas.Add(ClaveMulticonjunto(s.IdGen, s.IdAula, s.IdTercero, s.FechaSesion.Value, iniSes, finSes, s.Horas));
            }

            var previstasOrdenadas = previstas.OrderBy(k => k, StringComparer.Ordinal).ToList();
            var efectivasOrdenadas = efectivas.OrderBy(k => k, StringComparer.Ordinal).ToList();

            if (!previstasOrdenadas.SequenceEqual(efectivasOrdenadas))
            {
                return Result.Failure("Discrepancia en cobertura multiconjunto: las sesiones en staging no coinciden exactamente con la proyección de recursos (aula/docente), fechas, intervalos y horas del plan sellado.");
            }

            return Result.Success();
        }

        /// <summary>
        /// Ejecuta la compuerta estricta de promoción y retorna un Result con la causa exacta si la solución no es apta.
        /// </summary>
        /// <param name="corrida">Entidad de la corrida leída desde base de datos.</param>
        /// <param name="individuo">Individuo óptimo.</param>
        /// <param name="sesionesResult">Resultado de la lectura de sesiones desde staging.</param>
        /// <param name="contexto">Contexto sellado del AG.</param>
        /// <param name="permitirVacantes">Indica si se permite la existencia de vacantes bajo justificación explícita.</param>
        /// <returns>Result.Success() si cumple todos los requisitos de factibilidad y cobertura; de lo contrario Result.Failure.</returns>
        public static Result ValidarPromocion(
            AgCorrida corrida,
            AgIndividuo individuo,
            Result<IReadOnlyCollection<AgSesion>> sesionesResult,
            AgContexto contexto = null,
            bool permitirVacantes = false,
            ISet<int> autorizacionVacantes = null,
            ISet<string> autorizacionVacantesClaves = null)
        {
            if (individuo == null)
            {
                return Result.Failure("No hay individuo óptimo seleccionado.");
            }

            if (individuo.HardViolations > 0)
            {
                return Result.Failure($"La solución presenta {individuo.HardViolations} restricciones duras (HARD).");
            }

            // FASE 1874/1883/1885/1886, 2018, 2023, 2029, 2033, 2039, 2045 & 2049: Validación estricta y autorización explícita de vacantes no consentidas (solo lectura)
            var (asigsCorrida, clavesCorrida) = VacantesConsentidasEnCorrida(corrida);

            bool ConsentidaLocal(AgGen g) =>
                contexto != null ? contexto.EstaConsentida(g.IdAsignatura, g.IdSede)
                                 : asigsCorrida.Contains(g.IdAsignatura) || clavesCorrida.Contains($"{g.IdAsignatura}:{g.IdSede}");

            var noConsentidasGenes = (individuo.Genes ?? Enumerable.Empty<AgGen>())
                .Where(g => g.EsVacante == 1 || g.IdTercero <= 0)
                .Where(g => !ConsentidaLocal(g))
                .ToList();

            var noConsentidas = noConsentidasGenes
                .Select(g => g.IdAsignatura)
                .Distinct()
                .ToList();

            // Fase 2033 & 2049: Unificar predicado de consentimiento entre contexto y corrida
            bool tieneConsentidas = (contexto?.PermiteVacantes == true)
                                 || (corrida?.VacantesConsentidas?.Count ?? 0) > 0
                                 || asigsCorrida.Count > 0
                                 || clavesCorrida.Count > 0;
            bool permitir = permitirVacantes || tieneConsentidas;

            // Fase 2029, 2038 & 2045: La compuerta no autogenera consentimiento; exige autorización explícita sin mutar corrida
            if (noConsentidas.Count > 0)
            {
                var autorizadas = autorizacionVacantes ?? new HashSet<int>();
                var autorizadasClaves = autorizacionVacantesClaves ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (!permitirVacantes)
                {
                    return Result.Failure($"La solución contiene vacantes no consentidas en asignaturas: {string.Join(",", noConsentidas)}.");
                }

                // Fase 2045: Validar cobertura tanto por ID de asignatura como por clave granular asig:sede
                var genesNoCubiertos = noConsentidasGenes
                    .Where(g => !autorizadas.Contains(g.IdAsignatura) && !autorizadasClaves.Contains($"{g.IdAsignatura}:{g.IdSede}"))
                    .ToList();

                if (genesNoCubiertos.Count > 0)
                {
                    var faltantesAsigs = genesNoCubiertos.Select(g => g.IdAsignatura).Distinct().ToList();
                    return Result.Failure($"La autorización no cubre las asignaturas: {string.Join(",", faltantesAsigs)}. La justificación debe declarar 'asignaturas: {string.Join(",", noConsentidas)}'.");
                }

                LoggerHelper.Warning($"[Compuerta] Promoción autorizada con vacantes en: {string.Join(",", noConsentidas)}.");
            }
            else if (!permitir)
            {
                if (individuo.Vacantes > 0 || (individuo.Genes != null && individuo.Genes.Any(g => g.EsVacante == 1 || g.IdTercero <= 0)))
                {
                    return Result.Failure("La solución contiene genes en estado vacante sin consentimiento ni autorización.");
                }
            }

            if (corrida == null)
            {
                return Result.Failure("La información de la corrida en base de datos es nula.");
            }

            if (!permitir && corrida.Vacantes > 0)
            {
                return Result.Failure($"La corrida tiene {corrida.Vacantes} vacantes registradas.");
            }

            if (sesionesResult.IsFailure)
            {
                return Result.Failure($"Fallo al leer sesiones de staging: {sesionesResult.Error}");
            }

            var sesiones = sesionesResult.Value;
            if (sesiones == null || sesiones.Count == 0)
            {
                return Result.Failure("No se encontraron sesiones en staging para promover.");
            }

            if (!permitir)
            {
                int sesVac = sesiones.Count(s => s.EsVacante == 1);
                if (sesVac > 0)
                {
                    return Result.Failure($"Existen {sesVac} sesiones en staging en condición de vacante.");
                }
            }

            return ValidarCobertura(individuo.Genes, sesiones, contexto, permitir);
        }

        /// <summary>
        /// Evalúa si la solución es preliminarmente viable para iniciar el proceso de promoción
        /// (sin violaciones HARD y con sesiones de staging válidas), habilitando el botón de promoción
        /// en la UI para que solicite la autorización de vacantes si se requiere (Fase 2044).
        /// </summary>
        public static bool EsPromovibleSolicitable(
            AgCorrida corrida,
            AgIndividuo individuo,
            Result<IReadOnlyCollection<AgSesion>> sesionesResult)
        {
            if (corrida == null || individuo == null) return false;
            if (individuo.HardViolations > 0) return false;
            if (sesionesResult.IsFailure || sesionesResult.Value == null || sesionesResult.Value.Count == 0) return false;
            return true;
        }

        /// <summary>
        /// Evalúa si la solución es promovible a producción de acuerdo a la compuerta centralizada.
        /// </summary>
        public static bool EsPromovible(
            AgCorrida corrida,
            AgIndividuo individuo,
            Result<IReadOnlyCollection<AgSesion>> sesionesResult,
            AgContexto contexto = null,
            bool permitirVacantes = false,
            ISet<int> autorizacionVacantes = null,
            ISet<string> autorizacionVacantesClaves = null)
        {
            // FASE 1885/1886 & 2033: Ruta caliente sin deserialización JSON si la propiedad tipada ya está poblada
            bool permitir = permitirVacantes
                || (corrida?.VacantesConsentidas != null && corrida.VacantesConsentidas.Count > 0)
                || (contexto?.PermiteVacantes == true);
            if (!permitir)
            {
                var (asigsHist, clavesHist) = VacantesConsentidasEnCorrida(corrida);
                if (asigsHist.Count > 0 || clavesHist.Count > 0)
                {
                    permitir = true;
                    if (corrida != null)
                    {
                        corrida.VacantesConsentidas = asigsHist;
                    }
                    if (contexto != null && clavesHist.Count > 0 && (contexto.ClavesVacantesConsentidas == null || contexto.ClavesVacantesConsentidas.Count == 0))
                    {
                        contexto.ClavesVacantesConsentidas = clavesHist;
                    }
                }
            }

            if (corrida != null && corrida.VacantesConsentidas.Count > 0 && contexto != null &&
                (contexto.AsignaturasVacantesConsentidas == null || contexto.AsignaturasVacantesConsentidas.Count == 0))
            {
                contexto.AsignaturasVacantesConsentidas = corrida.VacantesConsentidas;
            }

            return ValidarPromocion(corrida, individuo, sesionesResult, contexto, permitirVacantes, autorizacionVacantes, autorizacionVacantesClaves).IsSuccess;
        }

        /// <summary>
        /// FASE 1874/1883/1886, 2034 y 2039: Obtiene las vacantes consentidas de la corrida (asignaturas y claves por sede).
        /// </summary>
        internal static (HashSet<int> Asignaturas, HashSet<string> Claves) VacantesConsentidasEnCorrida(AgCorrida corrida)
        {
            var asigs = new HashSet<int>();
            var claves = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (corrida == null)
            {
                return (asigs, claves);
            }

            if (corrida.VacantesConsentidas != null && corrida.VacantesConsentidas.Count > 0)
            {
                foreach (var id in corrida.VacantesConsentidas)
                {
                    asigs.Add(id);
                }
            }

            if (string.IsNullOrWhiteSpace(corrida.ParametrosJson))
            {
                return (asigs, claves);
            }

            try
            {
                var ser = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = 512 * 1024 };
                var dict = ser.DeserializeObject(corrida.ParametrosJson) as Dictionary<string, object>;
                if (dict != null && dict.TryGetValue("vacantesConsentidas", out var vacObj) && vacObj is System.Collections.IEnumerable list)
                {
                    foreach (var item in list)
                    {
                        if (item == null) continue;
                        string s = item.ToString().Trim();
                        if (s.Contains(':'))
                        {
                            claves.Add(s);
                            string[] partes = s.Split(':');
                            if (partes.Length > 0 && int.TryParse(partes[0], out int asigId))
                            {
                                asigs.Add(asigId);
                            }
                        }
                        else if (int.TryParse(s, out int id))
                        {
                            asigs.Add(id);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[AG] Error al deserializar vacantes consentidas de corrida #{corrida.Id}: {ex.Message}");
            }

            return (asigs, claves);
        }

        /// <summary>
        /// Intenta parsear cadenas de horario ("08:00 A 12:00", "8:00 AM - 12:00 PM") a intervalos de TimeSpan.
        /// </summary>
        private static bool TryParseIntervalo(string texto, out TimeSpan ini, out TimeSpan fin)
        {
            ini = TimeSpan.Zero;
            fin = TimeSpan.Zero;
            if (string.IsNullOrWhiteSpace(texto)) return false;

            var franjaRes = HorarioSesionParser.TryParse(texto);
            if (franjaRes.IsSuccess)
            {
                ini = franjaRes.Value.Inicio;
                fin = franjaRes.Value.Fin;
                return true;
            }

            string[] seps = new string[] { " A ", " a ", " - ", "-" };
            foreach (var sep in seps)
            {
                int idx = texto.IndexOf(sep, StringComparison.OrdinalIgnoreCase);
                if (idx > 0)
                {
                    string s1 = texto.Substring(0, idx).Trim();
                    string s2 = texto.Substring(idx + sep.Length).Trim();
                    if (TimeSpan.TryParse(s1, out ini) && TimeSpan.TryParse(s2, out fin))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
