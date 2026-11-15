using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Services.Genetica;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>
    /// Propone pasos de ventana y reclasificación de patrones con verificación de huellas transaccionales (Fases 1740-1750).
    /// </summary>
    internal sealed class AgEscaladorRanuraAula
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public AgEscaladorRanuraAula(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory;
        }

        internal sealed class Decision
        {
            public int IdAula { get; set; }
            public string AulaRef { get; set; }
            public string Subunidad { get; set; }
            public string Desde { get; set; }
            public string Hacia { get; set; }
            public int DemandaSesiones { get; set; }
            public int CapacidadAntes { get; set; }
            public int CapacidadDespues { get; set; }
            public string Motivo { get; set; }
            public bool Mejora { get; set; } = true;
            public int SemanasRequeridas { get; set; }
            public DateTime FechaInicioVentana { get; set; }
            public DateTime FechaLimite { get; set; }
            public DateTime? FechaFinProyectada { get; set; }
            public DateTime? FechaFinMinima { get { return FechaFinProyectada; } set { FechaFinProyectada = value; } }
            public int RanurasOciosas { get; set; }
            public int DemandaPendiente { get; set; }
            public double DeficitHoras { get; set; }
            public int MateriasFueraVentana { get; set; }
            public List<int> IdAsignaturasBottleneck { get; set; } = new List<int>();
            public List<AgReclasificacionPatron> ReclasificacionesPropuestas { get; set; } = new List<AgReclasificacionPatron>();
            public List<AgReclasificacionPatron> ReclasificacionesAutorizadas { get; set; } = new List<AgReclasificacionPatron>();
            public bool Aplica => (!string.IsNullOrWhiteSpace(Hacia) && !string.Equals(Desde, Hacia, StringComparison.OrdinalIgnoreCase) && Mejora && CapacidadDespues > CapacidadAntes)
                || (ReclasificacionesAutorizadas != null && ReclasificacionesAutorizadas.Count > 0);
            internal AgContexto ContextoOrigen { get; set; }
            internal int? PerfilPropuesto { get; set; }
            internal int? PerfilAplicado { get; set; }
            internal string HuellaDecision { get; set; }
            internal string HuellaContextoAntes { get; set; }
            internal string HuellaContextoDespues { get; set; }
            internal string HuellaPersistidaAntes { get; set; }
            internal string HuellaPersistidaDespues { get; set; }
        }

        public sealed class AgResultadoValidacionSeleccion
        {
            public bool EsValido { get; set; }
            public string Motivo { get; set; } = string.Empty;
            public string Mensaje { get; set; } = string.Empty;
            public int? IdAulaRechazada { get; set; }
            public int? IdAsignaturaRechazada { get; set; }

            public static AgResultadoValidacionSeleccion Valido() => new AgResultadoValidacionSeleccion { EsValido = true };
            public static AgResultadoValidacionSeleccion Rechazo(string motivo, string mensaje, int? idAula = null, int? idAsig = null) =>
                new AgResultadoValidacionSeleccion { EsValido = false, Motivo = motivo, Mensaje = mensaje, IdAulaRechazada = idAula, IdAsignaturaRechazada = idAsig };
        }

        /// <summary>
        /// Predicado canónico de elegibilidad de aula para propuesta y autorización de reclasificación.
        /// Aulas con Estado Aprobado (4), Inactiva (0), Mantenimiento (2) o Reservada (3) nunca son elegibles.
        /// </summary>
        public static bool EsAulaElegible(AgAula aula) => aula != null && aula.Id > 0 && !string.IsNullOrWhiteSpace(aula.AulaRef) && aula.Estado == AgEstadoAula.Disponible;

        /// <summary>
        /// Identifica patrones LEGADO con días asignados que pueden ser candidatos a redistribución.
        /// Acotado estrictamente a la identidad del aula y subunidad.
        /// </summary>
        public IReadOnlyList<AgReclasificacionPatron> IdentificarPatronesRedistribuibles(AgContexto contexto)
        {
            var resultado = new List<AgReclasificacionPatron>();
            if (contexto?.Aulas == null) return resultado;
            foreach (var aula in contexto.Aulas)
            {
                resultado.AddRange(IdentificarPatronesRedistribuibles(contexto, aula));
            }
            return resultado;
        }

        public IReadOnlyList<AgReclasificacionPatron> IdentificarPatronesRedistribuibles(AgContexto contexto, AgAula aula)
        {
            return IdentificarPatronesRedistribuibles(contexto, aula, null);
        }

        public IReadOnlyList<AgReclasificacionPatron> IdentificarPatronesRedistribuibles(AgContexto contexto, AgAula aula, AgResultadoEvaluacionVentana evaluacion = null)
        {
            var resultado = new List<AgReclasificacionPatron>();
            if (contexto == null || !EsAulaElegible(aula))
            {
                return resultado;
            }

            if (evaluacion == null)
            {
                if (contexto.Ventana == null)
                {
                    return resultado;
                }

                var prog = contexto.Programas != null ? contexto.Programas.FirstOrDefault(p => p.Id == aula.IdPrograma) : null;
                double u = AgUnidadesSemanales.HorasSesionAula(aula, AgNivelAcademicoCatalogo.Desde(prog?.Modalidad));
                evaluacion = AgEvaluadorDemandaVentana.Evaluar(aula.DiasVentana, aula.Jornada, u,
                    contexto, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin);
            }

            if (evaluacion == null || !string.IsNullOrEmpty(evaluacion.Error) || evaluacion.CabeEnPlazo)
            {
                return resultado;
            }

            if (contexto.AulaAsignaturaPorAulaAsig == null)
            {
                return resultado;
            }

            var relaciones = contexto.AulaAsignaturaPorAulaAsig.Values
                .Where(aa => aa.IdAula == aula.Id)
                .OrderBy(aa => aa.OrdenCurricular)
                .ThenBy(aa => aa.IdAsignatura);

            var programa = contexto.Programas != null ? contexto.Programas.FirstOrDefault(p => p.Id == aula.IdPrograma) : null;
            var sedeObj = (contexto.SedesPorId != null && contexto.SedesPorId.TryGetValue(aula.IdSede, out var s))
                ? s
                : contexto.Sedes?.FirstOrDefault(sd => sd.Id == aula.IdSede);
            string sedeNombre = !string.IsNullOrWhiteSpace(sedeObj?.Nombre) ? sedeObj.Nombre : (aula.Subunidad ?? string.Empty);

            foreach (var aa in relaciones)
            {
                string origen = (aa.OrigenPatron ?? string.Empty).Trim().ToUpperInvariant();
                if (origen != "LEGADO" || string.IsNullOrWhiteSpace(aa.DiasPatron))
                {
                    continue;
                }

                int sesionesPendientes = 0;
                if (evaluacion.SesionesPendientesPorAsignatura != null
                    && evaluacion.SesionesPendientesPorAsignatura.TryGetValue(aa.IdAsignatura, out int ses))
                {
                    sesionesPendientes = ses;
                }
                else if (evaluacion.IdAsignaturasSaturadas != null && evaluacion.IdAsignaturasSaturadas.Contains(aa.IdAsignatura))
                {
                    sesionesPendientes = Math.Max(1, evaluacion.DemandaPendiente);
                }

                bool saturada = (evaluacion.IdAsignaturasSaturadas == null || evaluacion.IdAsignaturasSaturadas.Contains(aa.IdAsignatura));
                if (!saturada || sesionesPendientes <= 0)
                {
                    continue;
                }

                double horasPendientes = 0;
                if (evaluacion.HorasPendientesPorAsignatura != null
                    && evaluacion.HorasPendientesPorAsignatura.TryGetValue(aa.IdAsignatura, out double h))
                {
                    horasPendientes = h;
                }
                else
                {
                    double u = AgUnidadesSemanales.HorasSesionAula(aula, AgNivelAcademicoCatalogo.Desde(programa?.Modalidad));
                    horasPendientes = sesionesPendientes * u;
                }

                var asig = contexto.Asignaturas != null ? contexto.Asignaturas.FirstOrDefault(a => a.Id == aa.IdAsignatura) : null;
                DateTime? fechaLimite = contexto.Ventana != null ? (DateTime?)contexto.Ventana.FechaFin : (DateTime?)evaluacion.FechaLimite;

                resultado.Add(new AgReclasificacionPatron
                {
                    IdAula = aula.Id,
                    AulaRef = aula.AulaRef,
                    Subunidad = aula.Subunidad ?? string.Empty,
                    IdAsignatura = aa.IdAsignatura,
                    NombreAsignatura = asig?.Nombre ?? $"Asig {aa.IdAsignatura}",
                    OrigenAnterior = "LEGADO",
                    OrigenPropuesto = "AUTOMATICO",
                    DiasPatronAnterior = aa.DiasPatron,
                    DiasPatronPropuesto = string.Empty,
                    SesionesPendientes = sesionesPendientes,
                    HorasPendientes = horasPendientes,
                    FechaLimite = fechaLimite,
                    Sede = sedeNombre,
                    Programa = programa?.Nombre ?? string.Empty,
                    SemestreGrupo = $"{aula.Semestre}{aula.Grupo}",
                    Motivo = $"Patrón histórico '{aa.DiasPatron}' en aula {aula.AulaRef}; reclasificar a AUTOMATICO para permitir redistribución.",
                    Autorizado = false
                });
            }

            return resultado;
        }

        /// <summary>
        /// Valida tipadamente las selecciones de reclasificación sin lanzar excepciones de control de flujo.
        /// Retorna un resultado estructurado diferenciando los motivos concretos de rechazo.
        /// </summary>
        public AgResultadoValidacionSeleccion ValidarSelecciones(AgContexto contexto, IReadOnlyCollection<AgReclasificacionPatron> reclasificacionesAutorizadas)
        {
            if (reclasificacionesAutorizadas == null || reclasificacionesAutorizadas.Count == 0)
            {
                return AgResultadoValidacionSeleccion.Valido();
            }

            if (contexto == null || contexto.Aulas == null || contexto.AulaAsignaturaPorAulaAsig == null)
            {
                return RechazoConLog("CONFIGURACION_MODIFICADA", "Contexto no inicializado o incompleto.", null, null);
            }

            var identidades = new HashSet<(int, int)>();
            foreach (var rec in reclasificacionesAutorizadas)
            {
                if (rec == null)
                {
                    return RechazoConLog("RECLASIFICACION_NULA", "Reclasificación nula.", null, null);
                }

                if (!rec.Autorizado)
                {
                    continue;
                }

                string error;
                if (!rec.EsTransicionValida(out error))
                {
                    return RechazoConLog("TRANSICION_INVALIDA", string.IsNullOrWhiteSpace(error) ? "Transición inválida." : error, rec.IdAula, rec.IdAsignatura);
                }

                if (!identidades.Add((rec.IdAula, rec.IdAsignatura)))
                {
                    return RechazoConLog("SELECCION_DUPLICADA", "Reclasificación duplicada para aula y asignatura.", rec.IdAula, rec.IdAsignatura);
                }

                var aulasCoincidentes = contexto.Aulas.Where(a => a != null
                    && a.Id == rec.IdAula
                    && string.Equals(a.AulaRef, rec.AulaRef, StringComparison.Ordinal)
                    && string.Equals(a.Subunidad ?? string.Empty, rec.Subunidad ?? string.Empty, StringComparison.Ordinal)).ToList();

                if (aulasCoincidentes.Count != 1 || !EsAulaElegible(aulasCoincidentes[0]))
                {
                    string mensajeAula = aulasCoincidentes.Count == 0
                        ? $"Aula {rec.AulaRef} (id={rec.IdAula}) no encontrada en el contexto."
                        : aulasCoincidentes.Count > 1
                            ? $"Aula {rec.AulaRef} (id={rec.IdAula}) ambigua en el contexto ({aulasCoincidentes.Count} encontradas)."
                            : $"Aula {rec.AulaRef} (id={rec.IdAula}) no elegible (estado={aulasCoincidentes[0].Estado}).";

                    return RechazoConLog("AULA_NO_ELEGIBLE", mensajeAula, rec.IdAula, rec.IdAsignatura);
                }

                AgAulaAsignatura relacion;
                if (!contexto.AulaAsignaturaPorAulaAsig.TryGetValue((rec.IdAula, rec.IdAsignatura), out relacion)
                    || !string.Equals((relacion.OrigenPatron ?? string.Empty).Trim(), (rec.OrigenAnterior ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(relacion.DiasPatron ?? string.Empty, rec.DiasPatronAnterior ?? string.Empty, StringComparison.Ordinal))
                {
                    return RechazoConLog("CONFIGURACION_MODIFICADA", "Reclasificación ajena o configuración modificada. Recalcular propuesta.", rec.IdAula, rec.IdAsignatura);
                }

                // Corrección [P2]: Comprobar pertenencia a las candidatas del diagnóstico vigente y sellado (materia con déficit positivo)
                var candidatasAula = IdentificarPatronesRedistribuibles(contexto, aulasCoincidentes[0]);
                var candidataValida = candidatasAula.FirstOrDefault(c => c.IdAsignatura == rec.IdAsignatura);
                if (candidataValida == null || candidataValida.SesionesPendientes <= 0)
                {
                    return RechazoConLog("MATERIA_SIN_DEFICIT",
                        $"La asignatura {rec.IdAsignatura} no presenta déficit de sesiones pendientes en el aula {rec.AulaRef} dentro del plazo evaluado.",
                        rec.IdAula, rec.IdAsignatura);
                }
            }

            return AgResultadoValidacionSeleccion.Valido();
        }

        private static AgResultadoValidacionSeleccion RechazoConLog(string motivo, string mensaje, int? idAula = null, int? idAsig = null)
        {
            var res = AgResultadoValidacionSeleccion.Rechazo(motivo, mensaje, idAula, idAsig);
            LoggerHelper.Warning($"[AG-RECLASIFICACION-WARN] Validación fallida: motivo={res.Motivo}, aula={res.IdAulaRechazada}, asig={res.IdAsignaturaRechazada}: {res.Mensaje}");
            return res;
        }

        public IReadOnlyList<Decision> Evaluar(AgContexto contexto)
        {
            return EvaluarInterno(contexto, null, null, CancellationToken.None);
        }

        public IReadOnlyList<Decision> Evaluar(AgContexto contexto, int idArbol)
        {
            return EvaluarInterno(contexto, idArbol, null, CancellationToken.None);
        }

        public IReadOnlyList<Decision> Evaluar(AgContexto contexto, int idArbol, CancellationToken ct)
        {
            return EvaluarInterno(contexto, idArbol, null, ct);
        }

        public IReadOnlyList<Decision> Evaluar(AgContexto contexto, int idArbol,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificacionesAutorizadas)
        {
            return EvaluarInterno(contexto, idArbol, reclasificacionesAutorizadas, CancellationToken.None);
        }

        public IReadOnlyList<Decision> Evaluar(AgContexto contexto, int idArbol,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificacionesAutorizadas, CancellationToken ct)
        {
            return EvaluarInterno(contexto, idArbol, reclasificacionesAutorizadas, ct);
        }

        private IReadOnlyList<Decision> EvaluarInterno(AgContexto contexto, int? perfil,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificacionesAutorizadas, CancellationToken ct)
        {
            var decisiones = new List<Decision>();
            if (contexto == null || contexto.Ventana == null) { return decisiones; }

            var validacion = ValidarSelecciones(contexto, reclasificacionesAutorizadas);
            if (!validacion.EsValido)
            {
                throw new ArgumentException(validacion.Mensaje);
            }

            var seleccionadas = new List<AgReclasificacionPatron>();
            if (reclasificacionesAutorizadas != null)
            {
                foreach (var rec in reclasificacionesAutorizadas)
                {
                    ct.ThrowIfCancellationRequested();
                    if (!rec.Autorizado) { continue; }

                    var copia = rec.Clone();
                    copia.OrigenAnterior = "LEGADO";
                    copia.OrigenPropuesto = rec.OrigenPropuesto.Trim().ToUpperInvariant();
                    copia.DiasPatronPropuesto = rec.DiasPatronPropuesto ?? string.Empty;
                    seleccionadas.Add(copia);
                }
            }

            foreach (var aula in contexto.Aulas.Where(EsAulaElegible).OrderBy(a => a.Id))
            {
                ct.ThrowIfCancellationRequested();
                LoggerHelper.Debug($"[AG-RANURAS-TMP] ESCALADOR_ENTRADA corrida={contexto.IdCorrida} traza={contexto.Traza?.Correlacion} perfil={perfil} aula={aula.Id} ref={aula.AulaRef} dias={aula.DiasVentana} plazoEfectivo={contexto.Ventana.FechaInicio:yyyy-MM-dd}..{contexto.Ventana.FechaFin:yyyy-MM-dd} plazoAula={aula.FechaInicio}..{aula.FechaFin}");

                string huellaInicial = AgHuellaPlanVentana.Calcular(contexto, aula);
                var programa = contexto.Programas.FirstOrDefault(p => p.Id == aula.IdPrograma);
                double unidad = AgUnidadesSemanales.HorasSesionAula(aula, AgNivelAcademicoCatalogo.Desde(programa?.Modalidad));

                var antes = AgEvaluadorDemandaVentana.Evaluar(aula.DiasVentana, aula.Jornada, unidad,
                    contexto, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct);

                // Paso 2: Siempre consultar candidatos de reclasificación independientemente del siguiente escalón y de la saturación
                var candidatos = IdentificarPatronesRedistribuibles(contexto, aula, antes);

                var autorizadasAula = seleccionadas.Where(r => r.IdAula == aula.Id).ToList();

                // Paso 4: Simulación en copia del contexto sin alterar el contexto original
                var ctxSimulado = ClonarContextoParaSimulacion(contexto);
                foreach (var rec in autorizadasAula)
                {
                    if (ctxSimulado.AulaAsignaturaPorAulaAsig.TryGetValue((aula.Id, rec.IdAsignatura), out var aaSim))
                    {
                        aaSim.OrigenPatron = rec.OrigenPropuesto;
                        aaSim.DiasPatron = rec.DiasPatronPropuesto ?? string.Empty;
                    }
                }
                ctxSimulado.AulaAsignaturas = ctxSimulado.AulaAsignaturaPorAulaAsig.Values.ToList();
                ctxSimulado.ConstruirIndices();

                // Simular primero en la misma ventana
                var evalMismaVentana = AgEvaluadorDemandaVentana.Evaluar(aula.DiasVentana, aula.Jornada, unidad,
                    ctxSimulado, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct);

                string paso = AgEscaleraVentanas.Siguiente(aula.DiasVentana);
                string destino = aula.DiasVentana;
                AgResultadoEvaluacionVentana despues = evalMismaVentana;

                if (autorizadasAula.Count > 0)
                {
                    // Si con reclasificación en la misma ventana aún quedan pendientes y hay escalón disponible, evaluar siguiente escalón
                    if (!evalMismaVentana.CabeEnPlazo && !string.IsNullOrEmpty(paso))
                    {
                        var evalSiguiente = AgEvaluadorDemandaVentana.Evaluar(paso, aula.Jornada, unidad,
                            ctxSimulado, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct);

                        if (AgEvaluadorDemandaVentana.PasoMejora(evalMismaVentana, evalSiguiente))
                        {
                            destino = paso;
                            despues = evalSiguiente;
                        }
                    }
                }
                else
                {
                    // Sin autorizaciones: flujo nominal de escalera si hay saturación
                    if (antes.HaySaturacion && !string.IsNullOrEmpty(paso))
                    {
                        var evalPaso = AgEvaluadorDemandaVentana.Evaluar(paso, aula.Jornada, unidad,
                            contexto, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct);

                        if (AgEvaluadorDemandaVentana.PasoMejora(antes, evalPaso))
                        {
                            destino = paso;
                            despues = evalPaso;
                        }
                    }
                }

                // Si no hay saturación, no hay candidatos y no hay autorizaciones, omitir
                if (!antes.HaySaturacion && candidatos.Count == 0 && autorizadasAula.Count == 0)
                {
                    LoggerHelper.Debug($"[AG-RANURAS-TMP] ESCALADOR_OMITIR aula={aula.Id} motivo=cabe_en_plazo_y_sin_candidatos");
                    continue;
                }

                // Si no hay cambio de ventana ni candidatos ni autorizaciones que mejoren, omitir
                bool hayCambioVentana = !string.Equals(aula.DiasVentana, destino, StringComparison.OrdinalIgnoreCase);
                bool hayAutorizaciones = autorizadasAula.Count > 0;
                bool hayCandidatos = candidatos.Count > 0;

                if (!hayCambioVentana && !hayAutorizaciones && !hayCandidatos)
                {
                    LoggerHelper.Debug($"[AG-RANURAS-TMP] ESCALADOR_OMITIR aula={aula.Id} motivo=sin_propuesta_viable");
                    continue;
                }

                var automaticas = contexto.AulaAsignaturaPorAulaAsig.Values.Where(aa => aa.IdAula == aula.Id
                    && string.Equals(aa.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase))
                    .Select(aa => aa.IdAsignatura).Distinct().OrderBy(id => id).ToList();

                var d = new Decision
                {
                    IdAula = aula.Id,
                    AulaRef = aula.AulaRef,
                    Subunidad = aula.Subunidad,
                    Desde = aula.DiasVentana,
                    Hacia = destino,
                    DemandaSesiones = antes.DemandaTotal,
                    CapacidadAntes = antes.CapacidadTotal,
                    CapacidadDespues = despues.CapacidadTotal,
                    Motivo = antes.Motivo,
                    SemanasRequeridas = despues.SemanasRequeridas,
                    FechaInicioVentana = contexto.Ventana.FechaInicio.Date,
                    FechaLimite = contexto.Ventana.FechaFin.Date,
                    FechaFinProyectada = despues.FechaFinProyectada,
                    RanurasOciosas = despues.RanurasOciosas,
                    DemandaPendiente = despues.DemandaPendiente,
                    DeficitHoras = despues.DeficitHoras,
                    MateriasFueraVentana = antes.MateriasFueraVentana.Count,
                    IdAsignaturasBottleneck = automaticas,
                    ReclasificacionesPropuestas = candidatos.Select(c => c.Clone()).ToList(),
                    ReclasificacionesAutorizadas = autorizadasAula.Select(a => a.Clone()).ToList(),
                    Mejora = despues.DemandaPendiente < antes.DemandaPendiente
                             || despues.CapacidadTotal > antes.CapacidadTotal
                             || (hayAutorizaciones && despues.DemandaPendiente <= antes.DemandaPendiente),
                    ContextoOrigen = contexto,
                    PerfilPropuesto = perfil,
                    HuellaContextoAntes = huellaInicial,
                    HuellaContextoDespues = AgHuellaPlanVentana.Calcular(contexto, aula, destino, automaticas, autorizadasAula)
                };

                d.HuellaDecision = HuellaAccion(d);

                using (var db = _connectionFactory())
                {
                    if (db != null)
                    {
                        if (db.State != ConnectionState.Open) { db.Open(); }
                        using (var tx = db.BeginTransaction())
                        {
                            if (!AgHuellaEscaladoPersistido.CoincideOrigen(db, tx, contexto, aula))
                            {
                                throw new InvalidOperationException("Configuración modificada antes de sellar. Recargue el contexto.");
                            }
                            d.HuellaPersistidaAntes = AgHuellaEscaladoPersistido.Calcular(db, tx, d.AulaRef, d.Subunidad, perfil ?? 0);
                            tx.Commit();
                        }
                    }
                }

                if (AgHuellaPlanVentana.Calcular(contexto, aula) != huellaInicial)
                {
                    throw new InvalidOperationException("Configuración modificada durante la evaluación. Recalcular propuesta.");
                }

                ct.ThrowIfCancellationRequested();
                decisiones.Add(d);
                LoggerHelper.Debug($"[AG-RANURAS-TMP] PROPUESTA aula={aula.Id} ref={aula.AulaRef} desde={d.Desde} hacia={d.Hacia} capacidad={d.CapacidadAntes}->{d.CapacidadDespues} pendientes={d.DemandaPendiente} autorizadas={autorizadasAula.Count} candidatas={candidatos.Count} sellada=true");
            }

            return decisiones;
        }

        private static AgContexto ClonarContextoParaSimulacion(AgContexto ctx)
        {
            var clon = new AgContexto
            {
                IdCorrida = ctx.IdCorrida,
                IdPeriodo = ctx.IdPeriodo,
                IdPrograma = ctx.IdPrograma,
                IdSede = ctx.IdSede,
                Subunidad = ctx.Subunidad,
                Ventana = ctx.Ventana,
                Programas = ctx.Programas,
                Asignaturas = ctx.Asignaturas,
                Aulas = ctx.Aulas,
                Bloques = ctx.Bloques,
                Festivos = ctx.Festivos,
                Presenciales = ctx.Presenciales,
                Docentes = ctx.Docentes,
                Traza = ctx.Traza
            };

            clon.AulaAsignaturaPorAulaAsig.Clear();
            clon.AulaAsignaturas = new List<AgAulaAsignatura>();

            foreach (var kv in ctx.AulaAsignaturaPorAulaAsig)
            {
                var aa = kv.Value;
                var nuevoAa = new AgAulaAsignatura
                {
                    IdAula = aa.IdAula,
                    IdAsignatura = aa.IdAsignatura,
                    HorasObjetivo = aa.HorasObjetivo,
                    OrdenCurricular = aa.OrdenCurricular,
                    Exclusivo = aa.Exclusivo,
                    Prioridad = aa.Prioridad,
                    Requerida = aa.Requerida,
                    Subunidad = aa.Subunidad,
                    DiasPatron = aa.DiasPatron,
                    OrigenPatron = aa.OrigenPatron
                };
                clon.AulaAsignaturaPorAulaAsig[kv.Key] = nuevoAa;
                clon.AulaAsignaturas.Add(nuevoAa);
            }

            return clon;
        }

        internal sealed class ResultadoAplicacion
        {
            public bool Exito { get; set; }
            public string Error { get; set; } = string.Empty;
            public string Advertencia { get; set; } = string.Empty;
            public int Confirmadas { get; set; }
            public int Modificadas { get; set; }
        }

        public Task<int> AplicarAsync(IReadOnlyList<Decision> decisiones, CancellationToken ct)
        {
            return AplicarAsync(decisiones, 0, ct);
        }

        public async Task<int> AplicarAsync(IReadOnlyList<Decision> decisiones, int idArbol, CancellationToken ct)
        {
            var resultado = await AplicarDetalladoAsync(decisiones, idArbol, ct).ConfigureAwait(false);
            return resultado.Confirmadas;
        }

        public Task<ResultadoAplicacion> AplicarDetalladoAsync(IReadOnlyList<Decision> decisiones, int idArbol, CancellationToken ct)
        {
            return AplicarDetalladoAsync(decisiones, idArbol, null, ct);
        }

        public async Task<ResultadoAplicacion> AplicarDetalladoAsync(IReadOnlyList<Decision> decisiones,
            int idArbol, AgContexto contextoVigente, CancellationToken ct)
        {
            var resultado = new ResultadoAplicacion();
            var posteriores = new Dictionary<Decision, string>();
            if (decisiones == null || decisiones.Count == 0) { resultado.Exito = true; return resultado; }

            // Validar el sello antes de filtrar: desactivar una acción tampoco permite ocultar su modificación.
            foreach (var decision in decisiones)
            {
                if (ct.IsCancellationRequested)
                {
                    return new ResultadoAplicacion { Error = "Operación cancelada." };
                }
                if (decision == null || string.IsNullOrEmpty(decision.HuellaDecision))
                {
                    return new ResultadoAplicacion { Error = "Decisión sin instantánea válida. Recalcular propuesta." };
                }
                if (decision.ReclasificacionesAutorizadas != null && decision.ReclasificacionesAutorizadas.Any(r => r == null)
                    || HuellaAccion(decision) != decision.HuellaDecision)
                {
                    return new ResultadoAplicacion { Error = "La propuesta fue modificada. Recalcular." };
                }
            }

            // Guardar exclusivamente decisiones aplicables (acciones confirmadas); selección vacía es SinCambios
            var decisionesAplicables = decisiones.Where(d => d != null && d.Aplica).ToList();
            if (decisionesAplicables.Count == 0)
            {
                resultado.Exito = true;
                resultado.Confirmadas = 0;
                resultado.Modificadas = 0;
                return resultado;
            }

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) { db.Open(); }
                    using (var tx = db.BeginTransaction())
                    {
                        foreach (var d in decisionesAplicables)
                        {
                            ct.ThrowIfCancellationRequested();
                            if (d == null || !d.Aplica || d.ContextoOrigen == null || string.IsNullOrEmpty(d.HuellaPersistidaAntes)
                                || string.IsNullOrEmpty(d.HuellaDecision))
                            {
                                throw new InvalidOperationException("Decisión sin instantánea válida. Recalcular propuesta.");
                            }
                            if (HuellaAccion(d) != d.HuellaDecision)
                            {
                                throw new InvalidOperationException("La propuesta fue modificada. Recalcular.");
                            }
                            if ((d.PerfilPropuesto.HasValue && d.PerfilPropuesto.Value != idArbol)
                                || (d.PerfilAplicado.HasValue && d.PerfilAplicado.Value != idArbol))
                            {
                                throw new InvalidOperationException("Cambió el perfil de la propuesta. Recalcular.");
                            }

                            var contexto = contextoVigente ?? d.ContextoOrigen;
                            var aulas = contexto.Aulas.Where(a => string.Equals(a.AulaRef, d.AulaRef, StringComparison.Ordinal)
                                && string.Equals(a.Subunidad ?? string.Empty, d.Subunidad ?? string.Empty, StringComparison.Ordinal)).ToList();
                            if (aulas.Count != 1) { throw new InvalidOperationException("Configuración modificada: cambió el aula."); }
                            if (d.IdAula <= 0 || d.IdAula != aulas[0].Id)
                            {
                                throw new InvalidOperationException($"La propuesta fue modificada: IdAula ({d.IdAula}) no coincide con el aula esperada ({aulas[0].Id}). Recalcular.");
                            }
                            int idAula = d.IdAula;

                            string huellaContexto = AgHuellaPlanVentana.Calcular(contexto, aulas[0]);
                            string huellaSql = AgHuellaEscaladoPersistido.Calcular(db, tx, d.AulaRef, d.Subunidad, idArbol);
                            bool contextoCompatible = huellaContexto == d.HuellaContextoAntes || huellaContexto == d.HuellaContextoDespues;
                            LoggerHelper.Debug($"[AG-RANURAS-TMP] GUARDAR_HUELLAS perfil={idArbol} ref={d.AulaRef} contextoAntes={huellaContexto == d.HuellaContextoAntes} contextoDespues={huellaContexto == d.HuellaContextoDespues} sqlAntes={huellaSql == d.HuellaPersistidaAntes} sqlDespues={huellaSql == d.HuellaPersistidaDespues}");

                            // Reaplicación idempotente: si ya está en el postestado exacto, confirmar sin nuevas escrituras
                            if (d.PerfilAplicado == idArbol && contextoCompatible && !string.IsNullOrEmpty(d.HuellaPersistidaDespues)
                                && huellaSql == d.HuellaPersistidaDespues
                                && AgHuellaEscaladoPersistido.EstadoCompleto(db, tx, d.AulaRef, d.Subunidad, idArbol, d.Hacia, d.IdAsignaturasBottleneck, d.ReclasificacionesAutorizadas))
                            {
                                resultado.Confirmadas++;
                                LoggerHelper.Debug($"[AG-RANURAS-TMP] GUARDAR_IDEMPOTENTE perfil={idArbol} ref={d.AulaRef}");
                                continue;
                            }

                            if (huellaContexto != d.HuellaContextoAntes || huellaSql != d.HuellaPersistidaAntes
                                || !AgHuellaEscaladoPersistido.CoincideOrigen(db, tx, contexto, aulas[0]))
                            {
                                throw new InvalidOperationException("Configuración modificada. Recalcular propuesta.");
                            }

                            string esperado = AgHuellaEscaladoPersistido.Calcular(db, tx, d.AulaRef, d.Subunidad,
                                idArbol, d.Hacia, d.IdAsignaturasBottleneck, d.ReclasificacionesAutorizadas);

                            int modificadas = 0;
                            if (!string.Equals(d.Desde, d.Hacia, StringComparison.OrdinalIgnoreCase))
                            {
                                modificadas = await db.ExecuteAsync(new CommandDefinition(@"
UPDATE ag_aula SET dias_ventana=@Hacia
WHERE aula_ref=@Ref AND COALESCE(NULLIF(TRIM(subunidad),''),@Sub)=@Sub
AND COALESCE(dias_ventana,'')=@Desde;",
                                    new { d.Hacia, d.Desde, Ref = d.AulaRef, Sub = d.Subunidad ?? string.Empty }, tx, cancellationToken: ct)).ConfigureAwait(false);
                                if (modificadas != 1) { throw new InvalidOperationException("Configuración modificada durante el guardado."); }

                                if (idArbol > 0)
                                {
                                    int nodos = await db.ExecuteAsync(new CommandDefinition(@"
UPDATE ag_arbol_nodo SET dias_ventana=@Hacia
WHERE idarbol=@Perfil AND aula_ref=@Ref AND COALESCE(NULLIF(TRIM(subunidad),''),@Sub)=@Sub
AND COALESCE(dias_ventana,'')=@Desde;",
                                        new { d.Hacia, d.Desde, Perfil = idArbol, Ref = d.AulaRef, Sub = d.Subunidad ?? string.Empty }, tx, cancellationToken: ct)).ConfigureAwait(false);
                                    if (nodos != 1) { throw new InvalidOperationException("Nodo del perfil ausente, duplicado o modificado."); }
                                }
                            }

                            if (d.IdAsignaturasBottleneck.Count > 0)
                            {
                                await db.ExecuteAsync(new CommandDefinition(@"
UPDATE ag_aula_asignatura SET dias_patron=''
WHERE id_aula = @IdAula
AND id_asignatura IN @Ids AND UPPER(COALESCE(origen_patron,''))='AUTOMATICO';",
                                    new { IdAula = idAula, Ids = d.IdAsignaturasBottleneck }, tx, cancellationToken: ct)).ConfigureAwait(false);
                            }

                            if (d.ReclasificacionesAutorizadas != null && d.ReclasificacionesAutorizadas.Count > 0)
                            {
                                foreach (var rec in d.ReclasificacionesAutorizadas)
                                {
                                    if (rec.IdAula <= 0 || rec.IdAula != idAula)
                                    {
                                        throw new InvalidOperationException($"La propuesta fue modificada: IdAula de la reclasificación ({rec.IdAula}) no coincide con el aula esperada ({idAula}). Recalcular.");
                                    }
                                    int targetIdAula = idAula;
                                    int afectoRec = await db.ExecuteAsync(new CommandDefinition(@"
UPDATE ag_aula_asignatura
SET origen_patron=@OrigenNuevo, dias_patron=@DiasNuevo
WHERE id_aula = @IdAula
AND id_asignatura=@IdAsig AND UPPER(TRIM(COALESCE(origen_patron,'')))=@OrigenAntes
AND COALESCE(dias_patron,'')=@DiasAntes;",
                                        new
                                        {
                                            OrigenNuevo = rec.OrigenPropuesto,
                                            DiasNuevo = rec.DiasPatronPropuesto ?? string.Empty,
                                            IdAula = targetIdAula,
                                            IdAsig = rec.IdAsignatura,
                                            OrigenAntes = (rec.OrigenAnterior ?? string.Empty).Trim().ToUpperInvariant(),
                                            DiasAntes = rec.DiasPatronAnterior ?? string.Empty
                                        }, tx, cancellationToken: ct)).ConfigureAwait(false);

                                    if (afectoRec != 1)
                                    {
                                        throw new InvalidOperationException($"Relación no autorizada, no coincidente o modificada para asignatura {rec.IdAsignatura} en aula {d.AulaRef}.");
                                    }
                                    modificadas += afectoRec;
                                }
                            }

                            if (AgHuellaEscaladoPersistido.Calcular(db, tx, d.AulaRef, d.Subunidad, idArbol) != esperado
                                || !AgHuellaEscaladoPersistido.EstadoCompleto(db, tx, d.AulaRef, d.Subunidad, idArbol, d.Hacia, d.IdAsignaturasBottleneck, d.ReclasificacionesAutorizadas))
                            {
                                throw new InvalidOperationException("Postcondiciones incompletas. Operación revertida.");
                            }

                            posteriores[d] = esperado;
                            resultado.Confirmadas++;
                            resultado.Modificadas += modificadas;
                        }

                        ct.ThrowIfCancellationRequested();
                        tx.Commit();
                        LoggerHelper.Debug($"[AG-RANURAS-TMP] GUARDAR_COMMIT perfil={idArbol} confirmadas={resultado.Confirmadas} modificadas={resultado.Modificadas}");
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Debug($"[AG-RANURAS-TMP] GUARDAR_REVERTIDO perfil={idArbol} tipo={ex.GetType().Name} motivo={ex.Message}");
                return new ResultadoAplicacion { Error = ex is OperationCanceledException ? "Operación cancelada." : ex.Message };
            }

            foreach (var par in posteriores)
            {
                par.Key.HuellaPersistidaDespues = par.Value;
                par.Key.PerfilAplicado = idArbol;
            }

            resultado.Exito = true;
            try
            {
                if (resultado.Modificadas > 0) { ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty(); }
                LoggerHelper.Info("[Escalado] Confirmadas=" + resultado.Confirmadas + "; modificadas=" + resultado.Modificadas);
            }
            catch (Exception ex) { resultado.Advertencia = "Guardado confirmado; falló la notificación posterior: " + ex.Message; }

            return resultado;
        }

        private static string HuellaAccion(Decision d)
        {
            var partes = new List<string>
            {
                "propuesta-escalado-v4", d.IdAula.ToString(CultureInfo.InvariantCulture), d.AulaRef, d.Subunidad, d.Desde, d.Hacia,
                d.Mejora.ToString(), d.CapacidadAntes.ToString(CultureInfo.InvariantCulture),
                d.CapacidadDespues.ToString(CultureInfo.InvariantCulture),
                d.DemandaPendiente.ToString(CultureInfo.InvariantCulture),
                d.FechaLimite.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                d.PerfilPropuesto.HasValue ? d.PerfilPropuesto.Value.ToString(CultureInfo.InvariantCulture) : null,
                d.IdAsignaturasBottleneck == null ? null : string.Join(",", d.IdAsignaturasBottleneck.OrderBy(id => id))
            };

            if (d.ReclasificacionesAutorizadas != null && d.ReclasificacionesAutorizadas.Count > 0)
            {
                partes.AddRange(d.ReclasificacionesAutorizadas
                    .OrderBy(r => r.IdAula).ThenBy(r => r.IdAsignatura)
                    .Select(r => $"rec:{r.IdAula}:{r.AulaRef}:{r.Subunidad}:{r.IdAsignatura}:{r.OrigenAnterior}->{r.OrigenPropuesto}:{r.DiasPatronAnterior}->{r.DiasPatronPropuesto}:{r.Autorizado}"));
            }

            return AgHuellaPlanVentana.Hash(partes);
        }
    }
}
