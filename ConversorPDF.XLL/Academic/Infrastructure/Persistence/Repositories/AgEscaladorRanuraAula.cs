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
            public string ParticionPreferente { get; set; } = string.Empty;
            public bool ParticionMejora { get; set; }
            public bool ParticionDifiereDeReferencia { get; set; }
            public bool ConservaPatronOrigen { get; set; } = true;
            public bool SoloParticion =>
                !Aplica
                && string.Equals(Desde, Hacia, StringComparison.OrdinalIgnoreCase)
                && (ReclasificacionesAutorizadas == null || ReclasificacionesAutorizadas.Count == 0)
                && (ParticionMejora || ParticionDifiereDeReferencia);
            /// <summary>
            /// Determina si la decisión requiere aplicar cambios persistentes en la base de datos (Fase 2256, Fase 2258, Fase 2284, Fase 2291).
            /// La partición de carriles es efímera y se recalcula en cada corrida por el planificador de oleadas;
            /// no posee campos persistidos en base de datos. Decisiones de solo partición (Desde == Hacia)
            /// sin reclasificaciones ni liberación de cuellos de botella no generan escrituras SQL.
            /// Si ConservaPatronOrigen es falso y existen asignaturas bottleneck, se requiere vaciar dias_patron, por ende Aplica es true.
            /// Si hay cambio de ventana y Mejora es true (incluso con igual capacidad por redistribución), Aplica es true.
            /// </summary>
            public bool Aplica =>
                (!string.IsNullOrWhiteSpace(Hacia)
                 && !string.Equals(Desde, Hacia, StringComparison.OrdinalIgnoreCase)
                 && Mejora)
                || (ReclasificacionesAutorizadas != null && ReclasificacionesAutorizadas.Count > 0)
                || (!ConservaPatronOrigen
                    && IdAsignaturasBottleneck != null
                    && IdAsignaturasBottleneck.Count > 0);
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

        public sealed class AgResultadoEvaluacionEscalado : IReadOnlyList<Decision>
        {
            private readonly IReadOnlyList<Decision> _decisiones;

            public bool EsValido { get; }
            public string Motivo { get; }
            public string Mensaje { get; }
            public int? IdAulaRechazada { get; }
            public int? IdAsignaturaRechazada { get; }
            public IReadOnlyList<Decision> Decisiones => _decisiones;

            public AgResultadoEvaluacionEscalado(bool esValido, string motivo, string mensaje, int? idAula, int? idAsig, IReadOnlyList<Decision> decisiones)
            {
                EsValido = esValido;
                Motivo = motivo ?? string.Empty;
                Mensaje = mensaje ?? string.Empty;
                IdAulaRechazada = idAula;
                IdAsignaturaRechazada = idAsig;
                _decisiones = decisiones ?? new List<Decision>();
            }

            public static AgResultadoEvaluacionEscalado Exito(IReadOnlyList<Decision> decisiones) =>
                new AgResultadoEvaluacionEscalado(true, string.Empty, string.Empty, null, null, decisiones);

            public static AgResultadoEvaluacionEscalado Rechazo(AgResultadoValidacionSeleccion val) =>
                new AgResultadoEvaluacionEscalado(false, val?.Motivo, val?.Mensaje, val?.IdAulaRechazada, val?.IdAsignaturaRechazada, new List<Decision>());

            public static AgResultadoEvaluacionEscalado Rechazo(string motivo, string mensaje, int? idAula = null, int? idAsig = null) =>
                new AgResultadoEvaluacionEscalado(false, motivo, mensaje, idAula, idAsig, new List<Decision>());

            public int Count => _decisiones.Count;
            public Decision this[int index] => _decisiones[index];
            public IEnumerator<Decision> GetEnumerator() => _decisiones.GetEnumerator();
            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => _decisiones.GetEnumerator();
        }

        /// <summary>
        /// Predicado canónico de admisión de escritura heredada, exclusivamente Disponible.
        /// Evaluar una aprobada no concede esta admisión ni autorización de reclasificación.
        /// </summary>
        public static bool EsAulaElegible(AgAula aula) => aula != null && aula.Id > 0 && !string.IsNullOrWhiteSpace(aula.AulaRef) && aula.Estado == AgEstadoAula.Disponible;

        /// <summary>Legado estricto: sólo Disponible. No usar como permiso de escritura.</summary>
        public static bool PuedeAplicarPropuestaHeredada(AgAula aula) => EsAulaElegible(aula);

        /// <summary>Legado estricto: redirige a PuedeAplicarPropuestaHeredada.</summary>
        public static bool PuedeAplicarPropuesta(AgAula aula) => PuedeAplicarPropuestaHeredada(aula);

        /// <summary>
        /// Única vía de escritura: Disponible o Aprobada verificada sin origen conservable
        /// (PendienteDeProgramar con 0 filas productivas de la cohorte). Las aprobadas con origen nunca admiten escritura.
        /// </summary>
        public static bool PuedeAplicarPropuesta(AgContexto contexto, AgAula aula)
        {
            if (aula == null || aula.Id <= 0 || string.IsNullOrWhiteSpace(aula.AulaRef)) return false;
            if (aula.Estado == AgEstadoAula.Disponible) return true;
            return aula.Estado == AgEstadoAula.Aprobado
                && contexto != null
                && contexto.PuedeGenerarAula(aula.Id)
                && contexto.PlanAprobadoSnapshot != null
                && !contexto.TieneOrigenConservable(aula.Id);
        }

        /// <summary>Evaluación informativa: aprobadas sólo con manifiesto/snapshot reconciliado y vigente.</summary>
        public static bool PuedeEvaluarPropuesta(AgContexto contexto, AgAula aula) => EsAulaElegible(aula)
            || AgSimuladorEscaladoAprobadas.OrigenIdentificable(contexto, aula);

        /// <summary>
        /// Identifica patrones LEGADO con días asignados que pueden ser candidatos a redistribución.
        /// Acotado estrictamente a la identidad del aula y subunidad.
        /// </summary>
        public IReadOnlyList<AgReclasificacionPatron> IdentificarPatronesRedistribuibles(AgContexto contexto)
        {
            return IdentificarPatronesRedistribuibles(contexto, CancellationToken.None);
        }

        public IReadOnlyList<AgReclasificacionPatron> IdentificarPatronesRedistribuibles(AgContexto contexto, CancellationToken ct)
        {
            var resultado = new List<AgReclasificacionPatron>();
            if (contexto?.Aulas == null) return resultado;
            foreach (var aula in contexto.Aulas)
            {
                ct.ThrowIfCancellationRequested();
                resultado.AddRange(IdentificarPatronesRedistribuibles(contexto, aula, null, ct));
            }
            return resultado;
        }

        public IReadOnlyList<AgReclasificacionPatron> IdentificarPatronesRedistribuibles(AgContexto contexto, AgAula aula)
        {
            return IdentificarPatronesRedistribuibles(contexto, aula, null, CancellationToken.None);
        }

        public IReadOnlyList<AgReclasificacionPatron> IdentificarPatronesRedistribuibles(AgContexto contexto, AgAula aula, AgResultadoEvaluacionVentana evaluacion)
        {
            return IdentificarPatronesRedistribuibles(contexto, aula, evaluacion, CancellationToken.None);
        }

        public IReadOnlyList<AgReclasificacionPatron> IdentificarPatronesRedistribuibles(AgContexto contexto, AgAula aula, AgResultadoEvaluacionVentana evaluacion, CancellationToken ct)
        {
            var resultado = new List<AgReclasificacionPatron>();
            if (contexto == null || !PuedeEvaluarPropuesta(contexto, aula))
            {
                return resultado;
            }

            ct.ThrowIfCancellationRequested();

            if (evaluacion == null)
            {
                if (contexto.Ventana == null)
                {
                    return resultado;
                }

                var prog = contexto.Programas != null ? contexto.Programas.FirstOrDefault(p => p.Id == aula.IdPrograma) : null;
                double u = AgUnidadesSemanales.HorasSesionAula(aula, AgNivelAcademicoCatalogo.Desde(prog?.Modalidad));
                evaluacion = AgEvaluadorDemandaVentana.Evaluar(aula.DiasVentana, aula.Jornada, u,
                    contexto, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct);
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
        /// Valida contra el conjunto de candidatas selladas si se proporciona; en su defecto realiza búsqueda acotada con caché por aula.
        /// </summary>
        public AgResultadoValidacionSeleccion ValidarSelecciones(AgContexto contexto, IReadOnlyCollection<AgReclasificacionPatron> reclasificacionesAutorizadas)
        {
            return ValidarSelecciones(contexto, reclasificacionesAutorizadas, null, CancellationToken.None);
        }

        public AgResultadoValidacionSeleccion ValidarSelecciones(
            AgContexto contexto,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificacionesAutorizadas,
            IReadOnlyCollection<AgReclasificacionPatron> candidatasSelladas,
            CancellationToken ct = default(CancellationToken))
        {
            if (reclasificacionesAutorizadas == null || reclasificacionesAutorizadas.Count == 0)
            {
                return AgResultadoValidacionSeleccion.Valido();
            }

            if (contexto == null || contexto.Aulas == null || contexto.AulaAsignaturaPorAulaAsig == null)
            {
                return RechazoConLog("CONFIGURACION_MODIFICADA", "Contexto no inicializado o incompleto.", null, null);
            }

            Dictionary<(int IdAula, int IdAsignatura), AgReclasificacionPatron> mapSelladas = null;
            if (candidatasSelladas != null && candidatasSelladas.Count > 0)
            {
                mapSelladas = new Dictionary<(int IdAula, int IdAsignatura), AgReclasificacionPatron>();
                foreach (var c in candidatasSelladas)
                {
                    if (c != null)
                    {
                        mapSelladas[(c.IdAula, c.IdAsignatura)] = c;
                    }
                }
            }

            var cacheCandidatasAula = new Dictionary<int, IReadOnlyList<AgReclasificacionPatron>>();
            var identidades = new HashSet<(int, int)>();

            foreach (var rec in reclasificacionesAutorizadas)
            {
                ct.ThrowIfCancellationRequested();

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

                if (aulasCoincidentes.Count != 1 || !PuedeAplicarPropuesta(contexto, aulasCoincidentes[0]))
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
                if (mapSelladas != null)
                {
                    if (!mapSelladas.TryGetValue((rec.IdAula, rec.IdAsignatura), out var candidataSellada) || candidataSellada.SesionesPendientes <= 0)
                    {
                        return RechazoConLog("MATERIA_SIN_DEFICIT",
                            $"La asignatura {rec.IdAsignatura} no presenta déficit de sesiones pendientes en el aula {rec.AulaRef} dentro del plazo evaluado.",
                            rec.IdAula, rec.IdAsignatura);
                    }
                }
                else
                {
                    if (!cacheCandidatasAula.TryGetValue(rec.IdAula, out var candidatasAula))
                    {
                        candidatasAula = IdentificarPatronesRedistribuibles(contexto, aulasCoincidentes[0], null, ct);
                        cacheCandidatasAula[rec.IdAula] = candidatasAula;
                    }

                    var candidataValida = candidatasAula.FirstOrDefault(c => c.IdAsignatura == rec.IdAsignatura);
                    if (candidataValida == null || candidataValida.SesionesPendientes <= 0)
                    {
                        return RechazoConLog("MATERIA_SIN_DEFICIT",
                            $"La asignatura {rec.IdAsignatura} no presenta déficit de sesiones pendientes en el aula {rec.AulaRef} dentro del plazo evaluado.",
                            rec.IdAula, rec.IdAsignatura);
                    }
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

        public AgResultadoEvaluacionEscalado Evaluar(AgContexto contexto)
        {
            return EvaluarInterno(contexto, null, null, null, CancellationToken.None);
        }

        public AgResultadoEvaluacionEscalado Evaluar(AgContexto contexto, int idArbol)
        {
            return EvaluarInterno(contexto, idArbol, null, null, CancellationToken.None);
        }

        public AgResultadoEvaluacionEscalado Evaluar(AgContexto contexto, int idArbol, CancellationToken ct)
        {
            return EvaluarInterno(contexto, idArbol, null, null, ct);
        }

        public AgResultadoEvaluacionEscalado Evaluar(AgContexto contexto, int idArbol,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificacionesAutorizadas)
        {
            return EvaluarInterno(contexto, idArbol, reclasificacionesAutorizadas, null, CancellationToken.None);
        }

        public AgResultadoEvaluacionEscalado Evaluar(AgContexto contexto, int idArbol,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificacionesAutorizadas, CancellationToken ct)
        {
            return EvaluarInterno(contexto, idArbol, reclasificacionesAutorizadas, null, ct);
        }

        public AgResultadoEvaluacionEscalado Evaluar(AgContexto contexto, int idArbol,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificacionesAutorizadas,
            IReadOnlyCollection<AgReclasificacionPatron> candidatasSelladas, CancellationToken ct)
        {
            return EvaluarInterno(contexto, idArbol, reclasificacionesAutorizadas, candidatasSelladas, ct);
        }

        private AgResultadoEvaluacionEscalado EvaluarInterno(AgContexto contexto, int? perfil,
            IReadOnlyCollection<AgReclasificacionPatron> reclasificacionesAutorizadas,
            IReadOnlyCollection<AgReclasificacionPatron> candidatasSelladas, CancellationToken ct)
        {
            var decisiones = new List<Decision>();
            if (contexto == null || contexto.Ventana == null) { return AgResultadoEvaluacionEscalado.Exito(decisiones); }

            var validacion = ValidarSelecciones(contexto, reclasificacionesAutorizadas, candidatasSelladas, ct);
            if (!validacion.EsValido)
            {
                return AgResultadoEvaluacionEscalado.Rechazo(validacion);
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

            // Decision es aplicable/persistible. Las aprobadas con origen recorren el simulador informativo separado.
            // Las pendientes verificadas (aprobadas sin origen) se escalan directamente para habilitar su factibilidad.
            foreach (var aula in contexto.Aulas.Where(a => PuedeAplicarPropuesta(contexto, a)).OrderBy(a => a.Id))
            {
                ct.ThrowIfCancellationRequested();
                LoggerHelper.Debug($"[AG-RANURAS-TMP] ESCALADOR_ENTRADA corrida={contexto.IdCorrida} traza={contexto.Traza?.Correlacion} perfil={perfil} aula={aula.Id} ref={aula.AulaRef} dias={aula.DiasVentana} plazoEfectivo={contexto.Ventana.FechaInicio:yyyy-MM-dd}..{contexto.Ventana.FechaFin:yyyy-MM-dd} plazoAula={aula.FechaInicio}..{aula.FechaFin}");

                string huellaInicial = AgHuellaPlanVentana.Calcular(contexto, aula);
                var programa = contexto.Programas.FirstOrDefault(p => p.Id == aula.IdPrograma);
                double unidad = AgUnidadesSemanales.HorasSesionAula(aula, AgNivelAcademicoCatalogo.Desde(programa?.Modalidad));

                var ranurasAntes = AgRanuras.PorVentana(aula.DiasVentana, contexto.Bloques, aula.Jornada, unidad);
                // FASE 2218: Obtener patrones efectivos (excluyendo AUTOMATICO) alineado con AgEvaluadorDemandaVentana
                var patronesAula = PatronesEfectivos(contexto, aula.Id);

                var opcionActual = AgOpcionesParticion.Resolver(aula.DiasVentana, ranurasAntes, patronesAula);
                var antes = AgEvaluadorDemandaVentana.Evaluar(aula.DiasVentana, aula.Jornada, unidad,
                    contexto, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct, opcion: opcionActual);

                // Paso 2: Siempre consultar candidatos de reclasificación independientemente del siguiente escalón y de la saturación
                var candidatos = IdentificarPatronesRedistribuibles(contexto, aula, antes, ct);

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

                // FASE 2218: Obtener patrones simulados efectivos (excluyendo AUTOMATICO)
                var patronesSimulados = PatronesEfectivos(ctxSimulado, aula.Id);
                var opcionMisma = AgOpcionesParticion.Resolver(aula.DiasVentana, ranurasAntes, patronesSimulados);

                // Simular primero en la misma ventana
                var evalMismaVentana = AgEvaluadorDemandaVentana.Evaluar(aula.DiasVentana, aula.Jornada, unidad,
                    ctxSimulado, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct, opcion: opcionMisma);

                string paso = AgEscaleraVentanas.Siguiente(aula.DiasVentana);
                string destino = aula.DiasVentana;
                AgResultadoEvaluacionVentana despues = evalMismaVentana;
                // Fase 2108 / Fase 2114: Retener evaluación del escalón para observabilidad fidedigna sin falsear capDespues
                AgResultadoEvaluacionVentana evalPaso = null;
                AgOpcionParticion opcionPaso = null;
                AgOpcionParticion opcionUsada = opcionMisma ?? opcionActual;

                if (autorizadasAula.Count > 0)
                {
                    // Si con reclasificación en la misma ventana aún quedan pendientes y hay escalón disponible, evaluar siguiente escalón
                    if (!evalMismaVentana.CabeEnPlazo && !string.IsNullOrEmpty(paso))
                    {
                        var ranurasPaso = AgRanuras.PorVentana(paso, ctxSimulado.Bloques, aula.Jornada, unidad);
                        var opcionesPasoSim = AgOpcionesParticion.Resolver(paso, ranurasPaso, patronesSimulados, "escalar");
                        evalPaso = AgEvaluadorDemandaVentana.Evaluar(paso, aula.Jornada, unidad,
                            ctxSimulado, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct, opcion: opcionesPasoSim);

                        if (AgEvaluadorDemandaVentana.PasoMejora(evalMismaVentana, evalPaso))
                        {
                            destino = paso;
                            despues = evalPaso;
                            opcionUsada = opcionesPasoSim;
                        }
                    }
                }
                else
                {
                    // Sin autorizaciones: flujo nominal de escalera si hay saturación
                    if (antes.HaySaturacion && !string.IsNullOrEmpty(paso))
                    {
                        var ranurasPaso = AgRanuras.PorVentana(paso, contexto.Bloques, aula.Jornada, unidad);
                        opcionPaso = AgOpcionesParticion.Resolver(paso, ranurasPaso, patronesAula, "escalar");
                        evalPaso = AgEvaluadorDemandaVentana.Evaluar(paso, aula.Jornada, unidad,
                            contexto, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct, opcion: opcionPaso);

                        if (AgEvaluadorDemandaVentana.PasoMejora(antes, evalPaso))
                        {
                            destino = paso;
                            despues = evalPaso;
                            opcionUsada = opcionPaso;
                        }
                    }
                }

                // FASE 2207 / FASE 2208: Evaluar referencia sin opción bajo la misma base de patrones (simulados si hubo reclasificación)
                // y reutilizar 'antes' como referencia si no hubo opción en la ventana actual.
                AgResultadoEvaluacionVentana referencia = autorizadasAula.Count > 0
                    ? AgEvaluadorDemandaVentana.Evaluar(aula.DiasVentana, aula.Jornada, unidad,
                          ctxSimulado, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct)
                    : (opcionActual == null
                        ? antes
                        : AgEvaluadorDemandaVentana.Evaluar(aula.DiasVentana, aula.Jornada, unidad,
                              contexto, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct));

                string particionPreferente = opcionUsada?.CarrilesTexto ?? string.Empty;
                var patronesConserva = autorizadasAula.Count > 0 ? patronesSimulados : patronesAula;
                // FASE 2219: Evaluar conservación y no-alojables sobre patrones efectivos
                bool conservaPatron = opcionUsada != null
                    && PatronesAlojablesEnOpcion(opcionUsada, patronesConserva)
                    && AgParticionCarriles.ContarPatronesNoAlojables(opcionUsada.Carriles, patronesConserva) == 0;

                // FASE 2209: Comparar franjas completas contra la partición canónica para detectar cambios en días con múltiples franjas (ej. 6M vs 6T)
                int cantCanonicas = AgCarriles.Cantidad(ranurasAntes);
                var canonicas = Enumerable.Range(0, cantCanonicas)
                    .Select(c => (IReadOnlyList<AgRanura>)AgCarriles.DelCarril(ranurasAntes, c))
                    .ToList();
                string canonicaTexto = string.Join("|", canonicas.Select(c => string.Join(",", c.Select(TokenRanura))));

                bool particionDifiere = opcionUsada != null
                    && !string.IsNullOrEmpty(particionPreferente)
                    && (!string.Equals(particionPreferente, canonicaTexto, StringComparison.OrdinalIgnoreCase)
                        || !ParticionesIguales(evalMismaVentana.Plan?.Particion, referencia.Plan?.Particion));

                // FASE 2215: Contar patrones no alojables sobre la referencia real (no sobre canónicos)
                int noAlojRef = AgParticionCarriles.ContarPatronesNoAlojables(
                    referencia.Plan?.Particion?.Select(b => b.Select(dia => new AgRanura(dia, AgFranjaJornada.NOCHE, unidad)).ToList())
                        .Cast<IReadOnlyList<AgRanura>>().ToList() ?? canonicas,
                    patronesConserva);
                int noAlojOpt = opcionUsada?.Carriles != null
                    ? AgParticionCarriles.ContarPatronesNoAlojables(opcionUsada.Carriles, patronesConserva)
                    : 0;

                // FASE 2206 / FASE 2215: particionDifiere es condición necesaria. Requiere reducción de no-alojables, infactibles o demanda.
                bool particionMejora = opcionUsada != null
                    && particionDifiere
                    && despues.DemandaPendiente <= antes.DemandaPendiente
                    && (noAlojOpt < noAlojRef
                        || (evalMismaVentana.Plan?.Infactibles?.Count ?? 0) < (referencia.Plan?.Infactibles?.Count ?? 0)
                        || evalMismaVentana.DemandaPendiente < referencia.DemandaPendiente);

                // FASE 2216: Referencia comparada de la ventana destino cuando el destino escala
                var referenciaComparada = string.Equals(destino, aula.DiasVentana, StringComparison.OrdinalIgnoreCase)
                    ? referencia
                    : (particionMejora
                        ? AgEvaluadorDemandaVentana.Evaluar(destino, aula.Jornada, unidad,
                              contexto, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct)
                        : null);

                bool hayMejoraParticion = !string.IsNullOrEmpty(particionPreferente)
                    && particionMejora
                    && particionDifiere
                    && (referenciaComparada == null || (despues.Plan?.Infactibles?.Count ?? 0) <= (referenciaComparada.Plan?.Infactibles?.Count ?? 0));

                // FASE 1856 / FASE 2108 / FASE 2114: Trazabilidad enriquecida y fidedigna de omisión del escalador en nivel Info
                // Si no hay saturación, no hay candidatos y no hay autorizaciones, omitir
                if (!antes.HaySaturacion && candidatos.Count == 0 && autorizadasAula.Count == 0 && !hayMejoraParticion)
                {
                    LoggerHelper.Info(
                        $"[ESCALADO] ESCALADOR_OMITIR aula={aula.Id} ref={aula.AulaRef} dias={aula.DiasVentana} " +
                        $"capAntes={antes.CapacidadTotal} pendAntes={antes.DemandaPendiente} antesError='{antes.Error}' " +
                        $"paso='{paso}' pasoCap={(evalPaso != null ? evalPaso.CapacidadTotal.ToString() : "n/a")} " +
                        $"pasoPend={(evalPaso != null ? evalPaso.DemandaPendiente.ToString() : "n/a")} " +
                        $"pasoError='{(evalPaso != null ? evalPaso.Error : string.Empty)}' candidatos={candidatos.Count} motivo=cabe_en_plazo_y_sin_candidatos");
                    contexto.Traza?.Dato("EscaladorOmitir", $"aula={aula.Id} ref={aula.AulaRef} motivo=cabe_en_plazo_y_sin_candidatos");
                    continue;
                }

                // Si no hay cambio de ventana ni candidatos ni autorizaciones que mejoren, omitir
                bool hayCambioVentana = !string.Equals(aula.DiasVentana, destino, StringComparison.OrdinalIgnoreCase);
                bool hayAutorizaciones = autorizadasAula.Count > 0;
                bool hayCandidatos = candidatos.Count > 0;

                if (!hayCambioVentana && !hayAutorizaciones && !hayCandidatos && !hayMejoraParticion)
                {
                    LoggerHelper.Info(
                        $"[ESCALADO] ESCALADOR_OMITIR aula={aula.Id} ref={aula.AulaRef} dias={aula.DiasVentana} " +
                        $"capAntes={antes.CapacidadTotal} pendAntes={antes.DemandaPendiente} antesError='{antes.Error}' " +
                        $"paso='{paso}' pasoCap={(evalPaso != null ? evalPaso.CapacidadTotal.ToString() : "n/a")} " +
                        $"pasoPend={(evalPaso != null ? evalPaso.DemandaPendiente.ToString() : "n/a")} " +
                        $"pasoError='{(evalPaso != null ? evalPaso.Error : string.Empty)}' candidatos={candidatos.Count} motivo=sin_propuesta_viable");
                    contexto.Traza?.Dato("EscaladorOmitir", $"aula={aula.Id} ref={aula.AulaRef} motivo=sin_propuesta_viable");
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
                    ParticionPreferente = particionPreferente,
                    ParticionMejora = particionMejora,
                    ParticionDifiereDeReferencia = particionDifiere,
                    ConservaPatronOrigen = conservaPatron,
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

                if (aula.Estado == AgEstadoAula.Aprobado)
                {
                    contexto.Traza?.Dato("ESCALON_SIN_ORIGEN",
                        $"aula={aula.Id} desde={aula.DiasVentana} hacia={destino} origen=AUSENTE filas=0");
                    LoggerHelper.Info($"[ESCALON SIN ORIGEN] corrida={contexto.IdCorrida} aula={aula.Id} ref={aula.AulaRef} subunidad={aula.Subunidad} desde={d.Desde} hacia={d.Hacia} demanda={d.DemandaSesiones} mejora={d.Mejora}");
                }

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

            return AgResultadoEvaluacionEscalado.Exito(decisiones);
        }

        internal static AgContexto ClonarContextoParaSimulacion(AgContexto ctx)
        {
            var clon = new AgContexto
            {
                IdCorrida = ctx.IdCorrida,
                IdPeriodo = ctx.IdPeriodo,
                IdPrograma = ctx.IdPrograma,
                IdSede = ctx.IdSede,
                Subunidad = ctx.Subunidad,
                Ventana = CopiarModelo(ctx.Ventana),
                Programas = CopiarModelo(ctx.Programas),
                Asignaturas = CopiarModelo(ctx.Asignaturas),
                Aulas = CopiarModelo(ctx.Aulas),
                Bloques = CopiarModelo(ctx.Bloques),
                Festivos = ctx.Festivos.ToList(),
                Presenciales = ctx.Presenciales.ToList(),
                Docentes = CopiarModelo(ctx.Docentes),
                Sedes = CopiarModelo(ctx.Sedes)
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

        // Sólo modelos de entrada, no contexto/índices, controles, conexiones o trazas compartidas.
        private static T CopiarModelo<T>(T modelo) => Newtonsoft.Json.JsonConvert.DeserializeObject<T>(
            Newtonsoft.Json.JsonConvert.SerializeObject(modelo));

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
                if (decisiones.Any(d => d != null && d.SoloParticion))
                {
                    resultado.Advertencia = "Sin cambios persistidos; la partición se recalcula en cada corrida.";
                }
                else if (decisiones.Any(d => d != null && !d.Aplica))
                {
                    resultado.Advertencia = "Sin cambios persistidos; hay decisiones no aplicables (revise demanda/capacidad).";
                }
                return resultado;
            }

            Decision dActual = null;
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) { db.Open(); }
                    using (var tx = db.BeginTransaction())
                    {
                        foreach (var d in decisionesAplicables)
                        {
                            dActual = d;
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
                            if (!PuedeAplicarPropuesta(contexto, aulas[0]))
                                throw new InvalidOperationException("AULA_NO_ELEGIBLE: sólo Disponible o Aprobada sin origen verificada admiten escalado.");

                            // Postcondición estricta: 0 filas productivas de la cohorte para aulas Aprobado
                            if (aulas[0].Estado == AgEstadoAula.Aprobado)
                            {
                                if (contexto.PlanAprobadoSnapshot == null || contexto.TieneOrigenConservable(aulas[0].Id))
                                {
                                    throw new InvalidOperationException("INVARIANTE_VIOLADO: Aula aprobada con origen existente no admite escalado de escritura.");
                                }
                                contexto.Traza?.Dato("ESCALON_SIN_ORIGEN",
                                    $"aula={aulas[0].Id} desde={d.Desde} hacia={d.Hacia} origen=AUSENTE filas=0");
                            }
                            int idAula = d.IdAula;

                            string huellaContexto = AgHuellaPlanVentana.Calcular(contexto, aulas[0]);
                            string huellaSql = AgHuellaEscaladoPersistido.Calcular(db, tx, d.AulaRef, d.Subunidad, idArbol);
                            bool contextoCompatible = huellaContexto == d.HuellaContextoAntes || huellaContexto == d.HuellaContextoDespues;
                            LoggerHelper.Debug($"[AG-RANURAS-TMP] GUARDAR_HUELLAS perfil={idArbol} ref={d.AulaRef} contextoAntes={huellaContexto == d.HuellaContextoAntes} contextoDespues={huellaContexto == d.HuellaContextoDespues} sqlAntes={huellaSql == d.HuellaPersistidaAntes} sqlDespues={huellaSql == d.HuellaPersistidaDespues}");

                            bool cambiaVentana = !string.IsNullOrWhiteSpace(d.Hacia)
                                && !string.Equals(d.Desde, d.Hacia, StringComparison.OrdinalIgnoreCase);
                            string destinoEsperado = cambiaVentana ? d.Hacia : null;

                            var liberarEsperado = d.ConservaPatronOrigen ? null : d.IdAsignaturasBottleneck;

                            // Reaplicación idempotente: si ya está en el postestado exacto, confirmar sin nuevas escrituras
                            if (d.PerfilAplicado == idArbol && contextoCompatible && !string.IsNullOrEmpty(d.HuellaPersistidaDespues)
                                && huellaSql == d.HuellaPersistidaDespues
                                && AgHuellaEscaladoPersistido.EstadoCompleto(db, tx, d.AulaRef, d.Subunidad, idArbol, destinoEsperado, liberarEsperado, d.ReclasificacionesAutorizadas))
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
                                idArbol, destinoEsperado, liberarEsperado, d.ReclasificacionesAutorizadas);

                            LoggerHelper.Info($"[ESCALADO] APLICAR ref={d.AulaRef} desde={d.Desde} hacia={d.Hacia} " +
                                $"conserva={d.ConservaPatronOrigen} liberar={(liberarEsperado?.Count ?? 0)} perfil={idArbol}");

                            int modificadas = 0;
                            if (cambiaVentana)
                            {
                                modificadas = await db.ExecuteAsync(new CommandDefinition(@"
UPDATE ag_aula SET dias_ventana=@Hacia
WHERE aula_ref=@Ref AND COALESCE(NULLIF(TRIM(subunidad),''),@Sub)=@Sub
AND COALESCE(dias_ventana,'')=@Desde;",
                                    new { d.Hacia, d.Desde, Ref = d.AulaRef, Sub = d.Subunidad ?? string.Empty }, tx, cancellationToken: ct)).ConfigureAwait(false);
                                if (modificadas != 1) { throw new InvalidOperationException("Configuración modificada durante el guardado."); }

                                if (idArbol > 0)
                                {
                                    var nodos = (await db.QueryAsync<string>(new CommandDefinition(@"
SELECT dias_ventana FROM ag_arbol_nodo
WHERE idarbol=@Perfil AND aula_ref=@Ref
AND COALESCE(NULLIF(TRIM(subunidad),''),@Sub)=@Sub;",
                                        new { Perfil = idArbol, Ref = d.AulaRef, Sub = d.Subunidad ?? string.Empty },
                                        tx, cancellationToken: ct)).ConfigureAwait(false)).ToList();

                                    if (nodos.Count != 1)
                                        throw new InvalidOperationException("Nodo del perfil ausente o duplicado.");

                                    string nodoActual = nodos[0] ?? string.Empty;

                                    int nodosAfectados = await db.ExecuteAsync(new CommandDefinition(@"
UPDATE ag_arbol_nodo SET dias_ventana=@Hacia
WHERE idarbol=@Perfil AND aula_ref=@Ref
AND COALESCE(NULLIF(TRIM(subunidad),''),@Sub)=@Sub
AND COALESCE(dias_ventana,'')=@NodoActual;",
                                        new { d.Hacia, Perfil = idArbol, Ref = d.AulaRef,
                                              Sub = d.Subunidad ?? string.Empty, NodoActual = nodoActual },
                                        tx, cancellationToken: ct)).ConfigureAwait(false);

                                    if (nodosAfectados != 1)
                                        throw new InvalidOperationException("Nodo del perfil ausente, duplicado o modificado.");

                                    modificadas += nodosAfectados;
                                }
                            }

                            if (!d.ConservaPatronOrigen
                                && d.IdAsignaturasBottleneck != null
                                && d.IdAsignaturasBottleneck.Count > 0)
                            {
                                int afectadasBottleneck = await db.ExecuteAsync(new CommandDefinition(@"
UPDATE ag_aula_asignatura SET dias_patron=''
WHERE id_aula = @IdAula
AND id_asignatura IN @Ids AND UPPER(COALESCE(origen_patron,''))='AUTOMATICO';",
                                    new { IdAula = idAula, Ids = d.IdAsignaturasBottleneck },
                                    tx, cancellationToken: ct)).ConfigureAwait(false);
                                modificadas += afectadasBottleneck;
                            }

                            if (d.ReclasificacionesAutorizadas != null && d.ReclasificacionesAutorizadas.Count > 0)
                            {
                                foreach (var rec in d.ReclasificacionesAutorizadas)
                                {
                                    if (rec == null || !rec.Autorizado
                                        || rec.IdAula <= 0 || rec.IdAula != idAula
                                        || !string.Equals(rec.AulaRef, d.AulaRef, StringComparison.Ordinal)
                                        || !string.Equals(rec.Subunidad ?? string.Empty, d.Subunidad ?? string.Empty, StringComparison.Ordinal))
                                    {
                                        throw new InvalidOperationException($"La propuesta fue modificada: reclasificación no autorizada o no coincidente para asignatura {rec?.IdAsignatura} en aula {d.AulaRef}. Recalcular.");
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
                                || !AgHuellaEscaladoPersistido.EstadoCompleto(db, tx, d.AulaRef, d.Subunidad, idArbol, destinoEsperado, liberarEsperado, d.ReclasificacionesAutorizadas))
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
                LoggerHelper.Info($"[ESCALADO] APLICAR_FALLO ref={dActual?.AulaRef} desde={dActual?.Desde} hacia={dActual?.Hacia} " +
                    $"conserva={dActual?.ConservaPatronOrigen} tipo={ex.GetType().Name} motivo={ex.Message}");
                LoggerHelper.Debug($"[AG-RANURAS-TMP] GUARDAR_REVERTIDO perfil={idArbol} tipo={ex.GetType().Name} motivo={ex.Message}");
                return new ResultadoAplicacion { Error = ex is OperationCanceledException ? "Operación cancelada." : ex.Message };
            }

            foreach (var par in posteriores)
            {
                par.Key.HuellaPersistidaDespues = par.Value;
                par.Key.PerfilAplicado = idArbol;
            }

            bool huboCambioVentana = decisionesAplicables.Any(x =>
                !string.IsNullOrWhiteSpace(x.Hacia)
                && !string.Equals(x.Desde, x.Hacia, StringComparison.OrdinalIgnoreCase));
            bool huboReclasificacion = decisionesAplicables.Any(x => x.ReclasificacionesAutorizadas != null && x.ReclasificacionesAutorizadas.Count > 0);
            bool huboLiberacionBottleneck = decisionesAplicables.Any(x => !x.ConservaPatronOrigen && x.IdAsignaturasBottleneck != null && x.IdAsignaturasBottleneck.Count > 0);

            if (resultado.Confirmadas > 0 && resultado.Modificadas == 0 && !huboCambioVentana && !huboReclasificacion && !huboLiberacionBottleneck)
            {
                resultado.Exito = true;
                resultado.Confirmadas = 0; // SinCambios
                resultado.Advertencia = "Sin cambios persistidos; la partición se recalcula en cada corrida.";
            }
            else
            {
                resultado.Exito = true;
            }
            try
            {
                if (resultado.Modificadas > 0) { ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty(); }
                LoggerHelper.Info("[Escalado] Confirmadas=" + resultado.Confirmadas + "; modificadas=" + resultado.Modificadas);
            }
            catch (Exception ex) { resultado.Advertencia = "Guardado confirmado; falló la notificación posterior: " + ex.Message; }

            return resultado;
        }

        /// <summary>
        /// Obtiene los patrones efectivos de un aula excluyendo orígenes AUTOMATICO (Fase 2218, Fase 2250, Fase 2251).
        /// Incluye fallback a AulaAsignaturas y construcción de índices si el índice por aula-asignatura no ha sido construido o está vacío.
        /// Misma semántica que AgEvaluadorDemandaVentana.cs:124-126.
        /// </summary>
        internal static IReadOnlyList<string> PatronesEfectivos(AgContexto ctx, int idAula)
        {
            return AgContexto.RelacionesAula(ctx)
                .Where(a => a.IdAula == idAula
                    && !string.Equals(a.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrWhiteSpace(a.DiasPatron))
                .Select(a => a.DiasPatron).ToList();
        }

        /// <summary>
        /// Indica si TODOS los patrones son alojables en algún carril de la opción (muchos-a-uno).
        /// Criterio estructural: día del patrón presente en el carril; si incluye sábado, el carril
        /// debe tener al menos una ranura de día 6. NO evalúa capacidad semanal ni demanda:
        /// eso lo valida AgPlanificadorOleadas. Se usa solo para decidir la limpieza de dias_patron.
        /// (Fases 2247, 2248; 2252 retirada por diseño, Fase 2254).
        /// </summary>
        internal static bool PatronesAlojablesEnOpcion(AgOpcionParticion opcion, IReadOnlyList<string> patrones)
        {
            if (opcion == null || opcion.Carriles == null || opcion.Carriles.Count == 0) return false;
            if (patrones == null || patrones.Count == 0) return true;

            var patronesDias = patrones
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Select(p => (IReadOnlyList<int>)AgCarriles.Ventana(p))
                .Where(d => d.Count > 0)
                .ToList();

            if (patronesDias.Count == 0) return true;

            // Fases 2247 / 2248: Modelo muchos-a-uno sin veto blanket de sábado.
            for (int i = 0; i < patronesDias.Count; i++)
            {
                var dias = patronesDias[i];
                bool alojable = false;
                for (int c = 0; c < opcion.Carriles.Count; c++)
                {
                    var carril = opcion.Carriles[c];
                    var diasCarril = carril.Select(r => r.Dia).Distinct().ToHashSet();
                    if (dias.All(d => diasCarril.Contains(d)))
                    {
                        // Si el patrón requiere sábado (día 6), el carril debe contener ranura de sábado
                        if (dias.Contains(6) && !carril.Any(r => r.Dia == 6))
                        {
                            continue;
                        }

                        alojable = true;
                        break;
                    }
                }

                if (!alojable)
                {
                    return false;
                }
            }

            return true;
        }

        private static string TokenRanura(AgRanura r)
        {
            if (r == null) return string.Empty;
            if (r.Dia == 6)
            {
                if (string.Equals(r.Franja, AgFranjaJornada.MANANA, StringComparison.OrdinalIgnoreCase)) return "6M";
                if (string.Equals(r.Franja, AgFranjaJornada.TARDE, StringComparison.OrdinalIgnoreCase)) return "6T";
            }
            return r.Dia.ToString(CultureInfo.InvariantCulture);
        }

        private static bool ParticionesIguales(IReadOnlyList<IReadOnlyList<int>> a, IReadOnlyList<IReadOnlyList<int>> b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null || a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++)
            {
                if (a[i] == null || b[i] == null || a[i].Count != b[i].Count) return false;
                for (int j = 0; j < a[i].Count; j++)
                {
                    if (a[i][j] != b[i][j]) return false;
                }
            }
            return true;
        }

        internal static IReadOnlyList<string> HuellaAccionPartes(Decision d)
        {
            var partes = new List<string>
            {
                "propuesta-escalado-v5", d.IdAula.ToString(CultureInfo.InvariantCulture), d.AulaRef, d.Subunidad, d.Desde, d.Hacia,
                d.Mejora.ToString(), d.CapacidadAntes.ToString(CultureInfo.InvariantCulture),
                d.CapacidadDespues.ToString(CultureInfo.InvariantCulture),
                d.DemandaPendiente.ToString(CultureInfo.InvariantCulture),
                d.FechaLimite.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                d.PerfilPropuesto.HasValue ? d.PerfilPropuesto.Value.ToString(CultureInfo.InvariantCulture) : null,
                d.IdAsignaturasBottleneck == null ? null : string.Join(",", d.IdAsignaturasBottleneck.OrderBy(id => id)),
                d.ParticionPreferente ?? string.Empty
            };

            partes.Add(d.ConservaPatronOrigen ? "cp=1" : "cp=0");

            if (d.ReclasificacionesAutorizadas != null && d.ReclasificacionesAutorizadas.Count > 0)
            {
                partes.AddRange(d.ReclasificacionesAutorizadas
                    .OrderBy(r => r.IdAula).ThenBy(r => r.IdAsignatura)
                    .Select(r => $"rec:{r.IdAula}:{r.AulaRef}:{r.Subunidad}:{r.IdAsignatura}:{r.OrigenAnterior}->{r.OrigenPropuesto}:{r.DiasPatronAnterior}->{r.DiasPatronPropuesto}:{r.Autorizado}"));
            }

            return partes;
        }

        internal static string HuellaAccion(Decision d) =>
            AgHuellaPlanVentana.Hash(HuellaAccionPartes(d));
    }
}
