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
    /// <summary>Propone un paso de ventana y verifica su instantánea antes de aplicar el lote transaccional.</summary>
    internal sealed class AgEscaladorRanuraAula
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public AgEscaladorRanuraAula(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory;
        }

        internal sealed class Decision
        {
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
            public DateTime? FechaFinProyectada { get; set; }
            public DateTime? FechaFinMinima { get { return FechaFinProyectada; } set { FechaFinProyectada = value; } }
            public int RanurasOciosas { get; set; }
            public int DemandaPendiente { get; set; }
            public int MateriasFueraVentana { get; set; }
            public List<int> IdAsignaturasBottleneck { get; set; } = new List<int>();
            public bool Aplica => !string.IsNullOrWhiteSpace(Hacia) && Mejora && CapacidadDespues > CapacidadAntes;
            internal AgContexto ContextoOrigen { get; set; }
            internal int? PerfilPropuesto { get; set; }
            internal int? PerfilAplicado { get; set; }
            internal string HuellaDecision { get; set; }
            internal string HuellaContextoAntes { get; set; }
            internal string HuellaContextoDespues { get; set; }
            internal string HuellaPersistidaAntes { get; set; }
            internal string HuellaPersistidaDespues { get; set; }
        }

        /// <summary>Compatibilidad: evaluación sin perfil explícito. Una conexión nula permite solo dry-run.</summary>
        public IReadOnlyList<Decision> Evaluar(AgContexto contexto)
        {
            return EvaluarInterno(contexto, null, CancellationToken.None);
        }

        /// <summary>Evalúa y sella el perfil seleccionado sin escribir en la base de datos.</summary>
        public IReadOnlyList<Decision> Evaluar(AgContexto contexto, int idArbol)
        {
            return EvaluarInterno(contexto, idArbol, CancellationToken.None);
        }

        public IReadOnlyList<Decision> Evaluar(AgContexto contexto, int idArbol, CancellationToken ct)
        {
            return EvaluarInterno(contexto, idArbol, ct);
        }

        private IReadOnlyList<Decision> EvaluarInterno(AgContexto contexto, int? perfil, CancellationToken ct)
        {
            var decisiones = new List<Decision>();
            if (contexto == null || contexto.Ventana == null) { return decisiones; }
            foreach (var aula in contexto.Aulas.Where(a => a.Estado == AgEstadoAula.Disponible
                && !string.IsNullOrWhiteSpace(a.AulaRef)).OrderBy(a => a.Id))
            {
                ct.ThrowIfCancellationRequested();
                LoggerHelper.Debug($"[AG-RANURAS-TMP] ESCALADOR_ENTRADA corrida={contexto.IdCorrida} traza={contexto.Traza?.Correlacion} perfil={perfil} aula={aula.Id} ref={aula.AulaRef} dias={aula.DiasVentana} plazoEfectivo={contexto.Ventana.FechaInicio:yyyy-MM-dd}..{contexto.Ventana.FechaFin:yyyy-MM-dd} plazoAula={aula.FechaInicio}..{aula.FechaFin}");
                string huellaInicial = AgHuellaPlanVentana.Calcular(contexto, aula);
                var programa = contexto.Programas.FirstOrDefault(p => p.Id == aula.IdPrograma);
                double unidad = AgUnidadesSemanales.HorasSesionAula(aula, AgNivelAcademicoCatalogo.Desde(programa?.Modalidad));
                var antes = AgEvaluadorDemandaVentana.Evaluar(aula.DiasVentana, aula.Jornada, unidad,
                    contexto, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct);
                if (!antes.HaySaturacion)
                {
                    LoggerHelper.Debug($"[AG-RANURAS-TMP] ESCALADOR_OMITIR aula={aula.Id} motivo=cabe_en_plazo_efectivo");
                    continue;
                }
                string paso = AgEscaleraVentanas.Siguiente(aula.DiasVentana);
                if (string.IsNullOrEmpty(paso))
                {
                    LoggerHelper.Debug($"[AG-RANURAS-TMP] ESCALADOR_OMITIR aula={aula.Id} motivo=ventana_terminal dias={aula.DiasVentana} pendientes={antes.DemandaPendiente} error={antes.Error}");
                    continue;
                }
                var despues = AgEvaluadorDemandaVentana.Evaluar(paso, aula.Jornada, unidad,
                    contexto, aula, contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin, ct);
                if (!AgEvaluadorDemandaVentana.PasoMejora(antes, despues))
                {
                    LoggerHelper.Debug($"[AG-RANURAS-TMP] ESCALADOR_OMITIR aula={aula.Id} motivo=paso_sin_mejora desde={aula.DiasVentana} hacia={paso} capacidad={antes.CapacidadTotal}->{despues.CapacidadTotal} pendientes={antes.DemandaPendiente}->{despues.DemandaPendiente} errorAntes={antes.Error} errorDespues={despues.Error}");
                    continue;
                }
                var automaticas = contexto.AulaAsignaturaPorAulaAsig.Values.Where(aa => aa.IdAula == aula.Id
                    && string.Equals(aa.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase))
                    .Select(aa => aa.IdAsignatura).Distinct().OrderBy(id => id).ToList();
                var d = new Decision
                {
                    AulaRef = aula.AulaRef, Subunidad = aula.Subunidad, Desde = aula.DiasVentana, Hacia = paso,
                    DemandaSesiones = antes.DemandaTotal, CapacidadAntes = antes.CapacidadTotal, CapacidadDespues = despues.CapacidadTotal,
                    Motivo = antes.Motivo, SemanasRequeridas = despues.SemanasRequeridas, FechaFinProyectada = despues.FechaFinProyectada,
                    RanurasOciosas = despues.RanurasOciosas, DemandaPendiente = despues.DemandaPendiente,
                    MateriasFueraVentana = antes.MateriasFueraVentana.Count, IdAsignaturasBottleneck = automaticas,
                    ContextoOrigen = contexto, PerfilPropuesto = perfil,
                    HuellaContextoAntes = huellaInicial,
                    HuellaContextoDespues = AgHuellaPlanVentana.Calcular(contexto, aula, paso, automaticas)
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
                LoggerHelper.Debug($"[AG-RANURAS-TMP] PROPUESTA aula={aula.Id} ref={aula.AulaRef} desde={d.Desde} hacia={d.Hacia} capacidad={d.CapacidadAntes}->{d.CapacidadDespues} pendientes={d.DemandaPendiente} automaticas={string.Join(",", automaticas)} sellada=true");
            }
            return decisiones;
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

        /// <summary>Devuelve aulas confirmadas; una reaplicación completa no vuelve a escribir.</summary>
        public async Task<int> AplicarAsync(IReadOnlyList<Decision> decisiones, int idArbol, CancellationToken ct)
        {
            var resultado = await AplicarDetalladoAsync(decisiones, idArbol, ct).ConfigureAwait(false);
            return resultado.Confirmadas;
        }

        public Task<ResultadoAplicacion> AplicarDetalladoAsync(IReadOnlyList<Decision> decisiones, int idArbol, CancellationToken ct)
        {
            return AplicarDetalladoAsync(decisiones, idArbol, null, ct);
        }

        /// <summary>La UI entrega un contexto recién leído. La instantánea SQL se verifica dentro de la transacción.</summary>
        public async Task<ResultadoAplicacion> AplicarDetalladoAsync(IReadOnlyList<Decision> decisiones,
            int idArbol, AgContexto contextoVigente, CancellationToken ct)
        {
            var resultado = new ResultadoAplicacion();
            var posteriores = new Dictionary<Decision, string>();
            if (decisiones == null || decisiones.Count == 0) { resultado.Exito = true; return resultado; }
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) { db.Open(); }
                    using (var tx = db.BeginTransaction())
                    {
                        foreach (var d in decisiones)
                        {
                            ct.ThrowIfCancellationRequested();
                            if (d == null || !d.Aplica || d.ContextoOrigen == null || string.IsNullOrEmpty(d.HuellaPersistidaAntes)
                                || string.IsNullOrEmpty(d.HuellaDecision))
                            {
                                throw new InvalidOperationException("Decisión sin instantánea válida. Recalcular propuesta.");
                            }
                            if (HuellaAccion(d) != d.HuellaDecision) { throw new InvalidOperationException("La propuesta fue modificada. Recalcular."); }
                            if ((d.PerfilPropuesto.HasValue && d.PerfilPropuesto.Value != idArbol)
                                || (d.PerfilAplicado.HasValue && d.PerfilAplicado.Value != idArbol))
                            {
                                throw new InvalidOperationException("Cambió el perfil de la propuesta. Recalcular.");
                            }
                            var contexto = contextoVigente ?? d.ContextoOrigen;
                            var aulas = contexto.Aulas.Where(a => string.Equals(a.AulaRef, d.AulaRef, StringComparison.Ordinal)
                                && string.Equals(a.Subunidad ?? string.Empty, d.Subunidad ?? string.Empty, StringComparison.Ordinal)).ToList();
                            if (aulas.Count != 1) { throw new InvalidOperationException("Configuración modificada: cambió el aula."); }
                            string huellaContexto = AgHuellaPlanVentana.Calcular(contexto, aulas[0]);
                            string huellaSql = AgHuellaEscaladoPersistido.Calcular(db, tx, d.AulaRef, d.Subunidad, idArbol);
                            bool contextoCompatible = huellaContexto == d.HuellaContextoAntes || huellaContexto == d.HuellaContextoDespues;
                            LoggerHelper.Debug($"[AG-RANURAS-TMP] GUARDAR_HUELLAS perfil={idArbol} ref={d.AulaRef} contextoAntes={huellaContexto == d.HuellaContextoAntes} contextoDespues={huellaContexto == d.HuellaContextoDespues} sqlAntes={huellaSql == d.HuellaPersistidaAntes} sqlDespues={huellaSql == d.HuellaPersistidaDespues}");
                            // Reaplicación: comprobar todo el postestado antes de exigir el estado anterior.
                            if (d.PerfilAplicado == idArbol && contextoCompatible && !string.IsNullOrEmpty(d.HuellaPersistidaDespues)
                                && huellaSql == d.HuellaPersistidaDespues
                                && AgHuellaEscaladoPersistido.EstadoCompleto(db, tx, d.AulaRef, d.Subunidad, idArbol, d.Hacia, d.IdAsignaturasBottleneck))
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
                                idArbol, d.Hacia, d.IdAsignaturasBottleneck);
                            int modificadas = await db.ExecuteAsync(new CommandDefinition(@"
UPDATE ag_aula SET dias_ventana=@Hacia
WHERE aula_ref=@Ref AND COALESCE(NULLIF(TRIM(subunidad),''),@Sub)=@Sub
AND COALESCE(dias_ventana,'')=@Desde;",
                                new { d.Hacia, d.Desde, Ref=d.AulaRef, Sub=d.Subunidad ?? string.Empty }, tx, cancellationToken: ct)).ConfigureAwait(false);
                            if (modificadas != 1) { throw new InvalidOperationException("Configuración modificada durante el guardado."); }
                            if (idArbol > 0)
                            {
                                int nodos = await db.ExecuteAsync(new CommandDefinition(@"
UPDATE ag_arbol_nodo SET dias_ventana=@Hacia
WHERE idarbol=@Perfil AND aula_ref=@Ref AND COALESCE(NULLIF(TRIM(subunidad),''),@Sub)=@Sub
AND COALESCE(dias_ventana,'')=@Desde;",
                                    new { d.Hacia, d.Desde, Perfil=idArbol, Ref=d.AulaRef, Sub=d.Subunidad ?? string.Empty }, tx, cancellationToken: ct)).ConfigureAwait(false);
                                if (nodos != 1) { throw new InvalidOperationException("Nodo del perfil ausente, duplicado o modificado."); }
                            }
                            if (d.IdAsignaturasBottleneck.Count > 0)
                            {
                                await db.ExecuteAsync(new CommandDefinition(@"
UPDATE ag_aula_asignatura SET dias_patron=''
WHERE id_aula IN (SELECT id FROM ag_aula WHERE aula_ref=@Ref AND COALESCE(NULLIF(TRIM(subunidad),''),@Sub)=@Sub)
AND id_asignatura IN @Ids AND UPPER(COALESCE(origen_patron,''))='AUTOMATICO';",
                                    new { Ref=d.AulaRef, Sub=d.Subunidad ?? string.Empty, Ids=d.IdAsignaturasBottleneck }, tx, cancellationToken: ct)).ConfigureAwait(false);
                            }
                            if (AgHuellaEscaladoPersistido.Calcular(db, tx, d.AulaRef, d.Subunidad, idArbol) != esperado
                                || !AgHuellaEscaladoPersistido.EstadoCompleto(db, tx, d.AulaRef, d.Subunidad, idArbol, d.Hacia, d.IdAsignaturasBottleneck))
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
                // Contadores nuevos: ningún resultado provisional sobrevive al rollback.
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
            return AgHuellaPlanVentana.Hash(new[]
            {
                "propuesta-escalado-v1", d.AulaRef, d.Subunidad, d.Desde, d.Hacia,
                d.Mejora.ToString(), d.CapacidadAntes.ToString(CultureInfo.InvariantCulture),
                d.CapacidadDespues.ToString(CultureInfo.InvariantCulture),
                d.PerfilPropuesto.HasValue ? d.PerfilPropuesto.Value.ToString(CultureInfo.InvariantCulture) : null,
                d.IdAsignaturasBottleneck == null ? null : string.Join(",", d.IdAsignaturasBottleneck.OrderBy(id => id))
            });
        }
    }
}
