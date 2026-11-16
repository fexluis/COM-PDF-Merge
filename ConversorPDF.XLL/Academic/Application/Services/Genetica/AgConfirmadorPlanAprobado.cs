using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Ratifica origen histórico y registra exclusivamente un candidato aislado; T5/T6 conservan sus guardas.</summary>
    internal sealed class AgConfirmadorPlanAprobado
    {
        internal sealed class Resultado
        {
            internal string Codigo { get; }
            internal AgCandidatoPlanAprobado Candidato { get; }
            internal AgPersistenciaRatificacion.Resultado Publicacion { get; }
            internal Resultado(string codigo, AgCandidatoPlanAprobado candidato = null, AgPersistenciaRatificacion.Resultado publicacion = null)
            { Codigo = codigo; Candidato = candidato; Publicacion = publicacion; }
        }

        private readonly AcademicDbSession _session;
        private readonly EncryptedSQLiteContext _vault;
        private readonly AgCandidatoPlanAprobadoRepository _repo = new AgCandidatoPlanAprobadoRepository();
        internal AgConfirmadorPlanAprobado(AcademicDbSession session, EncryptedSQLiteContext vault)
        { _session = session; _vault = vault; }

        /// <summary>Captura la huella SQL mostrada antes del diálogo, nunca una huella proporcionada por el cliente.</summary>
        internal async Task<string> CapturarConfiguracionAsync(AgContexto c, int idAula, AgPropuestaEscaladoAprobada p, CancellationToken ct)
        {
            if (_session == null || _vault == null || c?.PlanAprobadoSnapshot == null)
                return null;
            try
            {
                return await _session.ExecuteAsync(async db =>
                {
                    ct.ThrowIfCancellationRequested();
                    using (var tx = db.BeginTransaction())
                    {
                        var aula = c.Aulas.SingleOrDefault(a => a.Id == idAula);
                        if (!MismoTemporal(db, tx) || !OrigenContextual(c, p, aula)) return null;
                        if (!AgHuellaEscaladoPersistido.CoincideOrigen(db, tx, c, aula)) return null;
                        var leido = await AgPlanAprobadoRepository.ObtenerEnTransaccionAsync(db, tx, c.ManifiestoCobertura, ct);
                        if (leido.IsFailure || leido.Value.HuellaOrigen != c.PlanAprobadoSnapshot.HuellaOrigen
                            || leido.Value.EstadoOrigen != "REQUIERE_RATIFICACION") return null;
                        return AgHuellaEscaladoPersistido.Calcular(db, tx, aula.AulaRef, aula.Subunidad, 0);
                    }
                });
            }
            catch { return null; }
        }

        /// <summary>Encuentra un candidato publicado de otra corrida para mostrarlo, jamás para autorizar otro paso.</summary>
        internal async Task<AgCandidatoPlanAprobado> BuscarArchivadoAsync(AgContexto c, int idAula, AgPropuestaEscaladoAprobada p,
            string huellaSql, CancellationToken ct)
        {
            if (_session == null || _vault == null || string.IsNullOrEmpty(huellaSql)) return null;
            try
            {
                return await _session.ExecuteAsync(db =>
                {
                    ct.ThrowIfCancellationRequested();
                    using (var tx = db.BeginTransaction())
                    {
                        var aula = c.Aulas.SingleOrDefault(a => a.Id == idAula);
                        if (!MismoTemporal(db, tx) || !OrigenContextual(c, p, aula)
                            || !PermisoVigente(db, tx, App.DEFAULT_USER)
                            || !AgHuellaEscaladoPersistido.CoincideOrigen(db, tx, c, aula)
                            || AgHuellaEscaladoPersistido.Calcular(db, tx, aula.AulaRef, aula.Subunidad, 0) != huellaSql)
                            return Task.FromResult<AgCandidatoPlanAprobado>(null);
                        var operaciones = db.Query(@"SELECT c.*, a.actor AS actor_auditado, a.huella_origen AS origen_auditado
 FROM ag_candidato_plan_aprobado c JOIN ag_auditoria_ratificacion a ON a.operacion=c.operacion
 WHERE c.actor=@Actor AND c.id_aula=@Id AND c.huella_manifiesto=@Manifest
 AND c.huella_origen=@Origen AND c.huella_previa=@Previa AND c.desde=@Desde AND c.hacia=@Hacia
 ORDER BY c.creado_utc DESC LIMIT 2", new { Actor = App.DEFAULT_USER, Id = idAula,
                            Manifest = c.ManifiestoCobertura.Huella, Origen = c.PlanAprobadoSnapshot.HuellaOrigen,
                            Previa = huellaSql, Desde = aula.DiasVentana, Hacia = p?.Hacia ?? aula.DiasVentana }, tx).ToList();
                        if (operaciones.Count != 1) return Task.FromResult<AgCandidatoPlanAprobado>(null);
                        dynamic fila = operaciones[0];
                        var ids = ((string)fila.consentimiento).Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(int.Parse).ToArray();
                        var restaurado = new AgCandidatoPlanAprobado(Guid.Parse((string)fila.operacion),
                            Guid.Parse((string)fila.intento), idAula, App.DEFAULT_USER, (string)fila.motivo,
                            c.ManifiestoCobertura, c.PlanAprobadoSnapshot.HuellaOrigen, huellaSql,
                            aula.DiasVentana, p?.Hacia ?? aula.DiasVentana, ids);
                        return Task.FromResult(Coincide(fila, restaurado)
                            && (string)fila.actor_auditado == restaurado.Actor
                            && (string)fila.origen_auditado == restaurado.HuellaOrigen
                            && AgPersistenciaRatificacion.PublicadoContiene(_vault, restaurado) ? restaurado : null);
                    }
                });
            }
            catch { return null; }
        }

        /// <summary>Valida permiso y origen dentro de la misma transacción que inserta candidato+auditoría.
        /// Un rechazo, cancelación u origen obsoleto no realiza escrituras.</summary>
        internal async Task<Resultado> ConfirmarAsync(AgContexto c, AgPropuestaEscaladoAprobada p,
            int idAula, Guid intento, Guid operacion, string huellaSqlMostrada, string motivo, bool ratifica,
            bool autorizaPaso, IReadOnlyCollection<int> legado, CancellationToken ct)
        {
            if (_session == null || _vault == null) return new Resultado("PUBLICACION_NO_DISPONIBLE");
            if (ct.IsCancellationRequested) return new Resultado("RATIFICACION_CANCELADA");
            if (!ratifica || string.IsNullOrWhiteSpace(motivo) || (p != null && !autorizaPaso))
                return new Resultado("REVISION_RECHAZADA");
            if (intento == Guid.Empty || operacion == Guid.Empty || string.IsNullOrEmpty(huellaSqlMostrada)
                || c?.ManifiestoCobertura == null || c.PlanAprobadoSnapshot == null || c.ValidacionPlanAprobado == null)
                return new Resultado("RATIFICACION_INVALIDA");
            // Ratificación sola para un plan completo: el paso se representa por Desde=Hacia.
            var aula = c.Aulas.SingleOrDefault(a => a.Id == idAula);
            if (aula == null || !OrigenContextual(c, p, aula)) return new Resultado("ALCANCE_OBSOLETO");
            var ids = (legado ?? new int[0]).ToList();
            if (ids.Count != ids.Distinct().Count() || ids.Any(id => id <= 0)) return new Resultado("CONSENTIMIENTO_INVALIDO");
            var permitidos = p == null ? new int[0] : p.Relaciones.Where(r => r.RedistribuyeLegado && r.OrigenPatron == "LEGADO")
                .Select(r => r.Clave.IdAsignatura).ToArray();
            if (ids.Except(permitidos).Any()) return new Resultado("CONSENTIMIENTO_FIJO_O_EXTRANO");
            string actor = App.DEFAULT_USER;
            if (string.IsNullOrWhiteSpace(actor)) return new Resultado("PERMISO_DENEGADO");
            var candidato = new AgCandidatoPlanAprobado(operacion, intento, aula.Id, actor, motivo,
                c.ManifiestoCobertura, c.PlanAprobadoSnapshot.HuellaOrigen, huellaSqlMostrada,
                aula.DiasVentana, p?.Hacia ?? aula.DiasVentana, ids);
            string codigo;
            try
            {
                codigo = await _session.ExecuteAsync(async db =>
                {
                    ct.ThrowIfCancellationRequested();
                    // SQLite BeginTransaction acquires a writer reservation for this connection.
                    using (var tx = db.BeginTransaction())
                    {
                        if (!MismoTemporal(db, tx)) return "PUBLICACION_NO_DISPONIBLE";
                        if (!PermisoVigente(db, tx, actor)) return "PERMISO_DENEGADO";
                        if (!OrigenContextual(c, p, aula)) return "ALCANCE_OBSOLETO";
                        if (!AgHuellaEscaladoPersistido.CoincideOrigen(db, tx, c, aula)) return "CONFIGURACION_OBSOLETA";
                        string actual = AgHuellaEscaladoPersistido.Calcular(db, tx, aula.AulaRef, aula.Subunidad, 0);
                        if (actual != huellaSqlMostrada) return "CONFIGURACION_OBSOLETA";
                        var origen = await AgPlanAprobadoRepository.ObtenerEnTransaccionAsync(db, tx, c.ManifiestoCobertura, ct);
                        if (origen.IsFailure) return origen.Error;
                        if (origen.Value.EstadoOrigen != "REQUIERE_RATIFICACION" || origen.Value.HuellaOrigen != candidato.HuellaOrigen)
                            return "ORIGEN_OBSOLETO_O_NO_VERIFICABLE";
                        // La lectura canónica comprueba extras y ausencias; ninguna vinculación nueva queda silenciada.
                        var previas = db.Query("SELECT * FROM ag_candidato_plan_aprobado WHERE operacion=@Op OR (intento=@It AND id_aula=@A)",
                            new { Op = operacion.ToString("N"), It = intento.ToString("N"), A = aula.Id }, tx).ToList();
                        if (previas.Count != 0)
                        {
                            if (previas.Count != 1 || !Coincide(previas[0], candidato)
                                || db.QuerySingle<int>(@"SELECT COUNT(*) FROM ag_auditoria_ratificacion
WHERE operacion=@Op AND actor=@Actor AND huella_origen=@Origen AND motivo=@Motivo AND alcance=@Alcance",
                                    new { Op = operacion.ToString("N"), candidato.Actor, Origen = candidato.HuellaOrigen,
                                        candidato.Motivo, candidato.Alcance }, tx) != 1)
                                return "REPLAY_NO_COINCIDE";
                            return "REPLAY_VERIFICADO";
                        }
                        if (!_repo.Guardar(db, tx, candidato)) return "AUDITORIA_INCOMPLETA";
                        ct.ThrowIfCancellationRequested();
                        tx.Commit();
                        return "COMMIT_EN_TEMPORAL";
                    }
                });
            }
            catch (OperationCanceledException) { return new Resultado("RATIFICACION_CANCELADA"); }
            catch (Exception ex) { return new Resultado("RATIFICACION_TECNICA:" + ex.GetType().Name); }
            if (codigo != "COMMIT_EN_TEMPORAL" && codigo != "REPLAY_VERIFICADO") return new Resultado(codigo);
            if (codigo == "REPLAY_VERIFICADO" && AgPersistenciaRatificacion.PublicadoContiene(_vault, candidato))
                return new Resultado("REPLAY_PUBLICADO", candidato,
                    new AgPersistenciaRatificacion.Resultado(AgPersistenciaRatificacion.Estado.PublicadoEnVault));
            var publicado = await new AgPersistenciaRatificacion().PublicarAsync(_vault, candidato, ct);
            if (publicado.EstadoActual != AgPersistenciaRatificacion.Estado.PublicadoEnVault)
                return new Resultado("FLUSH_FALLIDO:" + publicado.Motivo, candidato, publicado);
            return new Resultado(codigo == "REPLAY_VERIFICADO" ? "REPLAY_PUBLICADO" : "CANDIDATO_PUBLICADO", candidato, publicado);
        }

        private bool MismoTemporal(IDbConnection db, IDbTransaction tx)
        {
            var rutas = db.Query<string>("SELECT file FROM pragma_database_list WHERE name='main'", transaction: tx).ToList();
            return rutas.Count == 1 && string.Equals(System.IO.Path.GetFullPath(rutas[0]),
                System.IO.Path.GetFullPath(_vault.TempDbPath), StringComparison.OrdinalIgnoreCase);
        }

        private static bool PermisoVigente(IDbConnection db, IDbTransaction tx, string actor)
        {
            var filas = db.Query<string>(@"SELECT p.opcionespermitidas FROM usuarios u
JOIN permisos p ON p.perfil=u.perfil WHERE u.nombreusuario=@Actor COLLATE NOCASE AND u.activo=1",
                new { Actor = actor }, tx).ToList();
            return filas.Count == 1 && (filas[0] ?? "").Split(',')
                .Any(t => t.Trim() == "AgHorariosAcademicos");
        }

        private static bool OrigenContextual(AgContexto c, AgPropuestaEscaladoAprobada p, AgAula aula = null)
        {
            if (c?.ManifiestoCobertura == null || c.PlanAprobadoSnapshot == null || c.Ventana == null) return false;
            aula = aula ?? (p == null ? null : c.Aulas.SingleOrDefault(a => a.Id == p.IdAula));
            if (p != null && aula != null && aula.Id != p.IdAula) return false;
            if (aula == null || aula.Estado != AgEstadoAula.Aprobado || c.PlanAprobadoSnapshot.EstadoOrigen != "REQUIERE_RATIFICACION"
                || c.PlanAprobadoSnapshot.Manifiesto.Huella != c.ManifiestoCobertura.Huella
                || c.ManifiestoCobertura.Fin > new DateTime(2027, 5, 30)
                || c.ManifiestoCobertura.Inicio != c.Ventana.FechaInicio.Date || c.ManifiestoCobertura.Fin != c.Ventana.FechaFin.Date
                || !c.ManifiestoCobertura.Relaciones.Any(r => r.Clave.IdAula == aula.Id)
                || c.Aulas.Count(a => a.Id == aula.Id) != 1) return false;
            if (p != null && (p.IdAula != aula.Id || p.HuellaOrigen != c.PlanAprobadoSnapshot.HuellaOrigen
                || p.HuellaConfiguracion != AgHuellaPlanVentana.Calcular(c, aula)
                || p.Desde != aula.DiasVentana || p.FechaLimite.Date != c.ManifiestoCobertura.Fin
                || string.IsNullOrEmpty(p.Hacia) || p.Hacia != AgEscaleraVentanas.Siguiente(p.Desde))) return false;
            return true;
        }

        private static bool Coincide(dynamic fila, AgCandidatoPlanAprobado c)
        {
            return (string)fila.operacion == c.Operacion.ToString("N") && (string)fila.intento == c.Intento.ToString("N")
                && Convert.ToInt32(fila.id_aula) == c.IdAula && (string)fila.actor == c.Actor
                && (string)fila.motivo == c.Motivo && (string)fila.alcance == c.Alcance
                && (string)fila.huella_manifiesto == c.HuellaManifiesto && (string)fila.huella_origen == c.HuellaOrigen
                && (string)fila.huella_previa == c.HuellaPrevia && (string)fila.huella_efecto == c.HuellaEfecto
                && (string)fila.desde == c.Desde && (string)fila.hacia == c.Hacia && (string)fila.consentimiento == c.Consentimiento;
        }
    }
}
