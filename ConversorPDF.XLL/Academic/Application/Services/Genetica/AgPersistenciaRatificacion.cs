using System;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Publica un backup online con comprobación de identidad del vault y lectura cifrada posterior; sin exclusión interproceso.</summary>
    internal sealed class AgPersistenciaRatificacion
    {
        internal enum Estado { CommitEnTemporal, PublicadoEnVault, FlushFallido }
        internal sealed class Resultado
        {
            internal Estado EstadoActual { get; }
            internal string Motivo { get; }
            internal Resultado(Estado estado, string motivo = null) { EstadoActual = estado; Motivo = motivo; }
        }

        /// <summary>Invocar solamente después del commit. Nunca toma MarkChangesPersisted como prueba de publicación.</summary>
        internal async Task<Resultado> PublicarAsync(EncryptedSQLiteContext contexto, AgCandidatoPlanAprobado c, CancellationToken ct)
        {
            if (contexto == null || c == null) return new Resultado(Estado.FlushFallido, "CONTEXTO_O_CANDIDATO_AUSENTE");
            string ruta = contexto.TempDbPath + ".ratif";
            try
            {
                ct.ThrowIfCancellationRequested();
                contexto.Open();
                var identidad = EncryptedSQLiteContext.CaptureVaultIdentity(contexto.EncryptedFilePath);
                contexto.BackupTo(ruta);
                byte[] bytes = await EncryptedSQLiteContext.LeerBytesConRetryStatic(ruta, ct: ct).ConfigureAwait(false);
                if (bytes == null || bytes.Length == 0) return new Resultado(Estado.FlushFallido, "BACKUP_VACIO");
                if (!ContieneOperacion(ruta, c))
                    return new Resultado(Estado.FlushFallido, "BACKUP_OPERACION_NO_COINCIDE");
                ct.ThrowIfCancellationRequested();
                if (!SqliteHelper.UpdateVaultIfUnchanged(bytes, contexto.EncryptedFilePath, identidad, ct))
                    return new Resultado(Estado.FlushFallido, "VAULT_POST_PUBLICACION_NO_COINCIDE");
                // También probar la fila desde una lectura independiente del vault recién publicado.
                return PublicadoContiene(contexto, c)
                    ? new Resultado(Estado.PublicadoEnVault)
                    : new Resultado(Estado.FlushFallido, "VAULT_OPERACION_NO_COINCIDE");
            }
            catch (OperationCanceledException) { return new Resultado(Estado.FlushFallido, "PUBLICACION_CANCELADA"); }
            catch (Exception ex) { return new Resultado(Estado.FlushFallido, "PUBLICACION_ERROR:" + ex.GetType().Name); }
            finally
            {
                try { if (File.Exists(ruta)) File.Delete(ruta); } catch (IOException) { /* informar falla de limpieza sin tapar error principal */ }
            }
        }

        /// <summary>Comprueba la operación desde bytes descifrados del .enc, no del temporal compartido.</summary>
        internal static bool PublicadoContiene(EncryptedSQLiteContext contexto, AgCandidatoPlanAprobado c)
        {
            string path = Path.Combine(Path.GetTempPath(), "ag_publicado_" + Guid.NewGuid().ToString("N") + ".db");
            try
            {
                var antes = EncryptedSQLiteContext.CaptureVaultIdentity(contexto.EncryptedFilePath);
                if (antes == null) return false;
                File.WriteAllBytes(path, contexto.ReadPublishedBytes());
                bool coincide = ContieneOperacion(path, c);
                return coincide && antes.SameAs(EncryptedSQLiteContext.CaptureVaultIdentity(contexto.EncryptedFilePath));
            }
            catch { return false; }
            finally { try { if (File.Exists(path)) File.Delete(path); } catch (IOException) { } }
        }

        private static bool ContieneOperacion(string path, AgCandidatoPlanAprobado c)
        {
            using (var db = new SQLiteConnection("Data Source=" + path + ";Pooling=False;Version=3;Read Only=True;"))
            {
                db.Open();
                if (db.ExecuteScalar<string>("PRAGMA integrity_check;") != "ok") return false;
                var fila = db.QuerySingleOrDefault<dynamic>(@"SELECT c.*, a.actor AS actor_auditado,
 a.huella_origen AS origen_auditado, a.motivo AS motivo_auditado, a.alcance AS alcance_auditado
 FROM ag_candidato_plan_aprobado c JOIN ag_auditoria_ratificacion a ON a.operacion=c.operacion
 WHERE c.operacion=@Operacion", new { Operacion = c.Operacion.ToString("N") });
                return fila != null && (string)fila.intento == c.Intento.ToString("N")
                    && Convert.ToInt32(fila.id_aula) == c.IdAula && (string)fila.actor == c.Actor
                    && (string)fila.motivo == c.Motivo && (string)fila.alcance == c.Alcance
                    && (string)fila.huella_manifiesto == c.HuellaManifiesto && (string)fila.huella_origen == c.HuellaOrigen
                    && (string)fila.huella_previa == c.HuellaPrevia && (string)fila.huella_efecto == c.HuellaEfecto
                    && (string)fila.desde == c.Desde && (string)fila.hacia == c.Hacia
                    && (string)fila.consentimiento == c.Consentimiento && (string)fila.actor_auditado == c.Actor
                    && (string)fila.origen_auditado == c.HuellaOrigen && (string)fila.motivo_auditado == c.Motivo
                    && (string)fila.alcance_auditado == c.Alcance;
            }
        }
    }
}
