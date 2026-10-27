using System;
using System.Data;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   CRUD sobre las tablas de control interno del sistema de sync:
    ///   - sync_control: rastrea el último cambio PUSH exitoso (seq de crsql_changes).
    ///   - sync_archive: registra archivos .crsync ya procesados (previene re-procesamiento).
    ///
    /// ESTAS TABLAS NO SON CRDT:
    ///   Son de infraestructura LOCAL. No se sincronizan entre sedes.
    ///   Cada DB (por subunidad) tiene su propia copia independiente.
    ///
    /// PERSISTENCIA:
    ///   Estas tablas se almacenan en la misma DB cifrada (vault). Se persisten al vault
    ///   junto con los datos de usuario cuando el contexto hace ReEncryptAndPersist.
    ///   NO se hace re-cifrado extra solo por actualizar sync_control.
    ///
    /// IDEMPOTENCIA:
    ///   EnsureTablesExist() usa CREATE TABLE IF NOT EXISTS — puede llamarse múltiples
    ///   veces sin efecto secundario. Se llama al inicio de cada operación de sync.
    /// </summary>
    internal sealed class SyncControlRepository
    {
        // ====================================================================
        // DDL — Definición de tablas
        // ====================================================================

        /// <summary>
        /// DDL de sync_control.
        /// PK compuesta (subunidad, device_id) porque el mismo vault podría ser usado
        /// por múltiples PCs de la misma sede (ej: PC principal + laptop de backup).
        /// </summary>
        private const string CreateSyncControlSql = @"
            CREATE TABLE IF NOT EXISTS sync_control (
                -- Código de subunidad a la que pertenece este registro
                subunidad       TEXT NOT NULL,

                -- Nombre del PC (Environment.MachineName), identifica el dispositivo
                device_id       TEXT NOT NULL,

                -- Último seq de crsql_changes enviado exitosamente a OneDrive.
                -- Al hacer PUSH, se consultan cambios usando la tupla (db_version, seq)
                last_pushed_db_version INTEGER DEFAULT 0,
                last_pushed_seq        INTEGER DEFAULT 0,

                -- Última vez que se hizo un PULL exitoso (para logging/diagnóstico).
                last_pulled_at  TEXT,

                PRIMARY KEY (subunidad, device_id)
            );";

        /// <summary>
        /// DDL de sync_archive.
        /// Registra nombres de archivos .crsync ya procesados para evitar
        /// aplicar el mismo changeset dos veces en caso de reinicio del add-in.
        /// </summary>
        private const string CreateSyncArchiveSql = @"
            CREATE TABLE IF NOT EXISTS sync_archive (
                -- Nombre del archivo .crsync procesado (clave única)
                filename        TEXT PRIMARY KEY,

                -- Subunidad que generó el archivo (para diagnóstico)
                subunidad       TEXT NOT NULL,

                -- SHA256 del contenido del archivo (para verificar integridad)
                checksum        TEXT,

                -- Timestamp UTC de cuando fue procesado por este nodo
                processed_at    TEXT DEFAULT (datetime('now'))
            );";

        // ====================================================================
        // SQL OPERACIONES
        // ====================================================================

        private const string GetLastPushedClockSql = @"
            SELECT COALESCE(last_pushed_db_version, 0) AS DbVersion,
                   COALESCE(last_pushed_seq, 0) AS Seq
            FROM   sync_control
            WHERE  subunidad = @Subunidad
              AND  device_id = @DeviceId
            LIMIT 1;";

        private const string UpsertLastPushedClockSql = @"
            INSERT INTO sync_control (subunidad, device_id, last_pushed_db_version, last_pushed_seq)
            VALUES (@Subunidad, @DeviceId, @DbVersion, @Seq)
            ON CONFLICT(subunidad, device_id)
            DO UPDATE SET 
                last_pushed_db_version = @DbVersion,
                last_pushed_seq = @Seq;";

        private const string UpdateLastPulledAtSql = @"
            INSERT INTO sync_control (subunidad, device_id, last_pulled_at)
            VALUES (@Subunidad, @DeviceId, @PulledAt)
            ON CONFLICT(subunidad, device_id)
            DO UPDATE SET last_pulled_at = @PulledAt;";

        private const string IsFileProcessedSql = @"
            SELECT COUNT(*)
            FROM   sync_archive
            WHERE  filename = @Filename;";

        private const string InsertArchiveSql = @"
            INSERT OR IGNORE INTO sync_archive (filename, subunidad, checksum)
            VALUES (@Filename, @Subunidad, @Checksum);";

        // ====================================================================
        // API PÚBLICA
        // ====================================================================

        /// <summary>
        /// Crea las tablas sync_control y sync_archive si no existen.
        ///
        /// CUÁNDO LLAMAR: Al inicio de cada operación de PUSH o PULL.
        /// IDEMPOTENTE: Puede llamarse múltiples veces sin efecto secundario.
        /// </summary>
        internal void EnsureTablesExist(IDbConnection connection, IDbTransaction transaction = null)
        {
            connection.Execute(CreateSyncControlSql, transaction: transaction);
            connection.Execute(CreateSyncArchiveSql, transaction: transaction);

            // Migración defensiva: Si la base de datos ya existía con la versión antigua de sync_control
            // intentamos agregar la nueva columna last_pushed_db_version sin destruir los datos.
            try
            {
                connection.Execute("ALTER TABLE sync_control ADD COLUMN last_pushed_db_version INTEGER DEFAULT 0;", transaction: transaction);
                LoggerHelper.Info("[SyncControlRepository] Migración de esquema: Añadida columna last_pushed_db_version a sync_control.");
            }
            catch
            {
                // Fallará silenciosamente si la columna ya existe, lo cual es el flujo normal
            }

            // Migración defensiva: asegurar que filas antiguas tengan 0 en lugar de NULL
            connection.Execute(@"
                UPDATE sync_control
                   SET last_pushed_db_version = COALESCE(last_pushed_db_version, 0),
                       last_pushed_seq = COALESCE(last_pushed_seq, 0)
                 WHERE last_pushed_db_version IS NULL OR last_pushed_seq IS NULL;",
                transaction: transaction);

            LoggerHelper.Debug("[SyncControlRepository] Tablas de control verificadas/creadas.");
        }

        /// <summary>
        /// Obtiene el reloj lógico (db_version, seq) del último envío a OneDrive.
        ///
        /// Este reloj se usa como punto de partida para el próximo PUSH:
        ///   WHERE (db_version > @LastDbVersion) OR (db_version = @LastDbVersion AND seq > @LastSeq)
        ///
        /// RETORNA SyncClock(0, 0) si no hay registro previo (primera vez que hace PUSH).
        /// </summary>
        /// <param name="connection">Conexión SQLite abierta.</param>
        /// <param name="subunidad">Código de subunidad (ej: "VALLE").</param>
        /// <param name="deviceId">Nombre del PC (Environment.MachineName).</param>
        internal SyncClock GetLastPushedClock(IDbConnection connection, string subunidad, string deviceId)
        {
            var result = connection.QueryFirstOrDefault(
                GetLastPushedClockSql,
                new { Subunidad = subunidad, DeviceId = deviceId });

            if (result == null)
            {
                LoggerHelper.Debug($"[SyncControlRepository] last_pushed_clock para {subunidad}/{deviceId}: v0.0 (Nuevo)");
                return new SyncClock(0, 0);
            }

            var clock = new SyncClock((long)result.DbVersion, (long)result.Seq);
            LoggerHelper.Debug($"[SyncControlRepository] last_pushed_clock para {subunidad}/{deviceId}: {clock}");
            return clock;
        }

        /// <summary>
        /// Actualiza el reloj lógico enviado exitosamente a OneDrive.
        ///
        /// USA UPSERT (INSERT OR UPDATE): si no existe el registro, lo crea.
        /// Se llama al final de cada PUSH exitoso.
        /// </summary>
        /// <param name="connection">Conexión SQLite abierta.</param>
        /// <param name="subunidad">Código de subunidad.</param>
        /// <param name="deviceId">Nombre del PC.</param>
        /// <param name="clock">Nuevo reloj lógico a guardar.</param>
        /// <param name="transaction">Transacción activa, si aplica.</param>
        internal void UpdateLastPushedClock(
            IDbConnection connection,
            string subunidad,
            string deviceId,
            SyncClock clock,
            IDbTransaction transaction = null)
        {
            connection.Execute(
                UpsertLastPushedClockSql,
                new { Subunidad = subunidad, DeviceId = deviceId, DbVersion = clock.DbVersion, Seq = clock.Seq },
                transaction: transaction);

            LoggerHelper.Debug($"[SyncControlRepository] last_pushed_clock actualizado a {clock} para {subunidad}/{deviceId}");
        }

        /// <summary>
        /// Registra el timestamp del último PULL exitoso para diagnóstico.
        /// Ayuda a responder "¿cuándo fue la última vez que recibimos cambios?"
        /// </summary>
        internal void UpdateLastPulledAt(
            IDbConnection connection,
            string subunidad,
            string deviceId,
            IDbTransaction transaction = null)
        {
            string utcNow = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
            connection.Execute(
                UpdateLastPulledAtSql,
                new { Subunidad = subunidad, DeviceId = deviceId, PulledAt = utcNow },
                transaction: transaction);
        }

        /// <summary>
        /// Verifica si un archivo .crsync ya fue procesado en esta sesión o sesiones anteriores.
        ///
        /// USO: Antes de procesar un archivo PULL para evitar aplicar el mismo changeset dos veces.
        /// Un archivo puede aparecer "dos veces" si el add-in se reinicia antes de que se mueva
        /// a la carpeta .processed.
        /// </summary>
        /// <param name="connection">Conexión SQLite abierta.</param>
        /// <param name="filename">Nombre del archivo .crsync (sin ruta).</param>
        internal bool IsFileAlreadyProcessed(IDbConnection connection, string filename)
        {
            int count = connection.QueryFirstOrDefault<int>(
                IsFileProcessedSql,
                new { Filename = filename });

            return count > 0;
        }

        /// <summary>
        /// Registra un archivo .crsync como procesado exitosamente.
        ///
        /// USA INSERT OR IGNORE para ser idempotente (si se llama dos veces con el mismo
        /// filename, no falla ni duplica).
        /// </summary>
        /// <param name="connection">Conexión SQLite abierta.</param>
        /// <param name="filename">Nombre del archivo .crsync procesado.</param>
        /// <param name="subunidad">Subunidad origen del archivo.</param>
        /// <param name="checksum">SHA256 del contenido (para auditoría).</param>
        /// <param name="transaction">Transacción activa, si aplica.</param>
        internal void MarkFileProcessed(
            IDbConnection connection,
            string filename,
            string subunidad,
            string checksum,
            IDbTransaction transaction = null)
        {
            connection.Execute(
                InsertArchiveSql,
                new { Filename = filename, Subunidad = subunidad, Checksum = checksum },
                transaction: transaction);

            LoggerHelper.Debug($"[SyncControlRepository] Archivo registrado en sync_archive: {filename}");
        }
    }
}
