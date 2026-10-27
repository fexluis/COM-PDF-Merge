using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Dapper;
using System.Web.Script.Serialization;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   Orquesta las operaciones de PUSH (local → OneDrive) y PULL (OneDrive → local)
    ///   para una subunidad específica.
    ///
    ///   PUSH: Lee los cambios CRDT pendientes de la DB local y los escribe como
    ///         archivo .crsync en la carpeta de OneDrive de la subunidad.
    ///
    ///   PULL: Lee un archivo .crsync entrante y aplica sus cambios a la DB local
    ///         insertando directamente en la tabla virtual crsql_changes de CR-SQLite.
    ///
    /// THREAD SAFETY:
    ///   NUNCA se ejecuta en el hilo principal de Excel (STA). Siempre es invocado
    ///   desde SyncScheduler o SyncFileWatcher vía Task.Run(). Crea su propio
    ///   EncryptedSQLiteContext independiente del que usa la UI.
    ///
    /// FILTRADO POR SUBUNIDAD:
    ///   Las tablas con columna `subunidad` se filtran al hacer PUSH:
    ///   solo se envían los registros que pertenecen a esta subunidad.
    ///   La tabla `divipola` (catálogo nacional) se replica completa sin filtro.
    ///
    /// VAULT:
    ///   - PUSH: MarkAsReadOnly (no persiste al vault, solo lee).
    ///   - PULL: Persiste al vault (se modificaron datos reales de usuario).
    /// </summary>
    internal sealed class SubunidadSyncService : IDisposable
    {
        // ====================================================================
        // TABLAS SINCRONIZABLES
        // ====================================================================

        /// <summary>
        /// Tablas que tienen columna `subunidad` y se filtran al hacer PUSH.
        /// Orden no importa para PUSH, pero sí para claridad del código.
        /// </summary>
        private static readonly string[] LocalTables = new[]
        {
            "config",
            "festivos",
            "sedes",
            "vinculaciones",
            "sesiones",
            "salvedades_conflictos",
            "cdpDependencias"
        };

        /// <summary>
        /// Tablas que se replican COMPLETAS (sin filtro de subunidad).
        /// </summary>
        private static readonly string[] GlobalTables = new[]
        {
            "programas",
            "asignaturas",
            "microcurriculo",
            "contenidos",
            "periodos",
            "terceros",
            "usuarios",
            "territoriales",
            "terminales",
            "divipola",
            "permisos"
        };

        // ====================================================================
        // DEPENDENCIAS Y ESTADO GLOBAL
        // ====================================================================

        private readonly string _subunidad;
        private readonly string _deviceId;
        private readonly string _vaultPath;
        private readonly SyncPathResolver _pathResolver;
        private readonly SyncControlRepository _controlRepo;
        private readonly CrSyncFileWriter _fileWriter;


        private static string GetOrCreateInstanceIdForVault(string vaultPath)
        {
            try
            {
                string folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ConversorPDF", "Data", "Devices");
                Directory.CreateDirectory(folder);

                string vaultHash;
                using (var sha = System.Security.Cryptography.SHA256.Create())
                {
                    string normalizedVaultPath = Path.GetFullPath(vaultPath).ToUpperInvariant();
                    var bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(normalizedVaultPath));
                    vaultHash = BitConverter.ToString(bytes, 0, 8).Replace("-", "").ToUpperInvariant();
                }

                string idFile = Path.Combine(folder, $"device_{vaultHash}.id");
                if (File.Exists(idFile))
                {
                    string stored = File.ReadAllText(idFile).Trim();
                    if (!string.IsNullOrEmpty(stored)) return stored;
                }

                string newId = Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant();
                File.WriteAllText(idFile, newId);
                return newId;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SubunidadSyncService] No se pudo leer/crear device.id para vault: {ex.Message}. Usando fallback temporal.");
                return Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant();
            }
        }

        // ====================================================================
        // ESTADO
        // ====================================================================

        private bool _disposed = false;

        // ====================================================================
        // CONSTRUCTOR
        // ====================================================================

        /// <summary>
        /// Inicializa el servicio para una subunidad específica.
        /// </summary>
        /// <param name="subunidad">Código de subunidad (ej: "VALLE").</param>
        /// <param name="vaultPath">Ruta al archivo .db.enc (vault cifrado).</param>
        /// <param name="pathResolver">Resolvedor de rutas OneDrive.</param>
        internal SubunidadSyncService(
            string subunidad,
            string vaultPath,
            SyncPathResolver pathResolver)
        {
            _subunidad = subunidad ?? throw new ArgumentNullException(nameof(subunidad));
            _vaultPath = vaultPath ?? throw new ArgumentNullException(nameof(vaultPath));
            _pathResolver = pathResolver ?? throw new ArgumentNullException(nameof(pathResolver));

            _deviceId = GetCompositeDeviceId(_vaultPath);
            _controlRepo = new SyncControlRepository();
            _fileWriter = new CrSyncFileWriter();

            LoggerHelper.Info($"[SubunidadSyncService] Inicializado para subunidad '{_subunidad}', device '{_deviceId}'");
        }

        internal static string GetCompositeDeviceId(string vaultPath)
        {
            string machine = SanitizeDeviceId(Environment.MachineName);
            string user = SanitizeDeviceId(Environment.UserName);
            return $"{machine}_{user}_{GetOrCreateInstanceIdForVault(vaultPath)}";
        }

        // ====================================================================
        // PUSH
        // ====================================================================

        private static bool EsEntornoDePruebas() =>
            SqliteHelper.IsTestVaultPathActive ||
            AppDomain.CurrentDomain.FriendlyName.IndexOf("test", StringComparison.OrdinalIgnoreCase) >= 0;

        /// <summary>
        /// PUSH: Lee cambios CRDT locales pendientes y los escribe a OneDrive.
        ///
        /// FLUJO:
        ///   1. Abrir EncryptedSQLiteContext (solo lectura).
        ///   2. Habilitar extensión CR-SQLite en la conexión.
        ///   3. Verificar tablas de control.
        ///   4. Leer last_pushed_seq del control.
        ///   5. Consultar crsql_changes por tabla (con y sin filtro subunidad).
        ///   6. Si hay cambios: serializar → escribir archivo .crsync → actualizar seq.
        ///   7. Cerrar contexto (sin re-cifrar, solo lectura).
        ///
        /// RETORNA: PushResult con información del resultado (cambios enviados, bytes, etc.)
        /// </summary>
        internal async Task<PushResult> PushChangesAsync(CancellationToken ct = default)
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(SubunidadSyncService));
            }

            // FASE 820 / FASE 822: Aislamiento estricto de entorno de pruebas para evitar contaminación de OneDrive
            if (EsEntornoDePruebas())
            {
                LoggerHelper.Info("[SubunidadSyncService] PUSH omitido: entorno de pruebas.");
                return PushResult.Skipped("Entorno de pruebas: PUSH deshabilitado.");
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();
            LoggerHelper.Info($"[SubunidadSyncService] Iniciando PUSH para subunidad '{_subunidad}'...");

            try
            {
                // Abrir el vault en modo solo lectura para no bloquear la UI más de lo necesario
                using (var context = SqliteHelper.GetContextForPath(_vaultPath))
                {
                    context.MarkAsReadOnly();
                    IDbConnection connection = context.Open();

                    // Habilitar CR-SQLite en esta conexión específica
                    var sqliteConn = connection as SQLiteConnection;
                    if (sqliteConn == null)
                    {
                        LoggerHelper.Warning("[SubunidadSyncService] La conexión no es SQLiteConnection. Saltando PUSH.");
                        return PushResult.Skipped("Conexión no compatible con CR-SQLite.");
                    }

                    if (!CrSqliteLoader.EnableOnConnection(sqliteConn))
                    {
                        LoggerHelper.Warning("[SubunidadSyncService] CR-SQLite no disponible. Saltando PUSH.");
                        return PushResult.Skipped("CR-SQLite no disponible en esta conexión.");
                    }

                    // Asegurar tablas de control existentes
                    _controlRepo.EnsureTablesExist(connection);

                    // Obtener el reloj lógico de partida del PUSH
                    SyncClock lastClock = _controlRepo.GetLastPushedClock(connection, _subunidad, _deviceId);
                    LoggerHelper.Debug($"[SubunidadSyncService] PUSH desde {lastClock}");

                    ct.ThrowIfCancellationRequested();

                    // Recolectar todos los cambios pendientes
                    var localChanges = new List<CrSyncChange>();
                    var globalChanges = new List<CrSyncChange>();
                    SyncClock maxClockFound = lastClock;

                    // Tablas con filtro de subunidad
                    foreach (string table in LocalTables)
                    {
                        ct.ThrowIfCancellationRequested();

                        if (!TableExists(connection, table))
                        {
                            LoggerHelper.Debug($"[SubunidadSyncService] Tabla '{table}' no existe, omitiendo.");
                            continue;
                        }

                        List<CrSyncChange> changes = ReadChangesWithSubunidadFilter(
                            connection, table, _subunidad, lastClock, ref maxClockFound);

                        localChanges.AddRange(changes);
                    }

                    // Tablas globales (sin filtro de subunidad)
                    foreach (string table in GlobalTables)
                    {
                        ct.ThrowIfCancellationRequested();

                        if (!TableExists(connection, table))
                        {
                            continue;
                        }

                        List<CrSyncChange> changes = ReadChangesGlobal(
                            connection, table, lastClock, ref maxClockFound);

                        globalChanges.AddRange(changes);
                    }

                    // Si no hay cambios, no generar archivo
                    if (localChanges.Count == 0 && globalChanges.Count == 0)
                    {
                        sw.Stop();
                        LoggerHelper.Debug($"[SubunidadSyncService] PUSH: No hay cambios pendientes desde {lastClock}. ({sw.ElapsedMilliseconds}ms)");
                        return PushResult.Success(0, 0, $"No hay cambios desde {lastClock}.");
                    }

                    int totalChanges = localChanges.Count + globalChanges.Count;
                    LoggerHelper.Info($"[SubunidadSyncService] PUSH: {totalChanges} cambios recolectados desde {lastClock} hasta {maxClockFound}.");

                    ct.ThrowIfCancellationRequested();

                    long totalBytes = 0;
                    
                    // Create base timestamp for both files to match
                    DateTime now = DateTime.UtcNow;

                    // Write Local changes
                    if (localChanges.Count > 0)
                    {
                        var localEnvelope = new CrSyncEnvelope
                        {
                            Version = "1.1",
                            Source = new CrSyncSource
                            {
                                DeviceId = _deviceId,
                                Subunidad = _subunidad,
                                TimestampUtc = now.ToString("yyyy-MM-ddTHH:mm:ssZ")
                            },
                            Changes = localChanges
                        };
                        string localFolder = _pathResolver.GetSubunidadInboxFolder(_subunidad);
                        string localFileName = SyncPathResolver.BuildFileName(_deviceId, _subunidad, now);
                        string localFilePath = Path.Combine(localFolder, localFileName);
                        _fileWriter.WriteToFile(localEnvelope, localFilePath);
                        totalBytes += new FileInfo(localFilePath).Length;
                    }

                    // Write Global changes
                    if (globalChanges.Count > 0)
                    {
                        var globalEnvelope = new CrSyncEnvelope
                        {
                            Version = "1.1",
                            Source = new CrSyncSource
                            {
                                DeviceId = _deviceId,
                                Subunidad = "_GT",
                                TimestampUtc = now.ToString("yyyy-MM-ddTHH:mm:ssZ")
                            },
                            Changes = globalChanges
                        };
                        string globalFolder = _pathResolver.GetGlobalInboxFolder();
                        string globalFileName = SyncPathResolver.BuildFileName(_deviceId, "_GT", now);
                        string globalFilePath = Path.Combine(globalFolder, globalFileName);
                        _fileWriter.WriteToFile(globalEnvelope, globalFilePath);
                        totalBytes += new FileInfo(globalFilePath).Length;
                    }

                    // Actualizar reloj lógico en control (si se completó bien)
                    _controlRepo.UpdateLastPushedClock(connection, _subunidad, _deviceId, maxClockFound);

                    sw.Stop();
                    
                    // FASE 23: Registrar el nombre real del archivo emitido (priorizando local, luego global)
                    string pushedFileName;
                    if (localChanges.Count > 0)
                        pushedFileName = SyncPathResolver.BuildFileName(_deviceId, _subunidad, now);
                    else
                        pushedFileName = SyncPathResolver.BuildFileName(_deviceId, "_GT", now);

                    UpdateSyncStateOnPush(pushedFileName);

                    LoggerHelper.Info($"[SubunidadSyncService] PUSH completado en {sw.ElapsedMilliseconds}ms. ({totalBytes} bytes en total)");

                    return PushResult.Success(totalChanges, totalBytes, $"Archivos de sync enviados a OneDrive exitosamente.");
                }
            }
            catch (Exception ex)
            {
                sw.Stop();
                LoggerHelper.Error($"[SubunidadSyncService] Error fatal durante PUSH en {sw.ElapsedMilliseconds}ms", ex);
                return PushResult.Skipped($"Error fatal durante PUSH: {ex.Message}");
            }
        }

        // ====================================================================
        // PULL
        // ====================================================================

        private static bool IsOneDriveTransientError(IOException ex)
        {
            int hr = ex.HResult;
            return hr == unchecked((int)0x8007016A)   // CLOUD_FILE_PROVIDER_NOT_RUNNING
                || hr == unchecked((int)0x80070179)   // CLOUD_FILE_TOO_MANY_THUMBNAILS
                || hr == unchecked((int)0x80070008)   // NOT_ENOUGH_MEMORY
                || hr == unchecked((int)0x80070194)   // CLOUD_FILE_NOT_IN_SYNC
                || ex.Message.Contains("proveedor de sincronización")
                || ex.Message.Contains("pocos recursos");
        }

        private static bool TryEnsureFileHydrated(string filePath)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                                               FileShare.ReadWrite, 1, FileOptions.SequentialScan))
                {
                    return true;
                }
            }
            catch (IOException ex) when (IsOneDriveTransientError(ex))
            {
                return false;
            }
        }

        private bool TableIsCrdtEnabled(IDbConnection connection, IDbTransaction transaction, string tableName)
        {
            try
            {
                var exists = connection.ExecuteScalar<int>(
                    "SELECT COUNT(1) FROM sqlite_master WHERE type='table' AND name=@TableName;",
                    new { TableName = tableName }, transaction: transaction);
                if (exists == 0) return false;

                // Verificar si existe la extensión CR-SQLite y si la tabla está registrada
                // La forma más segura de saber si una tabla es CRDT es buscar su tabla de reloj __crsql_clock
                string clockTableName = $"{tableName}__crsql_clock";
                var crdt = connection.ExecuteScalar<int>(
                    "SELECT COUNT(1) FROM sqlite_master WHERE type='table' AND name=@ClockTableName;",
                    new { ClockTableName = clockTableName }, transaction: transaction);
                return crdt > 0;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SubunidadSyncService] No se pudo verificar estado CRDT de {tableName}: {ex.Message}");
                return false;
            }
        }

        // ====================================================================

        /// <summary>
        /// PULL BATCH: Procesa una lista de archivos .crsync entrantes en una sola conexión SQLite 
        /// y realiza un único cifrado/persistencia al final. 
        /// Optimización crítica para arranque donde hay múltiples archivos pendientes.
        /// </summary>
        internal async Task ProcessIncomingFilesBatchAsync(string[] filePaths, CancellationToken ct = default)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(SubunidadSyncService));
            if (filePaths == null || filePaths.Length == 0) return;

            LoggerHelper.Info($"[SubunidadSyncService] PULL BATCH: Iniciando procesamiento de {filePaths.Length} archivos.");
            var sw = System.Diagnostics.Stopwatch.StartNew();
            
            var validEnvelopes = new List<(string FilePath, string FileName, CrSyncEnvelope Envelope)>();
            
            // FASE 1: Filtrar y leer todos los envelopes válidos
            foreach (var filePath in filePaths)
            {
                ct.ThrowIfCancellationRequested();
                string fileName = Path.GetFileName(filePath);
                
                if (IsOwnFile(fileName))
                {
                    LoggerHelper.Debug($"[SubunidadSyncService] PULL BATCH: Ignorando archivo propio: {fileName}");
                    continue;
                }

                CrSyncEnvelope envelope = null;
                try
                {
                    Exception lastEx = null;
                    const int maxAttempts = 3;
                    for (int attempt = 1; attempt <= maxAttempts; attempt++)
                    {
                        try
                        {
                            if (!TryEnsureFileHydrated(filePath))
                            {
                                await Task.Delay(TimeSpan.FromMilliseconds(500 * attempt), ct);
                                continue;
                            }
                            envelope = _fileWriter.ReadFromFile(filePath);
                            break;
                        }
                        catch (IOException ex) when (IsOneDriveTransientError(ex))
                        {
                            lastEx = ex;
                            await Task.Delay(TimeSpan.FromMilliseconds(500 * attempt), ct);
                        }
                    }

                    if (envelope == null)
                    {
                        LoggerHelper.Error($"[SubunidadSyncService] PULL BATCH: No se pudo leer {fileName} tras {maxAttempts} intentos. Se omitirá en este batch.", lastEx);
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error($"[SubunidadSyncService] PULL BATCH: Error leyendo {fileName}. Se moverá a .processed.", ex);
                    await MoveToProcessedAsync(filePath);
                    continue;
                }

                string envelopeSubunidad = envelope.Source?.Subunidad ?? string.Empty;
                bool isGlobal = string.Equals(envelopeSubunidad, "_GT", StringComparison.OrdinalIgnoreCase);

                if (!isGlobal && !string.Equals(envelopeSubunidad, _subunidad, StringComparison.OrdinalIgnoreCase))
                {
                    LoggerHelper.Warning($"[SubunidadSyncService] PULL BATCH: Archivo {fileName} pertenece a subunidad '{envelopeSubunidad}' y no coincide con la local '{_subunidad}'. Ignorando.");
                    await MoveToProcessedAsync(filePath);
                    continue;
                }

                if (!string.Equals(envelope.Version, "1.0", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(envelope.Version, "1.1", StringComparison.OrdinalIgnoreCase))
                {
                    LoggerHelper.Warning($"[SubunidadSyncService] PULL BATCH: Versión no soportada en {fileName}.");
                    await MoveToProcessedAsync(filePath);
                    continue;
                }

                if (envelope.Changes == null || envelope.Changes.Count == 0)
                {
                    LoggerHelper.Warning($"[SubunidadSyncService] PULL BATCH: Archivo {fileName} sin cambios.");
                    await MoveToProcessedAsync(filePath);
                    continue;
                }
                
                validEnvelopes.Add((filePath, fileName, envelope));
            }

            if (validEnvelopes.Count == 0)
            {
                sw.Stop();
                LoggerHelper.Info($"[SubunidadSyncService] PULL BATCH completado en {sw.ElapsedMilliseconds}ms: 0 archivos válidos para procesar.");
                return;
            }

            // FASE 2: Aplicar a SQLite
            bool changesPersisted = false;
            using (var context = SqliteHelper.GetContextForPath(_vaultPath))
            {
                IDbConnection connection = context.Open();
                _controlRepo.EnsureTablesExist(connection);

                var sqliteConn = connection as SQLiteConnection;
                if (sqliteConn == null || !CrSqliteLoader.EnableOnConnection(sqliteConn))
                {
                    LoggerHelper.Error("[SubunidadSyncService] PULL BATCH: CR-SQLite no disponible. No se pueden aplicar cambios.");
                    return;
                }

                foreach (var item in validEnvelopes)
                {
                    ct.ThrowIfCancellationRequested();
                    string fileName = item.FileName;
                    CrSyncEnvelope envelope = item.Envelope;

                    if (_controlRepo.IsFileAlreadyProcessed(connection, fileName))
                    {
                        LoggerHelper.Warning($"[SubunidadSyncService] PULL BATCH: {fileName} ya fue procesado.");
                        await MoveToProcessedAsync(item.FilePath);
                        continue;
                    }

                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            int aplicados = ApplyCrSqlChanges(connection, transaction, envelope.Changes, envelope.Version);
                            _controlRepo.MarkFileProcessed(connection, fileName, envelope.Source?.Subunidad ?? "UNKNOWN", envelope.Checksum, transaction);
                            _controlRepo.UpdateLastPulledAt(connection, _subunidad, _deviceId, transaction);
                            
                            if (aplicados > 0)
                            {
                                transaction.Commit();
                                changesPersisted = true;
                                LoggerHelper.Info($"[SubunidadSyncService] PULL BATCH: {aplicados} cambios aplicados de {fileName}.");
                            }
                            else
                            {
                                transaction.Commit(); // solo registros de sync_archive/control
                                LoggerHelper.Info($"[SubunidadSyncService] PULL BATCH: {fileName} procesado sin cambios aplicables.");
                            }
                            await MoveToProcessedAsync(item.FilePath);
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            LoggerHelper.Error($"[SubunidadSyncService] PULL BATCH: Error aplicando cambios de {fileName}. Reabriendo conexión para el siguiente archivo.", ex);
                            
                            try
                            {
                                // Abortamos la conexión actual para purgar cualquier estado corrupto o savepoints bloqueados
                                context.CloseConnectionResiliently();
                                connection = context.Open();
                                if (connection == null)
                                {
                                    LoggerHelper.Error("[SubunidadSyncService] PULL BATCH: No se pudo reabrir la conexión. Abortando batch.");
                                    changesPersisted = false;
                                    return; // Salir del método, no hay forma de continuar
                                }
                                _controlRepo.EnsureTablesExist(connection);
                                sqliteConn = connection as SQLiteConnection;
                                if (sqliteConn != null)
                                {
                                    CrSqliteLoader.EnableOnConnection(sqliteConn);
                                }
                            }
                            catch (Exception reopenEx)
                            {
                                LoggerHelper.Error("[SubunidadSyncService] PULL BATCH: Fallo al reabrir conexión. Abortando batch.", reopenEx);
                                changesPersisted = false;
                                return;
                            }
                        }
                    }
                }

                // FASE 3: Única persistencia al vault
                if (changesPersisted)
                {
                    ct.ThrowIfCancellationRequested();
                    context.CloseConnectionResiliently();
                    byte[] plainBytes = EncryptedSQLiteContext.LeerBytesConRetrySync(context.TempDbPath);
                    if (plainBytes != null)
                    {
                        SqliteHelper.UpdateVaultDirectly(plainBytes, _vaultPath);
                        context.MarkChangesPersisted();

                        // FASE 9.5: Actualizar estado de sincronización
                        UpdateSyncStateOnPull(validEnvelopes);

                        LoggerHelper.Info($"[SubunidadSyncService] PULL BATCH: Todos los cambios persistidos al vault exitosamente.");
                    }
                }
            }

            sw.Stop();
            LoggerHelper.Info($"[SubunidadSyncService] PULL BATCH completado en {sw.ElapsedMilliseconds}ms.");
        }

        /// <summary>
        /// PULL: Procesa un archivo .crsync entrante y aplica sus cambios a la DB local.
        ///
        /// FLUJO:
        ///   1. Verificar que el archivo no sea propio (mismo deviceId) y no esté procesado.
        ///   2. Leer y deserializar el archivo .crsync.
        ///   3. Verificar checksum de integridad.
        ///   4. Abrir EncryptedSQLiteContext con permiso de escritura.
        ///   5. Habilitar CR-SQLite.
        ///   6. Aplicar cambios insertando directamente en crsql_changes.
        ///   7. Persistir al vault (re-cifrar).
        ///   8. Mover archivo a .processed/.
        ///   9. Registrar en sync_archive.
        ///
        /// NOTA: Este método se ejecuta en background (Task.Run desde SyncFileWatcher).
        /// NUNCA en el hilo principal de Excel.
        /// </summary>
        /// <param name="filePath">Ruta completa al archivo .crsync entrante.</param>
        /// <param name="ct">Token de cancelación.</param>
        internal async Task ProcessIncomingFileAsync(string filePath, CancellationToken ct = default)
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(SubunidadSyncService));
            }

            string fileName = Path.GetFileName(filePath);
            LoggerHelper.Info($"[SubunidadSyncService] PULL: Procesando {fileName}...");

            var sw = System.Diagnostics.Stopwatch.StartNew();

            // Paso 1: Verificar que no es nuestro propio archivo
            // (podría aparecer en el FileSystemWatcher si el folder se monitorea)
            if (IsOwnFile(fileName))
            {
                LoggerHelper.Debug($"[SubunidadSyncService] PULL: Ignorando archivo propio: {fileName}");
                return;
            }

            // Paso 2: Leer y deserializar el envelope
            CrSyncEnvelope envelope = null;
            try
            {
                Exception lastEx = null;
                const int maxAttempts = 3;
                for (int attempt = 1; attempt <= maxAttempts; attempt++)
                {
                    try
                    {
                        if (!TryEnsureFileHydrated(filePath))
                        {
                            LoggerHelper.Warning($"[SubunidadSyncService] PULL: intento {attempt}/{maxAttempts} - archivo no hidratado aún. Reintentando...");
                            await Task.Delay(TimeSpan.FromMilliseconds(500 * attempt), ct);
                            continue;
                        }
                        envelope = _fileWriter.ReadFromFile(filePath);
                        break;
                    }
                    catch (IOException ex) when (IsOneDriveTransientError(ex))
                    {
                        lastEx = ex;
                        LoggerHelper.Warning($"[SubunidadSyncService] PULL: intento {attempt}/{maxAttempts} falló por recurso OneDrive. Reintentando...");
                        await Task.Delay(TimeSpan.FromMilliseconds(500 * attempt), ct);
                    }
                }

                if (envelope == null)
                {
                    LoggerHelper.Error($"[SubunidadSyncService] PULL: No se pudo leer {fileName} tras {maxAttempts} intentos por error transitorio. Se deja en inbox para reintento posterior.", lastEx);
                    return; // NO mover a .processed
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[SubunidadSyncService] PULL: Error leyendo/verificando {fileName}. Se moverá a .processed para evitar loops.", ex);
                await MoveToProcessedAsync(filePath);
                return;
            }

            ct.ThrowIfCancellationRequested();

            // Validación de subunidad
            string envelopeSubunidad = envelope.Source?.Subunidad ?? string.Empty;
            bool isGlobal = string.Equals(envelopeSubunidad, "_GT", StringComparison.OrdinalIgnoreCase);

            if (!isGlobal && !string.Equals(envelopeSubunidad, _subunidad, StringComparison.OrdinalIgnoreCase))
            {
                LoggerHelper.Warning($"[SubunidadSyncService] PULL: Archivo {fileName} pertenece a otra subunidad ({envelopeSubunidad}) en lugar de '{_subunidad}'. Se omite.");
                await MoveToProcessedAsync(filePath);
                return;
            }

            // Validación de versión del formato .crsync
            if (!string.Equals(envelope.Version, "1.0", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(envelope.Version, "1.1", StringComparison.OrdinalIgnoreCase))
            {
                LoggerHelper.Warning($"[SubunidadSyncService] PULL: Versión de archivo no soportada: '{envelope.Version}'. Ignorando {fileName}.");
                await MoveToProcessedAsync(filePath);
                return;
            }

            if (envelope.Changes == null || envelope.Changes.Count == 0)
            {
                LoggerHelper.Warning($"[SubunidadSyncService] PULL: Archivo {fileName} sin cambios. Ignorando.");
                await MoveToProcessedAsync(filePath);
                return;
            }

            LoggerHelper.Info(
                $"[SubunidadSyncService] PULL: Procesando {fileName} desde subunidad " +
                $"'{envelopeSubunidad}' para subunidad local '{_subunidad}'. " +
                $"Cambios: {envelope.Changes.Count}");

            // Paso 3: Abrir vault con permiso de escritura (los cambios modifican datos reales)
            using (var context = SqliteHelper.GetContextForPath(_vaultPath))
            {
                IDbConnection connection = context.Open();

                // Paso 4: Verificar que no está ya procesado (verificación dentro del vault)
                _controlRepo.EnsureTablesExist(connection);
                if (_controlRepo.IsFileAlreadyProcessed(connection, fileName))
                {
                    LoggerHelper.Warning($"[SubunidadSyncService] PULL: {fileName} ya fue procesado. Ignorando.");
                    await MoveToProcessedAsync(filePath);
                    return;
                }

                // Paso 5: Habilitar CR-SQLite
                var sqliteConn = connection as SQLiteConnection;
                if (sqliteConn == null || !CrSqliteLoader.EnableOnConnection(sqliteConn))
                {
                    LoggerHelper.Error("[SubunidadSyncService] PULL: CR-SQLite no disponible o conexión inválida. No se pueden aplicar cambios.");
                    return;
                }

                ct.ThrowIfCancellationRequested();

                // Paso 6: Aplicar los cambios en una transacción
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        int aplicados = ApplyCrSqlChanges(connection, transaction, envelope.Changes, envelope.Version);

                        // Registrar el archivo como procesado dentro de la misma transacción
                        _controlRepo.MarkFileProcessed(
                            connection, fileName, envelope.Source?.Subunidad ?? "UNKNOWN",
                            envelope.Checksum, transaction);

                        // Actualizar timestamp de último pull exitoso
                        _controlRepo.UpdateLastPulledAt(connection, _subunidad, _deviceId, transaction);

                        transaction.Commit();
                        LoggerHelper.Info($"[SubunidadSyncService] PULL: {aplicados} cambios aplicados de {fileName}.");
                    }
                    catch (System.Data.SQLite.SQLiteException sqlEx)
                    {
                        transaction.Rollback();
                        LoggerHelper.Error($"[SubunidadSyncService] PULL: Excepción SQLite al procesar {fileName}. ErrorCode={sqlEx.ErrorCode}, ResultCode={sqlEx.ResultCode}", sqlEx);
                        throw;
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        LoggerHelper.Error($"[SubunidadSyncService] PULL: Error aplicando cambios de {fileName}. Rollback ejecutado.", ex);
                        throw;
                    }
                }

                ct.ThrowIfCancellationRequested();

                // Paso 7: Persistir los cambios al vault cifrado
                context.CloseConnectionResiliently();
                byte[] plainBytes = EncryptedSQLiteContext.LeerBytesConRetrySync(context.TempDbPath);
                if (plainBytes != null)
                {
                    SqliteHelper.UpdateVaultDirectly(plainBytes, _vaultPath);
                    context.MarkChangesPersisted();
                    LoggerHelper.Info($"[SubunidadSyncService] PULL: Cambios de {fileName} persistidos al vault.");
                }
            }

            // Paso 8: Mover el archivo a .processed/ (fuera del contexto vault)
            await MoveToProcessedAsync(filePath);

            // FASE 9.5: Actualizar estado de sincronización (PULL) individual
            UpdateSyncStateOnPull(new[] { (filePath, fileName, envelope) }.ToList());

            sw.Stop();
            LoggerHelper.Info($"[SubunidadSyncService] PULL completado en {sw.ElapsedMilliseconds}ms: {fileName}");
        }

        // ====================================================================
        // MÉTODOS PRIVADOS DE LECTURA (PUSH)
        // ====================================================================

        /// <summary>
        /// Lee cambios de crsql_changes para una tabla que tiene columna `subunidad`,
        /// filtrando solo los registros que pertenecen a la subunidad actual.
        ///
        /// QUERY EXPLICADO:
        ///   crsql_changes es una tabla virtual de CR-SQLite que expone todos los cambios
        ///   rastreados. El JOIN con la tabla real filtra por subunidad.
        ///   La columna seq es secuencial y monotónica: filtramos con seq > lastSeq.
        ///
        /// NOTA: El CAST de pk a TEXT es necesario porque CR-SQLite almacena PKs como BLOB.
        ///
        ///
        /// FASE 12 (Revisión):
        ///   Se atrapa el 'SQL logic error' y se reintenta desde el clock 0. Si esto también falla 
        ///   (daño estructural irreversible), NO se omite la tabla, ya que causaría una pérdida silenciosa
        ///   de sincronización (Falso Positivo). En su lugar, se lanza una excepción crítica abortando el
        ///   PUSH para advertir que el archivo local debe ser reparado externamente.
        /// </summary>
        private static readonly Dictionary<string, string> TablePrimaryKeys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "config", "id" },
            { "usuarios", "id" },
            { "festivos", "idfestivo" },
            { "programas", "id" },
            { "asignaturas", "id" },
            { "microcurriculo", "id" },
            { "contenidos", "id" },
            { "terceros", "id" },
            { "sedes", "id" },
            { "periodos", "idperiodo" },
            { "vinculaciones", "id" },
            { "sesiones", "id" },
            { "salvedades_conflictos", "id" },
            { "BaseGastosTerrestre", "Id" },
            { "territoriales", "id" },
            { "terminales", "id" },
            { "divipola", "id" },
            { "cdpDependencias", "id" }
        };

        private List<CrSyncChange> ReadChangesWithSubunidadFilter(
            IDbConnection connection, string tableName, string subunidad,
            SyncClock lastClock, ref SyncClock maxClockFound)
        {
            if (!TablePrimaryKeys.TryGetValue(tableName, out string pkCol))
            {
                LoggerHelper.Warning($"[SubunidadSyncService] PK no encontrada para {tableName}, enviando sin filtro.");
                return ReadChangesGlobal(connection, tableName, lastClock, ref maxClockFound);
            }

            string sql = $@"
                SELECT c.""table"" AS Table_,
                       c.pk AS Pk, c.cid AS Cid, c.val AS Val,
                       c.col_version AS ColVersion, c.db_version AS DbVersion,
                       c.site_id AS SiteId, c.cl AS Cl, c.seq AS Seq_
                FROM crsql_changes c
                LEFT JOIN [{tableName}] t 
                    ON CAST(t.[{pkCol}] AS TEXT) = CAST(c.pk AS TEXT)
                WHERE c.""table"" = @TableName
                  AND (
                      c.cid = -1                                          -- FASE 17: Pasar DELETEs siempre
                      OR t.subunidad = @Subunidad 
                      OR t.subunidad = '' 
                      OR t.subunidad IS NULL
                  )
                  AND (c.db_version > @LastDbVersion OR (c.db_version = @LastDbVersion AND c.seq > @LastSeq))
                ORDER BY c.db_version, c.seq;";

            try
            {
                return MapToChangeRecords(connection, sql, tableName, subunidad, lastClock, ref maxClockFound);
            }
            catch (Exception ex) when (ex.Message.Contains("SQL logic error"))
            {
                LoggerHelper.Warning($"[SubunidadSyncService] Error PUSH (SQL logic error) en {tableName}. Reseteando clock local a 0 y reintentando...");
                try
                {
                    connection.Execute(
                        "UPDATE sync_control SET last_pushed_clock = '0.0' WHERE table_name = @TableName",
                        new { TableName = tableName });
                }
                catch { /* sync_control podría no existir aún o estar bloqueada */ }
                
                try
                {
                    // Reintentar una sola vez con clock=0
                    return MapToChangeRecords(connection, sql, tableName, subunidad, new SyncClock(0, 0), ref maxClockFound);
                }
                catch (Exception retryEx)
                {
                    LoggerHelper.Error($"[SubunidadSyncService] Reintento fatal en tabla {tableName}. DAÑO ESTRUCTURAL CR-SQLITE. Se requiere reparación externa. Error: {retryEx.Message}");
                    throw new InvalidOperationException($"La sincronización CRDT de la tabla '{tableName}' está corrupta irreparablemente. Ejecute script de reparación externa.", retryEx);
                }
            }
        }

        /// <summary>
        /// Lee cambios de crsql_changes para tablas globales (sin columna subunidad).
        /// Incluye misma validación estructural (Fase 12) que lanza excepción crítica si 
        /// los relojes CRDT están permanentemente corruptos.
        /// </summary>
        private List<CrSyncChange> ReadChangesGlobal(
            IDbConnection connection,
            string tableName,
            SyncClock lastClock,
            ref SyncClock maxClockFound)
        {
            string sql = $@"
                SELECT
                    c.""table"" AS Table_,
                    c.pk        AS Pk,
                    c.cid       AS Cid,
                    c.val       AS Val,
                    c.col_version AS ColVersion,
                    c.db_version  AS DbVersion,
                    c.site_id     AS SiteId,
                    c.cl          AS Cl,
                    c.seq         AS Seq_
                FROM crsql_changes c
                WHERE c.""table"" = @TableName
                  AND (c.db_version > @LastDbVersion OR (c.db_version = @LastDbVersion AND c.seq > @LastSeq))
                ORDER BY c.db_version, c.seq;";

            try
            {
                // Para globales, no hay filtro de subunidad
                return MapToChangeRecords(connection, sql, tableName, subunidad: null, lastClock, ref maxClockFound);
            }
            catch (Exception ex) when (ex.Message.Contains("SQL logic error"))
            {
                LoggerHelper.Warning($"[SubunidadSyncService] Error PUSH (SQL logic error) en {tableName}. Reseteando clock local a 0 y reintentando...");
                try
                {
                    connection.Execute(
                        "UPDATE sync_control SET last_pushed_clock = '0.0' WHERE table_name = @TableName",
                        new { TableName = tableName });
                }
                catch { /* sync_control podría no existir aún o estar bloqueada */ }
                
                try
                {
                    // Reintentar una sola vez con clock=0
                    return MapToChangeRecords(connection, sql, tableName, subunidad: null, new SyncClock(0, 0), ref maxClockFound);
                }
                catch (Exception retryEx)
                {
                    LoggerHelper.Error($"[SubunidadSyncService] Reintento fatal en tabla {tableName}. DAÑO ESTRUCTURAL CR-SQLITE. Se requiere reparación externa. Error: {retryEx.Message}");
                    throw new InvalidOperationException($"La sincronización CRDT de la tabla '{tableName}' está corrupta irreparablemente. Ejecute script de reparación externa.", retryEx);
                }
            }
        }

        /// <summary>
        /// Ejecuta el SQL de lectura de cambios y mapea los resultados a CrSyncChange.
        /// Actualiza maxClockFound con el mayor reloj encontrado.
        /// </summary>
        private List<CrSyncChange> MapToChangeRecords(
            IDbConnection connection,
            string sql,
            string tableName,
            string subunidad,
            SyncClock lastClock,
            ref SyncClock maxClockFound)
        {
            var result = new List<CrSyncChange>();

            // Dapper retorna dynamic para poder acceder a columnas calculadas (Seq_)
            object parameters = new 
            { 
                TableName = tableName, 
                Subunidad = subunidad,
                LastDbVersion = lastClock.DbVersion, 
                LastSeq = lastClock.Seq 
            };

            var rows = connection.Query(sql, parameters);

            foreach (dynamic row in rows)
            {
                long rowDbVersion = (long)row.DbVersion;
                long rowSeq = (long)row.Seq_;
                var rowClock = new SyncClock(rowDbVersion, rowSeq);

                if (rowClock > maxClockFound)
                {
                    maxClockFound = rowClock;
                }

                string pkForLog = row.Pk is byte[] pkBytesForLog
                    ? Convert.ToBase64String(pkBytesForLog)
                    : Convert.ToString(row.Pk);
                    
                if (Convert.ToString(row.Cid) == "-1")
                {
                    LoggerHelper.Debug($"[SubunidadSyncService] {tableName}: DELETE detectado para pk={pkForLog} (propagado sin verificación de subunidad).");
                }
                
                result.Add(new CrSyncChange
                {
                    Table = Convert.ToString(row.Table_),
                    Pk = pkForLog,
                    Cid = Convert.ToString(row.Cid),
                    Val = row.Val is byte[] valBytes 
                        ? GetSafeUtf8StringAndLog(valBytes, pkForLog ?? "<null>", tableName)
                        : Convert.ToString(row.Val),
                    ColVersion = ConvertDynamicToInt64(row.ColVersion),
                    DbVersion = ConvertDynamicToInt64(row.DbVersion),
                    SiteId = ConvertDynamicToInt64(row.SiteId),
                    Cl = ConvertDynamicToInt64(row.Cl)
                });
            }

            if (result.Count > 0)
            {
                LoggerHelper.Debug($"[SubunidadSyncService] {tableName}: {result.Count} cambios encontrados.");
            }
            // Eliminado bloque de diagnóstico de subunidad que requería lastSeq,
            // dado que el filtro se realiza delegando la lectura general.

            return result;
        }

        private string GetSafeUtf8StringAndLog(byte[] valBytes, string pk, string tableName)
        {
            string decoded = System.Text.Encoding.UTF8.GetString(valBytes);
            if (decoded.Contains("\uFFFD"))
            {
                LoggerHelper.Warning($"[SubunidadSyncService] {tableName}: Posible corrupción de val BLOB no textual detectada para pk={pk}. \uFFFD (Replacement Character) hallado.");
            }
            return decoded;
        }

        private static long ConvertDynamicToInt64(dynamic value)
        {
            if (value == null) return 0;
            if (value is byte[] bytes)
            {
                // CR-SQLite almacena enteros como BLOBs little-endian.
                // Soporta 1, 2, 4, 8 o 16 bytes; leemos solo los primeros 8 bytes efectivos.
                long result = 0;
                int length = Math.Min(bytes.Length, 8);
                for (int i = 0; i < length; i++)
                    result |= (long)bytes[i] << (i * 8);
                return result;
            }
            return Convert.ToInt64(value);
        }

        // ====================================================================
        // MÉTODOS PRIVADOS DE ESCRITURA (PULL)
        // ====================================================================

        /// <summary>
        /// Aplica una lista de cambios CRDT a la DB local realizando un INSERT masivo
        /// en la tabla virtual crsql_changes de CR-SQLite.
        ///
        /// CR-SQLite intercepta estas inserciones y aplica internamente el algoritmo LWW 
        /// (Last-Writer-Wins) comparando col_version, db_version y site_id para desempatar 
        /// o resolver conflictos.
        ///
        /// IMPORTANTE: Se llama DENTRO de una transacción explícita para garantizar
        /// atomicidad. Si falla, el Rollback en el llamador deshace todo.
        /// </summary>
        private int ApplyCrSqlChanges(
            IDbConnection connection,
            IDbTransaction transaction,
            List<CrSyncChange> changes,
            string envelopeVersion)
        {
            if (changes == null || changes.Count == 0) return 0;

            if (string.IsNullOrWhiteSpace(envelopeVersion))
            {
                throw new InvalidOperationException("[SubunidadSyncService] PULL: Archivo sin versión de formato. No se puede decodificar PK de forma segura.");
            }

            // ================================================================
            // FASE 5: DEFENSA DE PULL (Filtrar tablas no CRDT o inexistentes)
            // ================================================================
            var validChanges = new List<CrSyncChange>();
            var skippedTables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var change in changes)
            {
                if (skippedTables.Contains(change.Table))
                {
                    continue;
                }

                if (!TableIsCrdtEnabled(connection, transaction, change.Table))
                {
                    LoggerHelper.Warning($"[SubunidadSyncService] PULL: Tabla '{change.Table}' no existe o no está habilitada como CRDT en este nodo. Se omiten sus cambios.");
                    skippedTables.Add(change.Table);
                    continue;
                }
                validChanges.Add(change);
            }

            changes = validChanges;
            if (changes.Count == 0)
            {
                LoggerHelper.Debug("[SubunidadSyncService] PULL: No hay cambios válidos tras filtrar tablas no CRDT.");
                return 0;
            }
            // ================================================================

            bool isV1_1 = string.Equals(envelopeVersion, "1.1", StringComparison.OrdinalIgnoreCase);

            // 1. Diagnóstico del esquema de crsql_changes
            bool siteIdIsBlob = false;
            try
            {
                var cols = connection.Query<string>(
                    "SELECT name FROM pragma_table_info('crsql_changes') ORDER BY cid;",
                    transaction: transaction).ToList();
                LoggerHelper.Debug($"[SubunidadSyncService] crsql_changes esquema: {string.Join(",", cols)}");

                var siteIdCol = connection.QueryFirstOrDefault<string>(
                    "SELECT type FROM pragma_table_info('crsql_changes') WHERE name='site_id';",
                    transaction: transaction);
                siteIdIsBlob = string.Equals(siteIdCol, "BLOB", StringComparison.OrdinalIgnoreCase);
                LoggerHelper.Debug($"[SubunidadSyncService] site_id tipo detectado: {siteIdCol ?? "UNKNOWN"} (BLOB={siteIdIsBlob})");
            }
            catch (Exception schemaEx)
            {
                LoggerHelper.Warning($"[SubunidadSyncService] No se pudo leer esquema de crsql_changes: {schemaEx.Message}");
            }

            // 2. Muestra de cambios a aplicar
            int muestra = Math.Min(5, changes.Count);
            for (int i = 0; i < muestra; i++)
            {
                var c = changes[i];
                LoggerHelper.Debug(
                    $"[SubunidadSyncService] PULL cambio #{i}: tabla={c.Table}, " +
                    $"pk={c.Pk ?? "<null>"}, cid={c.Cid}, val={(c.Val == null ? "<null>" : "[len=" + c.Val.Length + "]")}, " +
                    $"col_version={c.ColVersion}, db_version={c.DbVersion}, site_id={c.SiteId}, cl={c.Cl}");
            }

            // 3. INSERT directo con objetos dinámicos para evitar conflictos de nombres
            string sql = @"
                INSERT INTO crsql_changes (""table"", pk, cid, val, col_version, db_version, site_id, cl)
                VALUES (@Table, @Pk, @Cid, @Val, @ColVersion, @DbVersion, @SiteId, @Cl);";

            var localColumnsByTable = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            int aplicados = 0;
            foreach (var change in changes)
            {
                if (string.IsNullOrEmpty(change.Pk))
                {
                    LoggerHelper.Warning($"[SubunidadSyncService] PULL: Cambio con PK nula/vacía omitido para tabla {change.Table}, cid {change.Cid}.");
                    continue;
                }

                if (!IsKnownCid(connection, transaction, localColumnsByTable, change.Table, change.Cid))
                {
                    LoggerHelper.Warning(
                        $"[SubunidadSyncService] PULL: cid '{change.Cid}' no existe en tabla local '{change.Table}'. Cambio omitido.");
                    continue;
                }

                byte[] pkBytes = null;
                if (isV1_1)
                {
                    try 
                    {
                        pkBytes = Convert.FromBase64String(change.Pk);
                    }
                    catch (FormatException)
                    {
                        LoggerHelper.Warning($"[SubunidadSyncService] PULL: PK en formato 1.1 no pudo decodificarse como Base64. Fallback a UTF-8 para tabla {change.Table}, cid {change.Cid}.");
                        pkBytes = System.Text.Encoding.UTF8.GetBytes(change.Pk);
                    }
                }
                else
                {
                    // Formato 1.0: texto plano
                    pkBytes = System.Text.Encoding.UTF8.GetBytes(change.Pk);
                }

                var p = new DynamicParameters();
                p.Add("Table", change.Table, DbType.String);
                p.Add("Pk", pkBytes, DbType.Binary); // CR-SQLite requiere BLOB exacto
                p.Add("Cid", change.Cid, DbType.String);
                p.Add("Val", change.Val, DbType.String);
                p.Add("ColVersion", change.ColVersion, DbType.Int64);
                p.Add("DbVersion", change.DbVersion, DbType.Int64);
                p.Add("Cl", change.Cl, DbType.Int64);

                if (siteIdIsBlob)
                {
                    // Enviamos el long como BLOB de 8 bytes little-endian.
                    // Si CR-SQLite espera 16 bytes (UUID), el log de ErrorCode nos lo indicará
                    // y podremos ajustar la expansión a 16 bytes con seguridad.
                    byte[] siteIdBytes = BitConverter.GetBytes(change.SiteId);
                    p.Add("SiteId", siteIdBytes, DbType.Binary);
                }
                else
                {
                    p.Add("SiteId", change.SiteId, DbType.Int64);
                }

                // La excepción se propaga directamente al invocador que ya loguea ErrorCode/ResultCode
                int rows = connection.Execute(sql, p, transaction: transaction);
                aplicados += rows;
            }

            if (aplicados > 0)
            {
                var stats = changes.GroupBy(c => c.Table).Select(g => $"{g.Key}={g.Count()}");
                LoggerHelper.Debug($"[SubunidadSyncService] PULL INSERT completado: {aplicados}/{changes.Count} filas afectadas. Detalle: {string.Join(", ", stats)}");
            }
            else
            {
                LoggerHelper.Debug($"[SubunidadSyncService] PULL INSERT completado: 0 filas afectadas.");
            }
            return aplicados;
        }

        // ====================================================================
        // UTILIDADES
        // ====================================================================

        /// <summary>
        /// Verifica si un cid (nombre de columna) existe en el esquema de la tabla local.
        /// Retorna true para marcadores numéricos o '-' de borrado de CR-SQLite.
        /// </summary>
        private static bool IsKnownCid(
            IDbConnection connection, IDbTransaction transaction,
            Dictionary<string, HashSet<string>> cache, string table, string cid)
        {
            if (string.IsNullOrEmpty(cid)) return false;

            // Marcadores internos CR-SQLite para DELETE
            if (cid == "-" || cid.TrimStart('-').Length == 0) return true;
            if (long.TryParse(cid, out _)) return true;

            if (!cache.TryGetValue(table, out var cols))
            {
                string escaped = table.Replace("'", "''");
                cols = new HashSet<string>(
                    connection.Query<string>($"SELECT name FROM pragma_table_info('{escaped}');",
                        transaction: transaction),
                    StringComparer.OrdinalIgnoreCase);
                cache[table] = cols;
            }
            return cols.Contains(cid);
        }

        /// <summary>
        /// Determina si un archivo .crsync fue generado por este mismo dispositivo.
        /// Los archivos propios nunca deben procesarse (loop de sync).
        ///
        /// FORMATO DEL NOMBRE: from_{DeviceId}_{Subunidad}_{timestamp}.crsync
        /// </summary>
        private bool IsOwnFile(string fileName)
        {
            string ownPrefix = $"from_{_deviceId}_";
            return fileName.StartsWith(ownPrefix, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Verifica si una tabla existe en la DB actual.
        /// Necesario porque algunas tablas sincronizables pueden no existir aún
        /// en DBs antiguas que no completaron la migración V24.
        /// </summary>
        private static bool TableExists(IDbConnection connection, string tableName)
        {
            int count = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@TableName;",
                new { TableName = tableName });
            return count > 0;
        }

        /// <summary>
        /// Mueve un archivo .crsync a la carpeta .processed/ de la subunidad.
        /// Si el archivo de destino ya existe, lo sobreescribe (puede pasar si hay reintento).
        /// Opera en background (Task.Run ya fue hecho por el llamador).
        /// </summary>
        private async Task MoveToProcessedAsync(string filePath)
        {
            await Task.Run(() =>
            {
                try
                {
                    string parentFolder = Path.GetDirectoryName(filePath);
                    string globalFolder = _pathResolver.GetGlobalInboxFolder();
                    
                    string processedFolder;
                    if (string.Equals(parentFolder, globalFolder, StringComparison.OrdinalIgnoreCase))
                    {
                        processedFolder = _pathResolver.GetGlobalProcessedFolder();
                    }
                    else
                    {
                        processedFolder = _pathResolver.GetProcessedFolder(_subunidad);
                    }

                    if (!Directory.Exists(processedFolder))
                    {
                        Directory.CreateDirectory(processedFolder);
                    }

                    string destPath = Path.Combine(processedFolder, Path.GetFileName(filePath));

                    // Si ya existe en .processed (reintento), borrar primero
                    if (File.Exists(destPath))
                    {
                        File.Delete(destPath);
                    }

                    File.Move(filePath, destPath);
                    LoggerHelper.Debug($"[SubunidadSyncService] Archivo movido a .processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"[SubunidadSyncService] No se pudo mover {Path.GetFileName(filePath)} a .processed: {ex.Message}");
                    // No relanzamos: el archivo quedó en la carpeta original.
                    // sync_archive lo marca como procesado, así que no se re-procesará.
                }
            });
        }

        /// <summary>
        /// Sanea el nombre del PC para usarlo en nombres de archivo.
        /// Reemplaza caracteres problemáticos con guiones bajos.
        /// </summary>
        private static string SanitizeDeviceId(string machineName)
        {
            if (string.IsNullOrEmpty(machineName)) { return "UNKNOWN"; }

            char[] chars = machineName.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] == ' ' || chars[i] == '\\' || chars[i] == '/' ||
                    chars[i] == ':' || chars[i] == '*' || chars[i] == '?' ||
                    chars[i] == '"' || chars[i] == '<' || chars[i] == '>' || chars[i] == '|')
                {
                    chars[i] = '_';
                }
            }

            return new string(chars).ToUpperInvariant();
        }

        // ====================================================================
        // IDISPOSABLE
        // ====================================================================

        public void Dispose()
        {
            if (_disposed) { return; }
            _disposed = true;
            LoggerHelper.Debug("[SubunidadSyncService] Disposed.");
        }

        // ====================================================================
        // FASE 9.5: ESTADO DE SINCRONIZACIÓN DESCENTRALIZADO
        // ====================================================================

        private string GetStateFilePath(string deviceId)
        {
            string inboxFolder = _pathResolver.GetSubunidadInboxFolder(_subunidad);
            return Path.Combine(inboxFolder, $"state_{deviceId}.json");
        }

        private SyncState ReadSyncState(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return null;
                string json = File.ReadAllText(filePath);
                var serializer = new JavaScriptSerializer();
                return serializer.Deserialize<SyncState>(json);
            }
            catch (Exception ex)
            {
                LoggerHelper.Debug($"[SubunidadSyncService] Error leyendo estado de sync en {filePath}: {ex.Message}");
                return null;
            }
        }

        private bool WriteSyncState(SyncState state)
        {
            string statePath = GetStateFilePath(_deviceId);
            try
            {
                var serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(state);
                
                // Retry-backoff básico para evitar bloqueos transitorios de OneDrive
                int retries = 3;
                while (retries > 0)
                {
                    try
                    {
                        File.WriteAllText(statePath, json);
                        return true;
                    }
                    catch (IOException)
                    {
                        retries--;
                        if (retries > 0) System.Threading.Thread.Sleep(500);
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SubunidadSyncService] No se pudo escribir estado de sync local: {ex.Message}");
                return false;
            }
        }

        private void UpdateSyncStateOnPush(string generatedFileName)
        {
            var state = ReadSyncState(GetStateFilePath(_deviceId)) ?? new SyncState { DeviceId = _deviceId };
            state.LastGeneratedFile = generatedFileName;
            state.LastUpdatedAt = DateTime.UtcNow;
            state.AppVersion = typeof(SubunidadSyncService).Assembly.GetName().Version.ToString();
            if (WriteSyncState(state))
            {
                LoggerHelper.Debug($"[SyncState] PUSH: LastGeneratedFile registrado -> {generatedFileName}");
            }
            else
            {
                LoggerHelper.Warning($"[SyncState] PUSH: No se pudo registrar LastGeneratedFile para {generatedFileName}");
            }
        }

        private void UpdateSyncStateOnPull(List<(string FilePath, string FileName, CrSyncEnvelope Envelope)> processed)
        {
            if (processed == null || processed.Count == 0) return;

            var state = ReadSyncState(GetStateFilePath(_deviceId)) ?? new SyncState { DeviceId = _deviceId };
            if (state.LastProcessedFiles == null) state.LastProcessedFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var sortedProcessed = processed.OrderBy(p => p.FileName, StringComparer.OrdinalIgnoreCase).ToList();

            var lastBySender = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in sortedProcessed)
            {
                if (!string.IsNullOrEmpty(item.Envelope?.Source?.DeviceId))
                {
                    string senderId = item.Envelope.Source.DeviceId;
                    state.LastProcessedFiles[senderId] = item.FileName;
                    lastBySender[senderId] = item.FileName;
                }
            }

            state.LastUpdatedAt = DateTime.UtcNow;
            state.AppVersion = typeof(SubunidadSyncService).Assembly.GetName().Version.ToString();
            
            if (WriteSyncState(state))
            {
                // Resumen eficiente: un solo log por batch
                var summaryParts = lastBySender.Select(kvp => $"{kvp.Key}={kvp.Value}");
                LoggerHelper.Debug($"[SyncState] PULL: {sortedProcessed.Count} sobres procesados. Últimos archivos por origen: {string.Join(", ", summaryParts)}");
            }
            else
            {
                LoggerHelper.Warning($"[SyncState] PULL: Falló actualización de estado tras procesar {sortedProcessed.Count} sobres.");
            }
        }

        /// <summary>
        /// Compara el estado local (último archivo generado) con los archivos procesados por otros nodos.
        /// Retorna advertencias si detecta nodos atrasados por más de 5 minutos.
        /// </summary>
        internal void CheckLaggingDevices()
        {
            try
            {
                var myState = ReadSyncState(GetStateFilePath(_deviceId));
                if (myState == null || string.IsNullOrEmpty(myState.LastGeneratedFile)) return;

                string inboxFolder = _pathResolver.GetSubunidadInboxFolder(_subunidad);
                if (!Directory.Exists(inboxFolder)) return;

                // Si mi último archivo generado lleva menos de 5 minutos, es normal que otros no lo hayan leído
                string myLastFilePath = Path.Combine(inboxFolder, myState.LastGeneratedFile);
                if (File.Exists(myLastFilePath))
                {
                    // Extraer fecha del nombre del archivo: from_DEVICEID_SUBUNIDAD_yyyyMMddHHmmss.crsync
                    string fileName = myState.LastGeneratedFile;
                    string tsPart = fileName.Replace(".crsync", "");
                    if (tsPart.Length >= 14)
                    {
                        tsPart = tsPart.Substring(tsPart.Length - 14);
                        if (DateTime.TryParseExact(tsPart, "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal, out DateTime fileTimeUtc))
                        {
                            var fileAge = DateTime.UtcNow - fileTimeUtc;
                            if (fileAge.TotalMinutes < 5) return;
                        }
                    }
                }

                var otherStateFiles = Directory.GetFiles(inboxFolder, "state_*.json")
                                               .Where(f => !Path.GetFileName(f).Equals($"state_{_deviceId}.json", StringComparison.OrdinalIgnoreCase))
                                               .ToList();

                int healthyCount = 0;
                int laggingCount = 0;

                foreach (var stateFile in otherStateFiles)
                {
                    var otherState = ReadSyncState(stateFile);
                    if (otherState != null && otherState.LastProcessedFiles != null)
                    {
                        otherState.LastProcessedFiles.TryGetValue(_deviceId, out string lastProcessedFromMe);
                        
                        bool isLagging = string.IsNullOrEmpty(lastProcessedFromMe) ||
                                         string.Compare(lastProcessedFromMe, myState.LastGeneratedFile, StringComparison.OrdinalIgnoreCase) < 0;

                        if (isLagging)
                        {
                            laggingCount++;
                            string procesado = lastProcessedFromMe ?? "N/A";
                            LoggerHelper.Warning($"[Sync] El dispositivo {otherState.DeviceId} podría estar atrasado. " +
                                $"Último archivo que procesó de nosotros: {procesado}. " +
                                $"Último archivo que empujamos: {myState.LastGeneratedFile}. (Su estado: {otherState.LastUpdatedAt:u})");
                        }
                        else
                        {
                            healthyCount++;
                        }
                    }
                }

                LoggerHelper.Debug($"[SyncState] Análisis de vecinos completado. {healthyCount} al día, {laggingCount} atrasados, {otherStateFiles.Count - healthyCount - laggingCount} sin estado válido.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Debug($"[SubunidadSyncService] Error verificando dispositivos atrasados: {ex.Message}");
            }
        }
    }

    // ====================================================================
    // RESULTADO DEL PUSH (Value Object)
    // ====================================================================

    /// <summary>
    /// Resultado de una operación PUSH.
    /// Permite al SyncScheduler loguear y notificar la UI sobre el resultado
    /// sin acoplar el servicio al scheduler.
    /// </summary>
    internal sealed class PushResult
    {
        internal bool WasSuccessful { get; private set; }
        internal bool HadChanges { get; private set; }
        internal int ChangesCount { get; private set; }
        internal long BytesSent { get; private set; }
        internal string FileName { get; private set; }
        internal string SkipReason { get; private set; }

        private PushResult() { }

        internal static PushResult Success(int changesCount, long bytesSent, string fileName)
        {
            return new PushResult
            {
                WasSuccessful = true,
                HadChanges = true,
                ChangesCount = changesCount,
                BytesSent = bytesSent,
                FileName = fileName
            };
        }

        internal static PushResult NoChanges()
        {
            return new PushResult
            {
                WasSuccessful = true,
                HadChanges = false
            };
        }

        internal static PushResult Skipped(string reason)
        {
            return new PushResult
            {
                WasSuccessful = false,
                HadChanges = false,
                SkipReason = reason
            };
        }
    }
}
